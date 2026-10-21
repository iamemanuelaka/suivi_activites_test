#!/usr/bin/env python3
"""
Synchronisation KoboToolbox -> MATRICE DE SUIVI DES ACTIVITES
(Mise à jour des nouvelles colonnes financières et physiques)
"""

import argparse
import datetime as dt
import json
import logging
import os
import re
import smtplib
import ssl
import sys
from copy import copy
from email.message import EmailMessage
from pathlib import Path
from typing import Dict, List, Optional

import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# --------------------------------------------------------------------------
# MAPPING DES DONNÉES (Kobo -> Excel)
# --------------------------------------------------------------------------

STATUT = {
    "non_demarre": "Non démarré",
    "demarrage_autorise": "Démarrage autorisé",
    "os_a_reprendre": "Ordre de service à reprendre (erreur sur l'OS)",
    "demarrage_effectif": "Démarrage effectif",
    "travaux_en_cours": "Travaux en cours",
    "travaux_acheves": "Travaux achevés",
    "suspendu": "Travaux suspendus",
    "abandonne": "Abandonné",
}

PGES = {
    "valide": "Validé",
    "transmis_validation": "Transmis en cours de validation",
    "non_transmis": "Non transmis",
    "pas_de_pges": "Pas de PGES",
}

ONG = {
    "ong_epc": "ONG EPC",
    "agir_odd": "AGIR pour ODD",
    "kera_sy_ire": "KERA-SY-IRÊ",
    "etoile_bounkani": "ONG Étoile du Bounkani",
    "cibes": "CIBES",
    "voix_femme": "ONG Voix de Femme",
    "pietro_bonili": "Centre de santé PIETRO BONILI",
    "solidarite": "Solidarité",
    "victoire_plus": "ONG Victoire Plus",
    "esperance_vie": "Espérance Vie",
    "mobilisee_autre": "ONG mobilisée",
    "non_mobilisee": "Non mobilisée",
}

# Mapping des en-têtes Excel (la fonction norm() les mettra en majuscules sans accents)
HEADERS_MATRIX = {
    "region": "REGION",
    "departement": "DEPARTEMENT",
    "entreprise": "ENTREPRISE",
    "remise": "DATE DE REMISE",
    "taux": "TAUX D",  # Match "TAUX D'EXECUTION"
    "obs_tech": "OBSERVATIONS TECHNIQUES",
    "pges": "OBSERVATIONS ENVIRONNEMENTALES",
    "ong": "ONG MOBILISEES",
    "recep_prov": "DATE DE RECEPTION PROVISOIRE",
    "recep_def": "DATE DE RECEPTION DEFINITIVE",
    "montant_decaisse": "MONTANT DECAISSE",
    "taux_financier": "TAUX DEXECUTION FINANCIERE", # Sans accent pour le match
    "volume_acheve": "VOLUME ACHEVE",
    "nom_agent_poste": "NOM PRENOMS + POSTE ET STRUCTURE",
}

SHEET_HISTORY = "HISTORIQUE"
MATRIX_PATTERN = "MATRICE_DE_SUIVI_DES_ACTIVITES_du_*.xlsx"

log = logging.getLogger("kobo_sync")

# --------------------------------------------------------------------------
# UTILITAIRES
# --------------------------------------------------------------------------

def load_env(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

def cfg(name: str, default: str = "", required: bool = False) -> str:
    val = os.environ.get(name, default)
    if required and not val:
        sys.exit(f"Variable manquante : {name}")
    return val

def norm(s) -> str:
    s = re.sub(r"\s+", " ", str(s or "")).strip().upper()
    return s.translate(str.maketrans("ÉÈÊÀÂÎÔÛÇ", "EEEAAIOUC"))

def pdate(value) -> Optional[dt.date]:
    if not value:
        return None
    try:
        return dt.date.fromisoformat(str(value)[:10])
    except (ValueError, TypeError):
        return None

# --------------------------------------------------------------------------
# 1. RÉCUPÉRATION KOBO
# --------------------------------------------------------------------------

def fetch_submissions() -> List[dict]:
    base = cfg("KOBO_URL", "https://kf.kobotoolbox.org").rstrip("/")
    uid = cfg("KOBO_ASSET_UID", required=True)
    headers = {"Authorization": f"Token {cfg('KOBO_TOKEN', required=True)}"}
    url = f"{base}/api/v2/assets/{uid}/data/?format=json&limit=1000"
    out = []
    while url:
        r = requests.get(url, headers=headers, timeout=60)
        r.raise_for_status()
        payload = r.json()
        out.extend(payload.get("results", []))
        url = payload.get("next")
    log.info("%d soumission(s) récupérée(s)", len(out))
    return out

def flatten(sub: dict) -> dict:
    flat = {}
    for k, v in sub.items():
        key = k.split("/")[-1]
        flat[key] = v
    flat["_id"] = sub.get("_id")
    flat["_submission_time"] = sub.get("_submission_time")
    flat["_validation_status"] = sub.get("_validation_status")
    return flat

def is_rejected(sub: dict) -> bool:
    vs = sub.get("_validation_status") or {}
    uid = vs.get("uid") if isinstance(vs, dict) else vs
    return uid == "validation_status_not_approved"

def lot_number(sub: dict) -> Optional[int]:
    try:
        # Gère "01", "1", etc.
        val = str(sub.get("id_projet", "")).strip()
        return int(val)
    except (ValueError, TypeError):
        return None

def latest_per_lot(subs: List[dict]) -> Dict[int, dict]:
    best: Dict[int, dict] = {}
    for s in subs:
        if is_rejected(s):
            continue
        n = lot_number(s)
        if n is None:
            continue
        key = (pdate(s.get("date_mission")) or dt.date.min, str(s.get("_submission_time") or ""))
        if n not in best or key > best[n]["_key"]:
            best[n] = {**s, "_key": key}
    return best

# --------------------------------------------------------------------------
# 2. MISE À JOUR MATRICE
# --------------------------------------------------------------------------

def find_columns(ws) -> tuple:
    for r in range(1, 11):
        found = {}
        for c in range(1, ws.max_column + 1):
            h = norm(ws.cell(r, c).value)
            for key, prefix in HEADERS_MATRIX.items():
                if key not in found and h.startswith(prefix):
                    found[key] = c
        if "taux" in found and "region" in found:
            return r, found
    sys.exit("En-têtes introuvables dans la matrice")

def update_matrix_sheet(ws, latest: Dict[int, dict]) -> dict:
    header_row, cols = find_columns(ws)
    rows = {}
    num_col = cols["region"] - 1 # La colonne A contient le numéro du projet
    
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, num_col).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            rows[int(v)] = r
    
    for n, s in latest.items():
        r = rows.get(n)
        if r is None:
            log.warning("Lot %s absent de la matrice", n)
            continue
        
        # 1. Statut + Observations
        statut = s.get("statut")
        obs = (s.get("obs_techniques") or "").strip()
        if statut:
            txt = STATUT.get(statut, statut)
            ws.cell(r, cols.get("obs_tech", 1)).value = f"{txt} – {obs}" if obs else txt
        
        # 2. Taux d'exécution physique
        taux = s.get("taux_execution")
        if taux not in (None, ""):
            ws.cell(r, cols.get("taux", 1)).value = int(float(taux))
        
        # 3. Date de remise de site
        d = pdate(s.get("date_remise_site"))
        if d and "remise" in cols:
            ws.cell(r, cols["remise"]).value = d
            ws.cell(r, cols["remise"]).number_format = "DD/MM/YYYY"
        
        # 4. PGES
        if "pges" in cols and s.get("pges_statut"):
            ws.cell(r, cols["pges"]).value = PGES.get(s["pges_statut"], s["pges_statut"])
        
        # 5. ONG
        if "ong" in cols:
            ong = s.get("Il_y_a_t_il_une_ONG_mobilis_e")
            if ong == "oui":
                nom_ong = s.get("ong_autre_nom") or "ONG mobilisée"
                ws.cell(r, cols["ong"]).value = nom_ong
            elif ong == "non":
                ws.cell(r, cols["ong"]).value = "Non mobilisée"
        
        # 6. Réceptions
        if "recep_prov" in cols and s.get("recep_prov") == "oui":
            d = pdate(s.get("date_recep_prov"))
            if d:
                ws.cell(r, cols["recep_prov"]).value = d
                ws.cell(r, cols["recep_prov"]).number_format = "DD/MM/YYYY"
        
        if "recep_def" in cols and s.get("recep_def") == "oui":
            d = pdate(s.get("date_recep_def"))
            if d:
                ws.cell(r, cols["recep_def"]).value = d
                ws.cell(r, cols["recep_def"]).number_format = "DD/MM/YYYY"
        
        # --- NOUVELLES COLONNES ---
        
        # 7. Montant décaissé
        if "montant_decaisse" in cols:
            val = s.get("Montant_d_caiss_sur_e_budget_total_FCFA")
            if val not in (None, ""):
                ws.cell(r, cols["montant_decaisse"]).value = int(float(val))
        
        # 8. Taux d'exécution financière
        if "taux_financier" in cols:
            val = s.get("Taux_d_ex_cution_financi_re_")
            if val not in (None, ""):
                ws.cell(r, cols["taux_financier"]).value = int(float(val))
        
        # 9. Volume achevé
        if "volume_acheve" in cols:
            val = s.get("Volume_de_travaux_achev_Km")
            if val not in (None, ""):
                ws.cell(r, cols["volume_acheve"]).value = float(val)
        
        # 10. Nom + Poste + Structure
        if "nom_agent_poste" in cols:
            nom = s.get("nom_agent", "")
            poste = s.get("agent", "")
            structure = s.get("Structure", "")
            autre = s.get("Autre_pr_ciser", "")
            struct_final = autre if structure == "autre" else structure
            
            texte_complet = f"{nom} - {poste} ({struct_final})".strip(" -()")
            if texte_complet:
                ws.cell(r, cols["nom_agent_poste"]).value = texte_complet

    return {"rows": rows}

def append_history(wb, subs: List[dict]) -> int:
    head = ["ID KOBO", "DATE VISITE", "N° LOT", "STATUT", "TAUX (%)", "RAPPORTEUR", "SOUMIS LE"]
    if SHEET_HISTORY not in wb.sheetnames:
        hs = wb.create_sheet(SHEET_HISTORY)
        hs.append(head)
        for c in hs[1]:
            c.font = Font(name="Arial", bold=True)
    
    hs = wb[SHEET_HISTORY]
    known = {str(hs.cell(r, 1).value) for r in range(2, hs.max_row + 1)}
    added = 0
    
    for s in sorted(subs, key=lambda x: str(x.get("_submission_time") or "")):
        if is_rejected(s) or str(s.get("_id")) in known:
            continue
        n = lot_number(s)
        hs.append([
            s.get("_id"), pdate(s.get("date_mission")), n, 
            STATUT.get(s.get("statut"), s.get("statut") or ""),
            int(float(s["taux_execution"])) if s.get("taux_execution") not in (None, "") else None,
            s.get("nom_agent") or "", str(s.get("_submission_time") or "")[:16].replace("T", " ")
        ])
        added += 1
    return added

# --------------------------------------------------------------------------
# 3. EMAIL (Optionnel)
# --------------------------------------------------------------------------

def send_mail(attachment: Path, today: dt.date) -> None:
    host = cfg("SMTP_HOST", required=True)
    port = int(cfg("SMTP_PORT", "587"))
    user = cfg("SMTP_USER", required=True)
    password = cfg("SMTP_PASSWORD", required=True)
    sender = cfg("MAIL_FROM", user)
    to = [a.strip() for a in cfg("MAIL_TO", required=True).split(",") if a.strip()]
    
    msg = EmailMessage()
    msg["Subject"] = f"Matrice de suivi – {today:%d/%m/%Y}"
    msg["From"], msg["To"] = sender, ", ".join(to)
    msg.set_content(f"Bonjour,\n\nVeuillez trouver ci-joint la matrice de suivi mise à jour.\n\nCordialement.")
    
    msg.add_attachment(
        attachment.read_bytes(), maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=attachment.name,
    )
    
    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as smtp:
            smtp.login(user, password)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=60) as smtp:
            smtp.starttls(context=ctx)
            smtp.login(user, password)
            smtp.send_message(msg)
    log.info("Email envoyé")

# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--from-json", help="Fichier JSON (test)")
    ap.add_argument("--base", help="Matrice de départ")
    ap.add_argument("--no-mail", action="store_true", help="Pas d'email")
    ap.add_argument("--env", default=".env", help="Fichier .env")
    args = ap.parse_args()
    
    load_env(Path(args.env))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    
    matrix_dir = Path(cfg("MATRIX_DIR", "./matrices"))
    matrix_dir.mkdir(parents=True, exist_ok=True)
    today = dt.date.today()
    
    if args.from_json:
        raw = json.loads(Path(args.from_json).read_text(encoding="utf-8"))
        raw = raw.get("results", raw) if isinstance(raw, dict) else raw
    else:
        raw = fetch_submissions()
    
    subs = [flatten(s) for s in raw]
    latest = latest_per_lot(subs)
    log.info("%d lot(s) avec visite valide", len(latest))
    
    base = Path(args.base) if args.base else None
    if not base:
        files = sorted(matrix_dir.glob(MATRIX_PATTERN), key=lambda p: p.stat().st_mtime, reverse=True)
        base = files[0] if files else None
    
    if not base:
        sys.exit("Aucune matrice trouvée")
    
    log.info("Matrice de départ: %s", base.name)
    wb = load_workbook(base)
    
    # Mise à jour de l'onglet principal
    ws_name = "MATRICE DE SUIVI" if "MATRICE DE SUIVI" in wb.sheetnames else wb.active.title
    update_matrix_sheet(wb[ws_name], latest)
    
    # Mise à jour de l'historique
    added = append_history(wb, subs)
    log.info("%d ligne(s) ajoutée(s) à l'historique", added)
    
    # Sauvegarde
    out = matrix_dir / f"MATRICE_DE_SUIVI_DES_ACTIVITES_du_{today:%d-%m-%y}.xlsx"
    wb.save(out)
    log.info("Matrice enregistrée: %s", out.name)
    
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
            f.write(f"matrix_path={out}\n")
    
    if not args.no_mail:
        send_mail(out, today)

if __name__ == "__main__":
    main()
