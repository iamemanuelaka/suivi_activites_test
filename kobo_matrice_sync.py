#!/usr/bin/env python3
"""
Synchronisation KoboToolbox -> MATRICE DE SUIVI DES ACTIVITES
+ Préparation Dynamic Data Attachments
+ Mise à jour Base de suivi détaillé
+ Envoi par e-mail

Usage:
    python kobo_matrice_sync.py
    python kobo_matrice_sync.py --no-mail
    python kobo_matrice_sync.py --prepare-dda
    python kobo_matrice_sync.py --from-json data.json --no-mail
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
import xml.etree.ElementTree as ET
from copy import copy
from email.message import EmailMessage
from pathlib import Path
from typing import Dict, List, Optional

import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# --------------------------------------------------------------------------
# Constantes de mapping
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

STATUTS_A_SURVEILLER = {"suspendu", "abandonne", "os_a_reprendre"}

HEADERS_MATRIX = {
    "region": "REGION",
    "departement": "DEPARTEMENT",
    "entreprise": "ENTREPRISE",
    "remise": "DATE DE REMISE",
    "taux": "TAUX D",
    "obs_tech": "OBSERVATIONS TECHNIQUES",
    "pges": "OBSERVATIONS ENVIRONNEMENTALES",
    "ong": "ONG MOBILISEES",
    "recep_prov": "DATE DE RECEPTION PROVISOIRE",
    "recep_def": "DATE DE RECEPTION DEFINITIVE",
    "montant_decaisse": "MONTANT DECAISSE",
    "taux_financier": "TAUX D'EXECUTION FINANCIERE",
    "volume_acheve": "VOLUME ACHEVE",
}

SHEET_HISTORY = "HISTORIQUE"
SHEET_BASE_SUIVI = "Base suivi"
MATRIX_PATTERN = "MATRICE_DE_SUIVI_DES_ACTIVITES_du_*.xlsx"

log = logging.getLogger("kobo_sync")

# --------------------------------------------------------------------------
# Configuration
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

# --------------------------------------------------------------------------
# 1. Récupération Kobo
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

def pdate(value) -> Optional[dt.date]:
    if not value:
        return None
    try:
        return dt.date.fromisoformat(str(value)[:10])
    except (ValueError, TypeError):
        return None

def is_rejected(sub: dict) -> bool:
    vs = sub.get("_validation_status") or {}
    uid = vs.get("uid") if isinstance(vs, dict) else vs
    return uid == "validation_status_not_approved"

def lot_number(sub: dict) -> Optional[int]:
    try:
        return int(str(sub.get("id_projet", "")).strip())
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
        key = (pdate(s.get("date_mission")) or dt.date.min, 
               str(s.get("_submission_time") or ""))
        if n not in best or key > best[n]["_key"]:
            best[n] = {**s, "_key": key}
    return best

# --------------------------------------------------------------------------
# 2. Préparation Dynamic Data Attachments
# --------------------------------------------------------------------------

def prepare_dda_data(subs: List[dict], output_dir: Path) -> Optional[Path]:
    """
    Prépare les données pour Dynamic Data Attachments.
    Crée un fichier XML que Kobo pourra récupérer pour le pré-remplissage.
    """
    if not subs:
        log.warning("Aucune soumission pour DDA")
        return None
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Regrouper par id_projet et garder la dernière soumission
    latest = latest_per_lot(subs)
    
    # Créer le XML
    root = ET.Element("data")
    
    for lot_num, sub in sorted(latest.items()):
        item = ET.SubElement(root, "item")
        
        # Identification
        ident = ET.SubElement(item, "identification")
        ET.SubElement(ident, "id_projet").text = str(lot_num)
        ET.SubElement(ident, "nom_agent").text = str(sub.get("nom_agent") or "")
        ET.SubElement(ident, "agent").text = str(sub.get("agent") or "")
        ET.SubElement(ident, "Structure").text = str(sub.get("Structure") or "")
        
        # Avancement
        avancement = ET.SubElement(item, "avancement")
        ET.SubElement(avancement, "date_remise_site").text = str(sub.get("date_remise_site") or "")
        ET.SubElement(avancement, "statut").text = str(sub.get("statut") or "")
        ET.SubElement(avancement, "taux_execution").text = str(sub.get("taux_execution") or "0")
        ET.SubElement(avancement, "Montant_d_caiss_sur_e_budget_total_FCFA").text = str(sub.get("Montant_d_caiss_sur_e_budget_total_FCFA") or "0")
        ET.SubElement(avancement, "Taux_d_ex_cution_financi_re_").text = str(sub.get("Taux_d_ex_cution_financi_re_") or "0")
        
        # PGES & ONG
        pges_ong = ET.SubElement(item, "pges_ong")
        ET.SubElement(pges_ong, "pges_statut").text = str(sub.get("pges_statut") or "")
        ET.SubElement(pges_ong, "Il_y_a_t_il_une_ONG_mobilis_e").text = str(sub.get("Il_y_a_t_il_une_ONG_mobilis_e") or "non")
        ET.SubElement(pges_ong, "ong_autre_nom").text = str(sub.get("ong_autre_nom") or "")
        
        # Réceptions
        receptions = ET.SubElement(item, "receptions")
        ET.SubElement(receptions, "recep_prov").text = str(sub.get("recep_prov") or "non")
        ET.SubElement(receptions, "date_recep_prov").text = str(sub.get("date_recep_prov") or "")
        ET.SubElement(receptions, "recep_def").text = str(sub.get("recep_def") or "non")
        ET.SubElement(receptions, "date_recep_def").text = str(sub.get("date_recep_def") or "")
    
    # Sauvegarder
    tree = ET.ElementTree(root)
    output_path = output_dir / "historique.xml"
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    
    log.info("Fichier DDA créé: %s (%d projets)", output_path, len(latest))
    return output_path

# --------------------------------------------------------------------------
# 3. Mise à jour Matrice
# --------------------------------------------------------------------------

def norm(s) -> str:
    s = re.sub(r"\s+", " ", str(s or "")).strip().upper()
    return s.translate(str.maketrans("ÉÈÊÀÂÎÔÛÇ", "EEEAAIOUC"))

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
    sys.exit("En-têtes introuvables")

def lot_rows(ws, header_row: int, cols: dict) -> Dict[int, int]:
    num_col = cols["region"] - 1
    rows = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, num_col).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            rows[int(v)] = r
        elif rows:
            break
    return rows

def update_matrix_sheet(ws_name: str, wb, latest: Dict[int, dict], is_base_suivi: bool = False) -> dict:
    ws = wb[ws_name] if ws_name in wb.sheetnames else wb.active
    
    if is_base_suivi:
        # Base suivi: une ligne par localité
        return update_base_suivi(ws, latest)
    else:
        # Matrice classique: une ligne par projet
        return update_matrice_classique(ws, latest)

def update_matrice_classique(ws, latest: Dict[int, dict]) -> dict:
    header_row, cols = find_columns(ws)
    rows = {}
    num_col = cols["region"] - 1
    
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, num_col).value
        if isinstance(v, (int, float)):
            rows[int(v)] = r
    
    for n, s in latest.items():
        r = rows.get(n)
        if r is None:
            continue
        
        # Statut + observations
        statut = s.get("statut")
        obs = (s.get("obs_techniques") or "").strip()
        if statut:
            txt = STATUT.get(statut, statut)
            ws.cell(r, cols.get("obs_tech", 1)).value = f"{txt} – {obs}" if obs else txt
        
        # Taux d'exécution
        taux = s.get("taux_execution")
        if taux not in (None, ""):
            ws.cell(r, cols.get("taux", 1)).value = int(float(taux))
        
        # Date de remise de site
        d = pdate(s.get("date_remise_site"))
        if d and "remise" in cols:
            ws.cell(r, cols["remise"]).value = d
            ws.cell(r, cols["remise"]).number_format = "DD/MM/YYYY"
        
        # PGES
        if "pges" in cols and s.get("pges_statut"):
            ws.cell(r, cols["pges"]).value = PGES.get(s["pges_statut"], s["pges_statut"])
        
        # ONG
        if "ong" in cols:
            ong = s.get("ong_mobilisee")
            if ong:
                ws.cell(r, cols["ong"]).value = ONG.get(ong, ong)
        
        # Réceptions
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
        
        # NOUVEAUX CHAMPS
        if "montant_decaisse" in cols:
            val = s.get("Montant_d_caiss_sur_e_budget_total_FCFA")
            if val not in (None, ""):
                ws.cell(r, cols["montant_decaisse"]).value = int(float(val))
        
        if "taux_financier" in cols:
            val = s.get("Taux_d_ex_cution_financi_re_")
            if val not in (None, ""):
                ws.cell(r, cols["taux_financier"]).value = int(float(val))
        
        if "volume_acheve" in cols:
            val = s.get("Volume_de_travaux_achev_Km")
            if val not in (None, ""):
                ws.cell(r, cols["volume_acheve"]).value = float(val)
    
    return {"rows": rows}

def update_base_suivi(ws, latest: Dict[int, dict]) -> dict:
    """
    Met à jour l'onglet Base suivi (une ligne par localité).
    Pour l'instant, on met à jour toutes les lignes d'un même projet
    avec les mêmes valeurs (taux, observations, etc.)
    """
    # Trouver la colonne id_projet ou numéro de projet
    header_row = 1
    cols_base = {}
    
    for c in range(1, ws.max_column + 1):
        h = norm(ws.cell(1, c).value)
        if "ID_PROJET" in h or "N°" in h or "NUMERO" in h:
            cols_base["id_projet"] = c
        elif "TAUX" in h and "FINANCIER" not in h and "EXECUTION" in h:
            cols_base["taux"] = c
        elif "OBS" in h and "TECH" in h:
            cols_base["obs_tech"] = c
        elif "PGES" in h:
            cols_base["pges"] = c
        elif "ONG" in h:
            cols_base["ong"] = c
    
    if "id_projet" not in cols_base:
        log.warning("Colonne id_projet introuvable dans Base suivi")
        return {"rows": {}}
    
    # Regrouper les lignes par id_projet
    rows_by_lot: Dict[int, List[int]] = {}
    for r in range(2, ws.max_row + 1):
        try:
            lot_num = int(ws.cell(r, cols_base["id_projet"]).value)
            if lot_num not in rows_by_lot:
                rows_by_lot[lot_num] = []
            rows_by_lot[lot_num].append(r)
        except (ValueError, TypeError):
            continue
    
    # Mettre à jour
    for lot_num, s in latest.items():
        if lot_num not in rows_by_lot:
            continue
        
        for r in rows_by_lot[lot_num]:
            # Taux d'exécution
            if "taux" in cols_base:
                taux = s.get("taux_execution")
                if taux not in (None, ""):
                    ws.cell(r, cols_base["taux"]).value = int(float(taux))
            
            # Observations
            if "obs_tech" in cols_base:
                statut = s.get("statut")
                obs = (s.get("obs_techniques") or "").strip()
                if statut:
                    txt = STATUT.get(statut, statut)
                    ws.cell(r, cols_base["obs_tech"]).value = f"{txt} – {obs}" if obs else txt
            
            # PGES
            if "pges" in cols_base and s.get("pges_statut"):
                ws.cell(r, cols_base["pges"]).value = PGES.get(s["pges_statut"], s["pges_statut"])
            
            # ONG
            if "ong" in cols_base:
                ong = s.get("ong_mobilisee")
                if ong:
                    ws.cell(r, cols_base["ong"]).value = ONG.get(ong, ong)
    
    return {"rows": rows_by_lot}

def append_history(wb, subs: List[dict], ident: Dict[int, dict]) -> int:
    head = ["ID KOBO", "DATE VISITE", "N° LOT", "REGION", "DEPARTEMENT", "ENTREPRISE",
            "STATUT", "TAUX (%)", "PGES", "ONG", "RAPPORTEUR", "SOUMIS LE"]
    
    if SHEET_HISTORY not in wb.sheetnames:
        hs = wb.create_sheet(SHEET_HISTORY)
        hs.append(head)
    
    hs = wb[SHEET_HISTORY]
    if hs.max_row == 1:
        for c in hs[1]:
            c.font = Font(name="Arial", bold=True)
            c.alignment = Alignment(wrap_text=True)
        hs.freeze_panes = "A2"
    
    known = {str(hs.cell(r, 1).value) for r in range(2, hs.max_row + 1)}
    added = 0
    
    for s in sorted(subs, key=lambda x: (str(x.get("date_mission")), str(x.get("_submission_time")))):
        if is_rejected(s) or str(s.get("_id")) in known:
            continue
        
        n = lot_number(s)
        i = ident.get(n, {})
        ong = s.get("ong_mobilisee")
        ong_txt = (s.get("ong_autre_nom") or "ONG mobilisée") if ong == "mobilisee_autre" else ONG.get(ong, "")
        
        hs.append([
            s.get("_id"), pdate(s.get("date_mission")), n, 
            i.get("region", ""), i.get("departement", ""), i.get("entreprise", ""),
            STATUT.get(s.get("statut"), s.get("statut") or ""),
            int(float(s["taux_execution"])) if s.get("taux_execution") not in (None, "") else None,
            PGES.get(s.get("pges_statut"), s.get("pges_statut") or ""), 
            ong_txt,
            s.get("nom_agent") or s.get("agent") or "", 
            str(s.get("_submission_time") or "")[:16].replace("T", " "),
        ])
        added += 1
    
    return added

# --------------------------------------------------------------------------
# 4. Email
# --------------------------------------------------------------------------

def build_summary(info: dict, latest: Dict[int, dict], today: dt.date) -> dict:
    since = today - dt.timedelta(days=7)
    reported, missing, watch = [], [], []
    
    for n in sorted(info.get("rows", {})):
        s = latest.get(n)
        d = pdate(s.get("date_mission")) if s else None
        if s and d and d >= since:
            reported.append(f"Lot {n:02d}")
        else:
            missing.append(f"Lot {n:02d}")
        
        if s and s.get("statut") in STATUTS_A_SURVEILLER:
            watch.append(f"Lot {n:02d}: {STATUT[s['statut']]}")
    
    return {"reported": reported, "missing": missing, "watch": watch, "mean": info.get("mean", 0)}

def send_mail(attachment: Path, summ: dict, today: dt.date) -> None:
    host = cfg("SMTP_HOST", required=True)
    port = int(cfg("SMTP_PORT", "587"))
    user = cfg("SMTP_USER", required=True)
    password = cfg("SMTP_PASSWORD", required=True)
    sender = cfg("MAIL_FROM", user)
    to = [a.strip() for a in cfg("MAIL_TO", required=True).split(",") if a.strip()]
    
    msg = EmailMessage()
    msg["Subject"] = f"Matrice de suivi – {today:%d/%m/%Y}"
    msg["From"], msg["To"] = sender, ", ".join(to)
    
    body = (
        f"Bonjour,\n\n"
        f"Matrice de suivi mise à jour au {today:%d/%m/%Y}.\n"
        f"Lots renseignés: {len(summ['reported'])}\n"
        f"\nMessage automatique."
    )
    msg.set_content(body)
    
    msg.add_attachment(
        attachment.read_bytes(), 
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        filename=attachment.name,
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
    
    log.info("Email envoyé à %s", ", ".join(to))

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--from-json", help="Fichier JSON (test hors ligne)")
    ap.add_argument("--base", help="Matrice de départ")
    ap.add_argument("--no-mail", action="store_true", help="Pas d'email")
    ap.add_argument("--prepare-dda", action="store_true", help="Préparer DDA")
    ap.add_argument("--env", default=".env", help="Fichier .env")
    args = ap.parse_args()
    
    load_env(Path(args.env))
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    
    matrix_dir = Path(cfg("MATRIX_DIR", "./matrices"))
    matrix_dir.mkdir(parents=True, exist_ok=True)
    
    today = dt.date.today()
    
    # Récupérer soumissions
    if args.from_json:
        raw = json.loads(Path(args.from_json).read_text(encoding="utf-8"))
        raw = raw.get("results", raw) if isinstance(raw, dict) else raw
    else:
        raw = fetch_submissions()
    
    subs = [flatten(s) for s in raw]
    latest = latest_per_lot(subs)
    log.info("%d lot(s) avec visite valide", len(latest))
    
    # Préparer Dynamic Data Attachments
    dda_path = None
    if args.prepare_dda:
        dda_dir = Path(cfg("DDA_OUTPUT_DIR", "./data/dda"))
        dda_path = prepare_dda_data(subs, dda_dir)
    
    # Mettre à jour matrice
    base = Path(args.base) if args.base else None
    if not base:
        files = sorted(matrix_dir.glob(MATRIX_PATTERN), 
                      key=lambda p: p.stat().st_mtime, reverse=True)
        base = files[0] if files else None
    
    if not base:
        sys.exit("Aucune matrice trouvée")
    
    log.info("Matrice de départ: %s", base.name)
    wb = load_workbook(base)
    
    # Mettre à jour onglet MATRICE DE SUIVI
    info_matrix = update_matrix_sheet("MATRICE DE SUIVI", wb, latest, is_base_suivi=False)
    
    # Mettre à jour onglet Base suivi (si existe)
    if SHEET_BASE_SUIVI in wb.sheetnames:
        info_base = update_matrix_sheet(SHEET_BASE_SUIVI, wb, latest, is_base_suivi=True)
        log.info("Base suivi mise à jour: %d localités", sum(len(v) for v in info_base.get("rows", {}).values()))
    
    # Historique
    added = append_history(wb, subs, {})
    log.info("%d ligne(s) ajoutée(s) à l'historique", added)
    
    # Sauvegarder
    out = matrix_dir / f"MATRICE_DE_SUIVI_DES_ACTIVITES_du_{today:%d-%m-%y}.xlsx"
    wb.save(out)
    log.info("Matrice enregistrée: %s", out.name)
    
    # Outputs pour GitHub Actions
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
            f.write(f"matrix_path={out}\n")
            if dda_path:
                f.write(f"dda_path={dda_path}\n")
    
    # Email
    if not args.no_mail:
        send_mail(out, build_summary(info_matrix, latest, today), today)
    else:
        log.info("Email désactivé")

if __name__ == "__main__":
    main()
