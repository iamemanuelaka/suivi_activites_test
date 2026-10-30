#!/usr/bin/env python3
"""
Synchronisation KoboToolbox -> MATRICE DE SUIVI DES ACTIVITES (+ tableau de bord) -> e-mail.

Étapes :
  1. Récupère les soumissions du formulaire via l'API KoboToolbox (v2), avec reprises automatiques.
  2. Garde, pour chaque lot (activité), la visite la plus récente.
  3. Met à jour les CELLULES de la dernière matrice de MATRIX_DIR (colonnes repérées par leur intitulé) et
     l'enregistre sous MATRICE_DE_SUIVI_DES_ACTIVITES_du_JJ-MM-AA.xlsx. Les graphiques, la carte, les formules,
     les filtres et le reste du tableau de bord sont conservés tels quels (voir xlsx_patch.py).
  4. Renseigne la feuille META (date d'actualisation + historique du taux moyen lus par le tableau de bord).
  5. Tient à jour le journal des visites matrices/historique_visites.csv (une ligne par soumission Kobo).
  6. Envoie la matrice par e-mail avec un résumé (mode hebdomadaire).

Usage :
  python kobo_matrice_sync.py                      # cycle complet (API + mail)
  python kobo_matrice_sync.py --no-mail            # sans envoi d'e-mail
  python kobo_matrice_sync.py --from-json data.json --no-mail   # test hors ligne
"""
import argparse
import csv
import datetime as dt
import hashlib
import json
import logging
import os
import re
import smtplib
import ssl
import sys
from email.message import EmailMessage
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from xlsx_patch import Package, idx_to_col

# --------------------------------------------------------------------------
# Libellés : valeurs Kobo (noms de choix) -> texte écrit dans la matrice
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
# Libellés de la colonne « STATUT DU CHANTIER » de la matrice (= liste PARAMETRES!I13:I19 du tableau de bord)
STATUT_MATRICE = {
    "non_demarre": "Non démarré", "demarrage_autorise": "Démarrage autorisé", "os_a_reprendre": "OS à reprendre",
    "demarrage_effectif": "En cours", "travaux_en_cours": "En cours", "travaux_acheves": "Achevé",
    "suspendu": "Suspendu", "abandonne": "Abandonné",
}
PGES = {
    "valide": "Validé",
    "transmis_validation": "Transmis en cours de validation",
    "non_transmis": "Non transmis",
    "pas_de_pges": "Pas de PGES",
}
ONG = {
    "ong_epc": "ONG EPC",
    "agir_odd": "AGIR pour les ODD",
    "kera_sy_ire": "KERA-SY-IRÊ",
    "etoile_bounkani": "ONG Étoile du Bounkani",
    "cibes": "CIBES",
    "voix_femme": "ONG Voix de Femme",
    "pietro_bonili": "Centre de santé PIETRO BONILI",
    "solidarite": "Solidarité",
    "victoire_plus": "ONG Victoire Plus",
    "esperance_vie": "Espérance Vie",
    "non_mobilisee": "Non mobilisée",
}
STATUTS_A_SURVEILLER = {"suspendu", "abandonne", "os_a_reprendre"}
FONCTION = {"gar_principal": "Suivi opérationnel", "gar_adjoint": "Sauvegarde sociale et environnementale", "superviseur": "Suivi évaluation", "autre": "Autre"}
STRUCTURE = {"pcr": "PCR", "missions_de_contr_le": "Missions de contrôle", "consultant": "Consultant", "autre": "Autre"}
MATRIX_SHEET = "MATRICE DE SUIVI"
META_SHEET = "META"
META_MAX_ROWS = 450          # le tableau de bord lit META!B4:B500
MATRIX_PATTERN = "MATRICE_DE_SUIVI_DES_ACTIVITES_du_*.xlsx"
HISTORY_FILE = "historique_visites.csv"
STATE_FILE = "sync_state.json"

log = logging.getLogger("kobo_sync")


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
def load_env(path: Path) -> None:
    """Charge un fichier .env simple (CLE=valeur) sans écraser l'environnement."""
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
        sys.exit(f"Variable de configuration manquante : {name} (voir .env.example)")
    return val


# --------------------------------------------------------------------------
# 1. Kobo
# --------------------------------------------------------------------------
def http_session() -> requests.Session:
    """Session avec reprises automatiques (erreurs réseau, 429, 5xx) et temporisation croissante."""
    s = requests.Session()
    retry = Retry(total=5, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504), allowed_methods=("GET",), respect_retry_after_header=True)
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://", HTTPAdapter(max_retries=retry))
    return s


def fetch_submissions() -> list[dict]:
    base = cfg("KOBO_URL", "https://kf.kobotoolbox.org").rstrip("/")
    uid = cfg("KOBO_ASSET_UID", required=True)
    headers = {"Authorization": f"Token {cfg('KOBO_TOKEN', required=True)}"}
    url = f"{base}/api/v2/assets/{uid}/data/?format=json&limit=1000"
    out, session = [], http_session()
    while url:
        try:
            r = session.get(url, headers=headers, timeout=60)
            r.raise_for_status()
        except requests.HTTPError as e:
            code = e.response.status_code if e.response is not None else None
            hint = {401: "jeton KOBO_TOKEN invalide ou expiré", 403: "ce jeton n'a pas accès au projet (droit « voir les données »)",
                    404: "KOBO_ASSET_UID ou KOBO_URL incorrect (serveur eu.kobotoolbox.org ou kf.kobotoolbox.org ?)"}.get(code, "erreur du serveur Kobo")
            raise SystemExit(f"Échec de la récupération Kobo (HTTP {code}) : {hint}.") from e
        except requests.RequestException as e:
            raise SystemExit(f"Kobo injoignable après plusieurs tentatives : {e}") from e
        payload = r.json()
        out.extend(payload.get("results", []))
        url = payload.get("next")
    log.info("%d soumission(s) récupérée(s) depuis Kobo", len(out))
    return out


def flatten(sub: dict) -> dict:
    """Retire les préfixes de groupes ('avancement/taux_execution' -> 'taux_execution')."""
    flat = {k.split("/")[-1]: v for k, v in sub.items()}
    flat["_id"] = sub.get("_id")
    flat["_submission_time"] = sub.get("_submission_time")
    flat["_validation_status"] = sub.get("_validation_status")
    return flat


def pdate(value) -> dt.date | None:
    if not value:
        return None
    try:
        return dt.date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def is_rejected(sub: dict) -> bool:
    vs = sub.get("_validation_status") or {}
    uid = vs.get("uid") if isinstance(vs, dict) else vs
    return uid == "validation_status_not_approved"


def lot_number(sub: dict) -> int | None:
    try:
        return int(str(sub.get("id_projet", "")).strip())
    except ValueError:
        return None


def latest_per_lot(subs: list[dict]) -> dict[int, dict]:
    """Dernière visite (date de visite, puis heure de soumission) par lot."""
    best: dict[int, dict] = {}
    for s in subs:
        if is_rejected(s):
            continue
        n = lot_number(s)
        if n is None:
            log.warning("Soumission %s ignorée : id_projet invalide", s.get("_id"))
            continue
        key = (pdate(s.get("date_mission")) or dt.date.min, str(s.get("_submission_time") or ""))
        if n not in best or key > best[n]["_key"]:
            best[n] = {**s, "_key": key}
    return best


# --------------------------------------------------------------------------
# 2. Matrice (modification des CELLULES uniquement : le reste du classeur est conservé tel quel)
# --------------------------------------------------------------------------
def norm(s) -> str:
    s = re.sub(r"\s+", " ", str(s or "")).strip().upper().replace("’", "'")
    return s.translate(str.maketrans("ÉÈÊËÀÂÎÏÔÛÙÇ", "EEEEAAIIOUUC"))


# Règles de reconnaissance des colonnes (sur l'intitulé normalisé). L'ordre n'a pas d'importance :
# chaque règle est exclusive, ce qui évite de confondre « TAUX D'EXECUTION FINANCIERE » et « TAUX D'EXECUTION (%) ».
HEADER_RULES = {
    "region":      lambda h: h.startswith("REGION"),
    "departement": lambda h: h.startswith("DEPARTEMENT"),
    "entreprise":  lambda h: h.startswith("ENTREPRISE"),
    "debut":       lambda h: h.startswith("DEBUT DES TRAVAUX") or h.startswith("DEBUT ET FIN"),
    "montant_dec": lambda h: h.startswith("MONTANT DECAISSE") or h.startswith("UTILISATION DU BUDGET"),
    "taux_fin":    lambda h: h.startswith("TAUX D") and "FINANC" in h,
    "taux":        lambda h: h.startswith("TAUX D") and "FINANC" not in h,
    "volume":      lambda h: h.startswith("VOLUME"),
    "remise":      lambda h: h.startswith("DATE DE REMISE"),
    "statut":      lambda h: h.startswith("STATUT DU CHANTIER") or h == "STATUT",
    "obs_tech":    lambda h: h.startswith("OBSERVATIONS TECHNIQUES"),
    "pges":        lambda h: h.startswith("OBSERVATIONS ENVIRONNEMENTALES"),
    "ong":         lambda h: h.startswith("ONG MOBILISEES"),
    "recep_prov":  lambda h: h.startswith("DATE DE RECEPTION PROVISOIRE"),
    "recep_def":   lambda h: h.startswith("DATE DE RECEPTION DEFINITIVE"),
    "photo":       lambda h: h.startswith("PHOTO"),
    "repondant":   lambda h: ("REPONDANT" in h) or h.startswith("NOM PRENOMS"),
}
REQUIRED_COLS = ("region", "departement", "entreprise", "taux")


def find_columns(ws) -> tuple[int, dict[str, int]]:
    for r in range(1, 11):
        found: dict[str, int] = {}
        for c in range(1, 60):
            h = norm(ws.get(r, c))
            if not h:
                continue
            for key, rule in HEADER_RULES.items():
                if key not in found and rule(h):
                    found[key] = c
        if all(k in found for k in REQUIRED_COLS):
            return r, found
    raise SystemExit("En-têtes de la matrice introuvables (REGIONS / DEPARTEMENT / ENTREPRISE / TAUX D'EXECUTION). "
                     "La structure de la feuille a-t-elle changé ?")


def lot_rows(ws, header_row: int, cols: dict[str, int]) -> dict[int, int]:
    num_col = cols["region"] - 1  # la colonne des numéros est juste avant 'REGIONS'
    rows: dict[int, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.get(r, num_col)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            rows[int(v)] = r
        elif rows:
            break  # fin du tableau
    return rows


def lot_identity(ws, rows: dict[int, int], cols: dict[str, int]) -> dict[int, dict]:
    ident, region = {}, ""
    for n, r in sorted(rows.items()):
        region = ws.get(r, cols["region"]) or region  # région portée par une cellule fusionnée
        ident[n] = {"region": str(region).strip(), "departement": str(ws.get(r, cols["departement"]) or "").strip(),
                    "entreprise": str(ws.get(r, cols["entreprise"]) or "").strip()}
    return ident


def ong_text(s: dict) -> str | None:
    """Nouveau formulaire : oui/non + nom ; ancien formulaire : liste de choix."""
    new = s.get("Il_y_a_t_il_une_ONG_mobilis_e")
    if new:
        if str(new).lower() in ("non", "no"):
            return ONG["non_mobilisee"]
        return (s.get("ong_autre_nom") or "").strip() or "ONG mobilisée"
    old = s.get("ong_mobilisee")
    if old:
        if old == "mobilisee_autre":
            return (s.get("ong_autre_nom") or "").strip() or "ONG mobilisée"
        return ONG.get(old, old)
    return None


def pretty(code) -> str:
    return str(code or "").replace("_", " ").strip().capitalize()


def respondent(s: dict) -> dict:
    structure = s.get("Structure") or s.get("structure") or ""
    if structure == "autre" and s.get("Autre_pr_ciser"):
        structure_txt = str(s["Autre_pr_ciser"]).strip()
    else:
        structure_txt = STRUCTURE.get(structure, pretty(structure))
    fonction = s.get("agent") or ""
    return {"nom": (s.get("nom_agent") or "").strip(), "fonction": FONCTION.get(fonction, pretty(fonction)), "structure": structure_txt}


def respondent_text(s: dict) -> str:
    r = respondent(s)
    if not r["nom"] and not r["structure"]:
        return ""
    txt = r["nom"] or "—"
    if r["structure"]:
        txt += f" - {r['structure']}"
    if r["fonction"]:
        txt += f" ({r['fonction']})"
    return txt


def photo_names(s: dict) -> list[str]:
    return [str(s[k]) for k in ("photo_1", "photo_2") if s.get(k)]


def photo_text(s: dict) -> str:
    n = len(photo_names(s))
    if not n:
        return ""
    d = pdate(s.get("date_mission"))
    return f"{n} photo(s) – visite du {d:%d/%m/%Y} – Kobo n° {s.get('_id')}" if d else f"{n} photo(s) – Kobo n° {s.get('_id')}"


def num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f.is_integer() else f


def update_matrix(pkg: Package, latest: dict[int, dict], sheet_name: str) -> dict:
    ws = pkg.sheet(sheet_name)
    header_row, cols = find_columns(ws)
    rows = lot_rows(ws, header_row, cols)
    ident = lot_identity(ws, rows, cols)
    missing = [k for k in HEADER_RULES if k not in cols]
    if missing:
        log.info("Colonnes absentes de cette matrice (ignorées) : %s", ", ".join(missing))
    skipped_formula = []

    def put(r, key, value, date=False):
        if key not in cols or value in (None, ""):
            return
        style_from = (r, cols.get("debut", cols.get("remise", cols[key]))) if date else None
        if not ws.set(r, cols[key], value, style_from=style_from):
            skipped_formula.append(f"{idx_to_col(cols[key])}{r}")

    for n, s in latest.items():
        r = rows.get(n)
        if r is None:
            log.warning("Lot %s absent de la matrice, soumission %s ignorée", n, s.get("_id"))
            continue
        statut = s.get("statut")
        obs = (s.get("obs_techniques") or "").strip()
        if "statut" in cols:
            if statut:
                put(r, "statut", STATUT_MATRICE.get(statut, STATUT.get(statut, statut)))
            if "obs_tech" in cols and ws.set(r, cols["obs_tech"], obs or None) is False:   # l'observation de la dernière visite remplace la précédente (effacée si vide)
                skipped_formula.append(f"{idx_to_col(cols['obs_tech'])}{r}")
        elif statut:                                                    # ancienne matrice sans colonne statut : « statut – observation »
            txt = STATUT.get(statut, statut)
            put(r, "obs_tech", f"{txt} – {obs}" if obs else txt)
        t = num(s.get("taux_execution"))
        if t is not None:
            put(r, "taux", t)
        put(r, "montant_dec", num(s.get("Montant_d_caiss_sur_e_budget_total_FCFA")))
        put(r, "taux_fin", num(s.get("Taux_d_ex_cution_financi_re_")))
        put(r, "volume", num(s.get("Volume_de_travaux_achev_Km")))
        put(r, "remise", pdate(s.get("date_remise_site")), date=True)
        if s.get("pges_statut"):
            put(r, "pges", PGES.get(s["pges_statut"], s["pges_statut"]))
        put(r, "ong", ong_text(s))
        if s.get("recep_prov") == "oui":
            put(r, "recep_prov", pdate(s.get("date_recep_prov")), date=True)
        if s.get("recep_def") == "oui":
            put(r, "recep_def", pdate(s.get("date_recep_def")), date=True)
        put(r, "photo", photo_text(s))
        put(r, "repondant", respondent_text(s))
    if skipped_formula:
        log.warning("Cellules contenant une formule, non modifiées : %s", ", ".join(skipped_formula))

    vals = [ws.get(r, cols["taux"]) for r in rows.values()]
    vals = [v * 100 if 0 < v < 1 else v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]  # même lecture que le tableau de bord
    mean = round(sum(vals) / len(vals), 2) if vals else 0
    return {"rows": rows, "ident": ident, "mean": mean, "cols": cols, "formula_cells": skipped_formula}


def update_meta(pkg: Package, now: dt.datetime, mean: float) -> None:
    """Alimente la feuille META lue par le tableau de bord (date d'actualisation + historique du taux moyen)."""
    if not pkg.has_sheet(META_SHEET):
        log.info("Feuille %s absente : horodatage non écrit.", META_SHEET)
        return
    m = pkg.sheet(META_SHEET)
    m.set(1, 2, now, style_from=(1, 2))
    r, last = 4, None
    while m.get(r, 1) is not None or m.get(r, 2) is not None:
        last = m.get(r, 2)
        r += 1
    if r > META_MAX_ROWS:                       # le tableau de bord lit META!B4:B500 : on garde les 300 dernières mesures
        keep = [(m.get(i, 1), m.get(i, 2)) for i in range(r - 300, r)]
        for i in range(4, r):
            m.set(i, 1, None); m.set(i, 2, None)
        for i, (a, b) in enumerate(keep):
            m.set(4 + i, 1, a, style_from=(4, 1)); m.set(4 + i, 2, b, style_from=(4, 2))
        r = 4 + len(keep)
    if last is None or abs(float(last) - mean) >= 0.05:
        m.set(r, 1, now, style_from=(4, 1))
        m.set(r, 2, mean, style_from=(4, 2))
        log.info("META : mesure ajoutée (ligne %d, taux moyen %s %%)", r, mean)


HISTORY_COLS = ["id_kobo", "date_visite", "n_lot", "region", "departement", "entreprise", "statut", "taux_physique", "montant_decaisse",
                "taux_financier", "pges", "ong", "repondant_nom", "repondant_fonction", "repondant_structure", "nb_photos", "photo_1", "photo_2", "soumis_le"]


def write_history(path: Path, subs: list[dict], ident: dict[int, dict]) -> tuple[int, int]:
    """Journal de toutes les visites (miroir de Kobo, une ligne par soumission, clé = id Kobo : ni doublon ni perte)."""
    old: dict[str, dict] = {}
    if path.exists():
        with path.open(encoding="utf-8", newline="") as fh:
            old = {row["id_kobo"]: row for row in csv.DictReader(fh)}
    new: dict[str, dict] = {}
    for s in subs:
        if is_rejected(s) or s.get("_id") is None:
            continue
        n = lot_number(s)
        i = ident.get(n, {})
        r = respondent(s)
        new[str(s["_id"])] = {
            "id_kobo": s["_id"], "date_visite": s.get("date_mission") or "", "n_lot": n or "", "region": i.get("region", ""),
            "departement": i.get("departement", ""), "entreprise": i.get("entreprise", ""), "statut": STATUT.get(s.get("statut"), s.get("statut") or ""),
            "taux_physique": s.get("taux_execution") or "", "montant_decaisse": s.get("Montant_d_caiss_sur_e_budget_total_FCFA") or "",
            "taux_financier": s.get("Taux_d_ex_cution_financi_re_") or "", "pges": PGES.get(s.get("pges_statut"), s.get("pges_statut") or ""),
            "ong": ong_text(s) or "", "repondant_nom": r["nom"], "repondant_fonction": r["fonction"], "repondant_structure": r["structure"],
            "nb_photos": len(photo_names(s)), "photo_1": s.get("photo_1") or "", "photo_2": s.get("photo_2") or "",
            "soumis_le": str(s.get("_submission_time") or "")[:16].replace("T", " "),
        }
    merged = {**old, **{k: {c: str(v) for c, v in row.items()} for k, row in new.items()}}   # une soumission supprimée dans Kobo reste dans le journal
    added = len(set(new) - set(old))
    updated = sum(1 for k in new if k in old and {c: str(v) for c, v in new[k].items()} != {c: old[k].get(c, "") for c in HISTORY_COLS})
    rows = sorted(merged.values(), key=lambda x: (x.get("date_visite", ""), int(x["id_kobo"]) if str(x["id_kobo"]).isdigit() else 0))
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=HISTORY_COLS, extrasaction="ignore")
        w.writeheader(); w.writerows(rows)
    return added, updated


def fingerprint(latest: dict[int, dict]) -> str:
    """Empreinte de tout ce qui peut modifier la matrice : détecte aussi les soumissions MODIFIÉES ou supprimées dans Kobo."""
    payload = {str(n): {k: v for k, v in s.items() if k != "_key"} for n, s in sorted(latest.items())}
    return hashlib.sha256(json.dumps(payload, sort_keys=True, default=str, ensure_ascii=False).encode("utf-8")).hexdigest()


def load_state(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def step_summary(lines: list[str]) -> None:
    """Écrit un résumé lisible dans l'onglet du run GitHub Actions."""
    p = os.environ.get("GITHUB_STEP_SUMMARY")
    if p:
        with open(p, "a", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")


def find_base_matrix(matrix_dir: Path, explicit: str | None) -> Path:
    if explicit:
        return Path(explicit)
    files = sorted(matrix_dir.glob(MATRIX_PATTERN), key=matrix_date)
    if not files:
        sys.exit(f"Aucune matrice '{MATRIX_PATTERN}' trouvée dans {matrix_dir}")
    return files[-1]


def matrix_date(p: Path) -> tuple:
    m = re.search(r"du_(\d{2})-(\d{2})-(\d{2})", p.name)
    d = dt.date(2000 + int(m[3]), int(m[2]), int(m[1])) if m else dt.date.min
    return (d, p.stat().st_mtime)


def prune_old(matrix_dir: Path, keep: int, current: Path) -> None:
    """Ne conserve que les `keep` matrices les plus récentes (le dépôt Git garde l'historique)."""
    files = sorted(matrix_dir.glob(MATRIX_PATTERN), key=matrix_date, reverse=True)
    for f in files[keep:]:
        if f != current:
            f.unlink()
            log.info("Ancienne matrice supprimée : %s", f.name)


# --------------------------------------------------------------------------
# 3. Résumé et e-mail
# --------------------------------------------------------------------------
def build_summary(info: dict, latest: dict[int, dict], today: dt.date) -> dict:
    since = today - dt.timedelta(days=7)
    reported, missing, watch = [], [], []
    for n in sorted(info["rows"]):
        s = latest.get(n)
        label = f"Lot {n:02d} – {info['ident'][n]['departement']} ({info['ident'][n]['entreprise']})"
        d = pdate(s.get("date_mission")) if s else None
        if s and d and d >= since:
            reported.append(label)
        else:
            missing.append(label)
        if s and s.get("statut") in STATUTS_A_SURVEILLER:
            watch.append(f"{label} : {STATUT[s['statut']]}")
    return {"reported": reported, "missing": missing, "watch": watch, "mean": info["mean"], "total": len(info["rows"])}


def mail_body(summ: dict, today: dt.date) -> str:
    lines = [
        "Bonjour,",
        "",
        f"Veuillez trouver ci-joint la matrice de suivi des activités mise à jour au {today:%d/%m/%Y}.",
        "",
        f"Taux d'exécution moyen : {summ['mean']} %",
        f"Lots renseignés ces 7 derniers jours : {len(summ['reported'])} sur {summ['total']}",
    ]
    if summ["missing"]:
        lines += ["", "Lots sans visite cette semaine :"] + [f"  - {x}" for x in summ["missing"]]
    if summ["watch"]:
        lines += ["", "Points d'attention :"] + [f"  - {x}" for x in summ["watch"]]
    lines += ["", "Message généré automatiquement à partir des données KoboToolbox.", "", "Cordialement"]
    return "\n".join(lines)


def send_mail(attachment: Path, summ: dict, today: dt.date) -> None:
    host = cfg("SMTP_HOST", required=True)
    port = int(cfg("SMTP_PORT", "587"))
    user = cfg("SMTP_USER", required=True)
    password = cfg("SMTP_PASSWORD", required=True)
    sender = cfg("MAIL_FROM", user)
    to = [a.strip() for a in cfg("MAIL_TO", required=True).split(",") if a.strip()]

    msg = EmailMessage()
    msg["Subject"] = f"Matrice de suivi des activités – {today:%d/%m/%Y}"
    msg["From"], msg["To"] = sender, ", ".join(to)
    msg.set_content(mail_body(summ, today))
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
    log.info("E-mail envoyé à %s", ", ".join(to))


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def run(args) -> int:
    load_env(Path(args.env))
    matrix_dir = Path(cfg("MATRIX_DIR", "./matrices"))
    matrix_dir.mkdir(parents=True, exist_ok=True)
    now = dt.datetime.now().replace(microsecond=0)
    today = now.date()

    if args.from_json:
        raw = json.loads(Path(args.from_json).read_text(encoding="utf-8"))
        raw = raw.get("results", raw) if isinstance(raw, dict) else raw
    else:
        raw = fetch_submissions()
    subs = [flatten(s) for s in raw]
    latest = latest_per_lot(subs)
    log.info("%d soumission(s), %d lot(s) avec au moins une visite valide", len(subs), len(latest))

    base = find_base_matrix(matrix_dir, args.base)
    log.info("Matrice de départ : %s", base.name)
    pkg = Package(base)
    info = update_matrix(pkg, latest, cfg("MATRIX_SHEET", MATRIX_SHEET))

    added, updated = write_history(matrix_dir / HISTORY_FILE, subs, info["ident"])
    log.info("Journal des visites : %d ajoutée(s), %d modifiée(s)", added, updated)

    fp, state_path = fingerprint(latest), matrix_dir / STATE_FILE
    if args.skip_if_unchanged and load_state(state_path).get("fingerprint") == fp:
        log.info("Aucune nouvelle soumission ni modification : matrice inchangée, rien n'est enregistré.")
        step_summary(["### Sync Kobo", "Aucun changement (aucune nouvelle soumission, aucune modification)."])
        return 0

    update_meta(pkg, now, info["mean"])
    pkg.force_full_calc()
    out = matrix_dir / f"MATRICE_DE_SUIVI_DES_ACTIVITES_du_{today:%d-%m-%y}.xlsx"
    pkg.save(out)
    state_path.write_text(json.dumps({"fingerprint": fp, "last_run": now.isoformat(), "matrix": out.name, "lots": len(latest)}, indent=1), encoding="utf-8")
    log.info("Matrice enregistrée : %s (taux moyen %s %%)", out, info["mean"])
    step_summary(["### Sync Kobo", f"- Matrice : `{out.name}`", f"- Lots mis à jour : **{len(latest)}**  ·  taux moyen **{info['mean']} %**",
                  f"- Journal : +{added} / ~{updated} ligne(s)"] + ([f"- ⚠ Cellules à formule non modifiées : {', '.join(info['formula_cells'])}"] if info["formula_cells"] else []))
    prune_old(matrix_dir, keep=int(cfg("MATRIX_KEEP", "8")), current=out)
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as fh:
            fh.write(f"matrix_path={out}\n")

    if args.no_mail:
        log.info("Envoi d'e-mail désactivé (--no-mail)")
        return 0
    send_mail(out, build_summary(info, latest, today), today)
    return 0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--from-json", help="Fichier JSON de soumissions (test hors ligne)")
    ap.add_argument("--base", help="Matrice de départ (par défaut : la plus récente de MATRIX_DIR)")
    ap.add_argument("--no-mail", action="store_true", help="Ne pas envoyer d'e-mail")
    ap.add_argument("--skip-if-unchanged", action="store_true", help="Ne rien enregistrer si aucune soumission n'a été ajoutée, modifiée ou supprimée")
    ap.add_argument("--env", default=".env", help="Fichier de configuration (défaut : .env)")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(), logging.FileHandler("sync.log", encoding="utf-8")])
    try:
        sys.exit(run(args))
    except SystemExit as e:
        if e.code not in (0, None):
            log.error("%s", e.code)
            step_summary(["### Sync Kobo : ÉCHEC", f"```\n{e.code}\n```"])
            if os.environ.get("GITHUB_ACTIONS"):
                print(f"::error title=Sync Kobo::{e.code}")
        raise
    except Exception as e:      # erreur inattendue : visible dans le journal et dans l'onglet du run, code de sortie 1
        log.exception("Erreur inattendue")
        step_summary(["### Sync Kobo : ÉCHEC", f"```\n{type(e).__name__}: {e}\n```"])
        if os.environ.get("GITHUB_ACTIONS"):
            print(f"::error title=Sync Kobo::{type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
