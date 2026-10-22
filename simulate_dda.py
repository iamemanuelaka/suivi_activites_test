#!/usr/bin/env python3
"""
Simulation des scénarios de test DDA (1-8, 10) sur le XForm *compilé* (pyxform) du questionnaire.

Ce que ce script exécute : les expressions XPath réelles du formulaire (calculate + setvalue) sur un moteur XPath 1.0
(lxml/libxml2) complété des fonctions ODK utilisées (max, int, decimal-date-time, if, current, instance).
Ce qu'il NE teste PAS : Enketo (Web), KoboCollect, le rafraîchissement réseau du DDA, le mode hors ligne, l'envoi réel.
Ces points doivent être validés sur Kobo (voir plan de tests).

Usage : python simulate_dda.py questionnaire.xlsx
"""
import copy, datetime as dt, math, sys, tempfile, os
from lxml import etree
from pyxform.xls2xform import xls2xform_convert

XF = "{http://www.w3.org/2002/xforms}"
CUR = None
PARENT = None  # <wrapper><root>...</root></wrapper>

# ---------- fonctions ODK ----------
def _txt(x):
    if isinstance(x, list):
        return (x[0].text or "") if x and hasattr(x[0], "text") else (str(x[0]) if x else "")
    return str(x)
def _num(x):
    if isinstance(x, bool): return 1.0 if x else 0.0
    if isinstance(x, (int, float)): return float(x)
    t = _txt(x).strip()
    try: return float(t)
    except ValueError: return float("nan")
def _truthy(x):
    if isinstance(x, list): return len(x) > 0
    if isinstance(x, bool): return x
    if isinstance(x, (int, float)): return x != 0 and not math.isnan(x)
    return str(x) != ""
def f_if(ctx, c, a, b): return a if _truthy(c) else b
def f_max(ctx, ns):
    v = [_num(n.text) for n in ns if (n.text or "").strip() != ""]
    v = [x for x in v if not math.isnan(x)]
    return max(v) if v else float("nan")
def f_int(ctx, x):
    n = _num(x); return float("nan") if math.isnan(n) else float(math.trunc(n))
def f_ddt(ctx, x):
    s = _txt(x).strip()
    if not s: return float("nan")
    try:
        d = dt.datetime.fromisoformat(s.replace("Z", "+00:00")) if "T" in s else dt.datetime.fromisoformat(s)
    except ValueError: return float("nan")
    if d.tzinfo is None: d = d.replace(tzinfo=dt.timezone.utc)
    return (d - dt.datetime(1970, 1, 1, tzinfo=dt.timezone.utc)).total_seconds() / 86400
def f_current(ctx): return [CUR]
def f_instance(ctx, name): return [PARENT]
ns = etree.FunctionNamespace(None)
import uuid as _uuid
def f_uuid(ctx): return str(_uuid.uuid4())
def f_today(ctx): return dt.date.today().isoformat()
def f_now(ctx): return dt.datetime.now(dt.timezone.utc).isoformat()
for n, f in (("uuid", f_uuid), ("today", f_today), ("now", f_now), ("if", f_if), ("max", f_max), ("int", f_int), ("decimal-date-time", f_ddt), ("current", f_current), ("instance", f_instance)):
    ns[n] = f

def to_str(r):
    if isinstance(r, list): return (r[0].text or "") if r and hasattr(r[0], "text") else ""
    if isinstance(r, bool): return "true" if r else "false"
    if isinstance(r, float):
        if math.isnan(r): return "NaN"
        return str(int(r)) if r == int(r) else repr(r)
    return str(r)

# ---------- chargement du XForm compilé ----------
class Form:
    def __init__(self, xlsx):
        out = os.path.join(tempfile.mkdtemp(), "f.xml")
        xls2xform_convert(xlsform_path=xlsx, xform_path=out, validate=False)
        t = etree.parse(out).getroot()
        inst = t.find(f".//{XF}instance/")
        self.skeleton = self._plain(inst)
        self.calcs = [(b.get("nodeset"), b.get("calculate")) for b in t.iter(f"{XF}bind") if b.get("calculate") and "/meta/" not in b.get("nodeset")]  # méta (instanceID/instanceName) hors périmètre
        self.sets = []  # (déclencheur, cible, expression)
        for sv in t.iter(f"{XF}setvalue"):
            if sv.get("event") == "xforms-value-changed":
                self.sets.append((sv.getparent().get("ref"), sv.get("ref"), sv.get("value")))
    def _plain(self, el):
        n = etree.Element(etree.QName(el).localname)
        n.text = el.text if len(el) == 0 else None
        for c in el: n.append(self._plain(c))
        return n
    def new(self): return copy.deepcopy(self.skeleton)

def node(doc, path):
    r = doc.xpath(path.replace(" ", "")); return r[0] if r else None
def ev(doc, expr, target):
    global CUR; CUR = target
    return doc.xpath(expr.strip())
def recalc(form, doc):
    for _ in range(6):
        changed = False
        for path, expr in form.calcs:
            t = node(doc, path); v = to_str(ev(doc, expr, t))
            if (t.text or "") != v: t.text = v; changed = True
        if not changed: break
def prefill(form, doc, trigger_path, stale=False):
    for trig, tgt, expr in form.sets:
        if trig.replace(" ", "") == trigger_path:
            t = node(doc, tgt); t.text = to_str(ev(doc, expr, t))

def set_(doc, name, val):
    for e in doc.iter(name): e.text = val; return
    raise KeyError(name)
def get(doc, name):
    for e in doc.iter(name): return e.text or ""
    raise KeyError(name)

def open_form(form, projet, date, start, nom, stale_setvalue=False):
    """Ouvre un formulaire vierge, choisit le projet (déclenche calculs puis préremplissage)."""
    d = form.new(); set_(d, "date_mission", date); set_(d, "start", start); set_(d, "nom_agent", nom)
    set_(d, "id_projet", projet)
    if stale_setvalue:                      # mauvais ordre : préremplissage AVANT recalcul (diagnostic)
        prefill(form, d, "/data/identification/id_projet"); recalc(form, d)
    else:                                   # ordre attendu : recalcul PUIS préremplissage
        recalc(form, d); prefill(form, d, "/data/identification/id_projet")
    return d
def change_project(form, d, projet, stale=False):
    set_(d, "id_projet", projet)
    if stale: prefill(form, d, "/data/identification/id_projet"); recalc(form, d)
    else: recalc(form, d); prefill(form, d, "/data/identification/id_projet")
_F = None
def submit(d, **edits):
    for k, v in edits.items(): set_(d, k, v)
    recalc(_F, d)
    row = copy.deepcopy(d); row.tag = "data"
    PARENT[0].append(row)
    return row
def reset_parent():
    global PARENT
    PARENT = etree.Element("wrapper"); etree.SubElement(PARENT, "root")

FIELDS = ["statut", "taux_execution", "obs_techniques", "pges_statut", "date_remise_site"]
def snap(d, fields=FIELDS): return {f: get(d, f) for f in fields}

# ---------- scénarios ----------
RESULTS = []
def check(test, desc, cond, detail=""):
    RESULTS.append((test, desc, "OK" if cond else "ÉCHEC", detail))
    print(f"[{'OK ' if cond else 'ÉCHEC'}] {test} – {desc} {detail}")

def run(xlsx):
    global _F
    form = Form(xlsx); _F = form
    S = lambda h, m=0: f"2026-10-0{1}T{h:02d}:{m:02d}:00.000+00:00"
    base = dict(statut="travaux_en_cours", taux_execution="40", obs_techniques="retard de livraison",
                pges_statut="valide", date_remise_site="2026-09-10")

    # T1 – projet jamais suivi
    reset_parent(); d = open_form(form, "03", "2026-10-05", S(8), "A")
    check("T1", "projet jamais suivi : aucun champ prérempli, aucune erreur", all(v == "" for v in snap(d).values()) and get(d, "nb_prec") == "0" and get(d, "numero_suivi") == "1" and get(d, "ordre_prec") == "", f"(ordre_suivi={get(d,'ordre_suivi')})")

    # T10 (partie) – lignes sans clé (anciennes versions / soumission technique INIT) ignorées
    reset_parent(); old = form.new(); set_(old, "id_projet", "26"); set_(old, "nom_agent", "INIT"); set_(old, "statut", "non_demarre"); set_(old, "taux_execution", "0")
    old.tag = "data"; PARENT[0].append(old)
    d = open_form(form, "26", "2026-10-06", S(9), "A")
    check("T10a", "ligne sans clé ordre_suivi (ex. soumission INIT / ancienne version) ignorée : lot 26 = jamais suivi", get(d, "nb_prec") == "0" and get(d, "statut") == "")

    # T2 – projet suivi une fois ; T3 – deux superviseurs
    reset_parent(); a = open_form(form, "01", "2026-10-05", S(8), "Superviseur A"); rowA = submit(a, **base)
    snapA = etree.tostring(rowA)
    b = open_form(form, "01", "2026-10-19", S(9), "Superviseur B")
    check("T2", "après un premier suivi, les réponses précédentes sont reprises", snap(b) == base and get(b, "numero_suivi") == "2", str(snap(b)))
    check("T3", "le nom de B n'est pas repris de A ; la note affiche A", get(b, "nom_agent") == "Superviseur B" and get(b, "prec_nom") == "Superviseur A")
    rowB = submit(b, taux_execution="65", obs_techniques="livraison partiellement effectuée")
    check("T6", "modification enregistrée dans la nouvelle ligne ; ligne A inchangée (simulation : ajout seul)", get(rowB, "taux_execution") == "65" and etree.tostring(PARENT[0][0]) == snapA)

    # T4 – plusieurs suivis, ordre d'arrivée différent de l'ordre des visites
    c = open_form(form, "01", "2026-10-26", S(10), "Superviseur C")
    check("T4a", "3e suivi : repris du suivi B (65 %), pas de A (40 %)", get(c, "taux_execution") == "65" and get(c, "prec_nom") == "Superviseur B")
    rowC = submit(c, taux_execution="75")
    late = open_form(form, "01", "2026-10-12", S(11), "Superviseur D"); submit(late, taux_execution="55")   # visite plus ancienne, envoyée plus tard
    n = open_form(form, "01", "2026-10-30", S(12), "Superviseur E")
    check("T4b", "visite ancienne envoyée tardivement n'écrase pas le dernier suivi (75 %)", get(n, "taux_execution") == "75" and get(n, "prec_nom") == "Superviseur C")
    sd1 = open_form(form, "01", "2026-10-26", S(15), "Superviseur F"); submit(sd1, taux_execution="80")   # même date que C, démarré plus tard
    n = open_form(form, "01", "2026-10-30", S(16), "Superviseur G")
    check("T4c", "même date de visite : départage par l'heure d'ouverture du formulaire (F 80 % > C 75 %)", get(n, "taux_execution") == "80" and get(n, "prec_nom") == "Superviseur F")
    reset_parent(); x = open_form(form, "02", "2026-10-05", S(8), "X"); submit(x, taux_execution="10")
    y = open_form(form, "02", "2026-10-05", S(8), "Y"); submit(y, taux_execution="20")      # clé strictement identique
    n = open_form(form, "02", "2026-10-30", S(9), "Z")
    check("T4d", "clés strictement identiques : résultat déterministe (1re soumission reçue)", get(n, "taux_execution") == "10")

    # T5 – plusieurs projets
    reset_parent()
    p1 = open_form(form, "01", "2026-10-05", S(8), "A"); submit(p1, statut="travaux_en_cours", taux_execution="40", obs_techniques="projet 01")
    p2 = open_form(form, "02", "2026-10-28", S(8), "B"); submit(p2, statut="suspendu", taux_execution="10", obs_techniques="projet 02")   # date plus récente
    o1 = open_form(form, "01", "2026-10-30", S(9), "C"); o2 = open_form(form, "02", "2026-10-30", S(9), "C"); o3 = open_form(form, "03", "2026-10-30", S(9), "C")
    check("T5", "aucune réponse d'un projet n'apparaît dans un autre (même si l'autre est plus récent)",
          get(o1, "obs_techniques") == "projet 01" and get(o2, "obs_techniques") == "projet 02" and get(o3, "obs_techniques") == "")

    # T7 – champ vide dans la dernière soumission
    reset_parent(); a = open_form(form, "01", "2026-10-05", S(8), "A"); submit(a, **base)
    b = open_form(form, "01", "2026-10-19", S(9), "B"); submit(b, obs_techniques="")
    c = open_form(form, "01", "2026-10-26", S(9), "C")
    check("T7", "champ vidé dans la dernière soumission : reste vide (pas de repli silencieux sur l'ancienne valeur)", get(c, "obs_techniques") == "" and get(c, "taux_execution") == "40")

    # T8 – changement de projet avant envoi
    reset_parent()
    p1 = open_form(form, "01", "2026-10-05", S(8), "A"); submit(p1, statut="travaux_en_cours", taux_execution="40", obs_techniques="projet 01")
    p2 = open_form(form, "02", "2026-10-05", S(8), "B"); submit(p2, statut="suspendu", taux_execution="10", obs_techniques="projet 02")
    d = open_form(form, "01", "2026-10-30", S(9), "C"); ok1 = get(d, "obs_techniques") == "projet 01"
    change_project(form, d, "02"); ok2 = get(d, "obs_techniques") == "projet 02" and get(d, "taux_execution") == "10" and get(d, "prec_nom") == "B"
    change_project(form, d, "03"); ok3 = get(d, "obs_techniques") == "" and get(d, "taux_execution") == "" and get(d, "prec_date") == ""
    check("T8", "01 -> 02 -> 03 : préremplissage suit le projet final, vidé pour un projet sans historique", ok1 and ok2 and ok3, "(ordre supposé : recalcul puis préremplissage)")
    d2 = open_form(form, "01", "2026-10-30", S(9), "C"); change_project(form, d2, "02", stale=True)
    check("T8-diag", "diagnostic : si le moteur préremplissait AVANT le recalcul, les valeurs du projet précédent resteraient",
          get(d2, "obs_techniques") == "projet 01", "(symptôme à rechercher lors du test réel)")
    return RESULTS

if __name__ == "__main__":
    res = run(sys.argv[1] if len(sys.argv) > 1 else "../docs/questionnaire_suivi_activites_v4_DDA.xlsx")
    bad = [r for r in res if r[2] != "OK"]
    print(f"\n{len(res)-len(bad)}/{len(res)} contrôles OK"); sys.exit(1 if bad else 0)
