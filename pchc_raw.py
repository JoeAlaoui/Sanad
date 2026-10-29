# -*- coding: utf-8 -*-
"""
pchc_raw.py — Reporting Métier PCHC : import des 5 feuilles brutes Tarkhiss et calcul unique des KPI.

- parse_workbook()  : lit le classeur consolidé (.xlsx/.xlsm, openpyxl read_only), en-têtes ligne 5,
                      données à partir de la ligne 6, et produit un rapport de contrôle qualité.
                      Fonction PURE : aucune écriture disque (le stockage est fait par app.py).
- pchc_compute()    : SEULE fonction de calcul des métriques. Elle reproduit à l'identique la
                      sémantique des formules du classeur Excel (COUNTA / COUNTIF insensible à la
                      casse / SUMIF par catégorie / IFERROR) ; Excel, PPTX, PDF, e-mail et UI
                      l'utilisent tous.

Les valeurs brutes sont conservées telles quelles (coquilles comprises, comme dans la référence) ;
les variantes connues (« Editon acte », « Sénace de travail ») sont regroupées au calcul, exactement
comme les formules Excel (COUNTIF(statut) + COUNTIF(coquille)).
"""
import hashlib
import io
import re
import unicodedata
import zipfile
from datetime import date, datetime

HEADER_ROW = 5          # ligne des en-têtes dans les feuilles brutes
FIRST_ROW = 6           # première ligne de données
LAST_ROW = 50000        # plage de lecture des formules Excel

# Catégories de statut (ordre d'affichage des légendes)
CATEGORIES = ["Délivré", "En cours", "Non conforme", "Refusé", "Autre"]

# Les 5 volets : clé interne, feuille brute canonique, feuille KPI, libellés, colonnes par défaut,
# thème (couleur titre de la feuille KPI). L'ordre est celui du classeur de référence.
MODULES = [
    {"key": "identification", "sheet": "identificationOpérateur", "kpi": "KPI Identification",
     "label": "Identification Opérateur", "short": "Identification", "dash": "1. Identification Opérateur",
     "cols": {"id": "B", "status": "F", "type": "D"}, "theme": "2E75B6",
     "match": ["identification", "operateur"]},
    {"key": "declaration", "sheet": "déclarationActivité", "kpi": "KPI Déclaration",
     "label": "Déclaration d'Activité", "short": "Déclaration", "dash": "2. Déclaration d'Activité",
     "cols": {"id": "B", "status": "H"}, "theme": "538135",
     "match": ["declaration"]},
    {"key": "ce", "sheet": "certificatEnregistrement", "kpi": "KPI Certificat Enreg.",
     "label": "Certificat d'Enregistrement (CE)", "short": "CE", "dash": "3. Certificat Enregistrement (CE)",
     "cols": {"id": "B", "status": "J"}, "theme": "C65911",
     "match": ["enregistrement"]},
    {"key": "clv", "sheet": "certificatLibreVente", "kpi": "KPI Certificat LV",
     "label": "Certificat Libre Vente (CLV)", "short": "CLV", "dash": "4. Certificat Libre Vente (CLV)",
     "cols": {"id": "B", "status": "H"}, "theme": "7030A0",
     "match": ["librevente", "clv"]},
    {"key": "aimp", "sheet": "autorisationImportationMatièreP", "kpi": "KPI Autor. Import.",
     "label": "Autorisation d'Importation MP (AIMP)", "short": "AIMP", "dash": "5. Autorisation Import. MP (AIMP)",
     "cols": {"id": "B", "status": "F"}, "theme": "1F4E79",
     "match": ["importation", "matiere", "aimp"]},
]
MOD = {m["key"]: m for m in MODULES}
DASHBOARD_THEME = "1F4E79"

# Statuts attendus par volet, dans l'ordre de traitement : (statut, catégorie, [coquilles regroupées])
STATUSES = {
    "declaration": [("En Vérification", "En cours"), ("En Evaluation", "En cours"), ("Evaluation Compléments", "En cours"),
                    ("Attente validation évaluation", "En cours"), ("Séance de travail", "En cours"), ("Validation acte", "En cours"),
                    ("Attente Signature", "En cours"), ("Non Conforme", "Non conforme"), ("Irrecevable", "Refusé"),
                    ("Rejeté", "Refusé"), ("Délivré", "Délivré")],
    "ce": [("En vérification", "En cours"), ("En Evaluation", "En cours"), ("Evaluation Compléments", "En cours"),
           ("Attente validation évaluation", "En cours"), ("Edition acte", "En cours", ["Editon acte"]), ("Attente paiement", "En cours"),
           ("Séance de travail", "En cours"), ("Validation acte", "En cours"), ("Attente Signature", "En cours"),
           ("Non conforme", "Non conforme"), ("Irrecevable", "Refusé"), ("Rejeté", "Refusé"), ("Délivré", "Délivré")],
    "clv": [("En vérification", "En cours"), ("En Evaluation", "En cours"), ("Attente validation évaluation", "En cours"),
            ("Edition acte", "En cours"), ("Validation acte", "En cours"), ("Attente signature", "En cours"),
            ("Non conforme", "Non conforme"), ("Irrecevable", "Refusé"), ("Rejeté", "Refusé"), ("Délivré", "Délivré")],
    "aimp": [("En vérification", "En cours"), ("En Evaluation", "En cours"), ("Evaluation Compléments", "En cours"),
             ("Attente validation évaluation", "En cours"), ("Séance de travail", "En cours", ["Sénace de travail"]),
             ("Edition acte", "En cours"), ("Validation acte", "En cours"), ("Signature acte", "En cours"),
             ("Non conforme", "Non conforme"), ("Irrecevable", "Refusé"), ("Rejeté", "Refusé"), ("Délivré", "Délivré")],
}
IDENT_STATUS = "Identifié"
# Types d'opérateur : (libellé affiché, préfixe COUNTIF « xxx* »)
OPERATOR_TYPES = [("Etablissement fabrication / import / export", "Etablissement de fabrication"),
                  ("Etablissement pharmaceutique industriel", "Etablissement pharmaceutique"),
                  ("Coopérative", "Coopérative"),
                  ("Etablissement compléments alimentaires", "Etablissement exerçant les activités portant sur les compléments alimentaires")]

# Synthèse exécutive : seuils voyant (🟢 si taux ≥ vert, 🟡 si ≥ orange) — valeurs du classeur de référence
VOYANT_THRESHOLDS = {"identification": (0.9, 0.5), "declaration": (0.9, 0.5), "ce": (0.1, 0.05),
                     "clv": (0.5, 0.3), "aimp": (0.2, 0.1)}
# Points d'alerte : (indicateur, module affiché, clé volet, type, statut | None, seuil attention, seuil critique)
ALERTS = [
    ("Taux de délivrance - Déclaration", "Déclaration", "declaration", "taux", None, 0.5, 0.3),
    ("Taux de délivrance - CE", "CE", "ce", "taux", None, 0.2, 0.1),
    ("Taux de délivrance - CLV", "CLV", "clv", "taux", None, 0.5, 0.3),
    ("Taux de délivrance - AIMP", "AIMP", "aimp", "taux", None, 0.2, 0.1),
    ("Déclarations Irrecevables", "Déclaration", "declaration", "nb", "Irrecevable", 5, 10),
    ("Déclarations Non Conformes", "Déclaration", "declaration", "nb", "Non Conforme", 50, 100),
    ("Dossiers Irrecevables CE", "CE", "ce", "nb", "Irrecevable", 30, 50),
    ("Dossiers en Attente Paiement CE", "CE", "ce", "nb", "Attente paiement", 50, 100),
    ("Dossiers Non Conformes AIMP", "AIMP", "aimp", "nb", "Non conforme", 10, 20),
]

# Signatures d'en-têtes (normalisées) pour détecter une feuille dont les colonnes ne correspondent pas
SIGNATURES = {
    "identification": {"id", "raisonsociale", "typeoperateur", "datemodification", "statut"},
    "declaration": {"refdossier", "etablissement", "typedemande", "typeactivites", "responsableevaluation", "datedepot", "statut"},
    "ce": {"refdossier", "nomproduit", "etablissement", "typedemande", "statutproduit", "natureproduit", "datedepot",
           "nce", "statut", "responsableevaluation", "nombrenonconforme"},
    "clv": {"refdossier", "etablissement", "nomproduit", "nce", "datedepot", "evaluateur", "statut"},
    "aimp": {"refdossier", "importateur", "datedepot", "evaluateur", "statut"},
}


# ----------------------------------------------------------------------------- utilitaires
def norm(s):
    """Normalisation tolérante : sans accents, minuscules, uniquement lettres/chiffres."""
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def col_letter(idx):
    """0 -> A, 1 -> B…"""
    out, n = "", idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def col_index(letter):
    n = 0
    for ch in letter:
        n = n * 26 + ord(ch) - 64
    return n - 1


def ci(s):
    """Égalité « COUNTIF » : insensible à la casse, sans suppression d'espaces (comme Excel)."""
    return s.lower() if isinstance(s, str) else None


def blank(v):
    return v is None or v == ""


def to_json_value(v):
    if isinstance(v, datetime):
        return {"__dt": v.isoformat()}
    if isinstance(v, date):
        return {"__dt": datetime(v.year, v.month, v.day).isoformat()}
    if v == "":
        return None
    if isinstance(v, (int, float, str)) or v is None:
        return v
    return str(v)


def from_json_value(v):
    if isinstance(v, dict) and "__dt" in v:
        return datetime.fromisoformat(v["__dt"])
    return v


def match_sheet(name, exact_only=False):
    """Associe un nom de feuille à un volet (insensible casse/accents, tolérant). Les feuilles de
    restitution (« KPI … », « Dashboard ») ne sont jamais prises pour des feuilles brutes."""
    n = norm(name)
    for m in MODULES:
        if n == norm(m["sheet"]):
            return m["key"]
    if exact_only or n.startswith("kpi") or n == "dashboard":
        return None
    for m in MODULES:
        if any(tok in n for tok in m["match"]):
            return m["key"]
    return None


def assign_sheets(titles):
    """{titre de feuille: clé volet} : noms exacts d'abord, puis correspondance tolérante."""
    out, used = {}, set()
    for t in titles:
        k = match_sheet(t, exact_only=True)
        if k and k not in used:
            out[t], used = k, used | {k}
    for t in titles:
        if t not in out:
            k = match_sheet(t)
            if k and k not in used:
                out[t], used = k, used | {k}
    return out


def status_aliases(key):
    """{statut canonique: [coquilles]} pour un volet."""
    return {s[0]: (s[2] if len(s) > 2 else []) for s in STATUSES.get(key, [])}


# ----------------------------------------------------------------------------- lecture
def _sheet_widths(data):
    """Largeurs de colonnes de chaque feuille (le mode read_only d'openpyxl ne les expose pas)."""
    out = {}
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        wbx = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        targets.update({k: v for v, k in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
        for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
            t = targets.get(rid, "")
            path = t.lstrip("/") if t.startswith("/") else "xl/" + t
            try:
                xml = z.read(path).decode("utf-8")
            except KeyError:
                continue
            widths = {}
            for tag in re.findall(r"<col [^>]*>", xml):
                a = dict(re.findall(r'(\w+)="([^"]*)"', tag))
                if "width" not in a or "min" not in a:
                    continue
                for i in range(int(a["min"]), min(int(a.get("max", a["min"])), 60) + 1):
                    widths[col_letter(i - 1)] = float(a["width"])
            out[_xml_unescape(name)] = widths
    except Exception:
        pass
    return out


def _xml_unescape(s):
    return s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'")


def _detect_cols(key, headers, report_sheet):
    """Colonnes par texte d'en-tête (« Statut », « Réf. »/« ID », « Type opérateur ») avec repli
    sur les lettres par défaut."""
    nh = [norm(h) for h in headers]
    defaults = MOD[key]["cols"]
    cols = {}

    def pick(field, finder):
        idx = next((i for i, h in enumerate(nh) if finder(h)), None)
        if idx is None:
            idx = col_index(defaults[field])
            report_sheet["notes"].append(f"Colonne « {field} » non trouvée par l'en-tête : repli sur la colonne {defaults[field]}.")
        cols[field] = col_letter(idx)

    pick("id", lambda h: h == "id" or h.startswith("ref"))
    pick("status", lambda h: h == "statut")
    if "type" in defaults:
        pick("type", lambda h: h == "typeoperateur")
    for field, letter in cols.items():
        if letter != defaults[field]:
            report_sheet["notes"].append(f"Colonne « {field} » détectée en {letter} (attendue en {defaults[field]}).")
    return cols


def parse_workbook(data, filename=""):
    """Analyse un classeur consolidé. `data` = contenu binaire du fichier.
    Retourne (parsed, report) : parsed = {clé volet: {...}} ; report = rapport d'import."""
    from openpyxl import load_workbook

    report = {"file": filename, "sha256": hashlib.sha256(data).hexdigest(), "size": len(data),
              "sheets": {}, "missing": [], "empty": [], "ignored": [], "blocking": [], "warnings": [],
              "unknown_statuses": {}, "typos": {}}
    parsed = {}
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as e:
        report["blocking"].append(f"Fichier illisible (.xlsx/.xlsm attendu) : {e}")
        return parsed, report
    widths = _sheet_widths(data)
    assignment = assign_sheets([ws.title for ws in wb.worksheets])

    for ws in wb.worksheets:
        key = assignment.get(ws.title)
        if not key:
            report["ignored"].append(ws.title)
            continue
        if key in parsed:
            report["warnings"].append(f"Feuille « {ws.title} » ignorée : le volet {MOD[key]['short']} est déjà fourni par « {parsed[key]['source_sheet']} ».")
            continue
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        rs = {"source_sheet": ws.title, "rows": 0, "header_row": HEADER_ROW, "notes": []}
        # En-têtes en ligne 5 ; tolérance : recherche d'une ligne contenant « Statut » dans les 10 premières.
        hdr_idx = HEADER_ROW - 1
        if not (len(rows) > hdr_idx and any(norm(v) == "statut" for v in rows[hdr_idx])):
            alt = next((i for i, r in enumerate(rows[:10]) if any(norm(v) == "statut" for v in r)), None)
            if alt is not None:
                hdr_idx = alt
                rs["header_row"] = alt + 1
                rs["notes"].append(f"En-têtes trouvés en ligne {alt + 1} (attendus en ligne {HEADER_ROW}).")
        headers = [("" if v is None else str(v)) for v in (rows[hdr_idx] if len(rows) > hdr_idx else [])]
        body = rows[hdr_idx + 1:]
        while body and all(blank(v) for v in body[-1]):
            body.pop()
        ncols = max([len(headers)] + [len(r) for r in body]) if (headers or body) else 0
        headers += [""] * (ncols - len(headers))
        # Préfixe (colonnes A… et lignes 1-4) conservé à l'identique pour la feuille brute de l'Excel
        rows_json = [[to_json_value(v) for v in r] + [None] * (ncols - len(r)) for r in body]
        cols = _detect_cols(key, headers, rs)
        rs["rows"] = sum(1 for r in rows_json if not blank(r[col_index(cols["id"])] if col_index(cols["id"]) < ncols else None))
        rs["cols"] = cols
        parsed[key] = {"key": key, "sheet": MOD[key]["sheet"], "source_sheet": ws.title, "headers": headers,
                       "header_prefix": [[to_json_value(v) for v in r] for r in rows[:hdr_idx]] if hdr_idx == HEADER_ROW - 1 else [],
                       "rows": rows_json, "cols": cols,
                       "widths": widths.get(ws.title, {})}
        report["sheets"][key] = rs
    wb.close()

    for m in MODULES:
        k = m["key"]
        if k not in parsed:
            report["missing"].append(m["sheet"])
            continue
        if report["sheets"][k]["rows"] == 0:
            report["empty"].append(m["sheet"])
        _check_signature(k, parsed[k], report)
        _check_statuses(k, parsed[k], report)

    if not parsed:
        report["blocking"].append("Aucune des 5 feuilles attendues n'a été reconnue (identificationOpérateur, déclarationActivité, "
                                  "certificatEnregistrement, certificatLibreVente, autorisationImportationMatièreP).")
    for s in report["missing"]:
        report["warnings"].append(f"Feuille « {s} » absente : le volet sera à 0 / N/D.")
    for s in report["empty"]:
        report["warnings"].append(f"Feuille « {s} » vide : le volet sera à 0 / N/D.")
    return parsed, report


def _check_signature(key, sheet, report):
    hs = {norm(h) for h in sheet["headers"] if norm(h)}
    if not hs:
        return

    def score(k):
        sig = SIGNATURES[k]
        return len(hs & sig) / len(hs | sig)
    own = score(key)
    best = max(SIGNATURES, key=score)
    if best != key and score(best) > own:
        found = ", ".join(h for h in sheet["headers"] if norm(h) in SIGNATURES[best] - SIGNATURES[key])[:120]
        report["blocking"].append(
            f"Feuille « {sheet['source_sheet']} » : les colonnes ({found}) correspondent aux données « {MOD[best]['label']} », "
            f"pas « {MOD[key]['label']} ». Vérifiez l'export avant d'importer.")
    elif own < 0.5:
        report["warnings"].append(f"Feuille « {sheet['source_sheet']} » : en-têtes inhabituels ({', '.join(h for h in sheet['headers'] if h.strip())[:160]}).")


def _check_statuses(key, sheet, report):
    """Statuts inconnus (→ « Autres ») et coquilles regroupées, comparaison insensible à la casse."""
    si = col_index(sheet["cols"]["status"])
    ii = col_index(sheet["cols"]["id"])
    if key == "identification":
        known = {ci(IDENT_STATUS): IDENT_STATUS}
        alias = {}
    else:
        known = {ci(s[0]): s[0] for s in STATUSES[key]}
        alias = {ci(a): canon for canon, al in status_aliases(key).items() for a in al}
    unknown, typos = {}, {}
    for r in sheet["rows"]:
        if blank(r[ii] if ii < len(r) else None):
            continue
        v = r[si] if si < len(r) else None
        lv = ci(v) if isinstance(v, str) else None
        if lv in known:
            continue
        if lv in alias:
            label = f"« {v} » → {alias[lv]}"
            typos[label] = typos.get(label, 0) + 1
            continue
        label = "(vide)" if blank(v) else str(v)
        unknown[label] = unknown.get(label, 0) + 1
    if unknown:
        report["unknown_statuses"][key] = unknown
        report["warnings"].append(f"{MOD[key]['short']} : {sum(unknown.values())} dossier(s) avec un statut non référencé "
                                  f"({', '.join(sorted(unknown))[:150]}) — comptés en « Autres statuts ».")
    if typos:
        report["typos"][key] = typos


# ----------------------------------------------------------------------------- calcul unique
def _column(sheet, letter):
    i = col_index(letter)
    return [(r[i] if i < len(r) else None) for r in sheet["rows"]]


def _countif(values, crit):
    c = ci(crit)
    return sum(1 for v in values if isinstance(v, str) and v.lower() == c)


def _countif_prefix(values, prefix):
    p = prefix.lower()
    return sum(1 for v in values if isinstance(v, str) and v.lower().startswith(p))


def _div(a, b):
    return a / b if b else 0


def voyant(total, taux, green, orange):
    if total == 0:
        return "⚪ N/D"
    return "🟢" if taux >= green else ("🟡" if taux >= orange else "🔴")


def alert_status(total, value, typ, warn, crit):
    if total == 0:
        return "⚪ N/D", "—"
    if (typ == "taux" and value < crit) or (typ == "nb" and value > crit):
        return "🔴 CRITIQUE", "HAUTE"
    if (typ == "taux" and value < warn) or (typ == "nb" and value > warn):
        return "🟡 ATTENTION", "MOYENNE"
    return "🟢 OK", "BASSE"


def pchc_compute(raw):
    """Calcule TOUTES les métriques du Reporting Métier à partir des feuilles brutes stockées.
    `raw` = {clé volet: {"rows": [...], "cols": {...}}} (volet absent = 0 dossier).
    Les résultats sont identiques aux formules du classeur Excel recalculé."""
    out = {"modules": {}, "synthesis": [], "totals": {}, "cards": [], "alerts": []}

    for m in MODULES:
        k = m["key"]
        sheet = raw.get(k) or {"rows": [], "cols": dict(m["cols"])}
        cols = sheet.get("cols") or m["cols"]
        ids = _column(sheet, cols["id"])
        sts = _column(sheet, cols["status"])
        total = sum(1 for v in ids if not blank(v))
        mod = {"key": k, "label": m["label"], "short": m["short"], "kpi": m["kpi"], "theme": m["theme"],
               "sheet": m["sheet"], "cols": cols, "total": total}
        if k == "identification":
            ident = _countif(sts, IDENT_STATUS)
            others = total - ident
            types_v = _column(sheet, cols["type"])
            types = [{"label": lab, "prefix": pre, "count": _countif_prefix(types_v, pre)} for lab, pre in OPERATOR_TYPES]
            types.append({"label": "Autres types / vide", "prefix": None, "count": total - sum(t["count"] for t in types)})
            ttot = sum(t["count"] for t in types)
            for t in types:
                t["pct"] = _div(t["count"], ttot)
            rows = [{"status": IDENT_STATUS, "category": "Délivré", "category_label": "Validé", "count": ident},
                    {"status": "Autres statuts / statut vide", "category": "Autre", "category_label": "Autre", "count": others}]
            stot = ident + others
            for r in rows:
                r["pct"] = _div(r["count"], stot)
            mod.update(delivered=ident, encours=others, nonconf=0, refused=0, bloques=None,
                       taux=_div(ident, total), taux_refus=0, statuses=rows, types=types, types_total=ttot,
                       status_total=stot, autre=others)
        else:
            rows = []
            for st in STATUSES[k]:
                name, cat = st[0], st[1]
                cnt = _countif(sts, name) + sum(_countif(sts, a) for a in (st[2] if len(st) > 2 else []))
                rows.append({"status": name, "category": cat, "category_label": cat, "count": cnt,
                             "aliases": list(st[2]) if len(st) > 2 else []})
            autre = total - sum(r["count"] for r in rows)
            rows.append({"status": "Autres statuts / statut vide", "category": "Autre", "category_label": "Autre", "count": autre})
            stot = sum(r["count"] for r in rows)
            for r in rows:
                r["pct"] = _div(r["count"], stot)

            def by(cat):
                return sum(r["count"] for r in rows if r["category"] == cat)
            delivered, nonconf, refused = by("Délivré"), by("Non conforme"), by("Refusé")
            encours = by("En cours") + by("Autre")
            mod.update(delivered=delivered, encours=encours, nonconf=nonconf, refused=refused, bloques=nonconf + refused,
                       taux=_div(delivered, total), taux_refus=_div(refused, total), statuses=rows,
                       status_total=stot, autre=autre)
        mod["control_ok"] = mod["status_total"] == total
        g, o = VOYANT_THRESHOLDS[k]
        # Taux de la synthèse (=IF(B>0,C/B,0)) — même valeur que taux de délivrance
        rate = mod["delivered"] / total if total > 0 else 0
        mod["voyant"] = voyant(total, rate, g, o)
        out["modules"][k] = mod
        out["synthesis"].append({"key": k, "module": m["dash"], "total": total, "delivered": mod["delivered"],
                                 "taux": rate, "encours": mod["encours"], "bloques": mod["bloques"],
                                 "seuil_vert": g, "seuil_orange": o, "voyant": mod["voyant"]})

    syn = out["synthesis"]
    tot = {"total": sum(r["total"] for r in syn), "delivered": sum(r["delivered"] for r in syn),
           "encours": sum(r["encours"] for r in syn), "bloques": sum(r["bloques"] or 0 for r in syn)}
    tot["taux"] = tot["delivered"] / tot["total"] if tot["total"] > 0 else 0
    out["totals"] = tot
    M = out["modules"]
    out["cards"] = [
        {"label": "OPÉRATEURS IDENTIFIÉS", "value": M["identification"]["delivered"], "kind": "ok"},
        {"label": "CE DÉLIVRÉS", "value": M["ce"]["delivered"], "kind": "ok"},
        {"label": "CLV DÉLIVRÉS", "value": M["clv"]["delivered"], "kind": "ok"},
        {"label": "AIMP DÉLIVRÉS", "value": M["aimp"]["delivered"], "kind": "ok"},
        {"label": "DOSSIERS EN COURS", "value": tot["encours"], "kind": "info"},
        {"label": "BLOQUÉS / REFUSÉS", "value": tot["bloques"], "kind": "bad"},
    ]
    for lab, modlab, k, typ, status, warn, crit in ALERTS:
        mm = M[k]
        if typ == "taux":
            value = mm["delivered"] / mm["total"] if mm["total"] > 0 else 0
        else:
            value = next(r["count"] for r in mm["statuses"] if r["status"] == status)
        st, prio = alert_status(mm["total"], value, typ, warn, crit)
        out["alerts"].append({"indicator": lab, "module": modlab, "key": k, "type": typ, "status_name": status,
                              "value": value, "warn": warn, "crit": crit, "status": st, "priority": prio})
    out["empty_modules"] = [M[k]["sheet"] for k in M if M[k]["total"] == 0]
    out["insights"] = {k: insights(M[k]) for k in M}
    return out


# ----------------------------------------------------------------------------- constats automatiques
def fmt_int(n):
    return f"{n:,}".replace(",", " ")


def fmt_pct(x, d=1):
    return f"{x * 100:.{d}f}".replace(".", ",") + " %"


VOYANT_WORD = {"🟢": "vert", "🟡": "orange", "🔴": "rouge", "⚪ N/D": "N/D"}


def insights(mod):
    """2 à 3 constats chiffrés générés automatiquement pour une section KPI (sans emoji : lisibles
    dans tous les livrables, y compris le PDF serveur)."""
    if mod["total"] == 0:
        return [f"Aucune donnée pour « {mod['sheet']} » : indicateurs à 0 / N/D."]
    out = []
    word = VOYANT_WORD.get(mod["voyant"], "N/D")
    if mod["key"] == "identification":
        out.append(f"{fmt_int(mod['delivered'])} opérateurs identifiés sur {fmt_int(mod['total'])}, "
                   f"soit un taux de réussite de {fmt_pct(mod['taux'])} (voyant {word}).")
        top = max(mod["types"][:-1], key=lambda t: t["count"])
        out.append(f"Type majoritaire : {top['label']} avec {fmt_int(top['count'])} opérateurs ({fmt_pct(top['pct'])}).")
        if mod["encours"]:
            out.append(f"{fmt_int(mod['encours'])} opérateur(s) en cours ou avec un autre statut.")
        return out
    main = max(mod["statuses"], key=lambda r: r["count"])
    out.append(f"Statut principal : « {main['status']} » avec {fmt_int(main['count'])} dossiers ({fmt_pct(main['pct'])} du total).")
    out.append(f"Taux de délivrance : {fmt_pct(mod['taux'])} ({fmt_int(mod['delivered'])} délivrés sur {fmt_int(mod['total'])}) — voyant {word}.")
    out.append(f"{fmt_int(mod['nonconf'])} non conforme(s) et {fmt_int(mod['refused'])} refusé(s) (irrecevables + rejetés), "
               f"soit {fmt_pct(_div(mod['bloques'], mod['total']))} des dossiers bloqués.")
    return out
