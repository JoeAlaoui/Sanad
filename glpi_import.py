# -*- coding: utf-8 -*-
"""
glpi_import.py — Parsing des exports GLPI (Moussanada).

Fonctions pures : ne font aucune I/O disque ni appel Flask. Elles reçoivent des
dictionnaires en mémoire (timeseries, sources) et les retournent modifiés, à charge de
l'appelant (app.py) de les charger/sauvegarder. Ce découpage permet de tester et de faire
évoluer le parsing GLPI indépendamment du reste de l'application.
"""
import csv as _csv
import io as _io
import re
import datetime as _dt


def glpi_cell_to_seconds(value):
    """Convertit une cellule de durée GLPI (datetime.time ou datetime.timedelta) en secondes."""
    if value is None:
        return 0
    if isinstance(value, _dt.timedelta):
        return value.total_seconds()
    if isinstance(value, _dt.time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (int, float)):
        return float(value) * 3600  # heures décimales
    if isinstance(value, str):
        v = value.strip().replace(",", ".")
        try:
            return float(v) * 3600
        except ValueError:
            return 0
    return 0


def decimal_hours_cell(value):
    """Convertit une cellule de la feuille 'Durée moyenne - Heure' (nombre décimal d'heures) en heures (float)."""
    if value is None:
        return 0.0
    if isinstance(value, str):
        v = value.strip().replace(",", ".")
        try:
            return float(v)
        except ValueError:
            return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def format_dh(hours):
    """Formate un nombre d'heures décimal à la française : 6,03 h"""
    hours = round(hours or 0, 2)
    return f"{hours:.2f}".replace(".", ",") + " h"


def normalize_header(s):
    """Normalise un en-tête de colonne (tolère les problèmes d'encodage/accents des exports GLPI)."""
    if not s:
        return ""
    s = str(s).lower()
    repl = {"é": "e", "è": "e", "ê": "e", "à": "a", "â": "a", "ô": "o", "î": "i",
            "ç": "c", "ù": "u", "û": "u", "ë": "e", "ï": "i", "ã©": "e", "ã´": "o", "ã ": "a"}
    for a, b in repl.items():
        s = s.replace(a, b)
    return "".join(ch for ch in s if ch.isalnum())


def match_header(headers, *keywords):
    """Trouve l'index de la première colonne dont l'en-tête normalisé contient un des mots-clés."""
    norm = [normalize_header(h) for h in headers]
    for kw in keywords:
        kwn = normalize_header(kw)
        for i, h in enumerate(norm):
            if kwn in h:
                return i
    return None


def parse_glpi_row_sheet(ws, headers, label_col_idx=0):
    """Parse une feuille GLPI type 'rawData - Ticket - X' (Catégorie/Service/Technicien/Demandeur)."""
    idx_ouverts = match_header(headers, "ouverts")
    idx_resolus = match_header(headers, "resolus")
    idx_retard = match_header(headers, "retard")
    idx_fermes = match_header(headers, "fermes", "ferme")
    idx_delai_prise = match_header(headers, "delai moyen de prise")
    idx_delai_resol = match_header(headers, "delai moyen de resolution")
    idx_delai_ferm = match_header(headers, "delai moyen de fermeture")

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or row[label_col_idx] is None:
            continue
        label = str(row[label_col_idx]).strip()
        if not label:
            continue

        def g(idx):
            return row[idx] if idx is not None and idx < len(row) else None

        item = {
            "label": label,
            "ouverts": int(g(idx_ouverts) or 0),
            "resolus": int(g(idx_resolus) or 0),
            "en_retard": int(g(idx_retard) or 0),
            "fermes": int(g(idx_fermes) or 0),
            "delai_prise_en_compte_s": glpi_cell_to_seconds(g(idx_delai_prise)),
            "delai_resolution_s": glpi_cell_to_seconds(g(idx_delai_resol)),
            "delai_fermeture_s": glpi_cell_to_seconds(g(idx_delai_ferm)),
        }
        rows.append(item)
    return rows


def import_moussanada_xlsx(file_storage, ts, src):
    """Parse un export GLPI consolidé (.xlsx). Mute et retourne (result, ts, src)."""
    from openpyxl import load_workbook
    wb = load_workbook(file_storage, data_only=True)
    result = {"tickets_months": 0, "durations_months": 0, "categories": 0, "services": 0,
              "techniciens": 0, "demandeurs": 0, "detected_month": None}
    detected_month = None

    for sheet_name in wb.sheetnames:
        parts = [p.strip() for p in sheet_name.split(" - ")]
        norm = normalize_header(sheet_name)
        p1 = normalize_header(parts[1]) if len(parts) > 1 else ""
        p2 = normalize_header(parts[2]) if len(parts) > 2 else ""
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [str(h) if h is not None else "" for h in rows[0]]

        is_tickets_series = len(parts) == 2 and p1 == "tickets"
        is_duration = "duree" in norm and "heure" in norm
        is_technicien = len(parts) == 2 and "technicien" in p1
        is_categorie = len(parts) >= 3 and p1 == "ticket" and "categorie" in p2
        is_service = len(parts) >= 3 and p1 == "ticket" and "service" in p2
        is_demandeur = len(parts) >= 3 and p1 == "ticket" and "demandeur" in p2

        if is_tickets_series:
            idx_mois = 0
            idx_ouverts = match_header(headers, "ouverts")
            idx_resolus = match_header(headers, "resolus")
            idx_retard = match_header(headers, "retard")
            idx_clos = match_header(headers, "clos")
            for r in rows[1:]:
                if not r or not r[idx_mois]:
                    continue
                m = str(r[idx_mois]).strip()
                if not m or m == "0":
                    continue
                ts.setdefault(m, {})
                ts[m]["ouverts"] = int(r[idx_ouverts] or 0) if idx_ouverts is not None else 0
                ts[m]["resolus"] = int(r[idx_resolus] or 0) if idx_resolus is not None else 0
                ts[m]["en_retard"] = int(r[idx_retard] or 0) if idx_retard is not None else 0
                ts[m]["clos"] = int(r[idx_clos] or 0) if idx_clos is not None else 0
                result["tickets_months"] += 1
                if re.match(r"^\d{4}-\d{2}$", m) and (ts[m]["ouverts"] or ts[m]["resolus"]):
                    if detected_month is None or m > detected_month:
                        detected_month = m

        elif is_duration:
            idx_mois = 0
            idx_clot = match_header(headers, "cloture")
            idx_resol = match_header(headers, "resolution")
            idx_reelle = match_header(headers, "duree reelle")
            for r in rows[1:]:
                if not r or not r[idx_mois]:
                    continue
                m = str(r[idx_mois]).strip()
                if not m or m == "0":
                    continue
                ts.setdefault(m, {})
                ts[m]["cloture_h"] = decimal_hours_cell(r[idx_clot]) if idx_clot is not None else 0
                ts[m]["resolution_h"] = decimal_hours_cell(r[idx_resol]) if idx_resol is not None else 0
                ts[m]["duree_reelle_h"] = decimal_hours_cell(r[idx_reelle]) if idx_reelle is not None else 0
                result["durations_months"] += 1

        elif is_categorie:
            src["categories"] = parse_glpi_row_sheet(ws, headers)
            result["categories"] = len(src["categories"])
        elif is_service:
            src["services"] = parse_glpi_row_sheet(ws, headers)
            result["services"] = len(src["services"])
        elif is_technicien:
            src["techniciens"] = parse_glpi_row_sheet(ws, headers)
            result["techniciens"] = len(src["techniciens"])
        elif is_demandeur:
            src["demandeurs"] = parse_glpi_row_sheet(ws, headers)
            result["demandeurs"] = len(src["demandeurs"])

    result["detected_month"] = detected_month
    return result, ts, src


def parse_csv_rows_ticket_series(rows, headers):
    """Couche 'parsing' — kind='tickets'. Ne fait aucune écriture, retourne une liste de
    dicts normalisés {"mois", "ouverts", "resolus", "en_retard", "clos"}. Réutilisée à la fois
    par l'import CSV (ci-dessous) et par le futur mode Direct GLPI (Lot C), qui construira ces
    mêmes dicts à partir de la réponse de l'API REST plutôt que d'un fichier."""
    idx_mois = 0
    idx_ouverts = match_header(headers, "ouverts")
    idx_resolus = match_header(headers, "resolus")
    idx_retard = match_header(headers, "retard")
    idx_clos = match_header(headers, "clos")
    records = []
    for r in rows:
        if not r or not r[idx_mois]:
            continue
        m = r[idx_mois].strip()
        if not m:
            continue
        records.append({
            "mois": m,
            "ouverts": int(r[idx_ouverts] or 0) if idx_ouverts is not None else 0,
            "resolus": int(r[idx_resolus] or 0) if idx_resolus is not None else 0,
            "en_retard": int(r[idx_retard] or 0) if idx_retard is not None else 0,
            "clos": int(r[idx_clos] or 0) if idx_clos is not None else 0,
        })
    return records


def parse_csv_rows_durations(rows, headers):
    """Couche 'parsing' — kind='durations'. Retourne une liste de dicts normalisés
    {"mois", "cloture_h", "resolution_h", "duree_reelle_h"}."""
    idx_mois = 0
    idx_clot = match_header(headers, "cloture")
    idx_resol = match_header(headers, "resolution")
    idx_reelle = match_header(headers, "duree reelle")
    records = []
    for r in rows:
        if not r or not r[idx_mois]:
            continue
        m = r[idx_mois].strip()
        if not m:
            continue
        records.append({
            "mois": m,
            "cloture_h": decimal_hours_cell(r[idx_clot]) if idx_clot is not None else 0,
            "resolution_h": decimal_hours_cell(r[idx_resol]) if idx_resol is not None else 0,
            "duree_reelle_h": decimal_hours_cell(r[idx_reelle]) if idx_reelle is not None else 0,
        })
    return records


def parse_csv_rows_source_items(rows, headers):
    """Couche 'parsing' — kind in (categories, services, techniciens, demandeurs). Retourne une
    liste de dicts normalisés {"label", "ouverts", "resolus", "en_retard", "fermes"}."""
    idx_ouverts = match_header(headers, "ouverts")
    idx_resolus = match_header(headers, "resolus")
    idx_retard = match_header(headers, "retard")
    idx_fermes = match_header(headers, "fermes", "ferme")
    records = []
    for r in rows:
        if not r or not r[0]:
            continue
        records.append({
            "label": r[0].strip(),
            "ouverts": int(r[idx_ouverts] or 0) if idx_ouverts is not None else 0,
            "resolus": int(r[idx_resolus] or 0) if idx_resolus is not None else 0,
            "en_retard": int(r[idx_retard] or 0) if idx_retard is not None else 0,
            "fermes": int(r[idx_fermes] or 0) if idx_fermes is not None else 0,
        })
    return records


def store_ticket_series(records, ts):
    """Couche 'stockage' — fusionne une liste de dicts {"mois", ...} dans le dict timeseries
    (ts), indexé par mois. Indépendante de l'origine des données (CSV ou API GLPI directe)."""
    for rec in records:
        m = rec["mois"]
        ts.setdefault(m, {})
        for k, v in rec.items():
            if k != "mois":
                ts[m][k] = v
    return ts


def store_duration_series(records, ts):
    """Couche 'stockage' — identique à store_ticket_series, séparée pour la lisibilité (les
    deux appels successifs, tickets puis durations, alimentent le même dict ts par mois)."""
    return store_ticket_series(records, ts)


def store_source_items(kind, records, src):
    """Couche 'stockage' — kind in (categories, services, techniciens, demandeurs). Remplace
    intégralement src[kind] (comportement identique à l'import historique : un nouvel import
    écrase la liste précédente, il ne la fusionne pas)."""
    items = []
    for rec in records:
        items.append({
            "label": rec.get("label", ""),
            "ouverts": rec.get("ouverts", 0), "resolus": rec.get("resolus", 0),
            "en_retard": rec.get("en_retard", 0), "fermes": rec.get("fermes", 0),
            "delai_prise_en_compte_s": rec.get("delai_prise_en_compte_s", 0),
            "delai_resolution_s": rec.get("delai_resolution_s", 0),
            "delai_fermeture_s": rec.get("delai_fermeture_s", 0),
        })
    src[kind] = items
    return src


def import_moussanada_csv(file_storage, kind, ts, src):
    """Parse un export GLPI séparé (.csv) d'un type donné, puis stocke. Conservé tel quel côté
    API publique (compatibilité totale avec l'import CSV existant) mais réécrit en interne pour
    appeler les fonctions de parsing/stockage séparées ci-dessus — ce sont ces mêmes fonctions
    de stockage que le mode Direct GLPI (Lot C) réutilisera avec des données venues de l'API
    plutôt que d'un fichier, garantissant une seule logique métier des deux côtés."""
    text = file_storage.read().decode("utf-8-sig", errors="replace")
    reader = _csv.reader(_io.StringIO(text), delimiter=";" if ";" in text.split("\n")[0] else ",")
    rows = list(reader)
    if not rows:
        return {"rows": 0}, ts, src
    headers, data_rows = rows[0], rows[1:]

    if kind == "tickets":
        records = parse_csv_rows_ticket_series(data_rows, headers)
        store_ticket_series(records, ts)
        return {"rows": len(records)}, ts, src

    if kind == "durations":
        records = parse_csv_rows_durations(data_rows, headers)
        store_duration_series(records, ts)
        return {"rows": len(records)}, ts, src

    if kind in ("categories", "services", "techniciens", "demandeurs"):
        records = parse_csv_rows_source_items(data_rows, headers)
        store_source_items(kind, records, src)
        return {"rows": len(records)}, ts, src

    return {"rows": 0}, ts, src


def parse_glpi_raw_tickets(file_storage):
    """Parse un export GLPI 'brut' ticket-par-ticket (ex. colonnes ID, Titre, Statut,
    Date d'ouverture, Priorité, Demandeur, Groupe demandeur, Technicien, Catégorie, TTR,
    Type, Dernière modification). Ce format, distinct des feuilles agrégées habituelles,
    permet de calculer des indicateurs que GLPI n'agrège pas nativement : répartition par
    Type/Priorité, et surtout la heatmap de charge (jour de semaine x heure d'ouverture),
    puisque la date d'ouverture inclut l'heure.
    Retourne une liste de tickets normalisés, dédupliqués par ID.
    """
    text = file_storage.read().decode("utf-8-sig", errors="replace")
    first_line = text.split("\n")[0]
    delim = ";" if first_line.count(";") >= first_line.count(",") else ","
    reader = _csv.reader(_io.StringIO(text), delimiter=delim)
    rows = list(reader)
    if not rows:
        return []
    headers = rows[0]

    idx_id = match_header(headers, "id")
    idx_titre = match_header(headers, "titre")
    idx_statut = match_header(headers, "statut")
    idx_ouverture = match_header(headers, "date d ouverture", "ouverture")
    idx_priorite = match_header(headers, "priorite")
    idx_demandeur = match_header(headers, "demandeur demandeur", "demandeur")
    idx_groupe = match_header(headers, "demandeur groupe", "groupe demandeur")
    idx_technicien = match_header(headers, "attribue a technicien", "technicien")
    idx_categorie = match_header(headers, "categorie")
    # "ttr" est un intitulé de colonne court et exact chez GLPI ; un match par sous-chaîne
    # (comme pour les autres colonnes) capturerait à tort "aTTRibué à - Technicien". On exige
    # donc une égalité stricte de l'en-tête normalisé pour cette seule colonne.
    norm_headers = [normalize_header(h) for h in headers]
    idx_ttr = norm_headers.index("ttr") if "ttr" in norm_headers else None
    idx_type = match_header(headers, "type")
    idx_modif = match_header(headers, "derniere modification", "modification")

    def g(row, idx):
        return row[idx].strip() if idx is not None and idx < len(row) and row[idx] else ""

    tickets = []
    for row in rows[1:]:
        if not row or not g(row, idx_id):
            continue
        tid = g(row, idx_id).replace(" ", "").replace("\u202f", "")
        tickets.append({
            "id": tid,
            "titre": g(row, idx_titre),
            "statut": g(row, idx_statut),
            "date_ouverture": g(row, idx_ouverture),
            "priorite": g(row, idx_priorite),
            "demandeur": g(row, idx_demandeur),
            "groupe_demandeur": g(row, idx_groupe),
            "technicien": g(row, idx_technicien),
            "categorie": g(row, idx_categorie),
            "ttr": g(row, idx_ttr),
            "type": g(row, idx_type),
            "derniere_modification": g(row, idx_modif),
        })
    return tickets


def merge_raw_tickets(existing, new_tickets):
    """Fusionne par ID (le plus récent écrase l'ancien) — permet des ré-imports incrémentaux
    sans dupliquer les tickets déjà connus."""
    by_id = {t["id"]: t for t in existing}
    for t in new_tickets:
        by_id[t["id"]] = t
    return list(by_id.values())


def compute_heatmap_matrix(tickets, date_field="date_ouverture"):
    """Construit une matrice 7 (jours, lundi=0) x 24 (heures) du nombre de tickets ouverts,
    à partir d'un champ date 'YYYY-MM-DD HH:MM[:SS]'. Utilisé pour la heatmap de charge."""
    matrix = [[0] * 24 for _ in range(7)]
    for t in tickets:
        raw = (t.get(date_field) or "").strip()
        if not raw:
            continue
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})", raw)
        if not m:
            continue
        y, mo, d, h, mi = (int(x) for x in m.groups())
        try:
            weekday = _dt.datetime(y, mo, d).weekday()  # 0=lundi
        except ValueError:
            continue
        matrix[weekday][h] += 1
    return matrix


def parse_glpi_datetime(raw):
    """Parse une date GLPI 'YYYY-MM-DD HH:MM[:SS]' en datetime, ou None si invalide/vide."""
    raw = (raw or "").strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?", raw)
    if not m:
        return None
    y, mo, d, h, mi, s = m.groups()
    try:
        return _dt.datetime(int(y), int(mo), int(d), int(h), int(mi), int(s or 0))
    except ValueError:
        return None


def compute_resolution_hours(tickets):
    """Calcule, pour chaque ticket ayant une date d'ouverture et une date de résolution (colonne
    TTR), la durée écoulée en heures. Ignore les tickets encore ouverts (TTR vide) ou aux dates
    incohérentes. Alimente l'histogramme de distribution des délais (plus parlant qu'une simple
    moyenne, qui masque les cas extrêmes)."""
    durations = []
    for t in tickets:
        d1 = parse_glpi_datetime(t.get("date_ouverture"))
        d2 = parse_glpi_datetime(t.get("ttr"))
        if not d1 or not d2:
            continue
        delta_h = (d2 - d1).total_seconds() / 3600
        if delta_h >= 0:
            durations.append(delta_h)
    return durations


def bucket_resolution_hours(durations):
    """Répartit des durées (heures) en tranches lisibles pour un histogramme."""
    buckets = [
        ("< 1h", 0, 1), ("1-4h", 1, 4), ("4-8h", 4, 8), ("8-24h", 8, 24),
        ("1-3j", 24, 72), ("3-7j", 72, 168), ("> 7j", 168, float("inf")),
    ]
    counts = {label: 0 for label, _, _ in buckets}
    for d in durations:
        for label, lo, hi in buckets:
            if lo <= d < hi:
                counts[label] += 1
                break
    return counts
