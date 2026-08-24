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


def import_moussanada_csv(file_storage, kind, ts, src):
    """Parse un export GLPI séparé (.csv) d'un type donné. Mute et retourne (result, ts, src)."""
    text = file_storage.read().decode("utf-8-sig", errors="replace")
    reader = _csv.reader(_io.StringIO(text), delimiter=";" if ";" in text.split("\n")[0] else ",")
    rows = list(reader)
    if not rows:
        return {"rows": 0}, ts, src
    headers = rows[0]

    if kind == "tickets":
        idx_mois = 0
        idx_ouverts = match_header(headers, "ouverts")
        idx_resolus = match_header(headers, "resolus")
        idx_retard = match_header(headers, "retard")
        idx_clos = match_header(headers, "clos")
        for r in rows[1:]:
            if not r or not r[idx_mois]:
                continue
            m = r[idx_mois].strip()
            if not m:
                continue
            ts.setdefault(m, {})
            ts[m]["ouverts"] = int(r[idx_ouverts] or 0) if idx_ouverts is not None else 0
            ts[m]["resolus"] = int(r[idx_resolus] or 0) if idx_resolus is not None else 0
            ts[m]["en_retard"] = int(r[idx_retard] or 0) if idx_retard is not None else 0
            ts[m]["clos"] = int(r[idx_clos] or 0) if idx_clos is not None else 0
        return {"rows": len(rows) - 1}, ts, src

    if kind == "durations":
        idx_mois = 0
        idx_clot = match_header(headers, "cloture")
        idx_resol = match_header(headers, "resolution")
        idx_reelle = match_header(headers, "duree reelle")
        for r in rows[1:]:
            if not r or not r[idx_mois]:
                continue
            m = r[idx_mois].strip()
            if not m:
                continue
            ts.setdefault(m, {})
            ts[m]["cloture_h"] = decimal_hours_cell(r[idx_clot]) if idx_clot is not None else 0
            ts[m]["resolution_h"] = decimal_hours_cell(r[idx_resol]) if idx_resol is not None else 0
            ts[m]["duree_reelle_h"] = decimal_hours_cell(r[idx_reelle]) if idx_reelle is not None else 0
        return {"rows": len(rows) - 1}, ts, src

    if kind in ("categories", "services", "techniciens", "demandeurs"):
        idx_ouverts = match_header(headers, "ouverts")
        idx_resolus = match_header(headers, "resolus")
        idx_retard = match_header(headers, "retard")
        idx_fermes = match_header(headers, "fermes", "ferme")
        items = []
        for r in rows[1:]:
            if not r or not r[0]:
                continue
            items.append({
                "label": r[0].strip(),
                "ouverts": int(r[idx_ouverts] or 0) if idx_ouverts is not None else 0,
                "resolus": int(r[idx_resolus] or 0) if idx_resolus is not None else 0,
                "en_retard": int(r[idx_retard] or 0) if idx_retard is not None else 0,
                "fermes": int(r[idx_fermes] or 0) if idx_fermes is not None else 0,
                "delai_prise_en_compte_s": 0, "delai_resolution_s": 0, "delai_fermeture_s": 0,
            })
        src[kind] = items
        return {"rows": len(items)}, ts, src

    return {"rows": 0}, ts, src
