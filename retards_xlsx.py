# -*- coding: utf-8 -*-
"""
retards_xlsx.py — Excel « Tickets en retard » (Moussanada), généré par gabarit.

Le gabarit static/charte/modele_tickets_en_retard.xlsx (construit par tools/build_retards_template.py)
porte toute la mise en page AMMPS : logo, bandeau, cartes KPI, colonnes, mises en forme conditionnelles,
impression. Ici on ne fait que le remplir : textes d'en-tête, valeurs des KPI, puis une ligne par ticket
en reprenant le style des deux lignes modèles du gabarit (zébrage). Les mêmes constantes de position sont
utilisées par le constructeur du gabarit et par le remplissage.
"""
import io
import os
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import retards

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "charte", "modele_tickets_en_retard.xlsx")
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

SHEET_T, SHEET_S = "Tickets en retard", "Synthèse par technicien"
# --- feuille 1 : tickets
TITLE_CELL, SUB_CELL, EDITION_CELL = "B7", "B8", "J3"
KPI_VALUE_CELLS, KPI_SUB_CELLS = ("B11", "D11", "G11", "I11"), ("B12", "D12", "G12", "I12")
HEADER_ROW, FIRST_ROW = 14, 15
T_HEAD = ["N°", "Titre", "Statut", "Priorité", "Technicien", "Demandeur", "Catégorie", "Ouvert le", "Échéance (TTR)", "Retard (j)"]
T_WIDTHS = [9, 54, 19, 11, 20, 24, 40, 17, 17, 11]
T_FIRST_COL = 2          # colonne B
# --- feuille 2 : synthèse par technicien
S_TITLE_CELL, S_SUB_CELL, S_HEADER_ROW, S_FIRST_ROW = "B6", "B7", 9, 10
S_HEAD = ["Technicien", "Tickets en retard", "Part du total", "Retard moyen (j)", "Retard max (j)", "Priorité haute"]
S_WIDTHS = [30, 17, 14, 17, 15, 15]
FOOTER_MODEL = "TOTAL"


def fr_days(v):
    return f"{v:.1f}".replace(".", ",")


def _put(ws, row, values, style_row, first_col=T_FIRST_COL):
    for i, v in enumerate(values):
        col = first_col + i
        dst, src = ws.cell(row, col), style_row[i]
        dst.value = v
        dst._style = copy(src._style)


def _models(ws, first_row, ncols, first_col=T_FIRST_COL):
    """Cellules de style des deux lignes modèles (blanche / zébrée) du gabarit."""
    return [[ws.cell(first_row + k, first_col + i) for i in range(ncols)] for k in (0, 1)]


def _clear_below(ws, first_row):
    """Supprime les lignes d'exemple du gabarit (valeurs, styles et hauteurs)."""
    if ws.max_row >= first_row:
        ws.delete_rows(first_row, ws.max_row - first_row + 1)
    for r in [r for r in ws.row_dimensions if r >= first_row]:
        del ws.row_dimensions[r]


def _ticket_row(t):
    return [t["id"], t["titre"], t["statut"], t["priorite"], t["technicien"], t["demandeur"], t["categorie"],
            t["date_ouverture"], t["echeance"], round(t["retard_h"] / 24, 1)]


def _synthese(late):
    by = {}
    for t in late:
        by.setdefault(t["technicien"], []).append(t)
    total = len(late) or 1
    rows = []
    for name, items in by.items():
        hours = [i["retard_h"] for i in items]
        rows.append([name, len(items), len(items) / total, round(sum(hours) / len(hours) / 24, 1),
                     round(max(hours) / 24, 1), sum(1 for i in items if (i["priorite"] or "").lower() in ("haute", "très haute", "majeure"))])
    rows.sort(key=lambda r: (-r[1], -r[4], r[0].lower()))
    return rows


def fill(v, edition=None, template=TEMPLATE):
    """Octets du classeur rempli. `v` = retards.view() ; les tickets des techniciens écartés (réglage du
    dashboard) n'y figurent pas, pas plus que dans les KPI et la synthèse."""
    late = v["shown"]
    total, avg_days, max_days, high = retards.stats(late)
    wb = load_workbook(template)
    ws, ss = wb[SHEET_T], wb[SHEET_S]
    now = v["as_of"]

    # ---- en-tête et KPI
    ws[TITLE_CELL] = "Tickets en retard"
    sub = (f"Situation au {now:%d/%m/%Y à %H:%M} — tickets non résolus dont l’échéance (TTR) est dépassée"
           + (f". {v['no_due']} ticket(s) ouvert(s) sans échéance renseignée ne sont pas comptés." if v["no_due"] else "."))
    ws[SUB_CELL] = sub
    ws[EDITION_CELL] = (edition or now).strftime("%d/%m/%Y")
    oldest = late[0] if late else None
    for cell, val in zip(KPI_VALUE_CELLS, (total, f"{fr_days(avg_days)} j", f"{fr_days(max_days)} j", high)):
        ws[cell] = val
    for cell, val in zip(KPI_SUB_CELLS, (f"sur {v['open_total']} ticket(s) non résolu(s)", "écart moyen à l’échéance",
                                          f"ticket n° {oldest['id']}" if oldest else "—", "tickets de priorité haute")):
        ws[cell] = val

    # ---- feuille 1 : une ligne par ticket
    models = _models(ws, FIRST_ROW, len(T_HEAD))
    _clear_below(ws, FIRST_ROW)
    for i, t in enumerate(late):
        _put(ws, FIRST_ROW + i, _ticket_row(t), models[i % 2])
    last = FIRST_ROW + max(len(late), 1) - 1
    if not late:
        ws.cell(FIRST_ROW, T_FIRST_COL + 1).value = "Aucun ticket en retard."
    first_l, last_l = get_column_letter(T_FIRST_COL), get_column_letter(T_FIRST_COL + len(T_HEAD) - 1)
    ws.auto_filter.ref = f"{first_l}{HEADER_ROW}:{last_l}{last}"
    ws.print_area = f"A1:{get_column_letter(T_FIRST_COL + len(T_HEAD))}{last + 1}"
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"

    # ---- feuille 2 : synthèse par technicien
    ss[S_SUB_CELL] = f"Répartition des {total} ticket(s) en retard au {now:%d/%m/%Y}."
    total_row = next((r for r in range(S_FIRST_ROW, ss.max_row + 1) if ss.cell(r, T_FIRST_COL).value == FOOTER_MODEL), None)
    smodels = _models(ss, S_FIRST_ROW, len(S_HEAD))
    footer = [ss.cell(total_row, T_FIRST_COL + i) for i in range(len(S_HEAD))] if total_row else smodels[0]
    _clear_below(ss, S_FIRST_ROW)
    rows = _synthese(late)
    for i, r in enumerate(rows):
        _put(ss, S_FIRST_ROW + i, r, smodels[i % 2])
    _put(ss, S_FIRST_ROW + len(rows), [FOOTER_MODEL, total, 1 if late else 0, round(avg_days, 1), round(max_days, 1), high], footer)
    ss.print_area = f"A1:{get_column_letter(T_FIRST_COL + len(S_HEAD))}{S_FIRST_ROW + len(rows) + 1}"

    wb.active = 0
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def filename(now=None):
    return f"Tickets_en_retard_Moussanada_{(now or datetime.now()):%Y-%m-%d}.xlsx"
