# -*- coding: utf-8 -*-
"""
Construit le gabarit Excel « Tickets en retard » (charte AMMPS) : static/charte/modele_tickets_en_retard.xlsx.

Les lignes d'exemple viennent d'un export GLPI brut (par défaut reference/glpi (1).csv, situation figée au
2026-11-02 18:00). Au quotidien, retards_xlsx.fill() charge ce gabarit, retire les exemples et injecte les
vrais tickets en reprenant le style des deux premières lignes. À relancer seulement pour changer la
mise en page :   python tools/build_retards_template.py [export.csv]
"""
import io
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from openpyxl import Workbook  # noqa: E402
from openpyxl.drawing.image import Image as XlImage  # noqa: E402
from openpyxl.formatting.rule import CellIsRule, DataBarRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

import ammps_charte as C  # noqa: E402
import retards  # noqa: E402
import retards_xlsx as X  # noqa: E402
from glpi_import import parse_glpi_raw_tickets  # noqa: E402

FOREST, GREEN, RED, INK, MUTED, LINE, PALE, LIGHT = (C.hx(c) for c in (C.FOREST, C.GREEN, C.RED, C.INK, C.MUTED, C.LINE, C.PALE, C.LIGHT))
FONT = "Arial"
SAMPLE_NOW = datetime(2026, 11, 2, 18, 0)
LAST = 5000   # plage des mises en forme conditionnelles : au-delà de tout export réaliste


def fill(color):
    return PatternFill("solid", fgColor=color)


def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def side(color=LINE, style="thin"):
    return Side(style=style, color=color)


def band(ws, first_col, last_col, logo_px, rows, rule_row, text_offset=1, indent=7, red_cols=1):
    """Logo + nom de l'agence (+ filet vert/rouge) en haut de feuille. Le texte démarre dans la colonne
    first_col + text_offset, décalé d'un retrait qui le place à droite du logo."""
    for r, h in rows.items():
        ws.row_dimensions[r].height = h
    img = XlImage(io.BytesIO(C.LOGO_PNG))
    img.width, img.height = logo_px, round(logo_px * 339 / 581)
    ws.add_image(img, f"{get_column_letter(first_col)}2")
    c = ws.cell(2, first_col + text_offset, C.AGENCY_LINES[0])
    c.font, c.alignment = font(12, True, FOREST), Alignment(vertical="bottom", indent=indent)
    c = ws.cell(3, first_col + text_offset, C.AGENCY_LINES[1])
    c.font, c.alignment = font(10, False, MUTED), Alignment(vertical="top", indent=indent)
    ws.row_dimensions[rule_row].height = 4
    for col in range(first_col, last_col + 1):
        ws.cell(rule_row, col).fill = fill(RED if col > last_col - red_cols else FOREST)


def header_row(ws, row, heads, first_col, height=26):
    ws.row_dimensions[row].height = height
    for i, h in enumerate(heads):
        c = ws.cell(row, first_col + i, h)
        c.font, c.fill = font(10, True, "FFFFFF"), fill(FOREST)
        c.alignment = Alignment(horizontal="left" if i in (1, 2, 4, 5, 6) else "center", vertical="center", wrap_text=True, indent=1 if i in (1, 2, 4, 5, 6) else 0)


def data_style(ws, row, ncols, first_col, left_cols, formats, zebra):
    for i in range(ncols):
        c = ws.cell(row, first_col + i)
        c.font = font(10, bold=(i == ncols - 1 and ncols == len(X.T_HEAD)))
        c.fill = fill(PALE if zebra else "FFFFFF")
        c.border = Border(bottom=side())
        left = i in left_cols
        c.alignment = Alignment(horizontal="left" if left else "center", vertical="center", wrap_text=left, indent=1 if left else 0)
        if i in formats:
            c.number_format = formats[i]


def sheet_common(ws, widths, first_col, tab):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = tab
    ws.column_dimensions["A"].width = 2
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(first_col + i)].width = w
    ws.column_dimensions[get_column_letter(first_col + len(widths))].width = 2
    ws.page_setup.orientation, ws.page_setup.paperSize = "landscape", 9
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top, ws.page_margins.bottom = 0.5, 0.6
    ws.oddFooter.center.text = "AMMPS · Tickets en retard · Page &P / &N"
    ws.oddFooter.center.size, ws.oddFooter.center.font = 8, "Arial,Regular"
    ws.oddFooter.center.color = MUTED


def title_block(ws, title_cell, sub_cell, last_col_letter, title="Tickets en retard", title_end=None):
    ws[title_cell] = title
    ws[title_cell].font = font(22, True, FOREST)
    ws[title_cell].alignment = Alignment(vertical="center")
    ws[sub_cell].font = font(10.5, False, MUTED)
    ws[sub_cell].alignment = Alignment(vertical="top", wrap_text=True)
    r = ws[sub_cell].row
    ws.merge_cells(f"{sub_cell}:{last_col_letter}{r}")
    tr = ws[title_cell].row
    ws.merge_cells(f"{title_cell}:{title_end or last_col_letter}{tr}")
    ws.row_dimensions[tr].height = 34
    ws.row_dimensions[r].height = 30


def build_tickets_sheet(ws, v):
    sheet_common(ws, X.T_WIDTHS, X.T_FIRST_COL, GREEN)
    last_col = X.T_FIRST_COL + len(X.T_HEAD) - 1
    band(ws, X.T_FIRST_COL, last_col, 104, {1: 6, 2: 22, 3: 18, 4: 12, 5: 4, 6: 6}, 5)
    ws.merge_cells(f"{get_column_letter(last_col - 1)}2:{get_column_letter(last_col)}2")
    ws.merge_cells(f"{get_column_letter(last_col - 1)}3:{get_column_letter(last_col)}3")
    ed_lbl, ed = ws.cell(2, last_col - 1, "ÉDITION"), ws[X.EDITION_CELL]
    ed_lbl.font, ed_lbl.alignment = font(8, True, MUTED), Alignment(horizontal="right", vertical="bottom")
    ed.font, ed.alignment = font(11, True, FOREST), Alignment(horizontal="right", vertical="top")
    title_block(ws, X.TITLE_CELL, X.SUB_CELL, get_column_letter(last_col), title_end="H")

    # cartes KPI : (colonnes fusionnées, couleur du filet haut, couleur de la valeur)
    cards = [("B", "C", RED, RED), ("D", "F", GREEN, FOREST), ("G", "H", GREEN, FOREST), ("I", "K", GREEN, FOREST)]
    labels = ["TICKETS EN RETARD", "RETARD MOYEN", "RETARD MAXIMAL", "PRIORITÉ HAUTE"]
    for (c1, c2, top, val), label in zip(cards, labels):
        for r, h in ((10, 18), (11, 32), (12, 18)):
            ws.row_dimensions[r].height = h
            ws.merge_cells(f"{c1}{r}:{c2}{r}")
            for col in range(ord(c1) - 64, ord(c2) - 63):
                cell = ws.cell(r, col)
                cell.fill = fill(LIGHT)
                cell.border = Border(top=side(top, "thick") if r == 10 else None,
                                     left=side("FFFFFF", "thick") if col == ord(c1) - 64 and c1 != "B" else None)
        a = ws[f"{c1}10"]
        a.value, a.font, a.alignment = label, font(8, True, MUTED), Alignment(horizontal="left", vertical="center", indent=1)
        a = ws[f"{c1}11"]
        a.font, a.alignment = font(22, True, val), Alignment(horizontal="left", vertical="center", indent=1)
        a = ws[f"{c1}12"]
        a.font, a.alignment = font(9, False, MUTED), Alignment(horizontal="left", vertical="top", indent=1)
    ws.row_dimensions[9].height = 6
    ws.row_dimensions[13].height = 8

    header_row(ws, X.HEADER_ROW, X.T_HEAD, X.T_FIRST_COL)
    formats = {7: "dd/mm/yyyy hh:mm", 8: "dd/mm/yyyy hh:mm", 9: "0.0"}
    for i, t in enumerate(v["late"]):
        r = X.FIRST_ROW + i
        data_style(ws, r, len(X.T_HEAD), X.T_FIRST_COL, {1, 2, 4, 5, 6}, formats, i % 2 == 1)
        for j, val in enumerate(X._ticket_row(t)):
            ws.cell(r, X.T_FIRST_COL + j, val)
    # mises en forme conditionnelles sur toute la plage utile (le remplissage ne les touche pas)
    prio, delay = f"E{X.FIRST_ROW}:E{LAST}", f"K{X.FIRST_ROW}:K{LAST}"
    ws.conditional_formatting.add(prio, CellIsRule(operator="equal", formula=['"Haute"'], font=Font(name=FONT, bold=True, color=RED), fill=fill("FDECED")))
    ws.conditional_formatting.add(prio, CellIsRule(operator="equal", formula=['"Très haute"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=fill(RED)))
    ws.conditional_formatting.add(delay, DataBarRule(start_type="num", start_value=0, end_type="max", color="E9A3A6", showValue=True))
    ws.conditional_formatting.add(delay, CellIsRule(operator="greaterThanOrEqual", formula=["7"], font=Font(name=FONT, bold=True, color=RED)))
    ws.auto_filter.ref = f"B{X.HEADER_ROW}:{get_column_letter(last_col)}{X.FIRST_ROW + max(len(v['late']), 1) - 1}"


def build_synthese_sheet(ws, v):
    sheet_common(ws, X.S_WIDTHS, X.T_FIRST_COL, FOREST)
    last_col = X.T_FIRST_COL + len(X.S_HEAD) - 1
    band(ws, X.T_FIRST_COL, last_col, 84, {1: 6, 2: 20, 3: 16, 4: 4, 5: 6}, 4, text_offset=0, indent=11)
    ws.row_dimensions[8].height = 8
    title_block(ws, X.S_TITLE_CELL, X.S_SUB_CELL, get_column_letter(last_col), "Synthèse par technicien")
    header_row(ws, X.S_HEADER_ROW, X.S_HEAD, X.T_FIRST_COL)
    ws.cell(X.S_HEADER_ROW, X.T_FIRST_COL).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    formats = {2: "0.0%", 3: "0.0", 4: "0.0"}
    rows = X._synthese(v["late"])
    for i, r in enumerate(rows):
        row = X.S_FIRST_ROW + i
        data_style(ws, row, len(X.S_HEAD), X.T_FIRST_COL, {0}, formats, i % 2 == 1)
        for j, val in enumerate(r):
            ws.cell(row, X.T_FIRST_COL + j, val)
    tot = X.S_FIRST_ROW + len(rows)
    hours = [t["retard_h"] for t in v["late"]]
    total = [X.FOOTER_MODEL, len(v["late"]), 1, round(sum(hours) / len(hours) / 24, 1), round(max(hours) / 24, 1), v["high"]]
    for j, val in enumerate(total):
        c = ws.cell(tot, X.T_FIRST_COL + j, val)
        c.font, c.fill = font(10, True, FOREST), fill(LIGHT)
        c.border = Border(top=side(FOREST, "medium"))
        c.alignment = Alignment(horizontal="left" if j == 0 else "center", vertical="center", indent=1 if j == 0 else 0)
        if j in formats:
            c.number_format = formats[j]
    ws.row_dimensions[tot].height = 22
    ws.conditional_formatting.add(f"C{X.S_FIRST_ROW}:C{tot - 1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="8CC9A3", showValue=True))
    ws.conditional_formatting.add(f"G{X.S_FIRST_ROW}:G{tot - 1}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, bold=True, color=RED)))


def build(csv_path, out_path):
    with open(csv_path, "rb") as f:
        class Upload:
            def read(self):
                return f.read()
        tickets = parse_glpi_raw_tickets(Upload())
    v = retards.view(tickets, None, SAMPLE_NOW)
    wb = Workbook()
    ws = wb.active
    ws.title = X.SHEET_T
    build_tickets_sheet(ws, v)
    build_synthese_sheet(wb.create_sheet(X.SHEET_S), v)
    # textes d'en-tête et KPI de l'exemple : mêmes valeurs que le remplissage réel
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), "retards_tpl_raw.xlsx")
    wb.properties.creator, wb.properties.title = "AMMPS — DSID", "Gabarit — Tickets en retard"
    wb.save(tmp)
    sample = X.fill(v, SAMPLE_NOW, template=tmp)
    with open(out_path, "wb") as f:
        f.write(sample)
    os.remove(tmp)
    return v


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "reference", "glpi (1).csv")
    out = os.path.join(ROOT, "static", "charte", "modele_tickets_en_retard.xlsx")
    v = build(src, out)
    print(f"Gabarit écrit : {out} ({v['total']} tickets en retard d'exemple sur {v['open_total']} ouverts)")
