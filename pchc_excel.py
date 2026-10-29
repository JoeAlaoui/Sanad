# -*- coding: utf-8 -*-
"""
pchc_excel.py — Classeur « Rapport Tarkhiss PCHC Métier » du Reporting Métier PCHC.

Copie conforme de la référence validée (reference/Rapport Tarkhiss_PCHC_Metier_YTD.xlsx, enregistrée
par Excel) : le classeur est produit à partir du GABARIT pchc_template/dashboard_gabarit.xlsx (la
référence elle-même, cf. tools/build_pchc_template.py). Graphiques, styles, couleurs, mises en forme
conditionnelles, fusions, largeurs et mise en page sont donc repris octet pour octet ; seuls sont
réécrits :
- les 5 feuilles brutes (masquées, comme dans la référence), depuis les données importées ;
- les valeurs en cache des formules (les FORMULES restent vivantes : COUNTA/COUNTIF/SUMIF jusqu'à la
  ligne 50 000) et des caches des graphiques, calculées par pchc_raw.pchc_compute() ;
- la date « Mise à jour » et les textes qui en dépendent.
Version « allégée » : sans les feuilles brutes ; les formules qui les lisaient deviennent des valeurs.

_Builder (port de reference/build_dashboard_pchc_reference.py, openpyxl) sert à calculer la valeur de
chaque cellule (et de repli si le gabarit est absent). recalc_check() : recalcul LibreOffice headless
si disponible, contrôle 0 erreur et égalité des valeurs avec le cache.
"""
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime, timezone
from xml.sax.saxutils import escape as xml_escape
from xml.sax.saxutils import unescape as xml_unescape

import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

import pchc_raw as R

NAVY, GREY, LGREY, LINE = "1F3864", "5B6B7F", "F2F5F9", "D0D7E2"
GREEN, AMBER, RED = "C6EFCE", "FFE699", "FFCCCC"
THEMES = {
    "KPI Identification": dict(title="2E75B6", light="D6E3F8", sec="4472C4", secfont="FFFFFF", hdr="8FAADC"),
    "KPI Déclaration": dict(title="538135", light="E2EFDA", sec="70AD47", secfont="FFFFFF", hdr="A9D08E"),
    "KPI Certificat Enreg.": dict(title="C65911", light="FCE4D6", sec="ED7D31", secfont="FFFFFF", hdr="F4B183"),
    # sec 9667CA : teinte du classeur de référence validé (le script d'origine indiquait 9966CC)
    "KPI Certificat LV": dict(title="7030A0", light="E2D1F2", sec="9667CA", secfont="FFFFFF", hdr="B4A7D6"),
    "KPI Autor. Import.": dict(title="1F4E79", light="D6DCE5", sec="2F5496", secfont="FFFFFF", hdr="8FAADC"),
    "Dashboard": dict(title="1F4E79", light="D6DCE4", sec="BDD7EE", secfont="1F4E79", hdr="2F5496"),
}
CAT_FILL = {"Délivré": GREEN, "En cours": "E4EEF9", "Non conforme": AMBER, "Refusé": RED, "Autre": "ECEFF3"}
CAT_PT = {"Délivré": ["2E8B57"], "En cours": ["1F6FB2", "4FA3D9", "8CC4EC", "0E4C7A", "5B7DB1", "A9C4E0", "3B5998"],
          "Non conforme": ["F4B400"], "Refusé": ["D64541", "A93226"], "Autre": ["9AA5B1"]}
TYPE_PAL = ["1F6FB2", "2E8B57", "F4B400", "8E6BBF"]
SUBTITLE = "Données Tarkhiss - Produits Cosmétiques et d’Hygiène Corporelle"
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
thin = Side(style="thin", color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
RAW_TAB = "8A97A8"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def F(sz=10, b=False, c="000000", i=False):
    return Font(name="Calibri", size=sz, bold=b, color=c, italic=i)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def pt_colors(cats):
    cnt, out = {}, []
    for c in cats:
        i = cnt.get(c, 0)
        cnt[c] = i + 1
        out.append(CAT_PT[c][i % len(CAT_PT[c])])
    return out


def _xl_sum(values):
    """SOMME() d'Excel : additions successives en double précision (ex. 0,99999999999999989)."""
    total = 0.0
    for v in values:
        total += v
    return total


def _sheet_ref(name):
    """Nom de feuille dans une formule : guillemets seulement si nécessaire (comme Excel)."""
    return name if re.fullmatch(r"[^\W\d]\w*", name) else f"'{name}'"


def _rng(raw_sheet, col):
    return f"{_sheet_ref(raw_sheet)}!${col}${R.FIRST_ROW}:${col}${R.LAST_ROW}"


class _Builder:
    def __init__(self, raw, comp, update, light):
        self.raw, self.c, self.update, self.light = raw, comp, update, light
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)
        self.cache = {}     # (titre feuille, cellule) -> valeur calculée de la formule
        self.cur = dict(THEMES["Dashboard"])
        self.rowmap, self.info = {}, {}

    # -------------------------------------------------------------- primitives (port du script)
    def style_range(self, ws, rng, font=None, fl=None, al=None, border=None):
        for row in ws[rng]:
            for c in row:
                if font: c.font = font
                if fl: c.fill = fl
                if al: c.alignment = al
                if border: c.border = border

    def put(self, ws, ref, v, font=None, fl=None, al=None, border=None, nf=None, cv=None):
        c = ws[ref]
        c.value = v
        if isinstance(v, str) and v.startswith("="):
            self.cache[(ws.title, ref)] = cv
        if font: c.font = font
        if fl: c.fill = fl
        if al: c.alignment = al
        if border: c.border = border
        if nf: c.number_format = nf
        return c

    def rawf(self, formula, value):
        """Formule qui lit une feuille brute : valeur figée en version allégée."""
        return value if self.light else formula

    def merge(self, ws, rng, v, font, fl=None, al=None, border=None, nf=None, cv=None):
        a, b = rng.split(":")
        if a != b:
            ws.merge_cells(rng)
        self.style_range(ws, rng, font, fl, al, border)
        ws[a].value = v
        if isinstance(v, str) and v.startswith("="):
            self.cache[(ws.title, a)] = cv
        if nf:
            ws[a].number_format = nf

    def banner(self, ws, title, subtitle, last_col="L"):
        cur = self.cur
        ws.sheet_view.showGridLines = False
        self.merge(ws, f"A1:{last_col}1", title, F(18, True, "FFFFFF"), fill(cur["title"]), Alignment(horizontal="left", vertical="center", indent=1))
        ws.row_dimensions[1].height = 38
        self.merge(ws, "A2:E2", subtitle, F(10, False, "000000", True), fill(cur["light"]), Alignment(horizontal="left", vertical="center", indent=1))
        self.style_range(ws, f"F2:{last_col}2", fl=fill(cur["light"]))
        self.put(ws, "F2", "Mise à jour :", F(10, True), fill(cur["light"]), RIGHT)
        self.merge(ws, "G2:H2", self.update, F(10, True), fill(cur["light"]), Alignment(horizontal="left", vertical="center", indent=1), nf="dd/mm/yyyy")
        ws.row_dimensions[2].height = 20

    def section(self, ws, rng, text):
        cur = self.cur
        self.merge(ws, rng, text, F(12, True, cur["secfont"]), fill(cur["sec"]), Alignment(horizontal="left", vertical="center", indent=1))
        ws.row_dimensions[ws[rng.split(":")[0]].row].height = 22
        for row in ws[rng]:
            for c in row:
                c.border = Border(bottom=Side(style="medium", color=cur["title"]))

    def header(self, ws, rng, labels=None):
        cur = self.cur
        for row in ws[rng]:
            for c in row:
                c.font = F(10, True, "000000" if cur["hdr"] == "FFC000" else "FFFFFF")
                c.fill = fill(cur["hdr"])
                c.alignment = CENTER
                c.border = BOX
        if labels:
            for ref, t in labels.items():
                ws[ref].value = t

    def new_sheet(self, name, tab):
        ws = self.wb.create_sheet(name)
        ws.sheet_properties.tabColor = tab
        for col, w in zip("ABCDEFGHIJKL", [40, 12, 10, 17, 3, 12, 12, 12, 12, 12, 12, 12]):
            ws.column_dimensions[col].width = w
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        return ws

    def make_pie(self, ws, kind, title, cats_ref, data_ref, colors, anchor, w=15.5, h=8.2):
        ch = DoughnutChart() if kind == "doughnut" else PieChart()
        ch.title = title
        ch.style = 10
        ch.add_data(data_ref, titles_from_data=False)
        ch.set_categories(cats_ref)
        s = ch.series[0]
        for i, col in enumerate(colors):
            p = DataPoint(idx=i)
            p.graphicalProperties.solidFill = col
            p.graphicalProperties.line.solidFill = "FFFFFF"
            s.dPt.append(p)
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True; s.dLbls.showPercent = False
        s.dLbls.showCatName = False; s.dLbls.showSerName = False; s.dLbls.showLegendKey = False
        s.dLbls.numFmt = "#,##0;;"
        ch.legend.position = "r"
        ch.width, ch.height = w, h
        ws.add_chart(ch, anchor)

    def make_bar(self, ws, title, cats_ref, data_ref, colors, anchor, horizontal=True, w=15.5, h=9):
        ch = BarChart()
        ch.type = "bar" if horizontal else "col"
        ch.style = 10
        ch.title = title
        ch.add_data(data_ref, titles_from_data=False)
        ch.set_categories(cats_ref)
        s = ch.series[0]
        for i, col in enumerate(colors):
            p = DataPoint(idx=i)
            p.graphicalProperties.solidFill = col
            s.dPt.append(p)
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True; s.dLbls.showCatName = False
        s.dLbls.showSerName = False; s.dLbls.showPercent = False; s.dLbls.showLegendKey = False
        ch.legend = None
        ch.gapWidth = 40
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        if horizontal:
            ch.x_axis.scaling.orientation = "maxMin"
        ch.y_axis.numFmt = "#,##0"
        ch.width, ch.height = w, h
        ws.add_chart(ch, anchor)

    # -------------------------------------------------------------- feuilles KPI par statut
    def status_sheet(self, key):
        m = R.MOD[key]
        name, raw = m["kpi"], m["sheet"]
        mod = self.c["modules"][key]
        idc, stc = mod["cols"]["id"], mod["cols"]["status"]
        self.cur.update(THEMES[name])
        ws = self.new_sheet(name, self.cur["title"])
        title = {"declaration": "📋 KPI DÉCLARATION D'ACTIVITÉ", "ce": "📋 KPI CERTIFICAT D'ENREGISTREMENT (CE)",
                 "clv": "📋 KPI CERTIFICAT LIBRE VENTE (CLV)", "aimp": "📋 KPI AUTORISATION D'IMPORTATION MATIÈRES PREMIÈRES"}[key]
        self.banner(ws, title, SUBTITLE)
        self.section(ws, "A4:D4", "📊 INDICATEURS CLÉS")
        ids = _rng(raw, idc)
        sts = [(s[0], s[1]) for s in R.STATUSES[key]]
        n = len(sts)
        als = R.status_aliases(key)
        t0 = 15; tl = t0 + n - 1; rA = tl + 1; rT = tl + 2
        cat = f"$D${t0}:$D${rA}"; num = f"$B${t0}:$B${rA}"
        tot = mod["total"]
        ind = [
            ("Total dossiers", self.rawf(f"=COUNTA({ids})", tot), "#,##0", tot),
            ("Délivrés", f'=SUMIF({cat},"Délivré",{num})', "#,##0", mod["delivered"]),
            ("En cours de traitement", f'=SUMIF({cat},"En cours",{num})+SUMIF({cat},"Autre",{num})', "#,##0", mod["encours"]),
            ("Non conformes (modifiables)", f'=SUMIF({cat},"Non conforme",{num})', "#,##0", mod["nonconf"]),
            ("Refusés (Rejeté + Irrecevable)", f'=SUMIF({cat},"Refusé",{num})', "#,##0", mod["refused"]),
            ("Taux de délivrance", "=IFERROR(B6/B5,0)", "0.0%", mod["taux"]),
            ("Taux de refus", "=IFERROR(B9/B5,0)", "0.0%", mod["taux_refus"]),
        ]
        fills = [LGREY, GREEN, "E4EEF9", AMBER, RED, GREEN, RED]
        for k, (lab, f, nf, v) in enumerate(ind):
            r = 5 + k
            self.put(ws, f"A{r}", lab, F(10, True, GREY), fill(LGREY), LEFT, BOX)
            self.put(ws, f"B{r}", f, F(12, True, NAVY), fill(fills[k]), RIGHT, BOX, nf, cv=v)
            ws.row_dimensions[r].height = 20
        self.section(ws, "A13:D13", "📈 RÉPARTITION PAR STATUT (ordre de traitement)")
        self.header(ws, "A14:D14", {"A14": "Statut", "B14": "Nombre", "C14": "%", "D14": "Catégorie"})
        sr = _rng(raw, stc)
        rows = mod["statuses"]
        for k, (s, c) in enumerate(sts):
            r = t0 + k
            self.rowmap[(name, s)] = r
            self.put(ws, f"A{r}", s, F(10), fill(CAT_FILL[c]), LEFT, BOX)
            f = f"=COUNTIF({sr},A{r})" + "".join(f'+COUNTIF({sr},"{a}")' for a in als.get(s, []))
            self.put(ws, f"B{r}", self.rawf(f, rows[k]["count"]), F(10, True), None, RIGHT, BOX, "#,##0", cv=rows[k]["count"])
            self.put(ws, f"C{r}", f"=IFERROR(B{r}/$B${rT},0)", F(10), None, RIGHT, BOX, "0.0%", cv=rows[k]["pct"])
            self.put(ws, f"D{r}", c, F(9, False, GREY), fill(CAT_FILL[c]), CENTER, BOX)
        ra = rows[-1]
        self.put(ws, f"A{rA}", "Autres statuts / statut vide", F(10, False, GREY, True), fill(CAT_FILL["Autre"]), LEFT, BOX)
        self.put(ws, f"B{rA}", f"=$B$5-SUM(B{t0}:B{tl})", F(10, True), None, RIGHT, BOX, "#,##0", cv=ra["count"])
        self.put(ws, f"C{rA}", f"=IFERROR(B{rA}/$B${rT},0)", F(10), None, RIGHT, BOX, "0.0%", cv=ra["pct"])
        self.put(ws, f"D{rA}", "Autre", F(9, False, GREY), fill(CAT_FILL["Autre"]), CENTER, BOX)
        self.rowmap[(name, "Autres")] = rA
        self.put(ws, f"A{rT}", "TOTAL", F(10, True), fill("D9E2F3"), LEFT, BOX)
        self.put(ws, f"B{rT}", f"=SUM(B{t0}:B{rA})", F(10, True), fill("D9E2F3"), RIGHT, BOX, "#,##0", cv=mod["status_total"])
        self.put(ws, f"C{rT}", f"=IFERROR(SUM(C{t0}:C{rA}),0)", F(10, True), fill("D9E2F3"), RIGHT, BOX, "0.0%", cv=_xl_sum(x["pct"] for x in rows))
        self.put(ws, f"D{rT}", f'=IF(B{rT}=B5,"✔ Contrôle OK","✖ Écart")', F(10, True), fill("D9E2F3"), CENTER, BOX,
                 cv="✔ Contrôle OK" if mod["control_ok"] else "✖ Écart")
        ws.conditional_formatting.add(f"C{t0}:C{rA}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="6FA8DC", showValue=True))
        ws.conditional_formatting.add(f"B{t0}:B{rA}", FormulaRule(formula=[f"B{t0}=0"], font=Font(name="Calibri", color="A6A6A6")))
        empty_msg = f"⚠ La feuille « {raw} » ne contient aucune donnée : indicateurs à 0 / N/D."
        self.put(ws, f"A{rT + 2}", f'=IF($B$5=0,"{empty_msg}","")', F(10, True, "B03A2E"), cv=empty_msg if tot == 0 else "")
        r = rT + 4
        self.put(ws, f"A{r}", "📌 LÉGENDE DES CATÉGORIES", F(10, True, NAVY))
        for k, c in enumerate(R.CATEGORIES):
            self.put(ws, f"A{r + 1 + k}", c, F(10), fill(CAT_FILL[c]), LEFT, BOX)
        if any(als.values()):
            self.put(ws, f"A{r + 8}", "Coquilles de saisie regroupées : " + " ; ".join(f"« {a} » → {k}" for k, v in als.items() for a in v) + ".", F(8, False, GREY, True))
        src = f"Source : feuille « {raw} » (col. {stc} = statut, col. {idc} = identifiant, lignes {R.FIRST_ROW}–{R.LAST_ROW})."
        if self.light:
            src = src[:-1] + " — version allégée : valeurs figées à la génération."
        self.put(ws, f"A{r + 7}", src, F(8, False, GREY, True))
        self.info[name] = dict(t0=t0, tl=tl, rA=rA, rT=rT, n=n, sts=sts)
        colors = pt_colors([c for _, c in sts])
        self.make_pie(ws, "doughnut", "Répartition par statut", Reference(ws, min_col=1, min_row=t0, max_row=tl),
                      Reference(ws, min_col=2, min_row=t0, max_row=tl), colors, "F4")
        self.make_bar(ws, "Dossiers par statut", Reference(ws, min_col=1, min_row=t0, max_row=tl),
                      Reference(ws, min_col=2, min_row=t0, max_row=tl), colors, f"F{t0 + 4}")
        ws.freeze_panes = "A4"
        return ws

    # -------------------------------------------------------------- KPI Identification
    def identification_sheet(self):
        name = "KPI Identification"
        mod = self.c["modules"]["identification"]
        cols = mod["cols"]
        self.cur.update(THEMES[name])
        ws = self.new_sheet(name, self.cur["title"])
        self.banner(ws, "📋 KPI IDENTIFICATION OPÉRATEUR", SUBTITLE)
        self.section(ws, "A4:D4", "📊 INDICATEURS CLÉS")
        raw = R.MOD["identification"]["sheet"]
        ids, sts, typ = _rng(raw, cols["id"]), _rng(raw, cols["status"]), _rng(raw, cols["type"])
        for k, (lab, f, nf, fl, v) in enumerate([
                ("Total opérateurs", self.rawf(f"=COUNTA({ids})", mod["total"]), "#,##0", LGREY, mod["total"]),
                ("Opérateurs identifiés", self.rawf(f'=COUNTIF({sts},"{R.IDENT_STATUS}")', mod["delivered"]), "#,##0", GREEN, mod["delivered"]),
                ("En cours / autres statuts", "=B5-B6", "#,##0", "E4EEF9", mod["encours"]),
                ("Taux de réussite (identifiés)", "=IFERROR(B6/B5,0)", "0.0%", GREEN, mod["taux"])]):
            r = 5 + k
            self.put(ws, f"A{r}", lab, F(10, True, GREY), fill(LGREY), LEFT, BOX)
            self.put(ws, f"B{r}", f, F(12, True, NAVY), fill(fl), RIGHT, BOX, nf, cv=v)
            ws.row_dimensions[r].height = 20
        self.section(ws, "A10:D10", "📈 RÉPARTITION PAR STATUT")
        self.header(ws, "A11:D11", {"A11": "Statut", "B11": "Nombre", "C11": "%", "D11": "Catégorie"})
        st = mod["statuses"]
        for r, (lab, f, c, cl), row in zip([12, 13], [(R.IDENT_STATUS, "=B6", "Délivré", "Validé"),
                                                       ("Autres statuts / statut vide", "=B7", "Autre", "Autre")], st):
            self.put(ws, f"A{r}", lab, F(10), fill(CAT_FILL[c]), LEFT, BOX)
            self.put(ws, f"B{r}", f, F(10, True), None, RIGHT, BOX, "#,##0", cv=row["count"])
            self.put(ws, f"C{r}", f"=IFERROR(B{r}/$B$14,0)", F(10), None, RIGHT, BOX, "0.0%", cv=row["pct"])
            self.put(ws, f"D{r}", cl, F(9, False, GREY), fill(CAT_FILL[c]), CENTER, BOX)
        self.put(ws, "A14", "TOTAL", F(10, True), fill("D9E2F3"), LEFT, BOX)
        self.put(ws, "B14", "=SUM(B12:B13)", F(10, True), fill("D9E2F3"), RIGHT, BOX, "#,##0", cv=mod["status_total"])
        self.put(ws, "C14", "=IFERROR(SUM(C12:C13),0)", F(10, True), fill("D9E2F3"), RIGHT, BOX, "0.0%", cv=_xl_sum(x["pct"] for x in st))
        self.put(ws, "D14", '=IF(B14=B5,"✔ Contrôle OK","✖ Écart")', F(10, True), fill("D9E2F3"), CENTER, BOX,
                 cv="✔ Contrôle OK" if mod["control_ok"] else "✖ Écart")
        self.section(ws, "A16:D16", "📈 RÉPARTITION PAR TYPE D'OPÉRATEUR")
        self.header(ws, "A17:D17", {"A17": "Type d'opérateur", "B17": "Nombre", "C17": "%", "D17": "Part"})
        types = mod["types"]
        for k, t in enumerate(types[:4]):
            r = 18 + k
            self.put(ws, f"A{r}", t["label"], F(10), None, LEFT, BOX)
            self.put(ws, f"B{r}", self.rawf(f'=COUNTIF({typ},"{t["prefix"]}*")', t["count"]), F(10, True), None, RIGHT, BOX, "#,##0", cv=t["count"])
            self.put(ws, f"C{r}", f"=IFERROR(B{r}/$B$23,0)", F(10), None, RIGHT, BOX, "0.0%", cv=t["pct"])
            self.put(ws, f"D{r}", f"=C{r}", F(9, False, GREY), None, CENTER, BOX, "0.0%", cv=t["pct"])
        t = types[4]
        self.put(ws, "A22", "Autres types / vide", F(10, False, GREY, True), fill("ECEFF3"), LEFT, BOX)
        self.put(ws, "B22", "=$B$5-SUM(B18:B21)", F(10, True), None, RIGHT, BOX, "#,##0", cv=t["count"])
        self.put(ws, "C22", "=IFERROR(B22/$B$23,0)", F(10), None, RIGHT, BOX, "0.0%", cv=t["pct"])
        self.put(ws, "D22", "=C22", F(9, False, GREY), None, CENTER, BOX, "0.0%", cv=t["pct"])
        self.put(ws, "A23", "TOTAL", F(10, True), fill("D9E2F3"), LEFT, BOX)
        self.put(ws, "B23", "=SUM(B18:B22)", F(10, True), fill("D9E2F3"), RIGHT, BOX, "#,##0", cv=mod["types_total"])
        self.put(ws, "C23", "=IFERROR(SUM(C18:C22),0)", F(10, True), fill("D9E2F3"), RIGHT, BOX, "0.0%", cv=_xl_sum(x["pct"] for x in types))
        self.put(ws, "D23", '=IF(B23=B5,"✔ Contrôle OK","✖ Écart")', F(10, True), fill("D9E2F3"), CENTER, BOX,
                 cv="✔ Contrôle OK" if mod["types_total"] == mod["total"] else "✖ Écart")
        for rr in ("C12:C13", "C18:C22"):
            ws.conditional_formatting.add(rr, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="6FA8DC", showValue=True))
        src = (f"Source : feuille « {raw} » (col. {cols['id']} = ID, {cols['type']} = type, {cols['status']} = statut, "
               f"lignes {R.FIRST_ROW}–{R.LAST_ROW}). Types mutuellement exclusifs (préfixes).")
        if self.light:
            src += " Version allégée : valeurs figées à la génération."
        self.put(ws, "A25", src, F(8, False, GREY, True))
        self.make_pie(ws, "pie", "Opérateurs par type", Reference(ws, min_col=1, min_row=18, max_row=21), Reference(ws, min_col=2, min_row=18, max_row=21), TYPE_PAL, "F4")
        self.make_bar(ws, "Opérateurs par type", Reference(ws, min_col=1, min_row=18, max_row=21), Reference(ws, min_col=2, min_row=18, max_row=21), TYPE_PAL, "F21", h=7.5)
        ws.freeze_panes = "A4"
        return ws

    # -------------------------------------------------------------- Dashboard
    def dashboard(self):
        c, M = self.c, self.c["modules"]
        self.cur.update(THEMES["Dashboard"])
        ws = self.new_sheet("Dashboard", self.cur["title"])
        for col, w in zip("ABCDEFGHIJKL", [36, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]):
            ws.column_dimensions[col].width = w
        self.banner(ws, "📊 TABLEAU DE BORD - REVUE DE DIRECTION", "Plateforme Tarkhiss - Produits Cosmétiques et d'Hygiène Corporelle")
        ws.sheet_view.zoomScale = 90
        self.section(ws, "A4:L4", "🎯 SYNTHÈSE EXÉCUTIVE")
        self.header(ws, "A5:I5", {"A5": "Module", "B5": "Total dossiers", "C5": "Délivrés / Validés", "D5": "Taux réussite",
                                  "E5": "En cours*", "F5": "Bloqués / Refusés", "G5": "Seuil 🟢 ≥", "H5": "Seuil 🟡 ≥", "I5": "Voyant"})
        ws.row_dimensions[5].height = 30
        for k, syn in enumerate(c["synthesis"]):
            r = 6 + k
            sh = R.MOD[syn["key"]]["kpi"]
            q = f"'{sh}'!"
            self.put(ws, f"A{r}", syn["module"], F(10, True), fill(LGREY), LEFT, BOX)
            self.put(ws, f"B{r}", f"={q}B5", F(10), None, RIGHT, BOX, "#,##0", cv=syn["total"])
            self.put(ws, f"C{r}", f"={q}B6", F(10), None, RIGHT, BOX, "#,##0", cv=syn["delivered"])
            self.put(ws, f"D{r}", f"=IF(B{r}>0,C{r}/B{r},0)", F(10, True), None, RIGHT, BOX, "0.0%", cv=syn["taux"])
            self.put(ws, f"E{r}", f"={q}B7", F(10), None, RIGHT, BOX, "#,##0", cv=syn["encours"])
            if syn["key"] == "identification":
                self.put(ws, f"F{r}", "–", F(10, False, GREY), None, RIGHT, BOX)
            else:
                self.put(ws, f"F{r}", f"={q}B8+{q}B9", F(10), None, RIGHT, BOX, "#,##0", cv=syn["bloques"])
            self.put(ws, f"G{r}", syn["seuil_vert"], F(10, False, "0000FF"), fill("FFFBE6"), RIGHT, BOX, "0%")
            self.put(ws, f"H{r}", syn["seuil_orange"], F(10, False, "0000FF"), fill("FFFBE6"), RIGHT, BOX, "0%")
            self.put(ws, f"I{r}", f'=IF(B{r}=0,"⚪ N/D",IF(D{r}>=G{r},"🟢",IF(D{r}>=H{r},"🟡","🔴")))', F(12), None, CENTER, BOX, cv=syn["voyant"])
            ws.row_dimensions[r].height = 21
        for col in "ABCDEFGHI":
            self.put(ws, f"{col}11", None, F(10, True), fill("D9E2F3"), RIGHT if col != "A" else LEFT, BOX)
        t = c["totals"]
        ws["A11"].value = "TOTAL"
        for col, v in zip("BCEF", (t["total"], t["delivered"], t["encours"], t["bloques"])):
            ws[f"{col}11"].value = f"=SUM({col}6:{col}10)"
            ws[f"{col}11"].number_format = "#,##0"
            self.cache[(ws.title, f"{col}11")] = v
        ws["D11"].value = "=IF(B11>0,C11/B11,0)"
        ws["D11"].number_format = "0.0%"
        self.cache[(ws.title, "D11")] = t["taux"]
        ws.conditional_formatting.add("D6:D10", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B", showValue=True))
        note = "* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté. Seuils en bleu = paramètres modifiables."
        ce_empty = "⚠ Feuille « certificatEnregistrement » vide : CE en N/D (aucune donnée). "
        self.put(ws, "A12", f'=IF(B8=0,"{ce_empty}","")&"{note}"', F(8, False, GREY, True),
                 cv=(ce_empty if M["ce"]["total"] == 0 else "") + note)
        self.cur.update(sec="7030A0", secfont="FFFFFF")
        self.section(ws, "A14:L14", "📋 RÉSUMÉ DÉCISIONNEL - CHIFFRES CLÉS")
        cards = [("A", "B", "=C6", GREEN), ("C", "D", "=C8", GREEN), ("E", "F", "=C9", GREEN), ("G", "H", "=C10", GREEN),
                 ("I", "J", "=E11", "E4EEF9"), ("K", "L", "=F11", RED)]
        for (a, b, f, col), card in zip(cards, c["cards"]):
            self.merge(ws, f"{a}15:{b}15", card["label"], F(8, True, GREY), fill(LGREY), CENTER, BOX)
            self.merge(ws, f"{a}16:{b}16", f, F(28, True, "375623" if col == GREEN else "1F4E79"), fill(col), CENTER, BOX, nf="#,##0", cv=card["value"])
        ws.row_dimensions[15].height = 22
        ws.row_dimensions[16].height = 40
        self.cur.update(sec="C00000", secfont="FFFFFF", hdr="FFC000")
        self.section(ws, "A18:L18", "⚠️ POINTS D'ALERTE - ACTIONS REQUISES")
        self.header(ws, "A19:G19", {"A19": "Indicateur", "B19": "Module", "C19": "Valeur actuelle", "D19": "Seuil attention",
                                    "E19": "Seuil critique", "F19": "Statut", "G19": "Priorité"})
        ws.row_dimensions[19].height = 30
        dash_row = {k: 6 + i for i, k in enumerate(m["key"] for m in R.MODULES)}
        for k, al in enumerate(c["alerts"]):
            r = 20 + k
            kpi = R.MOD[al["key"]]["kpi"]
            f = f"=D{dash_row[al['key']]}" if al["type"] == "taux" else f"='{kpi}'!B{self.rowmap[(kpi, al['status_name'])]}"
            tot = f"B{dash_row[al['key']]}"
            pctf = al["type"] == "taux"
            self.put(ws, f"A{r}", al["indicator"], F(10, True), fill(LGREY), LEFT, BOX)
            self.put(ws, f"B{r}", al["module"], F(9), None, CENTER, BOX)
            self.put(ws, f"C{r}", f, F(10, True), None, RIGHT, BOX, "0.0%" if pctf else "#,##0", cv=al["value"])
            self.put(ws, f"D{r}", al["warn"], F(10, False, "0000FF"), fill("FFFBE6"), RIGHT, BOX, "0%" if pctf else "#,##0")
            self.put(ws, f"E{r}", al["crit"], F(10, False, "0000FF"), fill("FFFBE6"), RIGHT, BOX, "0%" if pctf else "#,##0")
            op = "<" if pctf else ">"
            cond = f'IF(C{r}{op}E{r},"🔴 CRITIQUE",IF(C{r}{op}D{r},"🟡 ATTENTION","🟢 OK"))'
            self.put(ws, f"F{r}", f'=IF({tot}=0,"⚪ N/D",{cond})', F(10, True), None, CENTER, BOX, cv=al["status"])
            self.put(ws, f"G{r}", f'=IF(ISNUMBER(SEARCH("CRITIQUE",F{r})),"HAUTE",IF(ISNUMBER(SEARCH("ATTENTION",F{r})),"MOYENNE",IF(ISNUMBER(SEARCH("OK",F{r})),"BASSE","—")))',
                     F(10, True), None, CENTER, BOX, cv=al["priority"])
            ws.row_dimensions[r].height = 19
        aL = 20 + len(c["alerts"]) - 1
        for txt, colr in (("CRITIQUE", RED), ("ATTENTION", AMBER), ("OK", GREEN)):
            ws.conditional_formatting.add(f"F20:G{aL}", FormulaRule(formula=[f'ISNUMBER(SEARCH("{txt}",$F20))'], fill=fill(colr)))
        self.put(ws, f"A{aL + 1}", "Seuils : « taux » → alerte si la valeur est INFÉRIEURE au seuil ; « nombre » → SUPÉRIEURE. Seuils en bleu = paramètres (valeurs d’origine du fichier ou hypothèses à valider).", F(8, False, GREY, True))

        cr = aL + 3
        self.cur.update(sec="BDD7EE", secfont="1F4E79", hdr="2F5496")
        self.section(ws, f"A{cr}:L{cr}", "📊 GRAPHIQUES")
        c1 = Reference(ws, min_col=1, min_row=6, max_row=10)
        ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.style = 10
        ch.title = "Dossiers par module : délivrés / en cours / bloqués"
        for col, nm, colr in ((3, "Délivrés", "2E8B57"), (5, "En cours", "4FA3D9"), (6, "Bloqués / Refusés", "D64541")):
            s = Series(Reference(ws, min_col=col, min_row=6, max_row=10), title=nm)
            s.graphicalProperties.solidFill = colr
            ch.series.append(s)
        ch.set_categories(c1); ch.legend.position = "b"; ch.x_axis.delete = False; ch.y_axis.delete = False
        ch.y_axis.numFmt = "#,##0"; ch.width, ch.height = 17, 8.5
        ws.add_chart(ch, f"A{cr + 2}")
        tx = BarChart(); tx.type = "col"; tx.style = 10; tx.title = "Taux de réussite par module"
        tx.add_data(Reference(ws, min_col=4, min_row=6, max_row=10), titles_from_data=False); tx.set_categories(c1)
        s = tx.series[0]
        for i, colr in enumerate(["1F6FB2", "2E8B57", "8E6BBF", "F4B400", "D64541"]):
            p = DataPoint(idx=i); p.graphicalProperties.solidFill = colr; s.dPt.append(p)
        s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.numFmt = "0%"
        s.dLbls.showCatName = False; s.dLbls.showSerName = False; s.dLbls.showPercent = False; s.dLbls.showLegendKey = False
        tx.legend = None; tx.x_axis.delete = False; tx.y_axis.delete = False; tx.y_axis.numFmt = "0%"
        tx.y_axis.scaling.min = 0; tx.y_axis.scaling.max = 1; tx.width, tx.height = 17.5, 8.5
        ws.add_chart(tx, f"F{cr + 2}")

        def kref(sh, col):
            i = self.info[sh]
            return Reference(self.wb[sh], min_col=col, min_row=i["t0"], max_row=i["tl"])

        def kcolors(sh):
            return pt_colors([cc for _, cc in self.info[sh]["sts"]])
        r2 = cr + 2 + 18
        self.make_pie(ws, "doughnut", "Déclarations d'activité par statut", kref("KPI Déclaration", 1), kref("KPI Déclaration", 2), kcolors("KPI Déclaration"), f"A{r2}", w=17, h=8.5)
        self.make_pie(ws, "doughnut", "Certificats Libre Vente par statut", kref("KPI Certificat LV", 1), kref("KPI Certificat LV", 2), kcolors("KPI Certificat LV"), f"F{r2}", w=17.5, h=8.5)
        r3 = r2 + 18
        self.make_bar(ws, "Autorisations d’importation MP par statut", kref("KPI Autor. Import.", 1), kref("KPI Autor. Import.", 2), kcolors("KPI Autor. Import."), f"A{r3}", w=17, h=9)
        ki = self.wb["KPI Identification"]
        self.make_pie(ws, "pie", "Opérateurs par type", Reference(ki, min_col=1, min_row=18, max_row=21), Reference(ki, min_col=2, min_row=18, max_row=21), TYPE_PAL, f"F{r3}", w=17.5, h=9)

        dr = r3 + 20
        self.section(ws, f"A{dr}:L{dr}", "📈 DÉTAIL PAR STATUT (liens vers les feuilles KPI)")

        def detail(r0, a, b, numcol, pctcol, title, sh, rows_, vals):
            self.merge(ws, f"{a}{r0}:{pctcol}{r0}", title, F(10, True, "FFFFFF"), fill("4472C4"), CENTER, BOX)
            self.merge(ws, f"{a}{r0 + 1}:{b}{r0 + 1}", "Statut", F(9, True, NAVY), fill("D6DCE4"), CENTER, BOX)
            self.put(ws, f"{numcol}{r0 + 1}", "Nombre", F(9, True, NAVY), fill("D6DCE4"), CENTER, BOX)
            self.put(ws, f"{pctcol}{r0 + 1}", "%", F(9, True, NAVY), fill("D6DCE4"), CENTER, BOX)
            for k, (kr, (lab, cnt, pct)) in enumerate(zip(rows_, vals)):
                r = r0 + 2 + k
                q = f"'{sh}'!"
                if b != a and len(lab) > 30:
                    ws.row_dimensions[r].height = 22.35  # libellé long sur 2 lignes (comme la référence)
                self.merge(ws, f"{a}{r}:{b}{r}", f"={q}A{kr}", F(9), None, LEFT, BOX, cv=lab)
                self.put(ws, f"{numcol}{r}", f"={q}B{kr}", F(9, True), None, RIGHT, BOX, "#,##0", cv=cnt)
                self.put(ws, f"{pctcol}{r}", f"={q}C{kr}", F(9), None, RIGHT, BOX, "0.0%", cv=pct)
            ws.conditional_formatting.add(f"{pctcol}{r0 + 2}:{pctcol}{r0 + 1 + len(rows_)}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="6FA8DC", showValue=True))
            return r0 + 2 + len(rows_)

        def rows_for(key):
            i = self.info[R.MOD[key]["kpi"]]
            vals = [(x["status"], x["count"], x["pct"]) for x in M[key]["statuses"]]
            return list(range(i["t0"], i["rA"] + 1)), vals
        d0 = dr + 2
        e1 = detail(d0, "A", "A", "B", "C", "DÉCLARATION D’ACTIVITÉ", "KPI Déclaration", *rows_for("declaration"))
        e2 = detail(d0, "E", "F", "G", "H", "AIMP", "KPI Autor. Import.", *rows_for("aimp"))
        e3 = detail(d0, "I", "J", "K", "L", "CLV", "KPI Certificat LV", *rows_for("clv"))
        d1 = max(e1, e2, e3) + 1
        e4 = detail(d1, "A", "A", "B", "C", "CERTIFICAT D'ENREGISTREMENT (CE)", "KPI Certificat Enreg.", *rows_for("ce"))
        tvals = [(x["label"], x["count"], x["pct"]) for x in M["identification"]["types"]]
        e5 = detail(d1, "E", "F", "G", "H", "TYPES D'OPÉRATEURS", "KPI Identification", list(range(18, 23)), tvals)
        self.put(ws, f"A{max(e4, e5) + 1}", f"Sources : feuilles KPI (calculées depuis les feuilles brutes). Dernière actualisation : {self.update.strftime('%d/%m/%Y')}.", F(8, False, GREY, True))
        ws.freeze_panes = "A4"
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        return ws

    # -------------------------------------------------------------- feuilles brutes
    def raw_sheet(self, m):
        ws = self.wb.create_sheet(m["sheet"])
        ws.sheet_properties.tabColor = RAW_TAB
        data = self.raw.get(m["key"])
        if not data:
            return ws
        for r, row in enumerate(data.get("header_prefix") or [], start=1):
            for ci_, v in enumerate(row, start=1):
                v = R.from_json_value(v)
                if v is not None:
                    ws.cell(r, ci_, v)
        for ci_, h in enumerate(data["headers"], start=1):
            if h != "":
                ws.cell(R.HEADER_ROW, ci_, h)
        for r, row in enumerate(data["rows"], start=R.FIRST_ROW):
            for ci_, v in enumerate(row, start=1):
                if v is not None:
                    ws.cell(r, ci_, R.from_json_value(v))
        ncols = len(data["headers"])
        for col, w in (data.get("widths") or {}).items():
            if R.col_index(col) < ncols:
                ws.column_dimensions[col].width = w
        return ws

    def build_report_sheets(self):
        # Les feuilles KPI d'abord (le Dashboard y fait référence), puis le Dashboard placé en tête.
        self.identification_sheet()
        for key in ("declaration", "ce", "clv", "aimp"):
            self.status_sheet(key)
        ws = self.dashboard()
        self.wb.move_sheet(ws, offset=-(len(self.wb.sheetnames) - 1))

    def cell_values(self):
        """{feuille: {cellule: ("f" formule | "c" constante, valeur)}} des 6 feuilles de restitution."""
        self.build_report_sheets()
        out = {}
        for ws in self.wb.worksheets:
            cells = {}
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and v.startswith("="):
                        cells[c.coordinate] = ("f", self.cache[(ws.title, c.coordinate)])
                    elif v is not None:
                        cells[c.coordinate] = ("c", v)
            out[ws.title] = cells
        return out

    def build(self):
        self.build_report_sheets()
        if not self.light:
            for m in R.MODULES:
                self.raw_sheet(m)
        self.wb.active = 0
        for s in self.wb.worksheets:
            s.sheet_view.tabSelected = (s.title == "Dashboard")
        self.wb.calculation.fullCalcOnLoad = True
        buf = io.BytesIO()
        self.wb.save(buf)
        return inject_cached_values(buf.getvalue(), self.wb.sheetnames, self.cache)


TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pchc_template")
TEMPLATE_XLSX = os.path.join(TEMPLATE_DIR, "dashboard_gabarit.xlsx")
TEMPLATE_META = os.path.join(TEMPLATE_DIR, "dashboard_gabarit.json")


def build_xlsx(raw, comp, update, light=False):
    """Construit le classeur. raw = feuilles brutes stockées, comp = pchc_compute(raw),
    update = date « Mise à jour » (datetime). Retourne les octets .xlsx.
    Gabarit (copie conforme de la référence) si présent, sinon reconstruction openpyxl."""
    if not isinstance(update, datetime):
        update = datetime(update.year, update.month, update.day)
    if os.path.isfile(TEMPLATE_XLSX) and os.path.isfile(TEMPLATE_META):
        return _TemplateWriter(raw, comp, update, light).build()
    return _Builder(raw, comp, update, light).build()


# ----------------------------------------------------------------------------- valeurs en cache
def _xl_num(v):
    """Nombre tel qu'Excel l'écrit en cache : 17 chiffres significatifs (ex. 0.99999999999999989),
    notation scientifique sous 0,1 (ex. 8.641202254226675E-2)."""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int) or (isinstance(v, float) and v.is_integer() and abs(v) < 1e15):
        return str(int(v))
    if 0 < abs(v) < 0.1:
        mant, exp = format(v, ".16e").split("e")
        return f"{mant.rstrip('0').rstrip('.')}E{int(exp)}"
    return format(v, ".17g")


def _xl_serial(d):
    return (d - datetime(1899, 12, 30)).days + (d - datetime(d.year, d.month, d.day)).seconds / 86400


_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xml_text(s):
    return xml_escape(_XML_BAD.sub("", str(s)))


def _fmt_cached(v):
    if isinstance(v, bool):
        return ' t="b"', "1" if v else "0"
    if isinstance(v, (int, float)):
        return "", _xl_num(v)
    return ' t="str"', _xml_text(v)


# ----------------------------------------------------------------------------- gabarit (copie conforme)
_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)
_FORMULA = re.compile(r"(<f[^>]*/>|<f[^>]*>.*?</f>)", re.S)


class _SharedStrings:
    """Table des chaînes partagées : celles du gabarit (feuilles de restitution) + ajouts."""

    def __init__(self, xml):
        self.head = xml[:xml.index("<sst")]
        self.tag = re.search(r"<sst[^>]*>", xml).group(0)
        self.items = re.findall(r"<si>.*?</si>", xml, re.S)
        self.text = ["".join(xml_unescape(t) for t in re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)) for si in self.items]
        self.index = {}
        for i, t in enumerate(self.text):
            self.index.setdefault(t, i)

    def add(self, s):
        s = _XML_BAD.sub("", str(s))
        i = self.index.get(s)
        if i is None:
            sp = ' xml:space="preserve"' if s != s.strip() or "\n" in s else ""
            i = len(self.items)
            self.items.append(f"<si><t{sp}>{xml_escape(s)}</t></si>")
            self.text.append(s)
            self.index[s] = i
        return i

    def xml(self, count):
        tag = re.sub(r' count="\d+"', f' count="{count}"', self.tag)
        tag = re.sub(r' uniqueCount="\d+"', f' uniqueCount="{len(self.items)}"', tag)
        return self.head + tag + "".join(self.items) + "</sst>"


class _TemplateWriter:
    def __init__(self, raw, comp, update, light):
        self.raw, self.comp, self.update, self.light = raw, comp, update, light
        with open(TEMPLATE_META, encoding="utf-8") as f:
            self.meta = json.load(f)
        self.zin = zipfile.ZipFile(TEMPLATE_XLSX)
        self.sst = _SharedStrings(self.zin.read("xl/sharedStrings.xml").decode("utf-8"))
        self.values = _Builder(raw, comp, update, light).cell_values()
        self.raw_names = [m["sheet"] for m in R.MODULES]

    # Lettres de colonnes détectées à l'import ≠ lettres du gabarit → formules réécrites
    def _column_substitutions(self):
        subs = []
        for m in R.MODULES:
            cols = self.comp["modules"][m["key"]]["cols"]
            for field, default in m["cols"].items():
                new = cols.get(field, default)
                if new != default:
                    subs.append((f"{m['sheet']}!${default}${R.FIRST_ROW}:${default}${R.LAST_ROW}",
                                 f"{m['sheet']}!${new}${R.FIRST_ROW}:${new}${R.LAST_ROW}"))
        return subs

    def _template_value(self, attrs, inner):
        v = re.search(r"<v>(.*?)</v>", inner or "", re.S)
        if not v:
            return None
        if ' t="s"' in attrs:
            return self.sst.text[int(v.group(1))]
        if ' t="str"' in attrs or ' t="inlineStr"' in attrs:
            return xml_unescape(v.group(1))
        try:
            return float(v.group(1))
        except ValueError:
            return v.group(1)

    def _patch_sheet(self, title, xml):
        vals = self.values[title]
        seen, removed_si = set(), set()

        def repl(m):
            ref, attrs, inner = m.group(1) + m.group(2), m.group(3), m.group(5)
            seen.add(ref)
            if ref not in vals:
                return m.group(0)
            kind, v = vals[ref]
            fm = _FORMULA.match(inner or "")
            base = re.sub(r' t="[^"]*"', "", attrs)
            if kind == "f":
                if not fm:
                    raise ValueError(f"Gabarit : formule attendue en {title}!{ref}")
                t, txt = _fmt_cached(v)
                return f'<c r="{ref}"{base}{t}>{fm.group(1)}{"<v>" + txt + "</v>" if txt != "" else "<v/>"}</c>'
            if isinstance(v, datetime):
                v = _xl_serial(v)
            if not fm and self._template_value(attrs, inner) == (float(v) if isinstance(v, (int, float)) else v):
                return m.group(0)   # constante identique : cellule conservée à l'octet près
            if fm:   # version allégée : la formule lisait une feuille brute → valeur figée
                si = re.search(r'si="(\d+)"', fm.group(1))
                if si and "ref=" in fm.group(1):
                    removed_si.add(si.group(1))
            if isinstance(v, str):
                return f'<c r="{ref}"{base} t="s"><v>{self.sst.add(v)}</v></c>'
            return f'<c r="{ref}"{base}><v>{_xl_num(v)}</v></c>'

        out = _CELL.sub(repl, xml)
        missing = set(vals) - seen
        if missing:
            raise ValueError(f"Gabarit : cellules absentes dans {title} : {sorted(missing)[:10]}")
        for si in removed_si:
            if re.search(rf'<f t="shared" si="{si}"\s*/>', out):
                raise ValueError(f"Gabarit : formule partagée orpheline (si={si}) dans {title}")
        for old, new in self._column_substitutions():
            out = out.replace(old, new)
        return out

    def _raw_sheet(self, name, xml):
        key = next(m["key"] for m in R.MODULES if m["sheet"] == name)
        data = self.raw.get(key) or {}
        info = self.meta["raw"].get(name, {})
        styles, row_attrs = info.get("styles", {}), info.get("row_attrs", "")
        lines = [(r, row) for r, row in enumerate(data.get("header_prefix") or [], start=1)]
        lines.append((R.HEADER_ROW, [h if h != "" else None for h in data.get("headers", [])]))
        lines += [(r, row) for r, row in enumerate(data.get("rows", []), start=R.FIRST_ROW)]
        ncols = max([len(row) for _, row in lines] + [1])
        letters = [R.col_letter(i) for i in range(ncols)]
        parts, first, last = [], None, None
        for r, row in lines:
            cells = []
            for i, v in enumerate(row):
                v = R.from_json_value(v)
                if v is None or v == "":
                    continue
                col = letters[i]
                s = f' s="{styles[col]}"' if styles.get(col) else ""
                if isinstance(v, datetime):
                    v = v.strftime("%Y-%m-%d %H:%M:%S")   # stockées en texte, comme dans la référence
                if isinstance(v, str):
                    cells.append(f'<c r="{col}{r}"{s} t="s"><v>{self.sst.add(v)}</v></c>')
                else:
                    cells.append(f'<c r="{col}{r}"{s}><v>{_xl_num(v)}</v></c>')
            if cells:
                parts.append(f'<row r="{r}" spans="1:{ncols}"{row_attrs}>{"".join(cells)}</row>')
                first, last = first or r, r
        dim = f"A{first}:{letters[-1]}{last}" if parts else "A1"
        xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dim}"/>', xml, count=1)
        return xml.replace("<sheetData/>", "<sheetData>" + "".join(parts) + "</sheetData>", 1)

    def _value_at(self, sheet, ref):
        cell = self.values.get(sheet, {}).get(ref)
        return cell[1] if cell else None

    def _patch_chart(self, xml):
        """Caches des séries (valeurs et catégories) recalculés depuis les cellules sources."""
        def repl(m):
            kind, formula, fmt = m.group(1), xml_unescape(m.group(2)), m.group(4) or ""
            # attributs des points du gabarit (ex. formatCode="General" posé par Excel) conservés
            pt_attrs = dict(re.findall(r'<c:pt idx="(\d+)"([^>]*)>', m.group(0)))
            sheet, rng = formula.rsplit("!", 1)
            sheet = sheet.strip("'")
            a, b = rng.replace("$", "").split(":") if ":" in rng else (rng.replace("$", ""),) * 2
            ca, ra = re.match(r"([A-Z]+)(\d+)", a).groups()
            cb, rb = re.match(r"([A-Z]+)(\d+)", b).groups()
            refs = [f"{R.col_letter(c)}{r}" for r in range(int(ra), int(rb) + 1)
                    for c in range(R.col_index(ca), R.col_index(cb) + 1)]
            vals = [self._value_at(sheet, ref) for ref in refs]
            pts = []
            for i, v in enumerate(vals):
                if v is None:
                    continue
                if kind == "num" and not isinstance(v, (int, float)):
                    # cellule texte (ex. « – ») dans une série numérique : Excel la met en cache à 0
                    pt_attrs.setdefault(str(i), ' formatCode="General"')
                    v = 0
                pts.append(f'<c:pt idx="{i}"{pt_attrs.get(str(i), "")}><c:v>{_xl_num(v) if kind == "num" else _xml_text(v)}</c:v></c:pt>')
            return (f"<c:{kind}Ref><c:f>{m.group(2)}</c:f><c:{kind}Cache>{fmt}<c:ptCount val=\"{len(refs)}\"/>"
                    f"{''.join(pts)}</c:{kind}Cache></c:{kind}Ref>")
        return re.sub(r"<c:(num|str)Ref><c:f>(.*?)</c:f><c:(num|str)Cache>(<c:formatCode>.*?</c:formatCode>)?.*?</c:\3Cache></c:\1Ref>",
                      repl, xml, flags=re.S)

    def _drop_raw_from_package(self, name, data):
        """Version allégée : retire les feuilles brutes du classeur (workbook, relations, types, app.xml)."""
        x = data.decode("utf-8")
        raw_paths = [self.meta["sheets"][n] for n in self.raw_names]
        if name == "xl/workbook.xml":
            for n in self.raw_names:
                x = re.sub(rf'<sheet name="{re.escape(xml_escape(n))}"[^>]*/>', "", x)
        elif name == "xl/_rels/workbook.xml.rels":
            for p in raw_paths:
                x = re.sub(rf'<Relationship [^>]*Target="{re.escape(p[3:])}"[^>]*/>', "", x)
            x = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", x)   # chaîne de calcul périmée → reconstruite par Excel
        elif name == "[Content_Types].xml":
            for p in raw_paths:
                x = x.replace(f'<Override PartName="/{p}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', "")
            x = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", x)
        elif name == "docProps/app.xml":
            for n in self.raw_names:
                x = x.replace(f"<vt:lpstr>{xml_escape(n)}</vt:lpstr>", "")
            x = re.sub(r"<vt:i4>\d+</vt:i4>", "<vt:i4>6</vt:i4>", x, count=1)
            x = re.sub(r'(<TitlesOfParts><vt:vector size=")\d+', r"\g<1>6", x)
        return x.encode("utf-8")

    def build(self):
        sheets = self.meta["sheets"]
        by_path = {p: n for n, p in sheets.items()}
        raw_paths = {sheets[n] for n in self.raw_names}
        parts = {}
        for info in self.zin.infolist():
            name = info.filename
            data = self.zin.read(name)
            if name in by_path and by_path[name] not in self.raw_names:
                data = self._patch_sheet(by_path[name], data.decode("utf-8")).encode("utf-8")
            elif name == "xl/calcChain.xml" and self.light:
                continue
            elif name in raw_paths:
                if self.light:
                    continue
                data = self._raw_sheet(by_path[name], data.decode("utf-8")).encode("utf-8")
            elif name.startswith("xl/charts/chart"):
                data = self._patch_chart(data.decode("utf-8")).encode("utf-8")
            elif name == "docProps/core.xml":
                now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                x = re.sub(r"(<dcterms:modified[^>]*>)[^<]*", rf"\g<1>{now}", data.decode("utf-8"))
                data = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", "<cp:lastModifiedBy>AMMPS — DSID</cp:lastModifiedBy>", x).encode("utf-8")
            if self.light and name in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml", "docProps/app.xml"):
                data = self._drop_raw_from_package(name, data)
            parts[name] = (info, data)
        count = sum(len(re.findall(rb' t="s"', d)) for n, (_, d) in parts.items() if n.startswith("xl/worksheets/sheet"))
        info, _ = parts["xl/sharedStrings.xml"]
        parts["xl/sharedStrings.xml"] = (info, self.sst.xml(count).encode("utf-8"))
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, (info, data) in parts.items():
                zout.writestr(info, data)
        return out.getvalue()


def inject_cached_values(data, sheetnames, cache):
    """Ajoute <v>valeur</v> aux cellules-formules écrites par openpyxl (qui les laisse vides)."""
    by_sheet = {}
    for (title, ref), v in cache.items():
        by_sheet.setdefault(sheetnames.index(title) + 1, {})[ref] = v
    zin = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    pat = re.compile(r'<c r="([A-Z]+[0-9]+)"((?: [a-z]+="[^"]*")*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>', re.S)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", item.filename)
            if m and int(m.group(1)) in by_sheet:
                vals = by_sheet[int(m.group(1))]

                def repl(mo):
                    ref, attrs, f = mo.group(1), mo.group(2), mo.group(3)
                    if ref not in vals or vals[ref] is None:
                        return mo.group(0)
                    t, v = _fmt_cached(vals[ref])
                    attrs = re.sub(r' t="[^"]*"', "", attrs)
                    return f'<c r="{ref}"{attrs}{t}><f>{f}</f><v>{v}</v></c>'
                content = pat.sub(repl, content.decode("utf-8")).encode("utf-8")
            zout.writestr(item, content)
    return out.getvalue()


# ----------------------------------------------------------------------------- recalcul LibreOffice
ERROR_RE = re.compile(r"^(#(DIV/0!|REF!|NAME\?|VALUE!|N/A|NUM!|NULL!)|Err:\d+)")


def find_soffice():
    cand = [os.environ.get("SANAD_SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
            r"C:\Program Files\LibreOffice\program\soffice.exe", "/usr/bin/soffice", "/opt/libreoffice/program/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    return next((c for c in cand if c and os.path.isfile(c)), None)


def recalc_check(xlsx_bytes, timeout=240):
    """Recalcule le classeur avec LibreOffice headless (si présent) puis vérifie : 0 erreur de
    formule et valeurs identiques au cache. Retourne un dict de résultat (jamais d'exception)."""
    soffice = find_soffice()
    if not soffice:
        return {"available": False, "detail": "LibreOffice non installé : valeurs en cache calculées par SANAD (même logique que les formules, vérifiée par les tests)."}
    tmp = tempfile.mkdtemp(prefix="sanad_recalc_")
    try:
        src = os.path.join(tmp, "in.xlsx")
        with open(src, "wb") as f:
            f.write(xlsx_bytes)
        outdir = os.path.join(tmp, "out")
        os.makedirs(outdir)
        subprocess.run([soffice, "--headless", "--norestore", f"-env:UserInstallation=file://{tmp}/profile",
                        "--convert-to", "xlsx", "--outdir", outdir, src], check=True, timeout=timeout,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        recalced = openpyxl.load_workbook(os.path.join(outdir, "in.xlsx"), data_only=True)
        mine = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
        errors, mismatches = [], []
        for name in mine.sheetnames:
            if name not in recalced.sheetnames or name in [m["sheet"] for m in R.MODULES]:
                continue
            a, b = mine[name], recalced[name]
            for row in a.iter_rows():
                for c in row:
                    v2 = b[c.coordinate].value
                    if isinstance(v2, str) and ERROR_RE.match(v2):
                        errors.append(f"{name}!{c.coordinate}={v2}")
                    elif c.value != v2 and not (isinstance(c.value, (int, float)) and isinstance(v2, (int, float)) and abs(c.value - v2) < 1e-9):
                        if not isinstance(c.value, datetime):
                            mismatches.append(f"{name}!{c.coordinate}: {c.value!r} ≠ {v2!r}")
        return {"available": True, "errors": errors, "mismatches": mismatches[:50], "ok": not errors and not mismatches}
    except Exception as e:
        return {"available": True, "ok": False, "errors": [f"Recalcul LibreOffice en échec : {e}"], "mismatches": []}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
