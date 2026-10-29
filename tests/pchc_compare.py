# -*- coding: utf-8 -*-
"""Comparaison cellule par cellule d'un classeur généré avec le classeur de référence
(valeurs recalculées, formules, styles, fusions, largeurs, graphiques, mises en forme conditionnelles).
Utilisé par tests/test_pchc_export.py ; exécutable seul :
    python tests/pchc_compare.py généré.xlsx reference/dashboardTarkhiss_MAJ_v3_2026-10-28.xlsx
"""
import io
import sys
import warnings
from datetime import datetime

import openpyxl

RAW = ["identificationOpérateur", "déclarationActivité", "certificatEnregistrement", "certificatLibreVente",
       "autorisationImportationMatièreP"]


def _rgb(c):
    if c is None or c.rgb is None or not isinstance(c.rgb, str):
        return None
    return c.rgb[-6:].upper()


def _style(c):
    f, a, b = c.font, c.alignment, c.border
    fl = _rgb(c.fill.fgColor) if c.fill is not None and c.fill.fill_type == "solid" else None
    sides = [getattr(s, "style", None) for s in (b.left, b.right, b.top, b.bottom)]
    # Alignement absent ≡ « general / bottom » (valeurs par défaut écrites explicitement par Excel)
    return (f.name, float(f.sz or 0), bool(f.b), bool(f.i), _rgb(f.color), fl,
            a.horizontal or "general", a.vertical or "bottom", bool(a.wrap_text), float(a.indent or 0),
            *sides, _rgb(b.bottom.color) if sides[3] else None, c.number_format)


def _widths(ws):
    """Largeurs effectives par index de colonne (gère les groupes <col min max>)."""
    out = {}
    for key, dim in ws.column_dimensions.items():
        if dim.width is None or not dim.customWidth and dim.width == 13:
            continue
        lo = dim.min or openpyxl.utils.column_index_from_string(key)
        hi = dim.max or lo
        for i in range(lo, min(hi, 60) + 1):
            out[i] = dim.width
    return out


def _same(v1, v2):
    """Écart toléré = 0 (les valeurs en cache sont écrites sur 15 chiffres significatifs, comme Excel)."""
    return v1 == v2


def compare(gen, ref, skip_raw=False, max_diffs=60):
    """gen/ref : chemins ou octets. Retourne la liste des écarts (vide = identiques)."""
    def load(x, **kw):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return openpyxl.load_workbook(io.BytesIO(x) if isinstance(x, bytes) else x, **kw)
    gv, rv = load(gen, data_only=True), load(ref, data_only=True)
    gf, rf = load(gen), load(ref)
    diffs = []

    def d(msg):
        if len(diffs) < max_diffs:
            diffs.append(msg)
    if gv.sheetnames != rv.sheetnames:
        d(f"Feuilles : {gv.sheetnames} ≠ {rv.sheetnames}")
        return diffs
    for name in rv.sheetnames:
        raw = name in RAW
        if raw and skip_raw:
            continue
        a, b, af, bf = gv[name], rv[name], gf[name], rf[name]
        maxr, maxc = max(a.max_row, b.max_row), max(a.max_column, b.max_column)
        for row in range(1, maxr + 1):
            for col in range(1, maxc + 1):
                ca, cb = a.cell(row, col), b.cell(row, col)
                if not _same(ca.value, cb.value):
                    d(f"{name}!{cb.coordinate} valeur : {ca.value!r} ≠ réf {cb.value!r}")
                if raw:
                    continue
                fa, fb = af.cell(row, col).value, bf.cell(row, col).value
                if not _same(fa, fb) and not (isinstance(fa, datetime) and isinstance(fb, datetime)):
                    d(f"{name}!{cb.coordinate} formule : {fa!r} ≠ réf {fb!r}")
                if (cb.value is not None or cb.has_style) and _style(af.cell(row, col)) != _style(bf.cell(row, col)):
                    d(f"{name}!{cb.coordinate} style : {_style(af.cell(row, col))} ≠ réf {_style(bf.cell(row, col))}")
        if {str(x) for x in af.merged_cells.ranges} != {str(x) for x in bf.merged_cells.ranges}:
            d(f"{name} fusions différentes")
        last = maxc if raw else 12
        wma, wmb = _widths(af), _widths(bf)
        for i in range(1, last + 1):
            wa, wb_ = wma.get(i), wmb.get(i)
            if raw and not wb_:
                continue
            # tolérance 0,02 : Excel arrondit les largeurs à 2 décimales à l'enregistrement
            if wa is None or wb_ is None or abs(wa - wb_) > 0.02:
                d(f"{name} largeur {openpyxl.utils.get_column_letter(i)} : {wa} ≠ réf {wb_}")
        if raw:
            continue
        # Hauteurs : Excel arrondit au pixel (0,75 pt) à l'enregistrement et écrit explicitement les
        # hauteurs automatiques (15 pt par défaut, ajustement des lignes à retour à la ligne).
        for r in range(1, b.max_row + 1):
            ha, hb = af.row_dimensions[r].height, bf.row_dimensions[r].height
            if ha is None and (hb is None or hb == 15.0):
                continue
            if ha is None or hb is None or abs(ha - hb) > 0.75:
                d(f"{name} hauteur ligne {r} : {ha} ≠ réf {hb}")
        ga = sorted((type(c).__name__, c.anchor._from.col, c.anchor._from.row) for c in af._charts)
        gb = sorted((type(c).__name__, c.anchor._from.col, c.anchor._from.row) for c in bf._charts)
        if ga != gb:
            d(f"{name} graphiques : {ga} ≠ réf {gb}")
        cfa = sorted(str(r.sqref) for r in af.conditional_formatting)
        cfb = sorted(str(r.sqref) for r in bf.conditional_formatting)
        if cfa != cfb:
            d(f"{name} mises en forme conditionnelles : {cfa} ≠ réf {cfb}")
        for attr in ("freeze_panes",):
            if getattr(af, attr) != getattr(bf, attr):
                d(f"{name} {attr} : {getattr(af, attr)} ≠ réf {getattr(bf, attr)}")
        if _rgb(af.sheet_properties.tabColor) != _rgb(bf.sheet_properties.tabColor):
            d(f"{name} couleur d'onglet différente")
        if af.sheet_view.showGridLines != bf.sheet_view.showGridLines or (af.sheet_view.zoomScale or 100) != (bf.sheet_view.zoomScale or 100):
            d(f"{name} affichage (quadrillage/zoom) différent")
    return diffs


if __name__ == "__main__":
    res = compare(sys.argv[1], sys.argv[2])
    print("\n".join(res) if res else "IDENTIQUES")
    sys.exit(1 if res else 0)
