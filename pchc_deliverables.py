# -*- coding: utf-8 -*-
"""
pchc_deliverables.py — Livrables du Reporting Métier PCHC : PPTX (16:9), PDF (A4 paysage, serveur),
résumé HTML de l'e-mail et message .eml avec les 3 livrables en pièces jointes.

Toutes les valeurs viennent de pchc_raw.pchc_compute() (aucun calcul ici) : l'Excel, le PPTX, le PDF,
l'e-mail et l'UI affichent donc exactement les mêmes chiffres.
"""
import base64
import io
import os
from datetime import datetime
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from html import escape as esc

import pchc_raw as R
from pchc_excel import CAT_FILL, CAT_PT, TYPE_PAL, pt_colors

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF_MIME = "application/pdf"
NAVY = "1F4E79"
INK, MUTED, LINE, SOFT = "1F2937", "5B6B7F", "D0D7E2", "F2F5F9"
VOYANT = {"🟢": ("2E8B57", "OK"), "🟡": ("F4B400", "Attention"), "🔴": ("D64541", "Critique"), "⚪ N/D": ("9AA5B1", "N/D")}
ALERT_FILL = {"CRITIQUE": "FFCCCC", "ATTENTION": "FFE699", "OK": "C6EFCE", "N/D": "ECEFF3"}
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"]


def period_label(d):
    return f"{MOIS[d.month - 1]} {d.year}"


def voyant_info(v):
    return VOYANT.get(v, ("9AA5B1", "N/D"))


ALERT_LABEL = {"CRITIQUE": "Critique", "ATTENTION": "Attention", "OK": "OK", "N/D": "N/D"}


def alert_key(status):
    for k in ("CRITIQUE", "ATTENTION", "OK"):
        if k in status:
            return k
    return "N/D"


def light_hex(h, f=0.85):
    """Teinte claire d'une couleur (pour les fonds de section)."""
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "%02X%02X%02X" % tuple(int(c + (255 - c) * f) for c in (r, g, b))


def section_theme(key):
    return R.DASHBOARD_THEME if key is None else R.MOD[key]["theme"]


# ----------------------------------------------------------------------------- graphiques (PNG)
def _png(fig):
    import matplotlib.pyplot as plt
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return buf.getvalue()


def _plt():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": "#B8C2CF",
                         "axes.spines.top": False, "axes.spines.right": False})
    return plt


def chart_modules(comp, size=(7.2, 3.3)):
    """Barres empilées par volet : délivrés / en cours / bloqués (comme le Dashboard Excel)."""
    plt = _plt()
    syn = comp["synthesis"]
    labels = [R.MOD[s["key"]]["short"] for s in syn]
    series = [("Délivrés", [s["delivered"] for s in syn], "#2E8B57"),
              ("En cours", [s["encours"] for s in syn], "#4FA3D9"),
              ("Bloqués / Refusés", [s["bloques"] or 0 for s in syn], "#D64541")]
    fig, ax = plt.subplots(figsize=size)
    bottom = [0] * len(syn)
    top = max(1, max(s["total"] for s in syn))
    for name, vals, col in series:
        ax.bar(labels, vals, bottom=bottom, color=col, label=name, width=0.6)
        for i, v in enumerate(vals):
            if v and v / top > 0.04:
                ax.text(i, bottom[i] + v / 2, f"{v:,}".replace(",", " "), ha="center", va="center", fontsize=7.5, color="white", fontweight="bold")
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, s in enumerate(syn):
        ax.text(i, bottom[i] + top * 0.015, f"{s['taux'] * 100:.1f} %".replace(".", ","), ha="center", va="bottom", fontsize=8, color="#1F4E79", fontweight="bold")
    ax.set_ylim(0, top * 1.12)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", " ")))
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3, frameon=False, fontsize=8)
    ax.set_title("Dossiers par volet (taux de délivrance au-dessus)", fontsize=9.5, color="#1F4E79", loc="left")
    return _png(fig)


def chart_status(mod, size=(6.4, 3.6)):
    """Barres horizontales par statut (ordre de traitement, couleurs par catégorie)."""
    plt = _plt()
    rows = [r for r in mod["statuses"] if not (r["category"] == "Autre" and r["count"] == 0)]
    colors = ["#" + c for c in pt_colors([r["category"] for r in rows])]
    fig, ax = plt.subplots(figsize=size)
    y = list(range(len(rows)))
    ax.barh(y, [r["count"] for r in rows], color=colors, height=0.68)
    ax.set_yticks(y)
    ax.set_yticklabels([r["status"] for r in rows], fontsize=8)
    ax.invert_yaxis()
    mx = max([r["count"] for r in rows] + [1])
    for i, r in enumerate(rows):
        ax.text(r["count"] + mx * 0.01, i, f"{r['count']:,}".replace(",", " ") + f"  ({r['pct'] * 100:.1f} %)".replace(".", ","),
                va="center", fontsize=7.5, color="#1F2937")
    ax.set_xlim(0, mx * 1.28)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", " ")))
    ax.set_title("Dossiers par statut", fontsize=9.5, color="#" + mod["theme"], loc="left")
    return _png(fig)


def chart_types(mod, size=(6.4, 3.6)):
    """Anneau des types d'opérateurs (KPI Identification)."""
    import textwrap
    plt = _plt()
    types = [t for t in mod["types"][:4]]
    vals = [t["count"] for t in types]
    fig, (ax, lg) = plt.subplots(1, 2, figsize=size, gridspec_kw={"width_ratios": [1, 1.05]})
    lg.axis("off")
    if sum(vals) == 0:
        ax.text(0.5, 0.5, "Aucune donnée", ha="center", va="center", color="#5B6B7F")
        ax.axis("off")
        return _png(fig)
    wedges, _ = ax.pie(vals, colors=["#" + c for c in TYPE_PAL], startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.38, edgecolor="white"))
    ax.text(0, 0, f"{mod['types_total']:,}".replace(",", " ") + "\nopérateurs", ha="center", va="center", fontsize=10, color="#1F2937", fontweight="bold")
    labels = ["\n".join(textwrap.wrap(t["label"], 26)) + f"\n{t['count']:,}".replace(",", " ") + f" ({t['pct'] * 100:.1f} %)".replace(".", ",") for t in types]
    lg.legend(wedges, labels, loc="center left", frameon=False, fontsize=8.5, labelspacing=1.1, handlelength=1.2)
    ax.set_title("Opérateurs par type", fontsize=9.5, color="#" + mod["theme"], loc="left")
    ax.axis("equal")
    return _png(fig)


def module_chart(mod):
    return chart_types(mod) if mod["key"] == "identification" else chart_status(mod)


# ----------------------------------------------------------------------------- contenus partagés
def key_indicators(mod):
    """Cartes « indicateurs clés » d'une section KPI : (libellé, valeur formatée, couleur de fond)."""
    if mod["key"] == "identification":
        return [("Total opérateurs", R.fmt_int(mod["total"]), SOFT), ("Identifiés", R.fmt_int(mod["delivered"]), "C6EFCE"),
                ("En cours / autres", R.fmt_int(mod["encours"]), "E4EEF9"), ("Taux de réussite", R.fmt_pct(mod["taux"]), "C6EFCE")]
    return [("Total dossiers", R.fmt_int(mod["total"]), SOFT), ("Délivrés", R.fmt_int(mod["delivered"]), "C6EFCE"),
            ("En cours", R.fmt_int(mod["encours"]), "E4EEF9"), ("Non conformes", R.fmt_int(mod["nonconf"]), "FFE699"),
            ("Refusés", R.fmt_int(mod["refused"]), "FFCCCC"), ("Taux de délivrance", R.fmt_pct(mod["taux"]), "C6EFCE"),
            ("Taux de refus", R.fmt_pct(mod["taux_refus"]), "FFCCCC")]


def status_table_rows(mod):
    rows = [(r["status"], R.fmt_int(r["count"]), R.fmt_pct(r["pct"]), r["category_label"], CAT_FILL.get(r["category"], "FFFFFF"))
            for r in mod["statuses"]]
    rows.append(("TOTAL", R.fmt_int(mod["status_total"]), R.fmt_pct(1 if mod["status_total"] else 0),
                 "✔ Contrôle OK" if mod["control_ok"] else "✖ Écart", "D9E2F3"))
    return rows


def methodology(comp, meta, generated):
    """Contenu de la page « Méthodologie & qualité des données » (listes de lignes)."""
    rep = (meta or {}).get("report") or {}
    sources = [f"Fichier source : {meta.get('source_file', '—')} (SHA-256 {str(meta.get('sha256', ''))[:16]}…)" if meta else "Fichier source : —",
               f"Import du {fmt_dt(meta.get('imported_at'))} par {meta.get('user', '—')}" if meta else "Import : —",
               f"Génération du {generated.strftime('%d/%m/%Y à %H:%M')} — valeurs calculées par une fonction unique (identiques à l'Excel)."]
    for m in R.MODULES:
        mod = comp["modules"][m["key"]]
        sources.append(f"Feuille « {m['sheet']} » : {R.fmt_int(mod['total'])} dossiers (col. {mod['cols']['id']} = identifiant, "
                       f"col. {mod['cols']['status']} = statut{', col. ' + mod['cols']['type'] + ' = type' if 'type' in mod['cols'] else ''}).")
    controls = []
    for m in R.MODULES:
        mod = comp["modules"][m["key"]]
        controls.append(f"{m['short']} : somme des statuts = total ({'OK' if mod['control_ok'] else 'ÉCART'}) ; "
                        f"{R.fmt_int(mod['autre'])} dossier(s) en « Autres statuts / statut vide ».")
    for w in rep.get("warnings", []):
        controls.append("Avertissement à l'import : " + w)
    typos = [f"{R.MOD[k]['short']} : {lab} ({n})" for k, d in (rep.get("typos") or {}).items() for lab, n in d.items()]
    unknown = [f"{R.MOD[k]['short']} : « {lab} » ({n})" for k, d in (rep.get("unknown_statuses") or {}).items() for lab, n in d.items()]
    hyp = ["Voyant de synthèse : 🟢 si taux ≥ seuil vert, 🟡 si ≥ seuil orange, sinon 🔴 ; ⚪ N/D si aucune donnée."]
    hyp += [f"{R.MOD[k]['short']} : vert ≥ {g * 100:.0f} %, orange ≥ {o * 100:.0f} %." for k, (g, o) in R.VOYANT_THRESHOLDS.items()]
    hyp.append("Points d'alerte : « taux » → alerte si la valeur est inférieure au seuil ; « nombre » → supérieure. "
               "Seuils = valeurs d'origine du fichier de référence ou hypothèses à valider.")
    hyp.append("En cours = statuts de traitement + statuts non listés ; Bloqués/Refusés = Non conforme + Irrecevable + Rejeté. "
               "Comparaison des statuts insensible à la casse (comme NB.SI).")
    return {"sources": sources, "controls": controls, "typos": typos or ["Aucune coquille détectée."],
            "unknown": unknown or ["Aucun statut non référencé."], "hypotheses": hyp}


def fmt_dt(iso):
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m/%Y à %H:%M")
    except (TypeError, ValueError):
        return "—"


# ----------------------------------------------------------------------------- PPTX
def build_pptx(comp, meta, ctx):
    """Présentation 16:9 : garde, synthèse, chiffres clés & alertes, une section par KPI, méthodologie."""
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]
    gen = ctx["generated"]
    footer_txt = f"{ctx.get('agency', 'AMMPS — DSID')} · Reporting Métier PCHC · généré le {gen.strftime('%d/%m/%Y')}"

    def rgb(h):
        return RGBColor.from_string(h)

    def rect(slide, x, y, w, h, color, line=None, shape=MSO_SHAPE.RECTANGLE):
        s = slide.shapes.add_shape(shape, x, y, w, h)
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(color)
        if line:
            s.line.color.rgb = rgb(line)
            s.line.width = Pt(0.75)
        else:
            s.line.fill.background()
        s.shadow.inherit = False
        return s

    def text(slide, x, y, w, h, value, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        lines = value if isinstance(value, list) else [value]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = str(line)
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = "Calibri"
            r.font.color.rgb = rgb(color)
        return tb

    def header(slide, title, subtitle, theme):
        rect(slide, 0, 0, W, Inches(0.95), theme)
        rect(slide, 0, Inches(0.95), W, Inches(0.06), light_hex(theme, 0.55))
        text(slide, Inches(0.45), Inches(0.12), Inches(10.5), Inches(0.5), title, 24, True, "FFFFFF")
        text(slide, Inches(0.45), Inches(0.57), Inches(11), Inches(0.35), subtitle, 11.5, False, light_hex(theme, 0.75))

    def table(slide, x, y, w, col_w, rows, header_row, theme, font=10, row_h=0.28, fills=None, aligns=None, bold_last=False):
        nrows, ncols = len(rows) + 1, len(header_row)
        shape = slide.shapes.add_table(nrows, ncols, x, y, w, Inches(row_h * nrows))
        tbl = shape.table
        tbl.first_row = True
        tbl.horz_banding = False
        total = sum(col_w)
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = Emu(int(w * cw / total))
        for r in range(nrows):
            tbl.rows[r].height = Inches(row_h)
        for r, row in enumerate([header_row] + rows):
            for c in range(ncols):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(0.06)
                cell.margin_top = cell.margin_bottom = Inches(0.01)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                val = row[c] if c < len(row) else ""
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.alignment = (aligns[c] if aligns else (PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER))
                run = p.add_run()
                run.text = str(val)
                run.font.size = Pt(font)
                run.font.name = "Calibri"
                run.font.bold = r == 0 or (bold_last and r == nrows - 1)
                run.font.color.rgb = rgb("FFFFFF" if r == 0 else INK)
                cell.fill.solid()
                if r == 0:
                    cell.fill.fore_color.rgb = rgb(theme)
                else:
                    f = fills(r - 1, c) if fills else None
                    cell.fill.fore_color.rgb = rgb(f or ("FFFFFF" if r % 2 else "F8FAFC"))
        return shape

    def cards(slide, items, x, y, w, h, value_size=20, theme=NAVY):
        gap = Inches(0.12)
        cw = int((w - gap * (len(items) - 1)) / len(items))
        for i, (lab, val, bg) in enumerate(items):
            cx = x + i * (cw + gap)
            rect(slide, cx, y, cw, h, bg, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.08
            text(slide, cx, y + Inches(0.06), cw, Inches(0.3), lab.upper(), 8.5, True, MUTED, PP_ALIGN.CENTER)
            text(slide, cx, y + Inches(0.3), cw, h - Inches(0.34), val, value_size, True, theme, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

    def picture(slide, png, x, y, w=None, h=None):
        slide.shapes.add_picture(io.BytesIO(png), x, y, width=w, height=h)

    def bullets(slide, x, y, w, h, title, lines, theme, size=11):
        rect(slide, x, y, w, h, light_hex(theme, 0.9), line=light_hex(theme, 0.6))
        rect(slide, x, y, Inches(0.07), h, theme)
        text(slide, x + Inches(0.15), y + Inches(0.06), w - Inches(0.2), Inches(0.3), title, 11, True, theme)
        text(slide, x + Inches(0.15), y + Inches(0.36), w - Inches(0.25), h - Inches(0.4), [f"• {l}" for l in lines], size, False, INK)

    slides_meta = []

    # 1. Page de garde
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, H, NAVY)
    rect(s, 0, Inches(4.95), W, Inches(0.08), "2E75B6")
    stripe = [m["theme"] for m in R.MODULES]
    for i, col in enumerate(stripe):
        rect(s, Inches(0.8 + i * 0.5), Inches(1.35), Inches(0.42), Inches(0.1), col if col != NAVY else "8FAADC")
    text(s, Inches(0.8), Inches(1.7), Inches(11.5), Inches(1.0), ctx.get("title") or "Reporting Métier PCHC – Plateforme Tarkhiss", 36, True, "FFFFFF")
    text(s, Inches(0.8), Inches(2.75), Inches(11.5), Inches(0.5), "Revue de direction — Produits Cosmétiques et d'Hygiène Corporelle", 17, False, "BDD7EE")
    text(s, Inches(0.8), Inches(3.55), Inches(11.5), Inches(0.9),
         [f"Période : {ctx['period_label']}", f"Situation des dossiers au {fmt_dt(meta.get('imported_at')) if meta else '—'}"], 14, False, "FFFFFF")
    text(s, Inches(0.8), Inches(5.3), Inches(8), Inches(0.9), [ctx.get("agency") or "AMMPS — DSID", f"Document généré le {gen.strftime('%d/%m/%Y à %H:%M')}"], 13, False, "D6DCE4")
    if ctx.get("logo_path") and os.path.isfile(ctx["logo_path"]) and not ctx["logo_path"].lower().endswith(".svg"):
        try:
            badge = rect(s, Inches(10.35), Inches(5.25), Inches(2.3), Inches(1.4), "FFFFFF", shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            badge.adjustments[0] = 0.1
            s.shapes.add_picture(ctx["logo_path"], Inches(10.5), Inches(5.35), height=Inches(1.2))
        except Exception:
            pass
    slides_meta.append(None)

    # 2. Synthèse exécutive
    s = prs.slides.add_slide(blank)
    header(s, "Synthèse exécutive", f"Situation par volet — {ctx['period_label']}", NAVY)
    rows = []
    for r in comp["synthesis"]:
        rows.append([r["module"], R.fmt_int(r["total"]), R.fmt_int(r["delivered"]), R.fmt_pct(r["taux"]),
                     R.fmt_int(r["encours"]), "–" if r["bloques"] is None else R.fmt_int(r["bloques"]),
                     f"≥ {r['seuil_vert'] * 100:.0f} % / ≥ {r['seuil_orange'] * 100:.0f} %", "● " + voyant_info(r["voyant"])[1]])
    t = comp["totals"]
    rows.append(["TOTAL", R.fmt_int(t["total"]), R.fmt_int(t["delivered"]), R.fmt_pct(t["taux"]), R.fmt_int(t["encours"]), R.fmt_int(t["bloques"]), "", ""])
    syn = comp["synthesis"]

    def syn_fill(r, c):
        if r == len(syn):
            return "D9E2F3"
        if c == 7:
            return light_hex(voyant_info(syn[r]["voyant"])[0], 0.7)
        return None
    table(s, Inches(0.45), Inches(1.25), Inches(12.4), [3.3, 1.2, 1.4, 1.2, 1.2, 1.4, 1.8, 1.3],
          rows, ["Module", "Total", "Délivrés / Validés", "Taux", "En cours*", "Bloqués / Refusés", "Seuils vert / orange", "Voyant"],
          NAVY, font=11, row_h=0.36, fills=syn_fill, bold_last=True)
    picture(s, chart_modules(comp, size=(8.6, 2.9)), Inches(0.45), Inches(4.2), h=Inches(2.75))
    bullets(s, Inches(8.95), Inches(4.25), Inches(3.9), Inches(2.0), "À retenir",
            [f"{R.fmt_int(t['total'])} dossiers, {R.fmt_pct(t['taux'])} délivrés/validés.",
             f"{R.fmt_int(t['encours'])} dossiers en cours de traitement.",
             f"{R.fmt_int(t['bloques'])} bloqués ou refusés.",
             f"{sum(1 for a in comp['alerts'] if 'CRITIQUE' in a['status'])} alerte(s) critique(s) sur {len(comp['alerts'])}."], NAVY, 11)
    text(s, Inches(0.45), Inches(3.85), Inches(12), Inches(0.3), "* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté.", 8.5, False, MUTED, italic=True)
    slides_meta.append(NAVY)

    # 3. Chiffres clés + points d'alerte
    s = prs.slides.add_slide(blank)
    header(s, "Chiffres clés & points d'alerte", "Résumé décisionnel — actions requises", NAVY)
    card_bg = {"ok": "C6EFCE", "info": "E4EEF9", "bad": "FFCCCC"}
    cards(s, [(c["label"].title(), R.fmt_int(c["value"]), card_bg[c["kind"]]) for c in comp["cards"]],
          Inches(0.45), Inches(1.2), Inches(12.4), Inches(1.15), 24)
    arows = []
    for a in comp["alerts"]:
        pctf = a["type"] == "taux"
        fmt = (lambda v: f"{v * 100:.0f} %") if pctf else R.fmt_int
        arows.append([a["indicator"], a["module"], R.fmt_pct(a["value"]) if pctf else R.fmt_int(a["value"]),
                      fmt(a["warn"]), fmt(a["crit"]), "● " + ALERT_LABEL[alert_key(a["status"])], a["priority"]])
    al = comp["alerts"]
    table(s, Inches(0.45), Inches(2.6), Inches(12.4), [3.6, 1.4, 1.5, 1.5, 1.5, 1.6, 1.3], arows,
          ["Indicateur", "Module", "Valeur actuelle", "Seuil attention", "Seuil critique", "Statut", "Priorité"], "C00000",
          font=10.5, row_h=0.36, fills=lambda r, c: ALERT_FILL[alert_key(al[r]["status"])] if c >= 5 else None)
    text(s, Inches(0.45), Inches(6.35), Inches(12.4), Inches(0.4), "Seuils : « taux » → alerte si la valeur est INFÉRIEURE au seuil ; « nombre » → SUPÉRIEURE. Seuils = paramètres du classeur (en bleu dans l'Excel).", 9, False, MUTED, italic=True)
    slides_meta.append(NAVY)

    # 4. Une section par KPI
    for m in R.MODULES:
        mod = comp["modules"][m["key"]]
        th = m["theme"]
        s = prs.slides.add_slide(blank)
        header(s, f"KPI {m['label']}", f"Feuille « {m['sheet']} » — {R.fmt_int(mod['total'])} {'opérateurs' if m['key'] == 'identification' else 'dossiers'} · voyant {voyant_info(mod['voyant'])[1]}", th)
        cards(s, key_indicators(mod), Inches(0.45), Inches(1.18), Inches(12.4), Inches(0.95), 18, th)
        rows = status_table_rows(mod)
        if m["key"] == "identification":
            trows = [[t["label"], R.fmt_int(t["count"]), R.fmt_pct(t["pct"])] for t in mod["types"]]
            trows.append(["TOTAL", R.fmt_int(mod["types_total"]), R.fmt_pct(1 if mod["types_total"] else 0)])
            table(s, Inches(0.45), Inches(2.35), Inches(6.1), [3.2, 1.0, 1.0, 1.4], [r[:4] for r in rows],
                  ["Statut", "Nombre", "%", "Catégorie"], th, font=10.5, row_h=0.32,
                  fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bold_last=True)
            table(s, Inches(0.45), Inches(3.85), Inches(6.1), [3.8, 1.0, 1.0], trows, ["Type d'opérateur", "Nombre", "%"], th,
                  font=10.5, row_h=0.32, fills=lambda r, c, n=len(trows): "D9E2F3" if r == n - 1 else None, bold_last=True)
        else:
            rh = min(0.3, 4.1 / (len(rows) + 1))
            table(s, Inches(0.45), Inches(2.35), Inches(6.1), [3.0, 1.0, 1.0, 1.3], [r[:4] for r in rows],
                  ["Statut", "Nombre", "%", "Catégorie"], th, font=10 if len(rows) > 12 else 10.5, row_h=rh,
                  fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bold_last=True)
        picture(s, module_chart(mod), Inches(6.8), Inches(2.3), w=Inches(6.05))
        bullets(s, Inches(6.8), Inches(5.55), Inches(6.05), Inches(1.3), "Constats", comp["insights"][m["key"]], th, 10.5)
        slides_meta.append(th)

    # 5. Méthodologie & qualité des données
    s = prs.slides.add_slide(blank)
    header(s, "Méthodologie & qualité des données", "Sources, contrôles, coquilles normalisées, hypothèses des seuils", NAVY)
    md = methodology(comp, meta, gen)
    bullets(s, Inches(0.45), Inches(1.2), Inches(6.15), Inches(2.85), "Sources", md["sources"], NAVY, 9.5)
    bullets(s, Inches(6.75), Inches(1.2), Inches(6.1), Inches(2.85), "Contrôles", md["controls"][:8], NAVY, 9.5)
    bullets(s, Inches(0.45), Inches(4.2), Inches(6.15), Inches(2.65), "Coquilles normalisées & statuts non référencés",
            md["typos"] + md["unknown"], NAVY, 9.5)
    bullets(s, Inches(6.75), Inches(4.2), Inches(6.1), Inches(2.65), "Hypothèses des seuils", md["hypotheses"], NAVY, 9)
    slides_meta.append(NAVY)

    # Pied de page + numéros (sauf page de garde)
    n = len(prs.slides)
    for i, (slide, theme) in enumerate(zip(prs.slides, slides_meta), start=1):
        if theme is None:
            continue
        rect(slide, Inches(0.45), Inches(7.02), Inches(12.4), Emu(9525), LINE)
        text(slide, Inches(0.45), Inches(7.07), Inches(10.5), Inches(0.3), footer_txt, 8.5, False, MUTED)
        text(slide, Inches(11.35), Inches(7.07), Inches(1.5), Inches(0.3), f"{i} / {n}", 9, True, theme, PP_ALIGN.RIGHT)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------- PDF (HTML → WeasyPrint)
def _img(png, style=""):
    return f'<img src="data:image/png;base64,{base64.b64encode(png).decode()}" style="{style}">'


def _dot(hexc):
    return f'<span style="color:#{hexc};font-size:13px;">&#9679;</span>'


def _table(headers, rows, theme, widths=None, fills=None, font=10, bold_last=False):
    ws = widths or [None] * len(headers)
    h = "".join(f'<th style="background:#{theme};color:#fff;padding:5px 7px;font-size:{font}px;text-align:{"left" if i == 0 else "center"};'
                f'{f"width:{w}%;" if w else ""}border:1px solid #{theme};">{esc(str(x))}</th>' for i, (x, w) in enumerate(zip(headers, ws)))
    body = ""
    for r, row in enumerate(rows):
        last = bold_last and r == len(rows) - 1
        tds = ""
        for c, v in enumerate(row):
            bg = (fills(r, c) if fills else None) or ("FFFFFF" if r % 2 == 0 else "F8FAFC")
            tds += (f'<td style="background:#{bg};padding:4px 7px;font-size:{font}px;border:1px solid #{LINE};'
                    f'text-align:{"left" if c == 0 else "center"};{"font-weight:bold;" if last else ""}">{v}</td>')
        body += f"<tr>{tds}</tr>"
    return f'<table style="width:100%;border-collapse:collapse;">{"<tr>" + h + "</tr>"}{body}</table>'


def _cards_html(items, theme, size=17):
    w = 100 / len(items)
    tds = "".join(f'<td style="width:{w:.2f}%;padding:0 4px;"><div style="background:#{bg};border:1px solid #{LINE};border-radius:6px;padding:6px 4px;text-align:center;">'
                  f'<div style="font-size:8px;color:#{MUTED};font-weight:bold;text-transform:uppercase;letter-spacing:.3px;">{esc(lab)}</div>'
                  f'<div style="font-size:{size}px;color:#{theme};font-weight:bold;margin-top:2px;">{esc(val)}</div></div></td>'
                  for lab, val, bg in items)
    return f'<table style="width:100%;border-collapse:collapse;margin:0 -4px 10px;"><tr>{tds}</tr></table>'


def _band(title, subtitle, theme):
    return (f'<div style="background:#{theme};color:#fff;padding:10px 14px;border-radius:4px;margin-bottom:10px;">'
            f'<div style="font-size:19px;font-weight:bold;">{esc(title)}</div>'
            f'<div style="font-size:11px;color:#{light_hex(theme, 0.75)};margin-top:2px;">{esc(subtitle)}</div></div>')


def _box(title, lines, theme, font=10.5):
    items = "".join(f'<li style="margin:2px 0;">{esc(l)}</li>' for l in lines)
    return (f'<div style="background:#{light_hex(theme, 0.92)};border-left:4px solid #{theme};padding:7px 12px;margin-top:8px;">'
            f'<div style="font-weight:bold;color:#{theme};font-size:11px;margin-bottom:3px;">{esc(title)}</div>'
            f'<ul style="margin:0;padding-left:16px;font-size:{font}px;color:#{INK};">{items}</ul></div>')


def synthesis_table_html(comp, font=11):
    syn = comp["synthesis"]
    rows = []
    for r in syn:
        vc, vl = voyant_info(r["voyant"])
        rows.append([esc(r["module"]), R.fmt_int(r["total"]), R.fmt_int(r["delivered"]), R.fmt_pct(r["taux"]), R.fmt_int(r["encours"]),
                     "–" if r["bloques"] is None else R.fmt_int(r["bloques"]), f"{r['seuil_vert'] * 100:.0f} % / {r['seuil_orange'] * 100:.0f} %",
                     f"{_dot(vc)} {vl}"])
    t = comp["totals"]
    rows.append(["TOTAL", R.fmt_int(t["total"]), R.fmt_int(t["delivered"]), R.fmt_pct(t["taux"]), R.fmt_int(t["encours"]), R.fmt_int(t["bloques"]), "", ""])
    return _table(["Module", "Total", "Délivrés / Validés", "Taux", "En cours*", "Bloqués / Refusés", "Seuils vert / orange", "Voyant"],
                  rows, NAVY, [27, 9, 12, 8, 9, 12, 13, 10], font=font, bold_last=True,
                  fills=lambda r, c: "D9E2F3" if r == len(syn) else None)


def alerts_table_html(comp, font=10.5):
    al = comp["alerts"]
    rows = []
    for a in al:
        pctf = a["type"] == "taux"
        fmt = (lambda v: f"{v * 100:.0f} %") if pctf else R.fmt_int
        k = alert_key(a["status"])
        rows.append([esc(a["indicator"]), esc(a["module"]), R.fmt_pct(a["value"]) if pctf else R.fmt_int(a["value"]),
                     fmt(a["warn"]), fmt(a["crit"]), f"<b>{ALERT_LABEL[k]}</b>", a["priority"]])
    return _table(["Indicateur", "Module", "Valeur actuelle", "Seuil attention", "Seuil critique", "Statut", "Priorité"],
                  rows, "C00000", [30, 12, 13, 12, 12, 11, 10], font=font,
                  fills=lambda r, c: ALERT_FILL[alert_key(al[r]["status"])] if c >= 5 else None)


def report_html(comp, meta, ctx):
    """Document HTML complet du PDF (même structure et mêmes données que le PPTX)."""
    gen = ctx["generated"]
    pages = []
    logo = ""
    if ctx.get("logo_path") and os.path.isfile(ctx["logo_path"]) and not ctx["logo_path"].lower().endswith(".svg"):
        with open(ctx["logo_path"], "rb") as f:
            ext = "jpeg" if ctx["logo_path"].lower().endswith((".jpg", ".jpeg")) else "png"
            logo = f'<img src="data:image/{ext};base64,{base64.b64encode(f.read()).decode()}" style="height:70px;background:#fff;padding:8px;border-radius:6px;">'
    stripe = "".join(f'<span style="display:inline-block;width:34px;height:7px;background:#{m["theme"] if m["theme"] != NAVY else "8FAADC"};margin-right:6px;"></span>' for m in R.MODULES)
    pages.append(f"""<div class="cover" style="background:#{NAVY};color:#fff;padding:60px 50px 40px;height:15.2cm;">
      <div>{stripe}</div>
      <div style="font-size:32px;font-weight:bold;margin-top:22px;">{esc(ctx.get('title') or 'Reporting Métier PCHC – Plateforme Tarkhiss')}</div>
      <div style="font-size:16px;color:#BDD7EE;margin-top:8px;">Revue de direction — Produits Cosmétiques et d'Hygiène Corporelle</div>
      <div style="font-size:14px;margin-top:34px;">Période : {esc(ctx['period_label'])}<br>Situation des dossiers au {fmt_dt(meta.get('imported_at')) if meta else '—'}</div>
      <table style="width:100%;margin-top:60px;border-top:3px solid #2E75B6;"><tr>
        <td style="color:#D6DCE4;font-size:13px;padding-top:14px;">{esc(ctx.get('agency') or 'AMMPS — DSID')}<br>Document généré le {gen.strftime('%d/%m/%Y à %H:%M')}</td>
        <td style="text-align:right;padding-top:14px;">{logo}</td></tr></table>
    </div>""")
    t = comp["totals"]
    pages.append(_band("Synthèse exécutive", f"Situation par volet — {ctx['period_label']}", NAVY)
                 + synthesis_table_html(comp)
                 + f'<div style="font-size:8.5px;color:#{MUTED};font-style:italic;margin:4px 0 8px;">* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté.</div>'
                 + '<table style="width:100%;border-collapse:collapse;"><tr><td style="width:66%;vertical-align:top;">'
                 + _img(chart_modules(comp, size=(8.6, 3.0)), "width:100%;") + '</td><td style="vertical-align:top;padding-left:10px;">'
                 + _box("À retenir", [f"{R.fmt_int(t['total'])} dossiers, {R.fmt_pct(t['taux'])} délivrés/validés.",
                                      f"{R.fmt_int(t['encours'])} dossiers en cours de traitement.",
                                      f"{R.fmt_int(t['bloques'])} bloqués ou refusés.",
                                      f"{sum(1 for a in comp['alerts'] if 'CRITIQUE' in a['status'])} alerte(s) critique(s) sur {len(comp['alerts'])}."], NAVY)
                 + "</td></tr></table>")
    card_bg = {"ok": "C6EFCE", "info": "E4EEF9", "bad": "FFCCCC"}
    pages.append(_band("Chiffres clés & points d'alerte", "Résumé décisionnel — actions requises", NAVY)
                 + _cards_html([(c["label"], R.fmt_int(c["value"]), card_bg[c["kind"]]) for c in comp["cards"]], NAVY, 22)
                 + alerts_table_html(comp)
                 + f'<div style="font-size:8.5px;color:#{MUTED};font-style:italic;margin-top:6px;">Seuils : « taux » → alerte si la valeur est INFÉRIEURE au seuil ; « nombre » → SUPÉRIEURE. Seuils = paramètres du classeur (en bleu dans l\'Excel).</div>')
    for m in R.MODULES:
        mod = comp["modules"][m["key"]]
        th = m["theme"]
        rows = status_table_rows(mod)
        trows = [[esc(r[0]), r[1], r[2], esc(r[3])] for r in rows]
        tbl = _table(["Statut", "Nombre", "%", "Catégorie"], trows, th, [46, 16, 16, 22], font=10 if len(rows) > 12 else 10.5,
                     fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bold_last=True)
        if m["key"] == "identification":
            ty = [[esc(x["label"]), R.fmt_int(x["count"]), R.fmt_pct(x["pct"])] for x in mod["types"]]
            ty.append(["TOTAL", R.fmt_int(mod["types_total"]), R.fmt_pct(1 if mod["types_total"] else 0)])
            tbl += '<div style="height:10px;"></div>' + _table(["Type d'opérateur", "Nombre", "%"], ty, th, [62, 19, 19], font=10.5, bold_last=True,
                                                             fills=lambda r, c, n=len(ty): "D9E2F3" if r == n - 1 else None)
        unit = "opérateurs" if m["key"] == "identification" else "dossiers"
        pages.append(_band(f"KPI {m['label']}", f"Feuille « {m['sheet']} » — {R.fmt_int(mod['total'])} {unit} · voyant {voyant_info(mod['voyant'])[1]}", th)
                     + _cards_html(key_indicators(mod), th)
                     + '<table style="width:100%;border-collapse:collapse;"><tr>'
                     + f'<td style="width:49%;vertical-align:top;">{tbl}</td>'
                     + f'<td style="vertical-align:top;padding-left:14px;">{_img(module_chart(mod), "width:100%;")}'
                     + _box("Constats", comp["insights"][m["key"]], th) + "</td></tr></table>")
    md = methodology(comp, meta, gen)
    pages.append(_band("Méthodologie & qualité des données", "Sources, contrôles, coquilles normalisées, hypothèses des seuils", NAVY)
                 + '<table style="width:100%;border-collapse:collapse;"><tr><td style="width:50%;vertical-align:top;padding-right:7px;">'
                 + _box("Sources", md["sources"], NAVY, 9.5) + _box("Coquilles normalisées & statuts non référencés", md["typos"] + md["unknown"], NAVY, 9.5)
                 + '</td><td style="vertical-align:top;padding-left:7px;">'
                 + _box("Contrôles", md["controls"], NAVY, 9.5) + _box("Hypothèses des seuils", [h.replace("🟢", "vert").replace("🟡", "orange").replace("🔴", "rouge").replace("⚪ ", "") for h in md["hypotheses"]], NAVY, 9.5)
                 + "</td></tr></table>")
    return pages


def build_pdf(comp, meta, ctx):
    """PDF A4 paysage côté serveur : WeasyPrint (en-tête/pied de page, « page n / N »), repli xhtml2pdf."""
    pages = report_html(comp, meta, ctx)
    gen = ctx["generated"]
    head = f"{ctx.get('agency', 'AMMPS — DSID')} · Reporting Métier PCHC"
    foot = f"Généré le {gen.strftime('%d/%m/%Y à %H:%M')} · {ctx['period_label']}"
    body = "".join(f'<div class="page{" first" if i == 0 else ""}">{p}</div>' for i, p in enumerate(pages))
    try:
        import weasyprint
        css = f"""
          @page {{ size: A4 landscape; margin: 1.5cm 1.3cm 1.3cm 1.3cm;
                  @top-left {{ content: "{head}"; font: 8.5pt Arial, sans-serif; color: #{MUTED}; }}
                  @top-right {{ content: "{ctx['period_label']}"; font: 8.5pt Arial, sans-serif; color: #{MUTED}; }}
                  @bottom-left {{ content: "{foot}"; font: 8pt Arial, sans-serif; color: #{MUTED}; }}
                  @bottom-right {{ content: "Page " counter(page) " / " counter(pages); font: bold 8.5pt Arial, sans-serif; color: #{NAVY}; }} }}
          @page :first {{ margin: 0; @top-left {{ content: none; }} @top-right {{ content: none; }} @bottom-left {{ content: none; }} @bottom-right {{ content: none; }} }}
          body {{ font-family: Arial, 'DejaVu Sans', sans-serif; color: #{INK}; margin: 0; }}
          .page {{ page-break-after: always; }} .page:last-child {{ page-break-after: auto; }}
          .first .cover {{ height: 21cm !important; box-sizing: border-box; padding: 2.4cm 2cm !important; }}
          table {{ page-break-inside: avoid; }}"""
        doc = f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{body}</body></html>'
        return weasyprint.HTML(string=doc).write_pdf(), "weasyprint"
    except Exception:
        pass
    from xhtml2pdf import pisa
    css = f"""
      @page {{ size: A4 landscape; margin: 1.5cm 1.3cm 1.4cm 1.3cm;
        @frame header {{ -pdf-frame-content: hdr; top: 0.5cm; left: 1.3cm; right: 1.3cm; height: 0.8cm; }}
        @frame footer {{ -pdf-frame-content: ftr; bottom: 0.4cm; left: 1.3cm; right: 1.3cm; height: 0.8cm; }} }}
      body {{ font-family: Helvetica, sans-serif; color: #{INK}; }}"""
    hdr = f'<div id="hdr" style="font-size:8px;color:#{MUTED};">{esc(head)} · {esc(ctx["period_label"])}</div>'
    ftr = f'<div id="ftr" style="font-size:8px;color:#{MUTED};">{esc(foot)} — Page <pdf:pagenumber> / <pdf:pagecount></div>'
    body = "<pdf:nextpage/>".join(pages)
    doc = f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{hdr}{ftr}{body}</body></html>'
    buf = io.BytesIO()
    pisa.CreatePDF(src=doc, dest=buf, encoding="utf-8")
    return buf.getvalue(), "xhtml2pdf"


# ----------------------------------------------------------------------------- e-mail
def summary_html(comp, meta, ctx, wrap_style):
    """Résumé de la synthèse exécutive pour le corps de l'e-mail (tableau modules + alertes)."""
    t = comp["totals"]
    crit = sum(1 for a in comp["alerts"] if "CRITIQUE" in a["status"])
    return f"""
    <div style="{wrap_style}">
      <div style="background:#{NAVY};padding:20px 24px;border-radius:6px 6px 0 0;">
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">{esc(ctx.get('title') or 'Reporting Métier PCHC – Plateforme Tarkhiss')}</div>
        <div style="color:#BDD7EE;font-size:13px;margin-top:4px;">{esc(ctx.get('subtitle') or '')}</div>
      </div>
      <div style="padding:18px 22px;border:1px solid #{LINE};border-top:none;">
        <p style="font-size:13px;color:#{INK};margin:0 0 12px;">Situation des dossiers au {fmt_dt(meta.get('imported_at')) if meta else '—'} :
          <b>{R.fmt_int(t['total'])}</b> dossiers, <b>{R.fmt_pct(t['taux'])}</b> délivrés/validés, <b>{R.fmt_int(t['encours'])}</b> en cours,
          <b>{R.fmt_int(t['bloques'])}</b> bloqués ou refusés — {crit} alerte(s) critique(s).</p>
        <div style="font-size:14px;font-weight:bold;color:#{NAVY};margin:6px 0 8px;">Synthèse exécutive</div>
        {synthesis_table_html(comp, font=12)}
        <div style="font-size:14px;font-weight:bold;color:#C00000;margin:18px 0 8px;">Points d'alerte</div>
        {alerts_table_html(comp, font=12)}
        <p style="font-size:11.5px;color:#{MUTED};margin-top:14px;">Détail par volet, graphiques et méthodologie : voir les pièces jointes (Excel, PowerPoint, PDF).</p>
      </div>
    </div>"""


def _attach(msg, data, mime, filename):
    maintype, subtype = mime.split("/", 1)
    try:
        filename.encode("ascii")
        fn = filename
    except UnicodeEncodeError:
        fn = ("utf-8", "", filename)  # RFC 2231 : filename*=utf-8''…
    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=fn)


def build_eml(subject, to_list, cc_list, body_html, attachments):
    """Message multipart/mixed : corps multipart/alternative (texte + HTML) puis pièces jointes.
    attachments = [(octets, type MIME, nom de fichier)]. Retourne les octets .eml."""
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["To"] = "; ".join(to_list)
    if cc_list:
        msg["Cc"] = "; ".join(cc_list)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="ammps.local")
    msg["X-Unsent"] = "1"  # Outlook ouvre le message en brouillon, prêt à être envoyé
    msg.set_content("Ce message nécessite un client compatible HTML. Les livrables sont en pièces jointes.")
    msg.add_alternative(body_html, subtype="html")
    for data, mime, name in attachments:
        _attach(msg, data, mime, name)
    return bytes(msg)
