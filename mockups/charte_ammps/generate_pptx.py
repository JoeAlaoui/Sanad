"""Exemples : PPTX Tarkhiss et Moussanada (charte AMMPS), PCHC hybride (PDF + PPTX) et email PCHC.
PCHC hybride = cadre AMMPS (logo, filet, typographies) + ancien rapport (paysage, tableau à gauche,
graphique et constats à droite, couleur par volet, couleurs de statut).
Lancer depuis la racine du projet : .venv/bin/python mockups/charte_ammps/generate_pptx.py
"""
import base64
import html as H
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import generate as G  # noqa: E402
import generate_suite as S  # noqa: E402
import sandbox  # noqa: E402
from generate import R, esc  # noqa: E402

FOREST, GREEN, RED, INK, MUTED, LINE, PALE, LIGHT, DEEP, TINT = "0D4823", "058B44", "DF1E26", "183928", "5B6E62", "DCE6DF", "F3F7F4", "EAF4ED", "0B3420", "8CC9A3"
FONT = "Arial"  # police présente sur tous les postes : Montserrat/Inter ne sont pas embarquables en PPTX
GEN = S.GEN


def plain(v):
    return H.unescape(re.sub(r"<[^>]+>", "", str(v)))


def uri_bytes(uri):
    return uri if isinstance(uri, bytes) else base64.b64decode(uri.split(",", 1)[1])


class Deck:
    def __init__(self, title, period, footer):
        from pptx import Presentation
        from pptx.util import Inches
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(13.333), Inches(7.5)
        self.title, self.period, self.footer = title, period, footer
        self.logo = uri_bytes(G.LOGO)
        cp = self.prs.core_properties
        cp.author, cp.title, cp.last_modified_by = "AMMPS", f"{title} — {period}", "AMMPS"

    def slide(self):
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def rect(self, s, x, y, w, h, color, line=None, rounded=False):
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches, Pt
        sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        if rounded:
            sh.adjustments[0] = min(.5, .08 / min(w, h))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(color)
        if line:
            sh.line.color.rgb = RGBColor.from_string(line)
            sh.line.width = Pt(.75)
        else:
            sh.line.fill.background()
        sh.shadow.inherit = False
        return sh

    def text(self, s, x, y, w, h, value, size=12, bold=False, color=INK, align="l", anchor="t", italic=False):
        from pptx.dml.color import RGBColor
        from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
        from pptx.util import Inches, Pt
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        tf.margin_left = tf.margin_right = Inches(.04)
        tf.margin_top = tf.margin_bottom = Inches(.02)
        for i, line in enumerate(value if isinstance(value, list) else [value]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            if i:
                p.space_before = Pt(3)
            r = p.add_run()
            r.text = re.sub(r"(\d) %", "\\1\u00a0%", str(line))
            r.font.size, r.font.bold, r.font.italic, r.font.name = Pt(size), bold, italic, FONT
            r.font.color.rgb = RGBColor.from_string(color)
        return tb

    def filet(self, s, y, x=0, w=13.333, h=.05):
        self.rect(s, x, y, w * .9, h, FOREST)
        self.rect(s, x + w * .9, y, w * .1, h, RED)

    def pic(self, s, img, x, y, w, h):
        """Image ajustée et centrée dans la boîte, sans déformation."""
        from PIL import Image
        from pptx.util import Inches
        data = uri_bytes(img)
        iw, ih = Image.open(io.BytesIO(data)).size
        k = min(w / iw, h / ih)
        pw, ph = iw * k, ih * k
        s.shapes.add_picture(io.BytesIO(data), Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2), Inches(pw), Inches(ph))

    def box_pic(self, s, img, x, y, w, h, title):
        self.rect(s, x, y, w, h, "FFFFFF", LINE, True)
        self.text(s, x + .1, y + .05, w - .2, .3, title, 11, True, FOREST)
        self.pic(s, img, x + .1, y + .36, w - .2, h - .44)

    def table(self, s, x, y, w, weights, head, rows, theme=FOREST, font=10.5, row_h=.3, aligns=None, fills=None, bolds=None):
        from pptx.dml.color import RGBColor
        from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
        from pptx.util import Emu, Inches, Pt
        nr, nc = len(rows) + 1, len(head)
        tbl = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr)).table
        tbl.horz_banding = False
        tot = sum(weights)
        for i, cw in enumerate(weights):
            tbl.columns[i].width = Emu(int(Inches(w) * cw / tot))
        for r, row in enumerate([head] + rows):
            tbl.rows[r].height = Inches(row_h)
            for c in range(nc):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(.07)
                cell.margin_top = cell.margin_bottom = Inches(.015)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if (aligns[c] if aligns else ("l" if c == 0 else "c")) == "l" else PP_ALIGN.CENTER
                run = p.add_run()
                run.text = plain(row[c]) if c < len(row) else ""
                run.font.size, run.font.name = Pt(font), FONT
                run.font.bold = r == 0 or bool(bolds and bolds(r - 1, c))
                run.font.color.rgb = RGBColor.from_string("FFFFFF" if r == 0 else INK)
                cell.fill.solid()
                f = theme if r == 0 else ((fills(r - 1, c) if fills else None) or ("FFFFFF" if r % 2 else PALE))
                cell.fill.fore_color.rgb = RGBColor.from_string(f)

    def note(self, s, x, y, w, h, label, value, size=11, theme=GREEN, bg=LIGHT):
        self.rect(s, x, y, w, h, bg)
        self.rect(s, x, y, .05, h, theme)
        self.text(s, x + .14, y + .05, w - .22, .25, label.upper(), 8.5, True, theme)
        self.text(s, x + .14, y + .29, w - .22, h - .32, value, size)

    def bullets(self, s, x, y, w, h, label, lines, size=11, theme=GREEN, bg=LIGHT):
        self.note(s, x, y, w, h, label, [f"•  {l}" for l in lines], size, theme, bg)

    def kcards(self, s, items, x, y, w, h, vs=22):
        """Cartes KPI de la charte (blanche, bonne, alerte, marquante)."""
        gap = .12
        cw = (w - gap * (len(items) - 1)) / len(items)
        for i, it in enumerate(items):
            lab, val, sub, tone = (tuple(it) + ("", ""))[:4]
            cx = x + i * (cw + gap)
            bg, vc, lc, sc = {"good": (LIGHT, GREEN, MUTED, MUTED), "alert": ("FFFFFF", RED, MUTED, RED), "warn": (FOREST, "FFFFFF", "B1CFBA", "C6DCCD")}.get(tone, ("FFFFFF", FOREST, MUTED, MUTED))
            self.rect(s, cx, y, cw, h, bg, FOREST if tone == "warn" else LINE, True)
            if tone == "alert":
                self.rect(s, cx, y, cw, .04, RED)
            self.text(s, cx + .08, y + .07, cw - .16, .25, lab.upper(), 8.5, True, lc)
            self.text(s, cx + .08, y + .31, cw - .16, .45, val, vs, True, vc)
            if sub:
                self.text(s, cx + .08, y + h - .3, cw - .16, .25, sub, 9.5, False, sc)

    def cover(self, eyebrow, title, title2, subtitle, meta, stripe=()):
        s = self.slide()
        self.pic(s, self.logo, .7, .35, 2.5, 1.2)
        self.filet(s, 1.7, h=.07)
        self.rect(s, .75, 2.51, .32, .03, RED)
        self.text(s, 1.13, 2.38, 8, .3, eyebrow.upper(), 11, True, FOREST)
        self.text(s, .7, 2.72, 12, .75, title, 40, True, FOREST)
        self.text(s, .7, 3.42, 12, .75, title2, 40, True, GREEN)
        for i, col in enumerate(stripe):
            self.rect(s, .75 + i * .5, 4.2, .42, .08, col)
        self.text(s, .7, 4.33, 12, .4, subtitle, 16, False, MUTED)
        for i, (k, v) in enumerate(meta):
            self.text(s, .7, 4.86 + i * .3, 2, .3, k, 11, False, MUTED)
            self.text(s, 2.6, 4.86 + i * .3, 6, .3, v, 11, True, INK)
        self.rect(s, 0, 6.6, 13.333, .9, DEEP)
        self.text(s, .7, 6.72, 12, .3, R["agency"], 13, True, "FFFFFF")
        self.text(s, .7, 7.03, 12, .3, G.CONTACT, 10.5, False, "C2D3C8")

    def page(self, n, title, eyebrow):
        """Diapositive charte : en-tête logo + filet, carré d'index, eyebrow à tiret rouge, titre."""
        s = self.slide()
        self.pic(s, self.logo, .45, .12, 1, .66)
        self.rect(s, 1.52, .22, .01, .46, LINE)
        self.text(s, 1.6, .2, 4.5, .5, ["Agence Marocaine du Médicament", "et des Produits de Santé"], 9, False, FOREST, anchor="m")
        self.text(s, 8.3, .2, 4.6, .5, [self.title, self.period], 9.5, False, MUTED, "r", "m")
        self.filet(s, .85)
        self.rect(s, .45, 1.03, .5, .5, LIGHT)
        self.text(s, .45, 1.03, .5, .5, f"{n:02d}", 13, True, GREEN, "c", "m")
        self.rect(s, 1.07, 1.11, .22, .025, RED)
        self.text(s, 1.31, 1.0, 8, .24, eyebrow.upper(), 8.5, True, FOREST)
        self.text(s, 1.03, 1.17, 11.8, .42, title, 22, True, FOREST)
        return s

    def band(self, title, subtitle, theme):
        """Diapositive PCHC : bandeau de l'ancien rapport (couleur du volet), logo sur pastille blanche, filet AMMPS."""
        s = self.slide()
        self.rect(s, 0, 0, 13.333, .95, theme)
        self.rect(s, 11.75, .09, 1.2, .77, "FFFFFF", rounded=True)
        self.pic(s, self.logo, 11.82, .12, 1.06, .71)
        self.filet(s, .95)
        self.text(s, .45, .1, 11, .5, title, 24, True, "FFFFFF")
        self.text(s, .45, .57, 11, .3, subtitle, 11.5, False, "FFFFFF")
        return s

    def finish(self, path):
        n = len(self.prs.slides)
        for i, s in enumerate(self.prs.slides, 1):
            if i == 1:
                continue
            self.rect(s, .45, 7.02, 12.43, .01, LINE)
            self.text(s, .45, 7.06, 10.5, .3, self.footer, 9, False, MUTED)
            self.text(s, 11.3, 7.06, 1.58, .3, f"{i} / {n}", 9.5, True, FOREST, "r")
        self.prs.save(path)
        words = " ".join(sh.text_frame.text for s in self.prs.slides for sh in s.shapes if sh.has_text_frame)
        print(os.path.basename(path), n, "diapos", os.path.getsize(path) // 1024, "Ko | SANAD/GLPI :", bool(S.BANNED_MAIL.search(words)))


# ----------------------------------------------------------------------------- Moussanada
def deck_moussanada(c):
    ch, s = c["charts"], c["sections"]
    d = Deck("Rapport Support Moussanada", c["ml"], f"AMMPS · Moussanada — {c['ml']}")
    d.cover("Rapport mensuel", "Rapport Support", "Moussanada", "Support Helpdesk", [("Période", c["ml"]), ("Édition", GEN.strftime("%d/%m/%Y"))])

    p = d.page(1, "Volume global des tickets", "Volumes")
    d.kcards(p, c["k1"], .45, 1.72, 12.43, 1.1)
    d.box_pic(p, ch["s_volume"], .45, 2.93, 6.13, 2.85, "Volumes de tickets — comparatif")
    d.box_pic(p, ch["s_backlog"], 6.75, 2.93, 6.13, 2.85, "Backlog — tickets en instance")
    d.note(p, .45, 5.88, 12.43, 1.06, "Analyse", s.get("volume"), 11)

    p = d.page(2, "Évolution des délais de traitement", "Délais")
    d.kcards(p, c["k2"], .45, 1.72, 12.43, 1.1)
    d.box_pic(p, ch["s_trend"], .45, 2.93, 7.6, 3.98, "Tendance des tickets")
    d.note(p, 8.2, 2.93, 4.68, 3.98, "Analyse", s.get("delays"), 12)

    rows = lambda items: [[t] + r[1:] for t, r in zip(S.tail_labels(items), c["top_rows"](items))]
    bold = lambda r, k: k == 1
    p = d.page(3, f"Principales catégories de demandes en {c['ml']}", "Demandes")
    d.table(p, .45, 1.72, 6.05, [3.5, .9, .9, .9], ["Catégorie"] + S.TOP_HEAD, rows(c["cats"]), row_h=.46, bolds=bold, font=11)
    d.box_pic(p, ch["s_cat"], 6.65, 1.72, 6.23, 5.17, "Top catégories")

    p = d.page(4, "Pareto des catégories", "Demandes")
    d.box_pic(p, ch["s_pareto"], .45, 1.72, 12.43, 3.75, "Part cumulée des tickets ouverts")
    d.note(p, .45, 5.58, 12.43, 1.36, "Analyse", s.get("categories"), 11.5)

    p = d.page(5, "Services les plus demandeurs", "Demandes")
    d.table(p, .45, 1.72, 6.05, [3.5, .9, .9, .9], ["Service"] + S.TOP_HEAD, rows(c["srv"]), row_h=.35, bolds=bold, font=10.5)
    d.box_pic(p, ch["s_srv"], 6.65, 1.72, 6.23, 3.85, "Top services demandeurs")
    d.note(p, .45, 5.68, 12.43, 1.26, "Analyse", s.get("services"), 11)

    p = d.page(6, "Charge et mobilisation de l’équipe SI", "Équipe")
    rows_t = [r["cells"] if isinstance(r, dict) else r for r in c["tech_rows"]]
    n = len(rows_t)
    d.table(p, .45, 1.72, 7.0, [2.4, 1.2, .9, .9, 1, .9], S.TECH_HEAD, rows_t, row_h=.31, font=10.5,
            fills=lambda r, k: LIGHT if r == n - 1 else None, bolds=lambda r, k: r == n - 1 or k == 1)
    d.note(p, 7.6, 1.72, 5.28, 2.2, "Analyse", s.get("team"), 10.5)
    d.box_pic(p, ch["s_tech"], .45, 4.02, 4.45, 2.88, "Détail des sollicitations")
    d.box_pic(p, ch["s_techvol"], 5.0, 4.02, 4.45, 2.88, "Volume de sollicitation")
    d.box_pic(p, ch["s_radar"], 9.55, 4.02, 3.33, 2.88, "Radar de performance")

    p = d.page(7, f"Analyse comparative {c['pl']} / {c['ml']}", "Comparatif et synthèse")
    d.box_pic(p, ch["s_volume2"], .45, 1.72, 6.13, 2.85, "Volumes")
    d.box_pic(p, ch["s_delay"], 6.75, 1.72, 6.13, 2.85, "Délais (heures)")
    d.note(p, .45, 4.68, 12.43, 1.1, "Analyse", s.get("comparative"), 10.5)
    d.note(p, .45, 5.86, 12.43, 1.08, "Synthèse finale et conclusion", s.get("synthesis"), 10.5, "FFFFFF", FOREST)
    p.shapes[-1].text_frame.paragraphs[0].runs[0].font.color.rgb = __import__("pptx").dml.color.RGBColor.from_string("FFFFFF")
    d.finish(os.path.join(S.EX_DIR, "Moussanada_rapport_mensuel_2026-09.pptx"))


# ----------------------------------------------------------------------------- Tarkhiss
def deck_tarkhiss():
    per = "Octobre 2026"
    d = Deck("Rapport Support Tarkhiss", per, f"AMMPS · Tarkhiss — {per}")
    d.cover("Rapport mensuel", "Rapport Support", "Tarkhiss", "Support Email & Hotline", [("Période", per), ("Édition", GEN.strftime("%d/%m/%Y"))])

    p = d.page(1, "Indicateurs clés", "Vue d’ensemble")
    d.kcards(p, R["kpi_volume"], .45, 1.72, 12.43, 1.5)
    d.kcards(p, R["kpi_quali"], .45, 3.34, 12.43, 1.5)
    d.kcards(p, R["kpi_charge"], .45, 4.96, 12.43, 1.5)

    p = d.page(2, "Analyse graphique", "Tendances")
    d.box_pic(p, G.chart_weekly(figsize=(6.1, 4.3), short=True), .45, 1.72, 6.13, 5.17, "Évolution hebdomadaire — Bugs vs Demandes")
    d.box_pic(p, G.chart_rank(figsize=(6.1, 5.0)), 6.75, 1.72, 6.13, 5.17, "Répartition des problèmes techniques")

    cat = lambda items: [[l, n, q, t] for l, n, q, t in items]
    wt, al = [3.1, .6, .8, 6.5], ["l", "c", "c", "l"]
    bold = lambda r, k: k < 2
    p = d.page(3, "Problèmes techniques signalés (166)", "Qualification")
    d.table(p, .45, 1.72, 12.43, wt, ["Type", "Nb", "%", "Description"], cat(R["problems"]), font=12, row_h=.58, aligns=al, bolds=bold)
    p = d.page(4, "Demandes d’information reçues (91)", "Qualification")
    d.table(p, .45, 1.72, 12.43, wt, ["Type", "Nb", "%", "Description"], cat(R["demandes"]), font=12, row_h=.58, aligns=al, bolds=bold)

    p = d.page(5, "Évolution hebdomadaire", "Rythme du mois")
    d.table(p, .45, 1.72, 12.43, [1.7, .8, 1, 7.5], ["Semaine", "Bugs", "Demandes", "Observations"], [list(w) for w in R["weekly"]], font=12, row_h=.62, aligns=al,
            fills=lambda r, k: "FDECED" if r == R["crit"] else None, bolds=lambda r, k: k < 3)
    d.note(p, .45, 5.6, 12.43, 1.1, "Semaine critique", "19–25 octobre : 80 cas qualifiés, dont 55 bugs (33,1 % des incidents du mois).", 13, RED, "FDECED")

    p = d.page(6, "Activité hotline", "Appels")
    d.kcards(p, [(a, b) for a, b in R["calls_kpi"]], .45, 1.72, 12.43, 1.05)
    d.box_pic(p, G.chart_heat(figsize=(8.0, 3.3)), .45, 2.9, 8.2, 4.0, "Répartition des appels par jour et par semaine")
    d.note(p, 8.8, 2.9, 4.08, 4.0, "Rappel des appels manqués", R["callback"], 13)

    p = d.page(7, "Constats clés et recommandations", "Synthèse")
    d.bullets(p, .45, 1.72, 7.4, 5.18, "Constats clés", R["constats"], 14)
    for i, r in enumerate(R["recos"]):
        y = 1.72 + i * 1.75
        d.rect(p, 8.0, y, 4.88, 1.68, LIGHT)
        d.rect(p, 8.0, y, .05, 1.68, GREEN)
        d.text(p, 8.14, y + .08, 1, .4, f"{i + 1:02d}", 15, True, GREEN)
        d.text(p, 8.14, y + .48, 4.65, 1.15, r, 12.5)
    d.finish(os.path.join(S.EX_DIR, "Tarkhiss_rapport_mensuel_2026-10.pptx"))


# ----------------------------------------------------------------------------- PCHC hybride
def patch_charts(D):
    """Graphiques de l'ancien rapport, texte en Inter."""
    def _plt():
        plt = G.mpl()
        plt.rcParams.update({"font.size": 9, "axes.edgecolor": "#B8C2CF", "axes.spines.top": False, "axes.spines.right": False})
        return plt
    D._plt = _plt
    D.NAVY = FOREST


def chart_modules(comp, D, size):
    """Barres empilées de l'ancien rapport (vert / bleu / rouge), sans titre interne."""
    plt = D._plt()
    syn = comp["synthesis"]
    labels = [D.R.MOD[s["key"]]["short"] for s in syn]
    fig, ax = plt.subplots(figsize=size)
    bottom = [0] * len(syn)
    top = max(1, max(s["total"] for s in syn))
    for name, vals, col in (("Délivrés", [s["delivered"] for s in syn], "#2E8B57"), ("En cours", [s["encours"] for s in syn], "#4FA3D9"),
                            ("Bloqués / Refusés", [s["bloques"] or 0 for s in syn], "#D64541")):
        ax.bar(labels, vals, bottom=bottom, color=col, label=name, width=.6)
        for i, v in enumerate(vals):
            if v and v / top > .04:
                ax.text(i, bottom[i] + v / 2, f"{v:,}".replace(",", " "), ha="center", va="center", fontsize=7.5, color="white", fontweight="bold")
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, s in enumerate(syn):
        ax.text(i, bottom[i] + top * .015, f"{s['taux'] * 100:.1f} %".replace(".", ","), ha="center", va="bottom", fontsize=8.5, color="#" + FOREST, fontweight="bold")
    ax.set_ylim(0, top * 1.12)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.1), ncol=3, frameon=False, fontsize=8.5)
    return D._png(fig)


def keep_lines(comp, D):
    t = comp["totals"]
    return [f"{D.R.fmt_int(t['total'])} dossiers, {D.R.fmt_pct(t['taux'])} délivrés/validés.", f"{D.R.fmt_int(t['encours'])} dossiers en cours de traitement.",
            f"{D.R.fmt_int(t['bloques'])} bloqués ou refusés.", f"{sum(1 for a in comp['alerts'] if 'CRITIQUE' in a['status'])} alerte(s) critique(s) sur {len(comp['alerts'])}."]


def synth_rows(comp, D):
    rows = [[r["module"], D.R.fmt_int(r["total"]), D.R.fmt_int(r["delivered"]), D.R.fmt_pct(r["taux"]), D.R.fmt_int(r["encours"]), "–" if r["bloques"] is None else D.R.fmt_int(r["bloques"]),
             f"≥ {r['seuil_vert'] * 100:.0f} % / ≥ {r['seuil_orange'] * 100:.0f} %", "● " + D.voyant_info(r["voyant"])[1]] for r in comp["synthesis"]]
    t = comp["totals"]
    rows.append(["TOTAL", D.R.fmt_int(t["total"]), D.R.fmt_int(t["delivered"]), D.R.fmt_pct(t["taux"]), D.R.fmt_int(t["encours"]), D.R.fmt_int(t["bloques"]), "", ""])
    return rows


def alert_rows(comp, D):
    out = []
    for a in comp["alerts"]:
        pf = a["type"] == "taux"
        fmt = (lambda v: f"{v * 100:.0f} %") if pf else D.R.fmt_int
        out.append([a["indicator"], a["module"], D.R.fmt_pct(a["value"]) if pf else D.R.fmt_int(a["value"]), fmt(a["warn"]), fmt(a["crit"]), "● " + D.ALERT_LABEL[D.alert_key(a["status"])], a["priority"].capitalize()])
    return out


SYN_HEAD = ["Module", "Total", "Délivrés / Validés", "Taux", "En cours*", "Bloqués / Refusés", "Seuils vert / orange", "Voyant"]
ALERT_HEAD = ["Indicateur", "Module", "Valeur actuelle", "Seuil attention", "Seuil critique", "Statut", "Priorité"]
NOTE_SYN = "* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté."
NOTE_ALERT = "Seuils : « taux » → alerte si la valeur est inférieure au seuil ; « nombre » → supérieure. Seuils = paramètres du classeur."


def clean(s):
    return s.replace("🟢", "vert").replace("🟡", "orange").replace("🔴", "rouge").replace("⚪ ", "")


def deck_pchc(comp, meta, D, label):
    P = D.R
    d = Deck("Reporting Métier PCHC", label, f"AMMPS · Reporting Métier PCHC — {label}")
    d.cover("Revue de direction", "Reporting Métier", "PCHC", "Plateforme Tarkhiss — Produits Cosmétiques et d’Hygiène Corporelle",
            [("Période", label), ("Situation au", D.fmt_dt(meta.get("imported_at")))], stripe=[m["theme"] for m in P.MODULES])

    def pcards(s, items, y, h, vs, theme):
        gap, w = .12, 12.43
        cw = (w - gap * (len(items) - 1)) / len(items)
        for i, (lab, val, bg) in enumerate(items):
            cx = .45 + i * (cw + gap)
            d.rect(s, cx, y, cw, h, bg, LINE, True)
            d.text(s, cx, y + .06, cw, .25, lab.upper(), 8.5, True, MUTED, "c")
            d.text(s, cx, y + .3, cw, h - .34, val, vs, True, theme, "c", "m")

    syn = comp["synthesis"]
    p = d.band("Synthèse exécutive", f"Situation par volet — {label}", FOREST)
    d.table(p, .45, 1.18, 12.43, [3.3, 1.2, 1.4, 1.2, 1.2, 1.4, 1.8, 1.3], SYN_HEAD, synth_rows(comp, D), font=11, row_h=.36, bolds=lambda r, c: r == len(syn),
            fills=lambda r, c: "D9E2F3" if r == len(syn) else (D.light_hex(D.voyant_info(syn[r]["voyant"])[0], .7) if c == 7 else None))
    d.text(p, .45, 3.72, 12.4, .28, NOTE_SYN, 8.5, False, MUTED, italic=True)
    d.pic(p, chart_modules(comp, D, (8.6, 2.9)), .45, 4.02, 8.2, 2.9)
    d.bullets(p, 8.95, 4.1, 3.93, 2.3, "À retenir", keep_lines(comp, D), 11.5, FOREST)

    p = d.band("Chiffres clés & points d’alerte", "Résumé décisionnel — actions requises", FOREST)
    bg = {"ok": "C6EFCE", "info": "E4EEF9", "bad": "FFCCCC"}
    pcards(p, [(c["label"].title(), P.fmt_int(c["value"]), bg[c["kind"]]) for c in comp["cards"]], 1.18, 1.15, 24, FOREST)
    al = comp["alerts"]
    d.table(p, .45, 2.5, 12.43, [3.6, 1.4, 1.5, 1.5, 1.5, 1.6, 1.3], ALERT_HEAD, alert_rows(comp, D), RED, 10.5, .36, fills=lambda r, c: D.ALERT_FILL[D.alert_key(al[r]["status"])] if c >= 5 else None)
    d.text(p, .45, 6.2, 12.4, .3, NOTE_ALERT, 9, False, MUTED, italic=True)

    for m in P.MODULES:
        mod = comp["modules"][m["key"]]
        th = m["theme"]
        unit = "opérateurs" if m["key"] == "identification" else "dossiers"
        p = d.band(f"KPI {m['label']}", f"Feuille « {m['sheet']} » — {P.fmt_int(mod['total'])} {unit} · voyant {D.voyant_info(mod['voyant'])[1]}", th)
        pcards(p, D.key_indicators(mod), 1.18, .95, 18, th)
        rows = D.status_table_rows(mod)
        n = len(rows)
        lst = lambda r, c, n=n: r == n - 1
        if m["key"] == "identification":
            ty = [[x["label"], P.fmt_int(x["count"]), P.fmt_pct(x["pct"])] for x in mod["types"]]
            ty.append(["TOTAL", P.fmt_int(mod["types_total"]), P.fmt_pct(1 if mod["types_total"] else 0)])
            d.table(p, .45, 2.35, 6.1, [3.2, 1, 1, 1.4], ["Statut", "Nombre", "%", "Catégorie"], [r[:4] for r in rows], th, 10.5, .32, fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bolds=lst)
            d.table(p, .45, 3.85, 6.1, [3.8, 1, 1], ["Type d'opérateur", "Nombre", "%"], ty, th, 10.5, .32, fills=lambda r, c, k=len(ty): "D9E2F3" if r == k - 1 else None, bolds=lambda r, c, k=len(ty): r == k - 1)
        else:
            d.table(p, .45, 2.35, 6.1, [3, 1, 1, 1.3], ["Statut", "Nombre", "%", "Catégorie"], [r[:4] for r in rows], th, 10 if n > 12 else 10.5, min(.3, 4.5 / (n + 1)),
                    fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bolds=lst)
        d.pic(p, D.module_chart(mod), 6.8, 2.3, 6.08, 3.15)
        d.bullets(p, 6.8, 5.5, 6.08, 1.42, "Constats", comp["insights"][m["key"]], 10.5, th, D.light_hex(th, .9))

    md = D.methodology(comp, meta, GEN)
    p = d.band("Méthodologie & qualité des données", "Sources, contrôles, coquilles normalisées, hypothèses des seuils", FOREST)
    d.bullets(p, .45, 1.18, 6.15, 3.15, "Sources", md["sources"], 9, FOREST)
    d.bullets(p, 6.73, 1.18, 6.15, 2.25, "Contrôles", md["controls"][:8], 9, FOREST)
    d.bullets(p, .45, 4.43, 6.15, 2.48, "Coquilles normalisées & statuts non référencés", md["typos"] + md["unknown"], 9, FOREST)
    d.bullets(p, 6.73, 3.53, 6.15, 3.38, "Hypothèses des seuils", [clean(h) for h in md["hypotheses"]], 8.5, FOREST)
    path = os.path.join(S.EX_DIR, "PCHC_reporting_metier_2026-10.pptx")
    d.finish(path)
    return path


MIX_CSS = f"""
@page {{ size:A4 landscape; margin:23mm 13mm 14mm; }}
@page cover {{ margin:0; }}
.cover {{ height:209.5mm; }}
.cv-top {{ padding:9mm 20mm 0; }} .cv-top img {{ height:24mm; }} .cv-rule {{ margin:5mm 0 0; }} .cv-main {{ padding:15mm 20mm 0; }}
.cv-meta {{ margin-top:8mm; }} .cv-foot {{ padding:6mm 20mm; }}
.stripe {{ margin:8pt 0 6pt; }} .stripe span {{ display:inline-block; width:27pt; height:5pt; margin-right:5pt; }}
.pg {{ break-before:page; }}
.band {{ display:flex; align-items:center; gap:9pt; color:#fff; padding:6pt 11pt; border-radius:4pt; margin-bottom:7pt; }}
.bi {{ width:22pt; height:22pt; flex:none; border-radius:4pt; background:rgba(255,255,255,.2); font:700 9pt/22pt Montserrat; text-align:center; }}
.bt {{ font:700 14pt/1.15 Montserrat; }} .bs {{ font-size:8.2pt; opacity:.9; }}
.split {{ display:grid; grid-template-columns:49fr 51fr; gap:11pt; }}
.split img {{ width:100%; display:block; }}
.syn {{ display:grid; grid-template-columns:66fr 34fr; gap:11pt; margin-top:6pt; align-items:start; }} .syn img {{ width:100%; display:block; }}
.nt {{ font-size:7.2pt; color:{'#' + MUTED}; font-style:italic; margin-top:3pt; }}
table {{ break-inside:avoid; }}
.rh td, .rh th {{ height:var(--rh); vertical-align:middle; }}
.cw {{ margin:0 -4px; }}
"""


def pdf_pchc_mix(comp, meta, D, label):
    P = D.R
    R.update(title="Reporting Métier PCHC", short="Reporting Métier PCHC", period=label)
    uri = lambda png: "data:image/png;base64," + base64.b64encode(png).decode()
    cards = lambda items, th, size: '<div class="cw">' + D._cards_html(items, th, size).replace("margin:0 -4px 10px", "margin:0 0 10px") + "</div>"
    band = lambda n, t, sub, th: f'<div class="band" style="background:#{th}"><span class="bi">{n:02d}</span><div><div class="bt">{esc(t)}</div><div class="bs">{esc(sub)}</div></div></div>'
    cover = G.cover_a("Reporting Métier", "PCHC", "Plateforme Tarkhiss — Produits Cosmétiques et d’Hygiène Corporelle",
                      [("Période", label), ("Situation au", D.fmt_dt(meta.get("imported_at"))), ("Édition", GEN.strftime("%d/%m/%Y"))], "Revue de direction")
    stripe = '<div class="stripe">' + "".join(f'<span style="background:#{m["theme"]}"></span>' for m in P.MODULES) + "</div>"
    body = cover.replace('<div class="cv-sub">', stripe + '<div class="cv-sub">') + G.head_html()
    body += (f'<div class="pg">{band(1, "Synthèse exécutive", f"Situation par volet — {label}", FOREST)}<div class="rh" style="--rh:25px">{D.synthesis_table_html(comp, 12)}</div><div class="nt">{esc(NOTE_SYN)}</div>'
             f'<div class="syn"><div><img src="{uri(chart_modules(comp, D, (8.6, 3.2)))}"></div><div>{D._box("À retenir", keep_lines(comp, D), FOREST, 12.5)}</div></div></div>')
    bg = {"ok": "C6EFCE", "info": "E4EEF9", "bad": "FFCCCC"}
    body += (f'<div class="pg">{band(2, "Chiffres clés & points d’alerte", "Résumé décisionnel — actions requises", FOREST)}'
             + cards([(c["label"], P.fmt_int(c["value"]), bg[c["kind"]]) for c in comp["cards"]], FOREST, 22)
             + '<div class="rh" style="--rh:32px">' + re.sub(r">(HAUTE|MOYENNE|BASSE)<", lambda x: f">{x.group(1).capitalize()}<", D.alerts_table_html(comp, 13).replace("C00000", RED)) + f'</div><div class="nt">{esc(NOTE_ALERT)}</div></div>')
    for i, m in enumerate(P.MODULES, 3):
        mod = comp["modules"][m["key"]]
        th = m["theme"]
        rows = D.status_table_rows(mod)
        tr = [[esc(r[0]), r[1], r[2], esc(r[3])] for r in rows]
        tbl = D._table(["Statut", "Nombre", "%", "Catégorie"], tr, th, [46, 16, 16, 22], font=12, fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bold_last=True)
        if m["key"] == "identification":
            ty = [[esc(x["label"]), P.fmt_int(x["count"]), P.fmt_pct(x["pct"])] for x in mod["types"]]
            ty.append(["TOTAL", P.fmt_int(mod["types_total"]), P.fmt_pct(1 if mod["types_total"] else 0)])
            tbl += '<div style="height:8px"></div>' + D._table(["Type d'opérateur", "Nombre", "%"], ty, th, [62, 19, 19], font=12, bold_last=True, fills=lambda r, c, n=len(ty): "D9E2F3" if r == n - 1 else None)
        nb = len(rows) + 1 + (len(mod["types"]) + 2 if m["key"] == "identification" else 0)
        chart = D.chart_types(mod, (6.4, 4.1)) if m["key"] == "identification" else D.chart_status(mod, (6.4, 4.5))
        unit = "opérateurs" if m["key"] == "identification" else "dossiers"
        body += (f'<div class="pg">{band(i, "KPI " + m["label"], f"Feuille « {m['sheet']} » — {P.fmt_int(mod['total'])} {unit} · voyant {D.voyant_info(mod['voyant'])[1]}", th)}'
                 + cards(D.key_indicators(mod), th, 17) + f'<div class="split"><div class="rh" style="--rh:{min(27, 440 // nb - 9)}px">{tbl}</div><div><img src="{uri(chart)}">'
                 + D._box("Constats", comp["insights"][m["key"]], th, 11.5) + "</div></div></div>")
    md = D.methodology(comp, meta, GEN)
    n = len(P.MODULES) + 3
    body += (f'<div class="pg">{band(n, "Méthodologie & qualité des données", "Sources, contrôles, coquilles normalisées, hypothèses des seuils", FOREST)}<div class="split" style="grid-template-columns:1fr 1fr">'
             f'<div>{D._box("Sources", md["sources"], FOREST, 10.5)}{D._box("Coquilles normalisées & statuts non référencés", md["typos"] + md["unknown"], FOREST, 10.5)}</div>'
             f'<div>{D._box("Contrôles", md["controls"], FOREST, 10.5)}{D._box("Hypothèses des seuils", [clean(h) for h in md["hypotheses"]], FOREST, 10.5)}</div></div></div>')
    S.render_pdf("PCHC_reporting_metier_2026-10.pdf", MIX_CSS, body)


def mail_pchc(app, comp, meta, D, label, attach):
    P = D.R
    ms = app.load_module_settings("pchc")
    t = comp["totals"]
    m = S.Mail()
    tone = {"ok": "good", "bad": "alert"}
    cards = [(c["label"].capitalize(), P.fmt_int(c["value"]), "", tone.get(c["kind"], "")) for c in comp["cards"]]
    crit = sum(1 for a in comp["alerts"] if "CRITIQUE" in a["status"])
    dot = lambda col, txt: f'<span style="color:#{col};">&#9632;</span> {esc(txt)}'
    syn = []
    for r, row in zip(comp["synthesis"], synth_rows(comp, D)):
        col, name = D.voyant_info(r["voyant"])
        syn.append([f"<b>{esc(row[0])}</b>", row[1], row[2], f"<b>{row[3]}</b>", row[4], row[5], dot(col, name)])
    tr = synth_rows(comp, D)[-1]
    syn.append({"cls": "tot", "cells": [f"<b>{tr[0]}</b>"] + [f"<b>{x}</b>" for x in tr[1:6]] + [""]})
    colors = {"Critique": RED, "Attention": "C98A1B", "OK": "2E8B57"}
    al = [[esc(r[0]), esc(r[1]), f"<b>{r[2]}</b>", r[3], r[4], dot(colors.get(r[5][2:], MUTED), r[5][2:])] for r in alert_rows(comp, D)]
    b = (m.sec(1, "Synthèse exécutive")
         + m.note(f"Situation des dossiers au {D.fmt_dt(meta.get('imported_at'))} : {P.fmt_int(t['total'])} dossiers, {P.fmt_pct(t['taux'])} délivrés/validés, "
                  f"{P.fmt_int(t['encours'])} en cours, {P.fmt_int(t['bloques'])} bloqués ou refusés — {crit} alerte(s) critique(s).")
         + m.kpis(cards[:3], 3) + m.kpis(cards[3:], 3)
         + m.table(["Module", "Total", "Délivrés", "Taux", "En cours", "Bloqués", "Voyant"], syn, ["left", "center", "center", "center", "center", "center", "left"])
         + m.img("data:image/png;base64," + base64.b64encode(chart_modules(comp, D, (8.2, 2.6))).decode(), "Dossiers par volet")
         + m.sec(2, "Points d’alerte")
         + m.table(["Indicateur", "Module", "Valeur", "Attention", "Critique", "Statut"], al, ["left", "left", "center", "center", "center", "left"])
         + m.note("Détail par volet, graphiques et méthodologie : voir les pièces jointes (PDF, PowerPoint)."))
    intro = f"Veuillez trouver ci-dessous le reporting métier PCHC ({label})."
    html = app.strip_platform_mentions(m.shell("Reporting Métier PCHC", label, ms.get("greeting", "Bonjour,"), intro, b, S.signature(ms), "Revue de direction"))
    m.save("PCHC_email_2026-10", f"Reporting Métier PCHC – Plateforme Tarkhiss – {label}", html, attach)


if __name__ == "__main__":
    G.build_fonts()
    G.TARKHISS_BASE = {k: R[k] for k in ("title", "short", "period")}
    app, work = sandbox.load_app()
    print("copie isolée :", work)
    import pchc_deliverables as D
    import pchc_raw as P
    patch_charts(D)
    with app.app.app_context(), app.app.test_request_context():
        deck_tarkhiss()
        deck_moussanada(S.moussanada_ctx(app))
        meta = app.load_pchc_meta()
        comp = P.pchc_compute(app.load_pchc_raw())
        label = D.period_label(GEN).capitalize()
        pptx_path = deck_pchc(comp, meta, D, label)
        pdf_pchc_mix(comp, meta, D, label)
        pdf_path = os.path.join(S.EX_DIR, "PCHC_reporting_metier_2026-10.pdf")
        attach = [(open(pdf_path, "rb").read(), "application/pdf", "Reporting_PCHC_2026-10.pdf"), (open(pptx_path, "rb").read(), D.PPTX_MIME, "Reporting_PCHC_2026-10.pptx")]
        mail_pchc(app, comp, meta, D, label, attach)
