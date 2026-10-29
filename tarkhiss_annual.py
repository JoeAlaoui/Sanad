# -*- coding: utf-8 -*-
"""
tarkhiss_annual.py — Rapport annuel (ou de période) Tarkhiss « Support Email & Hotline ».

Rendus purs (aucun accès disque/Flask) à partir des données préparées par app.py :
- summary_html()  : éléments de la « Vue annuelle » (KPI, tendance mensuelle, charge par jour de
                    semaine, top 3 des catégories, détail par mois) — corps de l'e-mail et résumé du PDF ;
- cover_html()    : page de garde du PDF ;
- build_pptx()    : garde, résumé de la période, puis 2 diapositives par mois (éléments de l'envoi mensuel) ;
- write_summary_sheet() : feuille Excel « Synthèse » (tableaux + graphique natif).
Charte identique au rapport mensuel (bleu #0B4965, cartes KPI, tableaux à en-tête bleu).
"""
import io
import os
from html import escape as esc

NAVY, RED, GREEN, AMBER, GREY, LINE, SOFT, INK = "0B4965", "DC2828", "25935F", "F59F0A", "55616B", "DAE0E7", "F5F7F9", "0D1926"
HEAT = [(0.15, "E7F5EE"), (0.35, "CFEBDD"), (0.55, "6FC49A"), (0.75, "25935F"), (0.9, "F59F0A"), (9, "DC2828")]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


# Blocs de la Vue annuelle qui reçoivent un descriptif rédigé (et leur prompt Copilot)
BLOCKS = [
    ("indicateurs", "Indicateurs & tendance mensuelle"),
    ("charge", "Charge par jour de la semaine"),
    ("top3", "Top 3 catégories récurrentes"),
    ("detail", "Détail par mois"),
    ("synthese", "Synthèse & conclusion"),
]
BLOCK_LABEL = dict(BLOCKS)
_PROMPT_RULES = ("Rédige en français, ton sobre et factuel, orienté pilotage (Direction). Texte brut, sans titre, "
                 "sans puces ni tableau, prêt à être collé tel quel dans le rapport. Appuie-toi uniquement sur les chiffres ci-dessous.")


def build_prompts(data):
    """Prompts prêts à copier (Copilot / ChatGPT), un par bloc, avec les chiffres de la période."""
    p, t, ms = data["period_label"], data["totals"], data["months"]
    lines = "\n".join(f"- {m['label']} : {m['emails_received']} emails reçus, {m['emails_sent']} envoyés, {m['calls']} appels, "
                      f"{m['bugs']} problèmes techniques, {m['demandes']} demandes d'information" for m in ms) or "- (aucune donnée)"
    tot = (f"Totaux de la période : {t['emails_received']} emails reçus, {t['emails_sent']} emails envoyés, {t['calls']} appels, "
           f"{t['bugs']} problèmes techniques (bugs), {t['demandes']} demandes d'information, sur {len(ms)} mois avec données.")
    hm = ", ".join(f"{h['day']} {h['avg']}" for h in data["weekday_heatmap"])
    top = "\n".join(f"{i + 1}. {c['label']} : {c['count']} signalements" for i, c in enumerate(data["top_categories"])) or "(aucune catégorie saisie)"
    head = f"Tu es analyste qualité du support Tarkhiss (Support Email & Hotline, AMMPS). Période : {p}.\n\n"
    return {
        "indicateurs": head + f"{tot}\n\nDétail mensuel :\n{lines}\n\n"
                       "Rédige un paragraphe de 4 à 6 phrases qui commente les indicateurs de la période et la tendance "
                       "mensuelle : volumes globaux, mois de pic et mois creux, évolution des appels par rapport aux emails, "
                       "évolution des bugs et des demandes.\n" + _PROMPT_RULES,
        "charge": head + f"Charge moyenne par jour de la semaine (appels + emails reçus, moyenne par jour travaillé) : {hm}.\n\n"
                  "Rédige un paragraphe de 3 à 4 phrases qui commente la répartition de la charge sur la semaine : jours "
                  "les plus chargés, écart avec les jours calmes, conséquence pour l'organisation de la hotline.\n" + _PROMPT_RULES,
        "top3": head + f"Top 3 des catégories de problèmes techniques les plus signalées sur la période :\n{top}\n\n"
                f"Total des problèmes techniques sur la période : {t['bugs']}.\n\n"
                "Rédige un paragraphe de 3 à 5 phrases qui commente ces catégories récurrentes : poids dans le total, "
                "nature probable du problème, priorité de traitement suggérée.\n" + _PROMPT_RULES,
        "detail": head + f"Détail par mois :\n{lines}\n\n"
                  "Rédige un paragraphe de 3 à 5 phrases qui met en évidence les mois atypiques (hausse ou baisse marquée), "
                  "le ratio problèmes/demandes et sa variation d'un mois à l'autre.\n" + _PROMPT_RULES,
        "synthese": head + f"{tot}\n\nDétail mensuel :\n{lines}\n\nCharge par jour : {hm}.\n\nTop 3 catégories :\n{top}\n\n"
                    "Rédige une synthèse de 4 à 6 phrases pour la Direction : bilan de la période, points de vigilance, puis "
                    "2 à 3 recommandations concrètes et priorisées (formulées dans le texte, sans puces).\n" + _PROMPT_RULES,
    }


def _note_box(text):
    """Descriptif rédigé sous un bloc (e-mail / PDF) ; vide si non renseigné."""
    if not (text or "").strip():
        return ""
    paras = "".join(f'<p style="margin:0 0 6px;">{esc(p)}</p>' for p in text.strip().split("\n") if p.strip())
    return (f'<div style="background:#F5F7F9;border-left:4px solid #{NAVY};padding:10px 14px;margin:8px 0 4px;'
            f'font-family:Arial,sans-serif;font-size:12.5px;line-height:1.55;color:#{INK};">{paras}</div>')


def heat_color(avg, mx):
    """Même échelle que la heatmap de la Vue annuelle (heatColor côté écran)."""
    if not mx:
        return HEAT[0][1]
    ratio = avg / mx
    return next(c for t, c in HEAT if ratio < t)


def fmt_int(n):
    return f"{int(n or 0):,}".replace(",", " ")


def cut(s, n):
    s = str(s or "")
    return s if len(s) <= n else s[:n - 1].rstrip() + "…"


# ----------------------------------------------------------------------------- HTML (e-mail / PDF)
def _kpi(label, value, bg=SOFT, color=NAVY, width="20%"):
    return (f'<td style="width:{width};padding:10px;text-align:center;background:#{bg};border:1px solid #{LINE};">'
            f'<div style="font-size:11px;color:#{GREY};">{esc(label)}</div>'
            f'<div style="font-size:18px;font-weight:bold;color:#{color};">{esc(str(value))}</div></td>')


def _th(label, align="left"):
    return f'<td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:{align};">{esc(label)}</td>'


def _td(value, align="left", color=INK, bold=False):
    return (f'<td style="padding:8px 10px;border-bottom:1px solid #{LINE};font-family:Arial,sans-serif;font-size:13px;'
            f'color:#{color};text-align:{align};{"font-weight:bold;" if bold else ""}">{value}</td>')


def summary_html(data, ctx, trend_uri, wrap_style, notes=None):
    """Rapport « Vue annuelle » (même design que le rapport mensuel). trend_uri : graphique PNG en data URI ;
    notes : descriptifs rédigés {bloc: texte}, affichés sous chaque bloc."""
    notes = notes or {}
    t = data["totals"]
    months = data["months"]
    hm = data["weekday_heatmap"]
    mx = max([h["avg"] for h in hm] + [0])
    heat = "".join(
        f'<td style="width:14.28%;padding:10px 4px;text-align:center;background:#{heat_color(h["avg"], mx)};'
        f'color:#{"FFFFFF" if heat_color(h["avg"], mx) in ("25935F", "F59F0A", "DC2828") else INK};font-size:12px;font-weight:bold;'
        f'border:2px solid #FFFFFF;">{esc(h["day"][:3])}<br><span style="font-size:15px;">{h["avg"]}</span></td>' for h in hm)
    # Pas d'emoji dans le rapport : le moteur PDF serveur n'a pas de police emoji (carrés vides)
    top = "".join(f"<tr>{_td(f'<b>{i + 1}.</b> ' + esc(c['label']))}{_td(fmt_int(c['count']), 'center', bold=True)}</tr>"
                  for i, c in enumerate(data["top_categories"]))
    rows = "".join(f"<tr>{_td(esc(m['label']))}{_td(fmt_int(m['emails_received']), 'center')}{_td(fmt_int(m['emails_sent']), 'center')}"
                   f"{_td(fmt_int(m['calls']), 'center')}{_td(fmt_int(m['bugs']), 'center', RED, True)}{_td(fmt_int(m['demandes']), 'center', GREEN, True)}</tr>"
                   for m in months)
    total_row = (f'<tr style="background:#{SOFT};">{_td("<b>Total période</b>")}{_td(fmt_int(t["emails_received"]), "center", bold=True)}'
                 f'{_td(fmt_int(t["emails_sent"]), "center", bold=True)}{_td(fmt_int(t["calls"]), "center", bold=True)}'
                 f'{_td(fmt_int(t["bugs"]), "center", RED, True)}{_td(fmt_int(t["demandes"]), "center", GREEN, True)}</tr>')
    section = lambda title: f'<div style="font-size:14px;font-weight:bold;color:#{NAVY};margin:18px 0 8px;">{esc(title)}</div>'
    # Image centrée : cellule align=center (Outlook) + margin auto (navigateurs, WeasyPrint)
    trend = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:6px 0 0;"><tr>'
             f'<td align="center" style="text-align:center;padding:0;"><img src="{trend_uri}" style="max-width:100%;border-radius:8px;'
             f'border:1px solid #{LINE};display:block;margin:0 auto;"></td></tr></table>'
             f'<div style="font-size:11.5px;color:#{GREY};text-align:center;margin:4px 0 6px;">Évolution mensuelle des emails, appels, bugs et demandes sur la période sélectionnée.</div>'
             if trend_uri else f'<p style="font-size:12px;color:#{GREY};">Aucune donnée sur la période.</p>')
    return f"""
    <div style="{wrap_style}">
      <div style="background:#{NAVY};padding:22px 24px;border-radius:6px 6px 0 0;">
        {ctx.get('logos_html', '')}
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">{esc(ctx['title'])}</div>
        <div style="color:#B9D3E0;font-size:13px;margin-top:4px;">{esc(ctx['subtitle'])}</div>
      </div>
      <div style="padding:20px 24px;border:1px solid #{LINE};border-top:none;">
        <table style="width:100%;border-collapse:collapse;margin-bottom:14px;"><tr>
          {_kpi("EMAILS REÇUS", fmt_int(t["emails_received"]))}
          {_kpi("EMAILS ENVOYÉS", fmt_int(t["emails_sent"]))}
          {_kpi("APPELS", fmt_int(t["calls"]))}
          {_kpi("BUGS", fmt_int(t["bugs"]), "FCE9E9", RED)}
          {_kpi("DEMANDES", fmt_int(t["demandes"]), "E7F5EE", GREEN)}
        </tr></table>
        {_note_box(notes.get("indicateurs"))}
        {section("Tendance mensuelle")}
        {trend}
        {section("Charge moyenne par jour de la semaine (appels + emails reçus)")}
        <table style="width:100%;border-collapse:collapse;"><tr>{heat}</tr></table>
        {_note_box(notes.get("charge"))}
        {section("Top 3 catégories récurrentes sur la période")}
        {f'''<table style="width:100%;border-collapse:collapse;border:1px solid #{LINE};">
          <tr style="background:#{NAVY};">{_th("Catégorie")}{_th("Total signalé sur la période", "center")}</tr>{top}</table>'''
          if top else f'<p style="font-size:12px;color:#{GREY};">Aucune donnée d\'analyse pour la période.</p>'}
        {_note_box(notes.get("top3"))}
        {section("Détail par mois")}
        <table style="width:100%;border-collapse:collapse;border:1px solid #{LINE};">
          <tr style="background:#{NAVY};">{_th("Mois")}{_th("Emails reçus", "center")}{_th("Emails envoyés", "center")}{_th("Appels", "center")}{_th("Bugs", "center")}{_th("Demandes", "center")}</tr>
          {rows or f'<tr>{_td("Aucune donnée", color=GREY)}</tr>'}{total_row if months else ""}
        </table>
        {_note_box(notes.get("detail"))}
        {(section("Synthèse & conclusion") + _note_box(notes.get("synthese"))) if (notes.get("synthese") or "").strip() else ""}
      </div>
    </div>"""


def cover_html(ctx):
    """Page de garde du PDF (pleine page, sans en-tête ni numéro)."""
    logo = f'<img src="{ctx["logo_uri"]}" style="height:70px;background:#fff;padding:8px;border-radius:6px;">' if ctx.get("logo_uri") else ""
    return f"""
    <div style="background:#{NAVY};color:#FFFFFF;height:29.6cm;box-sizing:border-box;padding:4.5cm 2.2cm 2cm;font-family:Arial,sans-serif;">
      <div style="width:90px;height:6px;background:#{AMBER};margin-bottom:26px;"></div>
      <div style="font-size:32px;font-weight:bold;line-height:1.2;">{esc(ctx['title'])}</div>
      <div style="font-size:17px;color:#B9D3E0;margin-top:10px;">Support Email &amp; Hotline — synthèse de la période et rapports mensuels</div>
      <div style="font-size:15px;margin-top:46px;">Période : <b>{esc(ctx['period_label'])}</b></div>
      <div style="font-size:13px;color:#B9D3E0;margin-top:6px;">{ctx['months_count']} mois avec données</div>
      <table style="width:100%;margin-top:9cm;border-top:3px solid #{AMBER};"><tr>
        <td style="color:#D6DCE4;font-size:13px;padding-top:14px;">{esc(ctx['agency'])}<br>Document généré le {ctx['generated'].strftime('%d/%m/%Y à %H:%M')}</td>
        <td style="text-align:right;padding-top:14px;">{logo}</td></tr></table>
    </div>"""


# ----------------------------------------------------------------------------- PPTX
def build_pptx(ctx, data, months, trend_png, notes=None):
    """months : [{label, kpis:{...}, problems, demandes, weekly, constats, recommandations,
    chart_weekly (PNG), chart_problems (PNG)}] — mêmes éléments que l'envoi mensuel.
    notes : descriptifs {bloc: texte} → diapositives « Analyse de la période » (2 blocs par diapositive)."""
    notes = notes or {}
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Emu, Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W, H = prs.slide_width, prs.slide_height
    blank = prs.slide_layouts[6]
    rgb = RGBColor.from_string

    def rect(s, x, y, w, h, color, line=None, shape=MSO_SHAPE.RECTANGLE):
        r = s.shapes.add_shape(shape, x, y, w, h)
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(color)
        if line:
            r.line.color.rgb = rgb(line)
            r.line.width = Pt(0.75)
        else:
            r.line.fill.background()
        r.shadow.inherit = False
        return r

    def text(s, x, y, w, h, value, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
        tb = s.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        for i, line in enumerate(value if isinstance(value, list) else [value]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = str(line)
            r.font.size, r.font.bold, r.font.italic, r.font.name = Pt(size), bold, italic, "Arial"
            r.font.color.rgb = rgb(color)
        return tb

    def header(s, title, subtitle):
        rect(s, 0, 0, W, Inches(0.95), NAVY)
        rect(s, 0, Inches(0.95), W, Inches(0.06), AMBER)
        text(s, Inches(0.45), Inches(0.12), Inches(11.5), Inches(0.5), title, 24, True, "FFFFFF")
        text(s, Inches(0.45), Inches(0.57), Inches(12), Inches(0.35), subtitle, 12, False, "B9D3E0")

    def table(s, x, y, w, col_w, rows, head, font=10, row_h=0.28, fills=None, colors=None, bold_last=False):
        n, ncol = len(rows) + 1, len(head)
        tbl = s.shapes.add_table(n, ncol, x, y, w, Inches(row_h * n)).table
        tbl.first_row, tbl.horz_banding = True, False
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = Emu(int(w * cw / sum(col_w)))
        for r, row in enumerate([head] + rows):
            tbl.rows[r].height = Inches(row_h)
            for c in range(ncol):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(0.06)
                cell.margin_top = cell.margin_bottom = Inches(0.01)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                cell.text_frame.word_wrap = True
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
                run = p.add_run()
                run.text = str(row[c] if c < len(row) else "")
                run.font.size, run.font.name = Pt(font), "Arial"
                run.font.bold = r == 0 or (bold_last and r == n - 1)
                run.font.color.rgb = rgb("FFFFFF" if r == 0 else ((colors(r - 1, c) if colors else None) or INK))
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(NAVY if r == 0 else ((fills(r - 1, c) if fills else None) or ("FFFFFF" if r % 2 else "F8FAFC")))

    def cards(s, items, x, y, w, h, vsize=20):
        gap = Inches(0.12)
        cw = int((w - gap * (len(items) - 1)) / len(items))
        for i, (lab, val, bg, col) in enumerate(items):
            cx = x + i * (cw + gap)
            rect(s, cx, y, cw, h, bg, line=LINE)
            text(s, cx, y + Inches(0.06), cw, Inches(0.3), lab.upper(), 9, True, GREY, PP_ALIGN.CENTER)
            text(s, cx, y + Inches(0.3), cw, h - Inches(0.34), val, vsize, True, col, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

    def picture(s, png, x, y, w, h):
        """Image ajustée au cadre (w × h) en conservant ses proportions, centrée dans le cadre."""
        if not png:
            return
        from PIL import Image
        iw, ih = Image.open(io.BytesIO(png)).size
        scale = min(w / iw, h / ih)
        pw, ph = int(iw * scale), int(ih * scale)
        s.shapes.add_picture(io.BytesIO(png), x + (w - pw) // 2, y + (h - ph) // 2, width=pw, height=ph)

    def box(s, x, y, w, h, title, lines, size=10.5):
        rect(s, x, y, w, h, "F2F6F9", line=LINE)
        rect(s, x, y, Inches(0.07), h, NAVY)
        text(s, x + Inches(0.15), y + Inches(0.06), w - Inches(0.2), Inches(0.3), title, 11.5, True, NAVY)
        text(s, x + Inches(0.15), y + Inches(0.38), w - Inches(0.25), h - Inches(0.42), [f"• {l}" for l in lines] or ["—"], size)

    # 1. Page de garde
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, H, NAVY)
    rect(s, Inches(0.8), Inches(1.45), Inches(1.0), Inches(0.08), AMBER)
    text(s, Inches(0.8), Inches(1.7), Inches(11.5), Inches(1.0), ctx["title"], 36, True, "FFFFFF")
    text(s, Inches(0.8), Inches(2.75), Inches(11.5), Inches(0.5), "Support Email & Hotline — synthèse de la période et rapports mensuels", 17, False, "B9D3E0")
    text(s, Inches(0.8), Inches(3.6), Inches(11.5), Inches(0.9), [f"Période : {ctx['period_label']}", f"{ctx['months_count']} mois avec données"], 15, False, "FFFFFF")
    rect(s, 0, Inches(5.0), W, Inches(0.06), AMBER)
    text(s, Inches(0.8), Inches(5.3), Inches(8), Inches(0.9), [ctx["agency"], f"Document généré le {ctx['generated'].strftime('%d/%m/%Y à %H:%M')}"], 13, False, "D6DCE4")
    if ctx.get("logo_path") and os.path.isfile(ctx["logo_path"]) and not ctx["logo_path"].lower().endswith(".svg"):
        try:
            rect(s, Inches(10.35), Inches(5.25), Inches(2.3), Inches(1.4), "FFFFFF")
            s.shapes.add_picture(ctx["logo_path"], Inches(10.5), Inches(5.35), height=Inches(1.2))
        except Exception:
            pass

    # 2. Résumé de la période : KPI + tendance + top 3
    t = data["totals"]
    s = prs.slides.add_slide(blank)
    header(s, "Synthèse de la période", f"{ctx['period_label']} — vue annuelle")
    cards(s, [("Emails reçus", fmt_int(t["emails_received"]), SOFT, NAVY), ("Emails envoyés", fmt_int(t["emails_sent"]), SOFT, NAVY),
              ("Appels", fmt_int(t["calls"]), SOFT, NAVY), ("Bugs", fmt_int(t["bugs"]), "FCE9E9", RED),
              ("Demandes", fmt_int(t["demandes"]), "E7F5EE", GREEN)], Inches(0.45), Inches(1.2), Inches(12.4), Inches(1.05), 22)
    picture(s, trend_png, Inches(0.45), Inches(2.45), Inches(7.7), Inches(4.4))
    top = data["top_categories"]
    text(s, Inches(8.4), Inches(2.45), Inches(4.4), Inches(0.35), "Top 3 catégories récurrentes", 13, True, NAVY)
    if top:
        table(s, Inches(8.4), Inches(2.85), Inches(4.45), [3.2, 1.2], [[f"{i + 1}. {cut(c['label'], 40)}", fmt_int(c["count"])] for i, c in enumerate(top)],
              ["Catégorie", "Total"], 11, 0.38)
    else:
        text(s, Inches(8.4), Inches(2.85), Inches(4.4), Inches(0.4), "Aucune donnée d'analyse sur la période.", 11, False, GREY)

    # 3. Détail par mois + charge par jour de semaine
    s = prs.slides.add_slide(blank)
    header(s, "Détail par mois & charge hebdomadaire", ctx["period_label"])
    ms = data["months"]
    rows = [[m["label"], fmt_int(m["emails_received"]), fmt_int(m["emails_sent"]), fmt_int(m["calls"]), fmt_int(m["bugs"]), fmt_int(m["demandes"])] for m in ms]
    rows.append(["Total période", fmt_int(t["emails_received"]), fmt_int(t["emails_sent"]), fmt_int(t["calls"]), fmt_int(t["bugs"]), fmt_int(t["demandes"])])
    rh = min(0.34, 4.3 / (len(rows) + 1))
    table(s, Inches(0.45), Inches(1.2), Inches(12.4), [2.6, 1.6, 1.6, 1.4, 1.2, 1.4], rows,
          ["Mois", "Emails reçus", "Emails envoyés", "Appels", "Bugs", "Demandes"], 10.5 if len(rows) < 12 else 9.5, rh,
          fills=lambda r, c, n=len(rows): "E8EEF3" if r == n - 1 else None,
          colors=lambda r, c: RED if c == 4 else (GREEN if c == 5 else None), bold_last=True)
    hm = data["weekday_heatmap"]
    mx = max([h["avg"] for h in hm] + [0])
    y = Inches(1.35 + rh * (len(rows) + 1))
    text(s, Inches(0.45), y, Inches(12), Inches(0.35), "Charge moyenne par jour de la semaine (appels + emails reçus)", 12, True, NAVY)
    cw = Inches(12.4 / 7)
    for i, h in enumerate(hm):
        col = heat_color(h["avg"], mx)
        rect(s, Inches(0.45) + i * cw, y + Inches(0.4), cw - Inches(0.05), Inches(0.75), col)
        text(s, Inches(0.45) + i * cw, y + Inches(0.42), cw - Inches(0.05), Inches(0.72), [h["day"][:3], str(h["avg"])], 12, True,
             "FFFFFF" if col in ("25935F", "F59F0A", "DC2828") else INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

    # Analyse de la période : descriptifs rédigés, deux blocs par diapositive (texte long → coupé à 1 500 caractères)
    filled = [(k, lab, notes[k].strip()) for k, lab in BLOCKS if (notes.get(k) or "").strip()]
    for i in range(0, len(filled), 2):
        s = prs.slides.add_slide(blank)
        header(s, "Analyse de la période", ctx["period_label"])
        for j, (k, lab, txt) in enumerate(filled[i:i + 2]):
            x = Inches(0.45 + j * 6.3)
            rect(s, x, Inches(1.2), Inches(6.1), Inches(5.65), "F2F6F9", line=LINE)
            rect(s, x, Inches(1.2), Inches(0.07), Inches(5.65), NAVY)
            text(s, x + Inches(0.18), Inches(1.28), Inches(5.8), Inches(0.4), lab, 14, True, NAVY)
            text(s, x + Inches(0.18), Inches(1.75), Inches(5.75), Inches(5.0),
                 [p for p in cut(txt, 1500).split("\n") if p.strip()], 11)

    # 4. Deux diapositives par mois (éléments de l'envoi mensuel)
    for m in months:
        k = m["kpis"]
        s = prs.slides.add_slide(blank)
        header(s, f"{m['label']} — indicateurs", "Rapport mensuel du support Tarkhiss")
        cards(s, [("Emails reçus", fmt_int(k["received"]), SOFT, NAVY), ("Emails envoyés", fmt_int(k["sent"]), SOFT, NAVY),
                  ("Appels", fmt_int(k["calls"]), SOFT, NAVY), ("Temps comm.", k["talk_time"], SOFT, NAVY)],
              Inches(0.45), Inches(1.15), Inches(12.4), Inches(0.9), 18)
        cards(s, [("Problèmes tech.", fmt_int(k["problems"]), "FCE9E9", RED), ("Demandes info", fmt_int(k["demandes"]), "E7F5EE", GREEN),
                  ("Ratio bugs/demandes", k["ratio"], SOFT, NAVY), ("Semaine critique", cut(k["critical_week"], 18), "FEF3E0", "B9791F")],
              Inches(0.45), Inches(2.15), Inches(12.4), Inches(0.9), 16)
        picture(s, m.get("chart_weekly"), Inches(0.45), Inches(3.2), Inches(6.1), Inches(3.7))
        picture(s, m.get("chart_problems"), Inches(6.75), Inches(3.2), Inches(6.1), Inches(3.7))
        if not m.get("chart_weekly") and not m.get("chart_problems"):
            text(s, Inches(0.45), Inches(3.4), Inches(12), Inches(0.4), "Aucune analyse saisie pour ce mois.", 12, False, GREY)

        s = prs.slides.add_slide(blank)
        header(s, f"{m['label']} — analyse", "Problèmes, demandes, évolution hebdomadaire, constats")
        for x, title, items, color in ((Inches(0.45), "Problèmes techniques", m["problems"], RED), (Inches(6.75), "Demandes d'information", m["demandes"], GREEN)):
            tot = sum(i.get("count", 0) for i in items)
            rows = [[cut(i.get("label", ""), 46), fmt_int(i.get("count", 0)), f"{(i.get('count', 0) / tot * 100 if tot else 0):.1f} %".replace(".", ",")]
                    for i in sorted(items, key=lambda i: -i.get("count", 0))[:7]]
            text(s, x, Inches(1.15), Inches(6.1), Inches(0.35), f"{title} ({fmt_int(tot)})", 12.5, True, color)
            if rows:
                table(s, x, Inches(1.5), Inches(6.1), [4.0, 0.9, 1.0], rows, ["Type", "Nb", "%"], 10, 0.3)
            else:
                text(s, x, Inches(1.5), Inches(6.1), Inches(0.35), "Aucune donnée.", 11, False, GREY)
        wk = [[cut(w.get("week", ""), 22), fmt_int(w.get("bugs", 0)), fmt_int(w.get("demandes", 0))] for w in m["weekly"][:6]]
        text(s, Inches(0.45), Inches(4.05), Inches(6.1), Inches(0.35), "Évolution hebdomadaire", 12.5, True, NAVY)
        if wk:
            table(s, Inches(0.45), Inches(4.4), Inches(6.1), [3.0, 1.4, 1.4], wk, ["Semaine", "Bugs", "Demandes"], 10, 0.3,
                  colors=lambda r, c: RED if c == 1 else (GREEN if c == 2 else None))
        lines = [cut(c, 120) for c in m["constats"][:4]] + [cut("Reco. : " + r, 120) for r in m["recommandations"][:3]]
        box(s, Inches(6.75), Inches(4.05), Inches(6.1), Inches(2.8), "Constats & recommandations", lines, 10)

    n = len(prs.slides)
    for i, slide in enumerate(prs.slides, start=1):
        if i == 1:
            continue
        rect(slide, Inches(0.45), Inches(7.02), Inches(12.4), Emu(9525), LINE)
        text(slide, Inches(0.45), Inches(7.07), Inches(10.5), Inches(0.3), f"{ctx['agency']} · {ctx['title']} · {ctx['period_label']}", 9, False, GREY)
        text(slide, Inches(11.35), Inches(7.07), Inches(1.5), Inches(0.3), f"{i} / {n}", 9, True, NAVY, PP_ALIGN.RIGHT)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------- Excel (xlsxwriter)
def write_summary_sheet(wb, fmts, data, ctx, notes=None):
    """Feuille « Synthèse » : KPI de la période, détail par mois (+ graphique natif), top 3, charge par jour,
    puis les descriptifs rédigés (notes {bloc: texte})."""
    notes = notes or {}
    ws = wb.add_worksheet("Synthèse")
    ws.set_column("A:A", 30)
    ws.set_column("B:F", 15)
    ws.merge_range("A1:F1", f"{ctx['title']} — {ctx['period_label']}", fmts["title"])
    ws.write(1, 0, ctx["agency"])
    t = data["totals"]
    row = 3
    for lab, key in (("Emails reçus", "emails_received"), ("Emails envoyés", "emails_sent"), ("Appels", "calls"), ("Bugs", "bugs"), ("Demandes", "demandes")):
        ws.write(row, 0, lab, fmts["bold"])
        ws.write(row, 1, t[key])
        row += 1
    row += 1
    ws.write(row, 0, "Détail par mois", fmts["title"])
    row += 1
    head_row = row
    for i, h in enumerate(["Mois", "Emails reçus", "Emails envoyés", "Appels", "Bugs", "Demandes"]):
        ws.write(row, i, h, fmts["header"])
    for m in data["months"]:
        row += 1
        for i, k in enumerate(("label", "emails_received", "emails_sent", "calls", "bugs", "demandes")):
            ws.write(row, i, m[k])
    last = row
    if data["months"]:
        row += 1
        ws.write(row, 0, "Total période", fmts["bold"])
        for i, k in enumerate(("emails_received", "emails_sent", "calls", "bugs", "demandes"), start=1):
            col = chr(65 + i)
            ws.write_formula(row, i, f"=SUM({col}{head_row + 2}:{col}{last + 1})", fmts["bold"], t[k])
        ch = wb.add_chart({"type": "line"})
        for i, (name, color) in enumerate((("Emails reçus", "#0B4965"), (None, None), ("Appels", "#F59F0A"), ("Bugs", "#DC2828"), ("Demandes", "#25935F")), start=1):
            if not name:
                continue
            ch.add_series({"name": name, "categories": ["Synthèse", head_row + 1, 0, last, 0],
                           "values": ["Synthèse", head_row + 1, i, last, i], "line": {"color": color, "width": 2.25},
                           "marker": {"type": "circle", "size": 5, "fill": {"color": color}, "border": {"color": color}}})
        ch.set_title({"name": "Tendance mensuelle", "name_font": {"color": "#0B4965", "size": 12}})
        ch.set_legend({"position": "bottom"})
        ch.set_size({"width": 620, "height": 300})
        ws.insert_chart("H3", ch)
    row += 2
    ws.write(row, 0, "Top 3 catégories récurrentes", fmts["title"])
    row += 1
    ws.write(row, 0, "Catégorie", fmts["header"])
    ws.write(row, 1, "Total", fmts["header"])
    for c in data["top_categories"]:
        row += 1
        ws.write(row, 0, c["label"])
        ws.write(row, 1, c["count"])
    row += 2
    ws.write(row, 0, "Charge moyenne par jour de la semaine (appels + emails reçus)", fmts["title"])
    row += 1
    hm = data["weekday_heatmap"]
    mx = max([h["avg"] for h in hm] + [0])
    for i, h in enumerate(hm):
        col = heat_color(h["avg"], mx)
        f = wb.add_format({"bg_color": f"#{col}", "align": "center", "bold": True, "border": 1, "border_color": "#FFFFFF",
                           "font_color": "#FFFFFF" if col in ("25935F", "F59F0A", "DC2828") else f"#{INK}"})
        ws.write(row, i, h["day"][:3], f)
        ws.write(row + 1, i, h["avg"], f)
    filled = [(lab, notes[k].strip()) for k, lab in BLOCKS if (notes.get(k) or "").strip()]
    if filled:
        row += 3
        ws.write(row, 0, "Analyse de la période", fmts["title"])
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        for lab, txt in filled:
            row += 1
            ws.write(row, 0, lab, wb.add_format({"bold": True, "valign": "top", "font_color": "#0B4965"}))
            ws.merge_range(row, 1, row, 6, txt, wrap)
            ws.set_row(row, max(30, 15 * (len(txt) // 95 + txt.count("\n") + 1)))
    return ws
