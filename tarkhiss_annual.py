# -*- coding: utf-8 -*-
"""
tarkhiss_annual.py — Rapport annuel (ou de période) Tarkhiss « Support Email & Hotline ».

Rendus purs (aucun accès disque/Flask) à partir des données préparées par app.py. Charte AMMPS,
comme les rapports mensuels (months = contextes rapports_mensuels.tarkhiss_ctx) :
- build_pdf()     : garde, sommaire, résumé de la période, puis chaque mois ;
- build_pptx()    : garde, résumé de la période, analyse, puis 2 diapositives par mois ;
- mail_html()     : éléments de la « Vue annuelle » (KPI, tendance mensuelle, charge par jour de
                    semaine, top 3 des catégories, détail par mois) ;
- write_summary_sheet() : feuille Excel « Synthèse » (tableaux + graphique natif).
summary_html() et cover_html() : ancien rendu, repli du PDF quand WeasyPrint est indisponible.
"""
import ammps_charte as C
import rapports_mensuels as M
from ammps_charte import Deck, Mail, esc

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


# ----------------------------------------------------------------------------- charte AMMPS
DETAIL_HEAD = ["Mois", "Emails reçus", "Emails envoyés", "Appels", "Bugs", "Demandes"]
MONTH_HEAD = ["Type", "Nb", "%"]
SUBTITLE = "Synthèse de la période et rapports mensuels"
NO_ANALYSIS = "Aucune analyse saisie pour ce mois."


def title_parts(ctx):
    """« Rapport annuel — Support Tarkhiss » → les deux lignes de la page de garde."""
    first, _, second = ctx["title"].partition(" — ")
    return first, second


def period_charts(data):
    ms, hm = data["months"], data["weekday_heatmap"]
    multi = len({m["label"].split()[-1] for m in ms}) > 1
    labs = [m["label"].split()[0][:4] + "." + (" " + m["label"].split()[-1][2:] if multi else "") for m in ms]
    series = {"Emails reçus": [m["emails_received"] for m in ms], "Appels": [m["calls"] for m in ms],
              "Bugs": [m["bugs"] for m in ms], "Demandes": [m["demandes"] for m in ms]}
    trend = lambda size: C.lines(labs, series, figsize=size, colors=(C.FOREST, C.GREEN, C.RED, C.TINT))
    heat = lambda size: C.heat([[h["avg"] for h in hm]], [""], [h["day"][:3] for h in hm], size)
    makers = {}
    if ms:
        makers.update(trend=lambda: trend((8.2, 2.9)), trend_slide=lambda: trend((7.6, 4.2)))
    if any(h["avg"] for h in hm):
        makers.update(heat=lambda: heat((8.2, 1.1)), heat_slide=lambda: heat((12.2, 1.2)))
    return M.Charts(makers)


def detail_rows(data):
    n = fmt_int
    red = lambda v: f'<b style="color:{C.RED};">{n(v)}</b>'
    green = lambda v: f'<b style="color:{C.GREEN};">{n(v)}</b>'
    rows = [[esc(m["label"]), n(m["emails_received"]), n(m["emails_sent"]), n(m["calls"]), red(m["bugs"]), green(m["demandes"])] for m in data["months"]]
    t = data["totals"]
    if rows:
        rows.append({"cls": "tot", "cells": ["Total période", n(t["emails_received"]), n(t["emails_sent"]), n(t["calls"]), n(t["bugs"]), n(t["demandes"])]})
    return rows


def period_kpis(t):
    return [("Emails reçus", fmt_int(t["emails_received"]), ""), ("Emails envoyés", fmt_int(t["emails_sent"]), ""), ("Appels", fmt_int(t["calls"]), ""),
            ("Bugs", fmt_int(t["bugs"]), "", "alert"), ("Demandes", fmt_int(t["demandes"]), "", "good")]


def top_rows(data):
    return [[f"<b>{i}.</b> {esc(c['label'])}", f"<b>{fmt_int(c['count'])}</b>"] for i, c in enumerate(data["top_categories"], 1)]


def sorted_items(items, limit=7):
    return sorted(items, key=lambda i: -i[1])[:limit]


def month_pdf(c, n, anchor):
    ch = c["charts"]
    rows = lambda items: [[esc(l), f"<b>{k}</b>", p] for l, k, p, _ in sorted_items(items)]
    block = lambda title, items: f"<div><h4>{esc(title)}</h4>" + (C.table(MONTH_HEAD, rows(items), ["l", "c w11", "c w14"], "tight") if items else f'<p class="note">Aucune donnée.</p>') + "</div>"
    out = f"""<div class="newp">{C.sec(n, c['period'], 'Rapport mensuel', anchor)}<h4>Volumes</h4>{C.kpi_grid(c['kpi_volume'])}<h4>Qualification</h4>{C.kpi_grid(c['kpi_quali'])}"""
    if c["weekly"] or c["problems"]:
        out += f'<div class="two">{C.fig(ch["weekly_month"], "Bugs vs Demandes", "Semaine critique en rouge.")}{C.fig(ch["problems_month"], "Problèmes techniques", "Répartition par type.")}</div>'
    else:
        out += f'<p class="note">{NO_ANALYSIS}</p>'
    out += "</div>"
    if c["problems"] or c["demandes"] or c["constats"] or c["recos"]:
        out += f"""<div class="gap"><div class="two">{block(f"Problèmes techniques ({c['nb_prob']})", c['problems'])}{block(f"Demandes d’information ({c['nb_dem']})", c['demandes'])}</div>"""
        if c["constats"]:
            out += f"<h3>Constats clés</h3>{C.points(c['constats'])}"
        if c["recos"]:
            out += f'<div class="keep"><h3>Recommandations</h3>{C.recos(c["recos"])}</div>'
        out += "</div>"
    return out


def build_pdf(ctx, data, months, notes=None):
    notes = notes or {}
    ch = period_charts(data)
    first, second = title_parts(ctx)
    items = [("a1", "Indicateurs & tendance mensuelle"), ("a2", "Charge par jour de la semaine"), ("a3", "Top 3 catégories récurrentes"), ("a4", "Détail par mois")]
    synth = (notes.get("synthese") or "").strip()
    if synth:
        items.append(("a5", "Synthèse & conclusion"))
    base = len(items)
    items += [(f"m{i}", c["period"]) for i, c in enumerate(months)]
    body = C.cover(first, second, SUBTITLE, [("Période", ctx["period_label"]), ("Mois avec données", str(ctx["months_count"])), ("Édition", ctx["generated"].strftime("%d/%m/%Y"))], first)
    body += f"""{C.head_html(ctx['title'], ctx['period_label'])}<div>{C.sec(0, 'Sommaire', first)}{C.toc(items)}</div>
    <div class="newp">{C.sec(1, items[0][1], 'Vue d’ensemble', 'a1')}{C.kpi_grid(period_kpis(data['totals']), 5)}{C.notebox(notes.get('indicateurs'))}
      <div class="gap">{C.fig(ch['trend'], 'Tendance mensuelle', 'Évolution mensuelle des emails, appels, bugs et demandes sur la période.')}</div></div>
    <div class="gap keep">{C.sec(2, items[1][1], 'Rythme', 'a2')}{C.fig(ch['heat'], 'Charge moyenne par jour de la semaine (appels + emails reçus)', 'Moyenne par jour travaillé.')}{C.notebox(notes.get('charge'))}</div>
    <div class="gap keep">{C.sec(3, items[2][1], 'Qualification', 'a3')}{C.table(['Catégorie', 'Total signalé sur la période'], top_rows(data), ['l', 'c']) if data['top_categories'] else '<p class="note">Aucune donnée d’analyse pour la période.</p>'}{C.notebox(notes.get('top3'))}</div>
    <div class="gap">{C.sec(4, items[3][1], 'Consolidation', 'a4')}{C.table(DETAIL_HEAD, detail_rows(data), ['l', 'c', 'c', 'c', 'c', 'c'])}{C.notebox(notes.get('detail'))}</div>"""
    if synth:
        body += f"""<div class="gap keep">{C.sec(5, items[4][1], 'Synthèse', 'a5')}{C.notebox(synth, 'Conclusion')}</div>"""
    for i, c in enumerate(months):
        body += month_pdf(c, base + i + 1, f"m{i}")
    return C.render_pdf(ctx["title"], "Tarkhiss", ctx["period_label"], ctx["clean"](body), ".c5 {grid-template-columns:repeat(5,1fr);}")


def build_pptx(ctx, data, months, notes=None):
    """months : contextes mensuels ; notes : descriptifs {bloc: texte} → diapositives « Analyse de la période »."""
    notes = notes or {}
    per = ctx["period_label"]
    ch = period_charts(data)
    first, second = title_parts(ctx)
    d = Deck(ctx["title"], per, f"AMMPS · Tarkhiss — {per}", ctx["clean"])
    d.cover(first, first, second, SUBTITLE, [("Période", per), ("Mois avec données", str(ctx["months_count"]))])

    t = data["totals"]
    p = d.page(1, "Synthèse de la période", "Vue d’ensemble")
    d.kcards(p, period_kpis(t), .45, 1.72, 12.43, 1.0)
    d.box_pic(p, ch["trend_slide"], .45, 2.84, 7.7, 4.06, "Tendance mensuelle")
    d.text(p, 8.3, 2.84, 4.58, .32, "Top 3 catégories récurrentes", 12.5, True, C.FOREST)
    if data["top_categories"]:
        d.table(p, 8.3, 3.2, 4.58, [3.3, 1], ["Catégorie", "Total"], [[f"{i}. {cut(c['label'], 40)}", fmt_int(c["count"])] for i, c in enumerate(data["top_categories"], 1)],
                font=11, row_h=.5, bolds=lambda r, k: k == 1)
    else:
        d.text(p, 8.3, 3.2, 4.58, .4, "Aucune donnée d’analyse sur la période.", 11, False, C.MUTED)

    p = d.page(2, "Détail par mois & charge hebdomadaire", "Consolidation")
    rows = detail_rows(data)
    n = len(rows)
    plain = [r["cells"] if isinstance(r, dict) else r for r in rows]
    rh = min(.32, 3.6 / (n + 1)) if n else .32
    d.table(p, .45, 1.72, 12.43, [2.6, 1.6, 1.6, 1.4, 1.2, 1.4], DETAIL_HEAD, plain, font=10.5 if n < 12 else 9.5, row_h=rh,
            fills=lambda r, k: C.LIGHT if r == n - 1 else None, bolds=lambda r, k: r == n - 1 or k >= 4)
    y = 1.8 + rh * (n + 1)
    d.box_pic(p, ch["heat_slide"], .45, y, 12.43, 6.92 - y, "Charge moyenne par jour de la semaine (appels + emails reçus)")

    k = 2
    filled = [(key, lab, notes[key].strip()) for key, lab in BLOCKS if (notes.get(key) or "").strip()]
    for i in range(0, len(filled), 2):
        k += 1
        p = d.page(k, "Analyse de la période", "Commentaires")
        for j, (_, lab, txt) in enumerate(filled[i:i + 2]):
            d.note(p, .45 + j * 6.3, 1.72, 6.13, 5.18, lab, [x for x in cut(txt, 1500).split("\n") if x.strip()], 13)

    for c in months:
        ch_m = c["charts"]
        k += 1
        p = d.page(k, f"{c['period']} — indicateurs", "Rapport mensuel")
        d.kcards(p, c["kpi_volume"], .45, 1.72, 12.43, 1.0, 20)
        d.kcards(p, c["kpi_quali"], .45, 2.8, 12.43, 1.0, 20)
        if c["weekly"] or c["problems"]:
            d.box_pic(p, ch_m["weekly_month"], .45, 3.9, 6.13, 3.02, "Bugs vs Demandes")
            d.box_pic(p, ch_m["problems_month"], 6.75, 3.9, 6.13, 3.02, "Problèmes techniques")
        else:
            d.text(p, .45, 4.0, 12, .4, NO_ANALYSIS, 12, False, C.MUTED)

        k += 1
        p = d.page(k, f"{c['period']} — analyse", "Rapport mensuel")
        for x, title, items, col, nb in ((.45, "Problèmes techniques", c["problems"], C.RED, c["nb_prob"]), (6.75, "Demandes d’information", c["demandes"], C.GREEN, c["nb_dem"])):
            d.text(p, x, 1.66, 6.13, .3, f"{title} ({fmt_int(nb)})", 12, True, col)
            if items:
                d.table(p, x, 1.98, 6.13, [4.2, .8, 1], MONTH_HEAD, [[cut(l, 48), v, q] for l, v, q, _ in sorted_items(items)], font=10, row_h=.29, bolds=lambda r, j: j == 1)
            else:
                d.text(p, x, 1.98, 6.13, .3, "Aucune donnée.", 11, False, C.MUTED)
        d.text(p, .45, 4.42, 6.13, .3, "Évolution hebdomadaire", 12, True, C.FOREST)
        crit = c["crit"]
        if c["weekly"]:
            d.table(p, .45, 4.74, 6.13, [3, 1.4, 1.4], ["Semaine", "Bugs", "Demandes"], [[cut(w[0], 22), w[1], w[2]] for w in c["weekly"][:6]], font=10, row_h=.3,
                    fills=lambda r, j: M.ALERT_BG if r == crit else None)
        d.bullets(p, 6.75, 4.42, 6.13, 2.5, "Constats & recommandations", [cut(x, 120) for x in c["constats"][:4]] + [cut("Reco. : " + r, 120) for r in c["recos"][:3]] or ["—"], 10.5)
    return d.finish()


def mail_html(ctx, data, greeting, signature, notes=None, full_width=False):
    notes = notes or {}
    ch = period_charts(data)
    m = Mail
    b = m.sec(1, "Indicateurs") + m.kpis(period_kpis(data["totals"]), 5) + m.note(notes.get("indicateurs"))
    b += m.sec(2, "Tendance mensuelle") + m.img(ch["trend"], "Emails, appels, bugs et demandes par mois")
    b += m.sec(3, "Charge moyenne par jour de la semaine") + m.img(ch["heat"], "Appels + emails reçus") + m.note(notes.get("charge"))
    b += m.sec(4, "Top 3 catégories récurrentes")
    b += m.table(["Catégorie", "Total signalé"], top_rows(data), ["left", "center"]) if data["top_categories"] else m.note("Aucune donnée d’analyse pour la période.")
    b += m.note(notes.get("top3")) + m.sec(5, "Détail par mois")
    b += m.table(DETAIL_HEAD, detail_rows(data), ["left"] + ["center"] * 5) + m.note(notes.get("detail"))
    if (notes.get("synthese") or "").strip():
        b += m.sec(6, "Synthèse & conclusion") + m.note(notes["synthese"])
    intro = f"Veuillez trouver ci-dessous la synthèse du support Tarkhiss sur la période {ctx['period_label']}."
    return m.shell(ctx["title"], ctx["period_label"], greeting, intro, b, signature, title_parts(ctx)[0], full_width)


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
