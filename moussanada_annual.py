# -*- coding: utf-8 -*-
"""
moussanada_annual.py — Rapport annuel (ou de période) Moussanada « Helpdesk GLPI ».

Rendus purs (aucun accès disque/Flask) à partir des données préparées par app.py
(`moussanada_annual_data`). Charte AMMPS, comme les rapports mensuels. Contenu identique dans les
quatre sorties : KPI (total, en cours, résolus, clos, en instance), tendance mensuelle, catégories
les plus demandées (graphique + tableau), top des demandeurs (graphique + tableau), détail par mois.
- build_pdf()  : garde, sommaire, 5 sections ;
- build_pptx() : garde + 4 diapositives ;
- mail_html()  : corps de l'email ;
- write_summary_sheet() : feuille Excel « Synthèse » (graphiques natifs) + feuilles Catégories / Demandeurs.
Vocabulaire : « En instance » (jamais « en retard »).
"""
import ammps_charte as C
import rapports_mensuels as M
from ammps_charte import Deck, Mail, esc

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
SUBTITLE = "Synthèse de la période — Helpdesk GLPI"
DETAIL_HEAD = ["Mois", "Total tickets", "En cours", "Résolus", "Clos", "En instance", "Délai résolution", "Délai clôture"]
TOP_HEAD = ["Total", "Résolus", "%"]
TOP_N = 10


def fmt_int(n):
    return f"{int(n or 0):,}".replace(",", " ")


def dh(hours):
    return f"{(hours or 0):.2f}".replace(".", ",") + " h"


def cut(s, n):
    s = str(s or "")
    return s if len(s) <= n else s[:n - 1].rstrip() + "…"


def title_parts(ctx):
    first, _, second = ctx["title"].partition(" — ")
    return first, second


def period_kpis(t):
    return [("Total tickets", fmt_int(t["total"]), "ouverts, résolus, clos…"), ("En cours", fmt_int(t["en_cours"]), "non résolus", "warn"),
            ("Résolus", fmt_int(t["resolus"]), "", "good"), ("Clos", fmt_int(t["clos"]), ""),
            ("En instance", fmt_int(t["en_instance"]), "", "alert")]


def _top_rows(items):
    tot = sum(i["ouverts"] for i in items) or 1
    return [[f"<b>{n}.</b> {esc(i['label'])}", f"<b>{fmt_int(i['ouverts'])}</b>", fmt_int(i["resolus"]), C.fr_pct(i["ouverts"] * 100 / tot)]
            for n, i in enumerate(items, 1)]


def detail_rows(data):
    rows = [[esc(m["label"]), f"<b>{fmt_int(m['total'])}</b>", fmt_int(m["en_cours"]), fmt_int(m["resolus"]), fmt_int(m["clos"]),
             f'<b style="color:{C.RED};">{fmt_int(m["en_instance"])}</b>', dh(m["resolution_h"]), dh(m["cloture_h"])] for m in data["months"]]
    t = data["totals"]
    if rows:
        rows.append({"cls": "tot", "cells": ["Total période", fmt_int(t["total"]), fmt_int(t["en_cours"]), fmt_int(t["resolus"]), fmt_int(t["clos"]),
                                              fmt_int(t["en_instance"]), "—", "—"]})
    return rows


def period_charts(data):
    ms = data["months"]
    multi = len({m["label"].split()[-1] for m in ms}) > 1
    labs = [m["label"].split()[0][:4] + "." + (" " + m["label"].split()[-1][2:] if multi else "") for m in ms]
    trend = lambda size: C.lines(labs, {"Total": [m["total"] for m in ms], "Résolus": [m["resolus"] for m in ms], "Clos": [m["clos"] for m in ms]}, figsize=size)
    inst = lambda size: C.bars(labs, {"En instance": [m["en_instance"] for m in ms]}, figsize=size)
    two = lambda items: {"Total": [i["ouverts"] for i in items], "Résolus": [i["resolus"] for i in items]}
    makers = {}
    if ms:
        makers.update(trend=lambda: trend((8.2, 2.5)), trend_slide=lambda: trend((7.6, 4.1)),
                      inst=lambda: inst((8.2, 1.9)), inst_slide=lambda: inst((4.6, 4.1)))
    cats, dem, srv = data["top_categories"], data["top_demandeurs"], data["top_services"]
    if cats:
        cl = M.tail_labels(cats)
        makers.update(cat=lambda: C.hbars(cl, two(cats)), cat_slide=lambda: C.hbars(cl, two(cats), figsize=(6.2, 5.1), maxlen=34))
    if dem:
        dl = [d["label"] for d in dem]
        makers.update(dem=lambda: C.hbars(dl, two(dem)), dem_slide=lambda: C.hbars(dl, two(dem), figsize=(6.2, 5.1), maxlen=34))
    if srv:
        sl = M.tail_labels(srv)
        makers.update(srv=lambda: C.hbars(sl, two(srv)), srv_slide=lambda: C.hbars(sl, two(srv), figsize=(6.2, 5.1), maxlen=34))
    return M.Charts(makers)


# ----------------------------------------------------------------------------- PDF
def build_pdf(ctx, data):
    ch = period_charts(data)
    first, second = title_parts(ctx)
    items = [("a1", "Indicateurs & tendance mensuelle"), ("a2", "Catégories les plus demandées"), ("a3", "Top des demandeurs"), ("a4", "Top des services demandeurs"), ("a5", "Détail par mois")]
    align = ["l", "c w9", "c w9", "c w9"]
    body = C.cover(first, second, SUBTITLE, [("Période", ctx["period_label"]), ("Mois avec données", str(ctx["months_count"])), ("Source", data["source_label"]), ("Édition", ctx["generated"].strftime("%d/%m/%Y"))], first)
    body += f"""{C.head_html(ctx['title'], ctx['period_label'])}<div>{C.sec(0, 'Sommaire', first)}{C.toc(items)}</div>
    <div class="newp">{C.sec(1, items[0][1], 'Vue d’ensemble', 'a1')}{C.kpi_grid(period_kpis(data['totals']), 5)}
      <div class="gap">{C.fig(ch['trend'], 'Tendance mensuelle', 'Tickets reçus, résolus et clos, mois par mois.')}
      {C.fig(ch['inst'], 'Tickets en instance', 'Tickets en instance, mois par mois.')}</div></div>
    <div class="newp">{C.sec(2, items[1][1], 'Demandes', 'a2')}
      {C.table(['Catégorie'] + TOP_HEAD, _top_rows(data['top_categories']), align, 'tight') if data['top_categories'] else '<p class="note">Aucune catégorie sur la période.</p>'}
      <div class="gap">{C.fig(ch['cat'], 'Catégories les plus demandées', 'Dix premières catégories : tickets et résolus.')}</div></div>
    <div class="newp">{C.sec(3, items[2][1], 'Demandes', 'a3')}
      {C.table(['Demandeur'] + TOP_HEAD, _top_rows(data['top_demandeurs']), align, 'tight') if data['top_demandeurs'] else '<p class="note">Aucun demandeur sur la période.</p>'}
      <div class="gap">{C.fig(ch['dem'], 'Top des demandeurs', 'Dix premiers demandeurs : tickets et résolus.')}</div></div>
    <div class="newp">{C.sec(4, items[3][1], 'Demandes', 'a4')}
      {C.table(['Service'] + TOP_HEAD, _top_rows(data['top_services']), align, 'tight') if data['top_services'] else '<p class="note">Aucun service sur la période.</p>'}
      <div class="gap">{C.fig(ch['srv'], 'Top des services demandeurs', 'Dix premiers services : tickets et résolus.')}</div></div>
    <div class="newp">{C.sec(5, items[4][1], 'Consolidation', 'a5')}{C.table(DETAIL_HEAD, detail_rows(data), ['l'] + ['c'] * 7, 'tight')}</div>"""
    return C.render_pdf(ctx["title"], "Moussanada", ctx["period_label"], ctx["clean"](body), ".c5 {grid-template-columns:repeat(5,1fr);}")


# ----------------------------------------------------------------------------- PPTX
def build_pptx(ctx, data):
    per = ctx["period_label"]
    ch = period_charts(data)
    first, second = title_parts(ctx)
    d = Deck(ctx["title"], per, f"AMMPS · Moussanada — {per}", ctx["clean"])
    d.cover(first, first, second, SUBTITLE, [("Période", per), ("Mois avec données", str(ctx["months_count"])), ("Source", data["source_label"])])

    p = d.page(1, "Synthèse de la période", "Vue d’ensemble")
    d.kcards(p, period_kpis(data["totals"]), .45, 1.72, 12.43, 1.1)
    d.box_pic(p, ch["trend_slide"], .45, 2.95, 7.6, 3.95, "Tendance mensuelle")
    d.box_pic(p, ch["inst_slide"], 8.2, 2.95, 4.68, 3.95, "Tickets en instance")

    def top_page(n, title, first, items, labels, img, img_title):
        p = d.page(n, title, "Demandes")
        if items:
            rows = [[cut(lab, 44), fmt_int(i["ouverts"]), fmt_int(i["resolus"]), r[3]] for lab, i, r in zip(labels, items, _top_rows(items))]
            d.table(p, .45, 1.72, 6.05, [3.5, .9, .9, .9], [first] + TOP_HEAD, rows, row_h=.46, bolds=lambda r, k: k == 1, font=11)
        d.box_pic(p, img, 6.65, 1.72, 6.23, 5.17, img_title)

    cats, dem = data["top_categories"], data["top_demandeurs"]
    top_page(2, "Catégories les plus demandées", "Catégorie", cats, M.tail_labels(cats), ch["cat_slide"], "Catégories les plus demandées")
    top_page(3, "Top des demandeurs", "Demandeur", dem, [x["label"] for x in dem], ch["dem_slide"], "Top des demandeurs")
    srv = data["top_services"]
    top_page(4, "Top des services demandeurs", "Service", srv, M.tail_labels(srv), ch["srv_slide"], "Top des services demandeurs")

    p = d.page(5, "Détail par mois", "Consolidation")
    rows = [r["cells"] if isinstance(r, dict) else r for r in detail_rows(data)]
    n = len(rows)
    d.table(p, .45, 1.72, 12.43, [2.2, 1.4, 1.2, 1.2, 1.1, 1.3, 1.6, 1.6], DETAIL_HEAD, rows, font=11 if n < 12 else 9.5,
            row_h=min(.42, 5.0 / (n + 1)) if n else .4, fills=lambda r, k: C.LIGHT if r == n - 1 else None, bolds=lambda r, k: r == n - 1 or k in (1, 5))
    return d.finish()


# ----------------------------------------------------------------------------- email
def mail_html(ctx, data, greeting, signature, full_width=False):
    ch = period_charts(data)
    m = Mail
    al = ["left", "center", "center", "center"]
    b = m.sec(1, "Indicateurs") + m.kpis(period_kpis(data["totals"]), 5)
    b += m.sec(2, "Tendance mensuelle") + m.img(ch["trend"], "Tickets reçus, résolus et clos par mois") + m.img(ch["inst"], "Tickets en instance par mois")
    b += m.sec(3, "Catégories les plus demandées")
    b += (m.table(["Catégorie"] + TOP_HEAD, _top_rows(data["top_categories"]), al) + m.img(ch["cat"], "Catégories les plus demandées")
          if data["top_categories"] else m.note("Aucune catégorie sur la période."))
    b += m.sec(4, "Top des demandeurs")
    b += (m.table(["Demandeur"] + TOP_HEAD, _top_rows(data["top_demandeurs"]), al) + m.img(ch["dem"], "Top des demandeurs")
          if data["top_demandeurs"] else m.note("Aucun demandeur sur la période."))
    b += m.sec(5, "Top des services demandeurs")
    b += (m.table(["Service"] + TOP_HEAD, _top_rows(data["top_services"]), al) + m.img(ch["srv"], "Top des services demandeurs")
          if data["top_services"] else m.note("Aucun service sur la période."))
    b += m.sec(6, "Détail par mois") + m.table(DETAIL_HEAD, detail_rows(data), ["left"] + ["center"] * 7)
    intro = f"Veuillez trouver ci-dessous la synthèse du support Moussanada sur la période {ctx['period_label']}. Source des données : {data['source_label']}."
    return m.shell(ctx["title"], ctx["period_label"], greeting, intro, b, signature, title_parts(ctx)[0], full_width)


# ----------------------------------------------------------------------------- Excel (xlsxwriter)
def write_summary_sheet(wb, fmts, data, ctx):
    """Feuille « Synthèse » (KPI, détail par mois + graphique natif, tops) puis « Catégories » et « Demandeurs »."""
    ws = wb.add_worksheet("Synthèse")
    ws.set_column("A:A", 34)
    ws.set_column("B:H", 15)
    ws.merge_range("A1:H1", f"{ctx['title']} — {ctx['period_label']}", fmts["title"])
    ws.write(1, 0, ctx["agency"])
    ws.write(2, 0, f"Source : {data['source_label']}")
    t = data["totals"]
    row = 3
    for lab, key in (("Total tickets", "total"), ("En cours", "en_cours"), ("Résolus", "resolus"), ("Clos", "clos"), ("En instance", "en_instance")):
        ws.write(row, 0, lab, fmts["bold"])
        ws.write(row, 1, t[key])
        row += 1
    row += 1
    ws.write(row, 0, "Détail par mois", fmts["title"])
    row += 1
    head_row = row
    for i, h in enumerate(DETAIL_HEAD):
        ws.write(row, i, h, fmts["header"])
    for mth in data["months"]:
        row += 1
        for i, k in enumerate(("label", "total", "en_cours", "resolus", "clos", "en_instance", "resolution_h", "cloture_h")):
            ws.write(row, i, mth[k])
    last = row
    if data["months"]:
        row += 1
        ws.write(row, 0, "Total période", fmts["bold"])
        for i, k in enumerate(("total", "en_cours", "resolus", "clos", "en_instance"), start=1):
            col = chr(65 + i)
            ws.write_formula(row, i, f"=SUM({col}{head_row + 2}:{col}{last + 1})", fmts["bold"], t[k])
        ch = wb.add_chart({"type": "line"})
        for i, (name, color) in enumerate((("Total tickets", "#0B4965"), (None, None), ("Résolus", "#25935F"), ("Clos", "#F59F0A")), start=1):
            if not name:
                continue
            ch.add_series({"name": name, "categories": ["Synthèse", head_row + 1, 0, last, 0], "values": ["Synthèse", head_row + 1, i, last, i],
                           "line": {"color": color, "width": 2.25}, "marker": {"type": "circle", "size": 5, "fill": {"color": color}, "border": {"color": color}}})
        ch.set_title({"name": "Tendance mensuelle", "name_font": {"color": "#0B4965", "size": 12}})
        ch.set_legend({"position": "bottom"})
        ch.set_size({"width": 620, "height": 300})
        ws.insert_chart("J3", ch)
        ci = wb.add_chart({"type": "column"})
        ci.add_series({"name": "En instance", "categories": ["Synthèse", head_row + 1, 0, last, 0], "values": ["Synthèse", head_row + 1, 5, last, 5],
                       "fill": {"color": "#DC2828"}})
        ci.set_title({"name": "Tickets en instance", "name_font": {"color": "#0B4965", "size": 12}})
        ci.set_legend({"none": True})
        ci.set_size({"width": 620, "height": 260})
        ws.insert_chart("J19", ci)

    def top_sheet(name, first, items, anchor):
        sh = wb.add_worksheet(name)
        sh.set_column("A:A", 56)
        sh.set_column("B:D", 13)
        sh.merge_range("A1:D1", f"{name} — {ctx['period_label']}", fmts["title"])
        for i, h in enumerate([first, "Total", "Résolus", "%"]):
            sh.write(2, i, h, fmts["header"])
        tot = sum(i["ouverts"] for i in items) or 1
        for r, it in enumerate(items, start=3):
            sh.write(r, 0, it["label"])
            sh.write(r, 1, it["ouverts"])
            sh.write(r, 2, it["resolus"])
            sh.write(r, 3, it["ouverts"] / tot, fmts["pct"])
        if items:
            cb = wb.add_chart({"type": "bar"})
            n = len(items)
            cb.add_series({"name": "Total", "categories": [name, 3, 0, 2 + n, 0], "values": [name, 3, 1, 2 + n, 1], "fill": {"color": "#0B4965"}})
            cb.add_series({"name": "Résolus", "categories": [name, 3, 0, 2 + n, 0], "values": [name, 3, 2, 2 + n, 2], "fill": {"color": "#25935F"}})
            cb.set_y_axis({"reverse": True})
            cb.set_title({"name": name, "name_font": {"color": "#0B4965", "size": 12}})
            cb.set_legend({"position": "bottom"})
            cb.set_size({"width": 640, "height": max(300, 28 * n + 120)})
            sh.insert_chart(anchor, cb)

    top_sheet("Catégories", "Catégorie", data["top_categories"], "F3")
    top_sheet("Demandeurs", "Demandeur", data["top_demandeurs"], "F3")
    top_sheet("Services", "Service", data["top_services"], "F3")
    return ws
