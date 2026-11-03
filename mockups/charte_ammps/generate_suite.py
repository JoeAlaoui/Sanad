"""Exemples variante A : PDF Moussanada, PDF Reporting Métier PCHC, emails Tarkhiss et Moussanada.
Lancer depuis la racine du projet : .venv/bin/python mockups/charte_ammps/generate_suite.py
Lit une copie isolée de data/ (voir sandbox.py) ; les emails sont des brouillons sans destinataire.
"""
import base64
import os
import re
import sys
from datetime import datetime
from email.message import EmailMessage

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import generate as G  # noqa: E402
import sandbox  # noqa: E402
from generate import DEEP, FOREST, GREEN, INK, LIGHT, LINE, MUTED, PALE, R, RED, TINT, esc  # noqa: E402

AMBER = "#c98a1b"
EX_DIR = os.path.join(HERE, "exemples")
os.makedirs(EX_DIR, exist_ok=True)
YM_M = "2026-09"
GEN = datetime(2026, 10, 30, 12, 0)
BANNED = re.compile(r"\bSANAD\b", re.I)
BANNED_MAIL = re.compile(r"\b(SANAD|GLPI)\b", re.I)


def fr(v, d=1):
    return f"{v:.{d}f}".replace(".", ",") + " %"


# ----------------------------------------------------------------------------- Graphiques
def style(ax):
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0)


def bars(cats, series, figsize=(8.2, 2.4), colors=None, rotate=False):
    import numpy as np
    plt = G.mpl()
    colors = colors or ([TINT, FOREST] if len(series) == 2 else [FOREST])
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(len(cats))
    w = .8 / max(len(series), 1)
    top = max([max(v, default=0) for v in series.values()] + [1])
    for i, (name, vals) in enumerate(series.items()):
        bs = ax.bar(x + i * w - .4 + w / 2, vals, w * .92, color=colors[i % len(colors)], label=name)
        for b in bs:
            lab = f"{b.get_height():.0f}" if float(b.get_height()).is_integer() else f"{b.get_height():.1f}".replace(".", ",")
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + top * .02, lab, ha="center", fontsize=9, fontweight="bold")
    ax.set_xticks(x, cats, rotation=25 if rotate else 0, ha="right" if rotate else "center")
    style(ax)
    ax.set_ylim(0, top * (1.3 if len(series) > 1 else 1.15))
    if len(series) > 1:
        ax.legend(frameon=False, loc="upper left", ncol=len(series), fontsize=9)
    return G.png(fig)


def hbars(labels, series, figsize=None, colors=(FOREST, TINT), maxlen=40):
    import numpy as np
    plt = G.mpl()
    n = len(labels)
    fig, ax = plt.subplots(figsize=figsize or (8.2, .24 * n + .55))
    y = np.arange(n)
    h = .78 / len(series)
    top = max([max(v, default=0) for v in series.values()] + [1])
    for i, (name, vals) in enumerate(series.items()):
        bs = ax.barh(y + i * h - .39 + h / 2, vals, h * .92, color=colors[i % len(colors)], label=name)
        for b in bs:
            ax.text(b.get_width() + top * .012, b.get_y() + b.get_height() / 2, f"{b.get_width():.0f}", va="center", fontsize=8.5, fontweight="bold")
    ax.set_yticks(y, [(l[:maxlen] + "…") if len(l) > maxlen else l for l in labels])
    ax.invert_yaxis()
    ax.set_xticks([])
    ax.set_xlim(0, top * 1.1)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.tick_params(axis="y", length=0, labelcolor=INK, labelsize=9)
    if len(series) > 1:
        ax.legend(frameon=False, loc="lower right", fontsize=9)
    return G.png(fig)


def lines(cats, series, figsize=(8.2, 2.3)):
    plt = G.mpl()
    fig, ax = plt.subplots(figsize=figsize)
    for (name, vals), col in zip(series.items(), (FOREST, GREEN, TINT)):
        ax.plot(cats, vals, marker="o", linewidth=2.2, markersize=4.5, color=col, label=name)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color=LINE, linewidth=.6)
    ax.tick_params(length=0)
    ax.legend(frameon=False, loc="upper left", ncol=3, fontsize=9)
    return G.png(fig)


def pareto(cats, vals, cum, figsize=(8.2, 2.7)):
    plt = G.mpl()
    fig, ax = plt.subplots(figsize=figsize)
    ax.bar(range(len(cats)), vals, color=TINT, width=.7)
    ax.set_xticks(range(len(cats)), cats, rotation=30, ha="right", rotation_mode="anchor", fontsize=8)
    style(ax)
    ax2 = ax.twinx()
    ax2.plot(range(len(cats)), cum, color=FOREST, marker="o", markersize=4, linewidth=2)
    for i, c in enumerate(cum):
        ax2.text(i, c + 4, f"{c} %", ha="center", fontsize=8, fontweight="bold", color=FOREST)
    ax2.set_ylim(0, 118)
    ax2.set_yticks([])
    ax2.spines[["top", "right", "left"]].set_visible(False)
    return G.png(fig)


def radar(techs, volume_of, figsize=(4.4, 3.8)):
    import numpy as np
    plt = G.mpl()
    rt = techs[:6]
    mo = max(1, *[t.get("ouverts", 0) for t in rt])
    md = max(1, *[(t.get("delai_resolution_s") or 0) for t in rt])
    axes = ["Ouverts", "Résolus", "Taux\nrésolution", "Rapidité"]
    ang = np.linspace(0, 2 * np.pi, 4, endpoint=False).tolist()
    ang += ang[:1]
    fig, ax = plt.subplots(figsize=figsize, subplot_kw=dict(polar=True))
    for t, col, ls in zip(rt, (FOREST, GREEN, RED, MUTED, TINT, INK), ("-", "-", "-", "--", "--", "--")):
        o, r = t.get("ouverts", 0), t.get("resolus", 0)
        rate = min(100, round(r / o * 100)) if o else 0
        vals = [round(o / mo * 100), rate, rate, round(100 - (t.get("delai_resolution_s") or 0) / md * 100) if t.get("delai_resolution_s") else 100]
        ax.plot(ang, vals + vals[:1], color=col, linewidth=1.8, linestyle=ls, label=t["label"])
    ax.set_xticks(ang[:-1], axes, fontsize=8.5)
    ax.tick_params(axis="x", pad=13)
    ax.set_yticklabels([])
    ax.set_ylim(0, 100)
    ax.spines["polar"].set_color(LINE)
    ax.grid(color=LINE)
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.17), ncol=2, frameon=False, fontsize=7.5)
    return G.png(fig)


def status_chart(mod):
    plt = G.mpl()
    col = {"Délivré": FOREST, "En cours": TINT, "Non conforme": AMBER, "Refusé": RED, "Autre": "#b8c4bc"}
    rows = [r for r in mod["statuses"] if not (r["category"] == "Autre" and r["count"] == 0)]
    fig, ax = plt.subplots(figsize=(8.2, .19 * len(rows) + .35))
    ax.barh(range(len(rows)), [r["count"] for r in rows], color=[col.get(r["category"], MUTED) for r in rows], height=.68)
    ax.set_yticks(range(len(rows)), [r["status"][:44] for r in rows])
    ax.invert_yaxis()
    mx = max([r["count"] for r in rows] + [1])
    for i, r in enumerate(rows):
        ax.text(r["count"] + mx * .01, i, f"{r['count']:,}".replace(",", " ") + "  ·  " + fr(r["pct"] * 100), va="center", fontsize=8.5, fontweight="bold")
    ax.set_xticks([])
    ax.set_xlim(0, mx * 1.3)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.tick_params(axis="y", length=0, labelcolor=INK, labelsize=8.5)
    return G.png(fig)


def modules_chart(comp, short):
    plt = G.mpl()
    syn = comp["synthesis"]
    labels = [short[s["key"]] for s in syn]
    fig, ax = plt.subplots(figsize=(8.2, 2.5))
    bottom = [0] * len(syn)
    top = max(1, max(s["total"] for s in syn))
    for name, vals, col, tc in (("Délivrés", [s["delivered"] for s in syn], FOREST, "white"), ("En cours", [s["encours"] for s in syn], TINT, INK),
                                ("Bloqués / refusés", [s["bloques"] or 0 for s in syn], RED, "white")):
        ax.bar(labels, vals, bottom=bottom, color=col, label=name, width=.58)
        for i, v in enumerate(vals):
            if v and v / top > .05:
                ax.text(i, bottom[i] + v / 2, f"{v:,}".replace(",", " "), ha="center", va="center", fontsize=8.5, color=tc, fontweight="bold")
        bottom = [b + v for b, v in zip(bottom, vals)]
    for i, s in enumerate(syn):
        ax.text(i, bottom[i] + top * .02, fr(s["taux"] * 100), ha="center", fontsize=9.5, fontweight="bold", color=FOREST)
    style(ax)
    ax.set_ylim(0, top * 1.17)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(.5, -.08), ncol=3, fontsize=9)
    return G.png(fig)


# ----------------------------------------------------------------------------- PDF : composants ajoutés
EXTRA_CSS = f"""
.c3 {{ grid-template-columns:repeat(3,1fr); }}
.c7 {{ grid-template-columns:repeat(7,1fr); gap:4pt; }} .c7 .kpi {{ padding:4pt 5pt; }} .c7 .kl {{ font-size:5.8pt; min-height:14pt; }} .c7 .kv {{ font-size:10.5pt; }}
td.c {{ white-space:nowrap; }} .nw {{ white-space:nowrap; }} .keep {{ break-inside:avoid; }}
.mod table.data.tight td {{ padding:1.5pt 7pt; font-size:8pt; }} .mod figure {{ margin-bottom:4pt; padding:5pt 8pt 4pt; }} .mod h3 {{ margin:6pt 0 2pt; }}
.mod ul.pts li {{ padding:2pt 0 2pt 11pt; font-size:8.3pt; }} .mod ul.pts li::before {{ top:5.5pt; }} .mod .gap {{ margin-top:7pt; }} .mod .sec {{ margin-bottom:3pt; }}
.notebox {{ background:{LIGHT}; border-left:2.4pt solid {GREEN}; padding:7pt 11pt; margin:6pt 0 4pt; break-inside:avoid; font-size:9pt; }}
.notebox p {{ margin:0 0 4pt; }} .notebox p:last-child {{ margin:0; }}
.notebox .nl {{ font:700 7pt Montserrat; letter-spacing:.8pt; text-transform:uppercase; color:{GREEN}; margin-bottom:3pt; }}
table.data tbody tr.tot td {{ background:{LIGHT}; font-weight:700; color:{FOREST}; }}
.st {{ display:inline-block; width:6pt; height:6pt; margin-right:4pt; }}
.small {{ font-size:8pt; }} .small ul.pts li {{ padding:2.5pt 0 2.5pt 11pt; }} .small ul.pts li::before {{ top:6.5pt; width:4pt; height:4pt; }}
figure.half img {{ width:100%; }}
"""


def notebox(text, label="Analyse"):
    if not (text or "").strip():
        return ""
    paras = "".join(f"<p>{esc(p)}</p>" for p in text.split("\n") if p.strip())
    return f'<div class="notebox"><div class="nl">{esc(label)}</div>{paras}</div>'


def render_pdf(name, css, body):
    import weasyprint
    from pypdf import PdfReader
    doc = (f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>{esc(R["title"])} — {esc(R["period"])}</title>'
           f'<style>{G.base_css()}{G.cover_a_css()}{EXTRA_CSS}{css}</style></head><body>{body}</body></html>')
    doc = re.sub(r"(\d)[ \u202f]%", "\\1\u00a0%", doc)
    out = os.path.join(EX_DIR, name)
    weasyprint.HTML(string=doc.replace(" ", " "), base_url=HERE).write_pdf(out)
    r = PdfReader(out)
    print(name, len(r.pages), "pages", os.path.getsize(out) // 1024, "Ko", "SANAD:", bool(BANNED.search(" ".join(p.extract_text() for p in r.pages))))


def toc(items):
    rows = "".join(f'<li><a href="#{a}"><span class="tn">{i:02d}</span><span class="tl">{esc(l)}</span><span class="tp" data-t="#{a}"></span></a></li>'
                   for i, (a, l) in enumerate(items, 1))
    return f'<ol class="toc">{rows}</ol>'


# ----------------------------------------------------------------------------- Moussanada : données communes PDF + email
def tail_labels(items):
    """Libellés de graphique : dernier niveau de l'arborescence, deux niveaux si homonymes."""
    parts = [[p.strip() for p in i.get("label", "").split(">")] for i in items]
    last = [p[-1] for p in parts]
    return [" > ".join(p[-2:]) if last.count(p[-1]) > 1 else p[-1] for p in parts]


def moussanada_ctx(app):
    ym = YM_M
    d = app.moussanada_data(ym).get_json()
    m, ts, notes = d["month"], d["timeseries"], d["notes"]
    tk, du = m["tickets"], m["durations"]
    py = app.prev_ym(ym)
    prev = ts.get(py, {})
    ml, pl = app.month_label_fr(ym), app.month_label_fr(py)
    plow = pl.split()[0].lower()
    lbl = app.get_kpi_labels("moussanada")
    inst = lbl["tickets_retard"] if lbl["tickets_retard"] != app.MOUSSANADA_DEFAULT_KPI_LABELS["tickets_retard"] else "Tickets en instance"

    def delta(cur, pv, pol, fmt=str):
        if not pv:
            return "", ""
        diff = (cur - pv) / pv * 100
        tone = ""
        if abs(diff) >= 1 and pol != "neutral":
            tone = "good" if (diff > 0) == (pol == "positive") else "alert"
        return f"{'+' if diff >= 0 else '−'}{abs(diff):.0f} % vs {plow} ({fmt(pv)})", tone

    def k(label, val, cur, pv, pol, fmt=str):
        s, tone = delta(cur, pv, pol, fmt)
        return (label, str(val), s, tone)

    pc = app.pct
    tr, tc = pc(tk["resolus"], tk["ouverts"]), pc(tk["clos"], tk["ouverts"])
    ptr, ptc = pc(prev.get("resolus", 0), prev.get("ouverts", 0)), pc(prev.get("clos", 0), prev.get("ouverts", 0))
    pf = lambda v: fr(v)
    dh = app.format_dh
    months = app.real_months(ts, require_data=True)
    labs = [app.month_label_fr(x).split()[0][:4] + "." for x in months]
    cats = sorted(m["categories"], key=lambda x: x.get("ouverts", 0), reverse=True)
    srv = sorted(m["services"], key=lambda x: x.get("ouverts", 0), reverse=True)
    techs = sorted(m["techniciens"], key=app.tech_volume, reverse=True)
    tot = sum(c.get("ouverts", 0) for c in cats) or 1
    cum, cp = 0, []
    for c in cats[:12]:
        cum += c.get("ouverts", 0)
        cp.append(round(cum / tot * 100))
    charts = dict(
        volume=bars(["Ouverts", "Résolus", "En instance", "Clos"], {pl: [prev.get(x, 0) or 0 for x in ("ouverts", "resolus", "en_retard", "clos")],
                                                                     ml: [tk[x] for x in ("ouverts", "resolus", "en_retard", "clos")]}),
        backlog=bars(labs, {"En instance": [ts[x].get("en_retard", 0) or 0 for x in months]}, figsize=(8.2, 1.9)),
        trend=lines(labs, {"Ouverts": [ts[x].get("ouverts", 0) or 0 for x in months], "Résolus": [ts[x].get("resolus", 0) or 0 for x in months],
                           "Clos": [ts[x].get("clos", 0) or 0 for x in months]}),
        volume_half=bars(["Ouverts", "Résolus", "En instance", "Clos"], {pl: [prev.get(x, 0) or 0 for x in ("ouverts", "resolus", "en_retard", "clos")],
                                                                          ml: [tk[x] for x in ("ouverts", "resolus", "en_retard", "clos")]}, figsize=(4.4, 2.8)),
        cat=hbars(tail_labels(cats[:10]), {"Ouverts": [c.get("ouverts", 0) for c in cats[:10]], "Résolus": [c.get("resolus", 0) for c in cats[:10]]}),
        pareto=pareto(tail_labels(cats[:12]), [c.get("ouverts", 0) for c in cats[:12]], cp),
        srv=hbars(tail_labels(srv[:10]), {"Ouverts": [c.get("ouverts", 0) for c in srv[:10]], "Résolus": [c.get("resolus", 0) for c in srv[:10]]}),
        techvol=bars([t["label"] for t in techs], {"Sollicitations": [app.tech_volume(t) for t in techs]}, figsize=(4.4, 3.8), rotate=True),
        tech=bars([t["label"] for t in techs], {"Ouverts": [t.get("ouverts", 0) for t in techs], "Résolus": [t.get("resolus", 0) for t in techs]}, figsize=(8.2, 2.2), colors=[FOREST, TINT]),
        techvol_wide=bars([t["label"] for t in techs], {"Sollicitations": [app.tech_volume(t) for t in techs]}, figsize=(8.2, 2.0)),
        delay_wide=bars(["Délai résolution", "Délai clôture"], {pl: [round(prev.get("resolution_h", 0) or 0, 1), round(prev.get("cloture_h", 0) or 0, 1)],
                                                                ml: [round(du["resolution_h"] or 0, 1), round(du["cloture_h"] or 0, 1)]}, figsize=(8.2, 2.2)),
        radar=radar(techs, app.tech_volume),
        delay=bars(["Délai résolution", "Délai clôture"], {pl: [round(prev.get("resolution_h", 0) or 0, 1), round(prev.get("cloture_h", 0) or 0, 1)],
                                                           ml: [round(du["resolution_h"] or 0, 1), round(du["cloture_h"] or 0, 1)]}, figsize=(4.4, 2.8)),
    )
    vs = {pl: [prev.get(x, 0) or 0 for x in ("ouverts", "resolus", "en_retard", "clos")], ml: [tk[x] for x in ("ouverts", "resolus", "en_retard", "clos")]}
    two = lambda rows: {"Ouverts": [r.get("ouverts", 0) for r in rows], "Résolus": [r.get("resolus", 0) for r in rows]}
    charts.update(  # formats diapositive
        s_volume=bars(["Ouverts", "Résolus", "En instance", "Clos"], vs, figsize=(6.0, 2.7)),
        s_backlog=bars(labs, {"En instance": [ts[x].get("en_retard", 0) or 0 for x in months]}, figsize=(6.0, 2.7)),
        s_trend=lines(labs, {"Ouverts": [ts[x].get("ouverts", 0) or 0 for x in months], "Résolus": [ts[x].get("resolus", 0) or 0 for x in months],
                             "Clos": [ts[x].get("clos", 0) or 0 for x in months]}, figsize=(7.6, 3.9)),
        s_cat=hbars(tail_labels(cats[:10]), two(cats[:10]), figsize=(6.2, 5.1), maxlen=34),
        s_pareto=pareto(tail_labels(cats[:12]), [c.get("ouverts", 0) for c in cats[:12]], cp, figsize=(12.2, 3.3)),
        s_srv=hbars(tail_labels(srv[:10]), two(srv[:10]), figsize=(6.2, 3.7), maxlen=46),
        s_techvol=bars([t["label"] for t in techs], {"Sollicitations": [app.tech_volume(t) for t in techs]}, figsize=(4.4, 2.6), rotate=True),
        s_tech=bars([t["label"] for t in techs], two(techs), figsize=(4.4, 2.6), colors=[FOREST, TINT], rotate=True),
        s_radar=radar(techs, app.tech_volume, figsize=(3.6, 2.9)),
        s_volume2=bars(["Ouverts", "Résolus", "En instance", "Clos"], vs, figsize=(6.0, 2.6)),
        s_delay=bars(["Délai résolution", "Délai clôture"], {pl: [round(prev.get("resolution_h", 0) or 0, 1), round(prev.get("cloture_h", 0) or 0, 1)],
                                                             ml: [round(du["resolution_h"] or 0, 1), round(du["cloture_h"] or 0, 1)]}, figsize=(6.0, 2.6)),
    )
    n = len(techs) or 1

    def top_rows(items):
        total = sum(i.get("ouverts", 0) for i in items) or 1
        return [[esc(i.get("label", "")), f"<b>{i.get('ouverts', 0)}</b>", i.get("resolus", 0), fr(pc(i.get("ouverts", 0), total))] for i in items]

    avg_v = sum(app.tech_volume(t) for t in techs) / n
    tech_rows = [[esc(t.get("label", "")), f"<b>{app.tech_volume(t)}</b>", t.get("ouverts", 0), t.get("resolus", 0), t.get("en_retard", 0),
                  "≥ moy." if app.tech_volume(t) >= avg_v else "&lt; moy."] for t in techs]
    if techs:
        tech_rows.append({"cls": "tot", "cells": ["Moyenne équipe", f"{avg_v:.1f}".replace(".", ","), f"{sum(t.get('ouverts', 0) for t in techs) / n:.1f}".replace(".", ","),
                                                  f"{sum(t.get('resolus', 0) for t in techs) / n:.1f}".replace(".", ","), "", ""]})
    return dict(
        ml=ml, pl=pl, notes=notes, sections=notes.get("sections", {}), charts=charts, cats=cats[:10], srv=srv[:10], top_rows=top_rows, tech_rows=tech_rows,
        k1=[k(lbl["tickets_ouverts"], tk["ouverts"], tk["ouverts"], prev.get("ouverts", 0), "neutral"), k(lbl["tickets_resolus"], tk["resolus"], tk["resolus"], prev.get("resolus", 0), "positive"),
            k(lbl["tickets_clos"], tk["clos"], tk["clos"], prev.get("clos", 0), "positive"), k(inst, tk["en_retard"], tk["en_retard"], prev.get("en_retard", 0), "negative")],
        k2=[k(lbl["taux_resolution"], fr(tr), tr, ptr, "positive", pf), k(lbl["taux_cloture"], fr(tc), tc, ptc, "positive", pf),
            k(lbl["delai_resolution"], dh(du["resolution_h"]), du["resolution_h"], prev.get("resolution_h", 0), "negative", dh),
            k(lbl["delai_cloture"], dh(du["cloture_h"]), du["cloture_h"], prev.get("cloture_h", 0), "negative", dh)],
    )


TOP_HEAD = ["Ouverts", "Résolus", "%"]
TECH_HEAD = ["Technicien", "Sollicitations", "Ouverts", "Résolus", "En instance", "vs équipe"]


def pdf_moussanada(app, c):
    R.update(title="Rapport Support Moussanada", short="Moussanada", period=c["ml"])
    s = c["sections"]
    ch = c["charts"]
    title3, title6 = f"Principales catégories de demandes en {c['ml']}", f"Analyse comparative {c['pl']} / {c['ml']}"
    items = [("m1", "Volume global des tickets"), ("m2", "Évolution des délais de traitement"), ("m3", title3), ("m4", "Services les plus demandeurs"),
             ("m5", "Charge et mobilisation de l’équipe SI"), ("m6", title6), ("m7", "Synthèse finale et conclusion")]
    A = ["l", "c w9", "c w9", "c w9"]
    tb = lambda rows, first: G.table([first] + TOP_HEAD, rows, A, "tight")
    body = G.cover_a("Rapport Support", "Moussanada", "Support Helpdesk", [("Période", c["ml"]), ("Volet", "Moussanada"), ("Édition", GEN.strftime("%d/%m/%Y"))])
    body += f"""{G.head_html()}<div>{G.sec(0, 'Sommaire', 'Rapport mensuel')}{toc(items)}</div>
    <div class="gap">{G.sec(1, items[0][1], 'Volumes', 'm1')}{G.kpi_grid(c['k1'])}
      {G.fig(ch['volume'], 'Volumes de tickets — comparatif', f"Comparatif {c['pl']} / {c['ml']}.")}
      {G.fig(ch['backlog'], 'Backlog — tickets en instance', 'Tickets en instance, mois par mois.')}{notebox(s.get('volume'))}</div>
    <div class="gap">{G.sec(2, items[1][1], 'Délais', 'm2')}{G.kpi_grid(c['k2'])}
      {G.fig(ch['trend'], 'Tendance des tickets', 'Tickets ouverts, résolus et clos, mois par mois.')}{notebox(s.get('delays'))}</div>
    <div class="gap"><div class="keep">{G.sec(3, title3, 'Demandes', 'm3')}{tb(c['top_rows'](c['cats']), 'Catégorie')}</div>
      <div class="gap">{G.fig(ch['cat'], 'Top catégories', 'Dix premières catégories : tickets ouverts et résolus.')}</div>
      {G.fig(ch['pareto'], 'Pareto des catégories', 'Part cumulée des tickets ouverts.')}{notebox(s.get('categories'))}</div>
    <div class="gap"><div class="keep">{G.sec(4, items[3][1], 'Demandes', 'm4')}{tb(c['top_rows'](c['srv']), 'Service')}</div>
      <div class="gap">{G.fig(ch['srv'], 'Top services demandeurs', 'Dix premiers services : tickets ouverts et résolus.')}</div>{notebox(s.get('services'))}</div>
    <div class="gap"><div class="keep">{G.sec(5, items[4][1], 'Équipe', 'm5')}{G.table(TECH_HEAD, c['tech_rows'], ['l', 'c', 'c', 'c', 'c', 'c'], 'tight')}</div>
      <div class="gap">{G.fig(ch['tech'], 'Détail des sollicitations', 'Tickets ouverts et résolus par technicien.')}</div>
      <div class="two">{G.fig(ch['techvol'], 'Volume de sollicitation', 'Ouverts + résolus.')}{G.fig(ch['radar'], 'Radar de performance', 'Principaux techniciens.')}</div>{notebox(s.get('team'))}</div>
    <div class="gap"><div class="keep">{G.sec(6, title6, 'Comparatif', 'm6')}
      <div class="two">{G.fig(ch['volume_half'], 'Volumes', 'Comparatif des volumes.')}{G.fig(ch['delay'], 'Délais (heures)', 'Délais moyens de traitement.')}</div></div>{notebox(s.get('comparative'))}</div>
    <div class="gap">{G.sec(7, items[6][1], 'Synthèse', 'm7')}{notebox(s.get('synthesis'), 'Conclusion')}</div>"""
    render_pdf("Moussanada_rapport_mensuel_2026-09.pdf", "", body)


# ----------------------------------------------------------------------------- PCHC
def pdf_pchc(app):
    import pchc_deliverables as D
    import pchc_raw as P
    meta = app.load_pchc_meta()
    comp = P.pchc_compute(app.load_pchc_raw())
    label = D.period_label(GEN).capitalize()
    R.update(title="Reporting Métier PCHC", short="Reporting Métier PCHC", period=label)
    t, syn = comp["totals"], comp["synthesis"]
    vcol = {"OK": GREEN, "Attention": AMBER, "Critique": RED}
    vcell = lambda name: f'<span class="nw"><span class="st" style="background:{vcol.get(name, MUTED)}"></span>{esc(name)}</span>'
    rows = [[f"<b>{esc(r['module'])}</b>", P.fmt_int(r["total"]), P.fmt_int(r["delivered"]), f"<b>{P.fmt_pct(r['taux'])}</b>", P.fmt_int(r["encours"]),
             "–" if r["bloques"] is None else P.fmt_int(r["bloques"]), f"{r['seuil_vert'] * 100:.0f} / {r['seuil_orange'] * 100:.0f} %", vcell(D.voyant_info(r["voyant"])[1])] for r in syn]
    rows.append({"cls": "tot", "cells": ["TOTAL", P.fmt_int(t["total"]), P.fmt_int(t["delivered"]), P.fmt_pct(t["taux"]), P.fmt_int(t["encours"]), P.fmt_int(t["bloques"]), "", ""]})
    synt = G.table(["Module", "Total", "Délivrés / validés", "Taux", "En cours*", "Bloqués / refusés", "Seuils vert / orange", "Voyant"], rows, ["l", "c", "c", "c", "c", "c", "c", "l"])
    crit = sum(1 for a in comp["alerts"] if "CRITIQUE" in a["status"])
    keep = "".join(f"<li>{esc(x)}</li>" for x in (f"{P.fmt_int(t['total'])} dossiers, {P.fmt_pct(t['taux'])} délivrés/validés.", f"{P.fmt_int(t['encours'])} dossiers en cours de traitement.",
                                                  f"{P.fmt_int(t['bloques'])} bloqués ou refusés.", f"{crit} alerte(s) critique(s) sur {len(comp['alerts'])}."))
    tone = {"ok": "good", "bad": "alert"}
    cards = [(c["label"].capitalize(), P.fmt_int(c["value"]), "", tone.get(c["kind"], "")) for c in comp["cards"]]
    al = []
    for a in comp["alerts"]:
        pf = a["type"] == "taux"
        fmt = (lambda v: f"{v * 100:.0f} %") if pf else P.fmt_int
        al.append([esc(a["indicator"]), esc(a["module"]), f"<b>{P.fmt_pct(a['value']) if pf else P.fmt_int(a['value'])}</b>", fmt(a["warn"]), fmt(a["crit"]),
                   vcell(D.ALERT_LABEL[D.alert_key(a["status"])]), a["priority"].capitalize()])
    alerts = G.table(["Indicateur", "Module", "Valeur", "Attention", "Critique", "Statut", "Priorité"], al, ["l", "l", "c", "c", "c", "l", "c"], "tight")
    items = [("p1", "Synthèse exécutive"), ("p2", "Chiffres clés & points d’alerte")]
    secs = ""
    for i, m in enumerate(P.MODULES, 3):
        mod = comp["modules"][m["key"]]
        items.append((f"p{i}", m["label"]))
        kis = [(l, v, "", "") for l, v, _ in D.key_indicators(mod)]
        st = [[esc(r[0]), f"<b>{r[1]}</b>", r[2], esc(r[3]).replace("✔ ", "").replace("✖ ", "")] for r in D.status_table_rows(mod)]
        st[-1] = {"cls": "tot", "cells": st[-1]}
        tbl = G.table(["Statut", "Nombre", "%", "Catégorie"], st, ["l", "c w9", "c w9", "l w18"], "tight")
        if m["key"] == "identification":
            ty = [[esc(x["label"]), P.fmt_int(x["count"]), P.fmt_pct(x["pct"])] for x in mod["types"]]
            ty.append({"cls": "tot", "cells": ["TOTAL", P.fmt_int(mod["types_total"]), P.fmt_pct(1 if mod["types_total"] else 0)]})
            tbl += '<div class="gap"></div>' + G.table(["Type d’opérateur", "Nombre", "%"], ty, ["l", "c w14", "c w14"], "tight")
            chart = hbars([x["label"] for x in mod["types"][:4]], {"Opérateurs": [x["count"] for x in mod["types"][:4]]}, maxlen=46)
            ctitle = "Opérateurs par type"
        else:
            chart = status_chart(mod)
            ctitle = "Dossiers par statut"
        sub = f"Feuille « {m['sheet']} » · {P.fmt_int(mod['total'])} {'opérateurs' if m['key'] == 'identification' else 'dossiers'} · voyant {D.voyant_info(mod['voyant'])[1]}"
        ins = "".join(f"<li>{esc(x)}</li>" for x in comp["insights"][m["key"]])
        cols = 7 if len(kis) == 7 else 4
        secs += f"""<div class="newp mod">{G.sec(i, m['label'], m['kpi'], f'p{i}')}<p class="note">{esc(sub)}</p><div class="mini">{G.kpi_grid(kis, cols)}</div>{tbl}
          <div class="gap">{G.fig(chart, ctitle, 'Situation à la date d’import.')}</div><h3>Constats</h3><ul class="pts">{ins}</ul></div>"""
    md = D.methodology(comp, meta, GEN)
    clean = lambda s: s.replace("🟢", "vert").replace("🟡", "orange").replace("🔴", "rouge").replace("⚪ ", "")
    lst = lambda ttl, xs: f'<h4>{esc(ttl)}</h4><ul class="pts">' + "".join(f"<li>{esc(clean(x))}</li>" for x in xs) + "</ul>"
    n = len(P.MODULES) + 3
    items.append((f"p{n}", "Méthodologie & qualité des données"))
    short = {m["key"]: m["short"] for m in P.MODULES}
    body = G.cover_a("Reporting Métier", "PCHC", "Plateforme Tarkhiss — Produits Cosmétiques et d’Hygiène Corporelle",
                     [("Période", label), ("Situation au", D.fmt_dt(meta.get("imported_at"))), ("Édition", GEN.strftime("%d/%m/%Y"))], "Revue de direction")
    body += f"""{G.head_html()}<div>{G.sec(0, 'Sommaire', 'Revue de direction')}{toc(items)}</div>
    <div class="gap">{G.sec(1, 'Synthèse exécutive', f'Situation par volet — {label}', 'p1')}{synt}
      <p class="note">* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté.</p>
      <h3>À retenir</h3><ul class="pts">{keep}</ul>
      {G.fig(modules_chart(comp, short), 'Dossiers par volet', 'Taux de délivrance au-dessus de chaque barre.')}</div>
    <div class="gap keep">{G.sec(2, items[1][1], 'Résumé décisionnel', 'p2')}<div class="mini">{G.kpi_grid(cards, 3)}</div>{alerts}
      <p class="note">Seuils : « taux » → alerte si la valeur est inférieure au seuil ; « nombre » → supérieure. Seuils = paramètres du classeur.</p></div>
    {secs}
    <div class="newp small">{G.sec(n, items[-1][1], 'Sources et contrôles', f'p{n}')}<div class="two" style="break-inside:auto">
      <div>{lst('Sources', md['sources'])}{lst('Coquilles normalisées & statuts non référencés', md['typos'] + md['unknown'])}</div>
      <div>{lst('Contrôles', md['controls'])}{lst('Hypothèses des seuils', md['hypotheses'])}</div></div></div>"""
    render_pdf("PCHC_reporting_metier_2026-10.pdf", "", body)


# ----------------------------------------------------------------------------- Emails (tableaux + styles en ligne, Arial, images CID)
class Mail:
    W = 720

    def __init__(self):
        self.images = {}

    def img(self, uri, alt):
        cid = f"img{len(self.images) + 1}"
        self.images[cid] = base64.b64decode(uri.split(",", 1)[1])
        return (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {LINE};margin:0 0 12px;"><tr><td style="padding:10px;">'
                f'<div style="font:bold 13px Arial,sans-serif;color:{FOREST};margin:0 0 6px;">{esc(alt)}</div>'
                f'<img src="cid:{cid}" width="{self.W - 42}" alt="{esc(alt)}" style="display:block;width:100%;max-width:{self.W - 42}px;height:auto;border:0;"></td></tr></table>')

    @staticmethod
    def sec(n, title):
        idx = (f'<td width="30" align="center" bgcolor="{LIGHT}" style="font:bold 12px Arial,sans-serif;color:{GREEN};width:30px;height:30px;">{n:02d}</td><td width="10" style="width:10px;"></td>' if n else "")
        return (f'<table role="presentation" cellpadding="0" cellspacing="0" style="margin:22px 0 10px;"><tr>{idx}'
                f'<td style="font:bold 17px Arial,sans-serif;color:{FOREST};">{esc(title)}</td></tr></table>')

    @staticmethod
    def kpis(items, cols=4):
        tone = {"alert": (("#FFFFFF", RED), RED), "good": ((LIGHT, GREEN), MUTED), "warn": ((FOREST, "#FFFFFF"), "#C6DCCD")}
        tds = ""
        for i, it in enumerate(items):
            label, val, sub, t = (tuple(it) + ("", ""))[:4]
            if i:
                tds += '<td width="6" style="width:6px;font-size:0;line-height:0;">&nbsp;</td>'
            (bg, vc), sc = tone.get(t, (("#FFFFFF", FOREST), MUTED))
            lc = "#B1CFBA" if t == "warn" else MUTED
            tds += (f'<td width="{100 // cols}%" valign="top" bgcolor="{bg}" style="background:{bg};border:1px solid {FOREST if t == "warn" else LINE};'
                    f'{f"border-top:3px solid {RED};" if t == "alert" else ""}padding:9px 10px;">'
                    f'<div style="font:bold 10px Arial,sans-serif;color:{lc};text-transform:uppercase;">{esc(label)}</div>'
                    f'<div style="font:bold 21px Arial,sans-serif;color:{vc};padding:3px 0 2px;">{esc(val)}</div>'
                    f'<div style="font:11px Arial,sans-serif;color:{sc};">{esc(sub) or "&nbsp;"}</div></td>')
        return f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 6px;"><tr>{tds}</tr></table>'

    @staticmethod
    def table(heads, rows, aligns):
        th = "".join(f'<th align="{a}" bgcolor="{FOREST}" style="background:{FOREST};color:#FFFFFF;font:bold 12px Arial,sans-serif;padding:8px 9px;text-align:{a};">{esc(h)}</th>' for h, a in zip(heads, aligns))
        body = ""
        for r, row in enumerate(rows):
            cells = row["cells"] if isinstance(row, dict) else row
            cls = row.get("cls", "hl") if isinstance(row, dict) else ""
            bg = {"hl": "#FDECED", "tot": LIGHT}.get(cls, "#FFFFFF" if r % 2 == 0 else PALE)
            body += "<tr>" + "".join(f'<td align="{a}" bgcolor="{bg}" style="background:{bg};border-bottom:1px solid {LINE};font:13px Arial,sans-serif;color:{INK};padding:7px 9px;text-align:{a};{"white-space:nowrap;" if a == "center" else ""}">{c}</td>' for c, a in zip(cells, aligns)) + "</tr>"
        return f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 6px;"><tr>{th}</tr>{body}</table>'

    @staticmethod
    def note(text):
        if not (text or "").strip():
            return ""
        ps = "".join(f'<p style="margin:0 0 6px;">{esc(p)}</p>' for p in text.split("\n") if p.strip())
        return (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:6px 0 4px;"><tr><td bgcolor="{LIGHT}" '
                f'style="background:{LIGHT};border-left:3px solid {GREEN};padding:10px 14px;font:13px/1.5 Arial,sans-serif;color:{INK};">{ps}</td></tr></table>')

    @staticmethod
    def bullets(items):
        return "".join(f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td width="16" valign="top" style="font:13px Arial,sans-serif;color:{GREEN};padding:6px 0;border-bottom:1px solid {LINE};">&#9632;</td>'
                       f'<td style="font:13px/1.45 Arial,sans-serif;color:{INK};padding:6px 0;border-bottom:1px solid {LINE};">{esc(i)}</td></tr></table>' for i in items)

    def shell(self, title, period, greeting, intro, blocks, signature, kind="Rapport mensuel"):
        rule = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td width="90%" height="4" bgcolor="{FOREST}" style="background:{FOREST};font-size:0;line-height:0;">&nbsp;</td>'
                f'<td width="10%" height="4" bgcolor="{RED}" style="background:{RED};font-size:0;line-height:0;">&nbsp;</td></tr></table>')
        self.images["logo"] = base64.b64decode(G.LOGO.split(",", 1)[1])
        return f"""<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr><td align="left" style="padding:0 0 14px;font:14px/1.5 Arial,sans-serif;color:{INK};text-align:left;">
          <p style="margin:0 0 8px;">{esc(greeting)}</p><p style="margin:0;">{esc(intro)}</p></td></tr>
        <tr><td align="center">
        <table role="presentation" width="{self.W}" cellpadding="0" cellspacing="0" style="width:{self.W}px;max-width:100%;">
        <tr><td><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
          <td width="150" style="padding:6px 0 10px;"><img src="cid:logo" width="130" alt="AMMPS" style="display:block;width:130px;height:auto;border:0;"></td>
          <td style="font:12px/1.4 Arial,sans-serif;color:{FOREST};border-left:1px solid {LINE};padding-left:14px;">Agence Marocaine du Médicament<br>et des Produits de Santé</td>
          <td align="right" style="font:bold 11px Arial,sans-serif;color:{MUTED};text-transform:uppercase;letter-spacing:1px;">{esc(kind)}<br><span style="font-size:14px;color:{FOREST};text-transform:none;letter-spacing:0;">{esc(period)}</span></td>
        </tr></table>{rule}</td></tr>
        <tr><td style="padding:16px 0 2px;font:bold 25px/1.2 Arial,sans-serif;color:{FOREST};text-align:left;">{esc(title)}</td></tr>
        <tr><td style="text-align:left;">{blocks}</td></tr>
        </table></td></tr>
        <tr><td align="left" style="padding:22px 0 0;font:14px/1.5 Arial,sans-serif;color:{INK};text-align:left;">Cordialement,{'<br>' + signature if signature else ''}</td></tr>
        </table>"""

    def save(self, base, subject, html, attachments=()):
        html = re.sub(r"(\d)[  ]%", "\\1&nbsp;%", html)
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["X-Unsent"] = "1"
        msg.set_content("Ce message nécessite un client compatible HTML.")
        msg.add_alternative(html, subtype="html")
        part = msg.get_payload()[1]
        for cid, data in self.images.items():
            part.add_related(data, maintype="image", subtype="png", cid=f"<{cid}>", filename=f"{cid}.png", disposition="inline")
        for data, mime, name in attachments:
            msg.add_attachment(data, maintype=mime.split("/")[0], subtype=mime.split("/")[1], filename=name)
        with open(os.path.join(EX_DIR, base + ".eml"), "wb") as f:
            f.write(bytes(msg))
        prev = html
        for cid, data in self.images.items():
            prev = prev.replace(f'"cid:{cid}"', '"data:image/png;base64,' + base64.b64encode(data).decode() + '"')
        with open(os.path.join(EX_DIR, base + "_apercu.html"), "w", encoding="utf-8") as f:
            f.write(f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{esc(subject)}</title></head>'
                    f'<body style="margin:0;padding:16px;background:#fff;">{prev}</body></html>')
        print(base, "objet :", subject, "|", os.path.getsize(os.path.join(EX_DIR, base + ".eml")) // 1024, "Ko | SANAD/GLPI :", bool(BANNED_MAIL.search(html)))


def signature(ms):
    return "<br>".join(esc(p) for p in (ms.get("signature_name"), ms.get("signature_function"), ms.get("signature_phone")) if p)


def mail_tarkhiss(app):
    R.update(G.TARKHISS_BASE)
    ms = app.load_module_settings("tarkhiss")
    es = ms.get("email_sections", {})
    on = lambda k: es.get(k, True)
    m = Mail()
    b = m.kpis(R["kpi_volume"]) + m.kpis(R["kpi_quali"]) + m.kpis(R["kpi_charge"], 2)
    b = m.sec(1, "Indicateurs clés") + b
    if on("graphs"):
        b += m.sec(2, "Analyse graphique")
        if on("chart_weekly"):
            b += m.img(G.chart_weekly(), "Évolution hebdomadaire — Bugs vs Demandes")
        if on("chart_problems"):
            b += m.img(G.chart_rank(), "Répartition des problèmes techniques")
    cat = lambda items: [[f"<b>{esc(l)}</b>", f"<b>{n}</b>", p, f'<span style="color:{MUTED};font-size:12px;">{esc(d)}</span>'] for l, n, p, d in items]
    al = ["left", "center", "center", "left"]
    if on("problems"):
        b += m.sec(3, "Problèmes techniques signalés (166)") + m.table(["Type", "Nb", "%", "Description"], cat(R["problems"]), al)
    if on("demandes"):
        b += m.sec(4, "Demandes d’information reçues (91)") + m.table(["Type", "Nb", "%", "Description"], cat(R["demandes"]), al)
    if on("weekly"):
        rows = []
        for i, (w, x, y, o) in enumerate(R["weekly"]):
            c = [f'<b style="white-space:nowrap;">{w}</b>', f"<b>{x}</b>", f"<b>{y}</b>", f'<span style="color:{MUTED};font-size:12px;">{esc(o)}</span>']
            rows.append({"cells": c} if i == R["crit"] else c)
        b += m.sec(5, "Évolution hebdomadaire") + m.table(["Semaine", "Bugs", "Demandes", "Observations"], rows, al)
    if on("synth"):
        b += m.sec(6, "Constats clés") + m.bullets(R["constats"]) + m.sec(0, "Recommandations") + "".join(m.note(f"{i:02d}   {r}") for i, r in enumerate(R["recos"], 1))
    intro = f"Veuillez trouver ci-dessous le rapport mensuel du support Tarkhiss ({R['period']})."
    html = app.strip_platform_mentions(m.shell("Rapport Support Tarkhiss", R["period"], ms.get("greeting", "Bonjour,"), intro, b, signature(ms)))
    m.save("Tarkhiss_email_2026-10", app.build_subject("tarkhiss", R["period"]), html)


def mail_moussanada(app, c):
    ms = app.load_module_settings("moussanada")
    es = ms.get("email_sections", {})
    on = lambda k: es.get(k, True)
    ch, s = c["charts"], c["sections"]
    m = Mail()
    im = lambda k, t: m.img(ch[k], t) if on(f"chart_{k.removesuffix('_wide')}") else ""
    al = ["left", "center", "center", "center"]
    b = ""
    if on("volume"):
        b += m.sec(1, "Volume global des tickets") + m.kpis(c["k1"]) + im("volume", "Volumes de tickets — comparatif") + im("backlog", "Backlog — tickets en instance") + m.note(s.get("volume"))
    if on("delays"):
        b += m.sec(2, "Évolution des délais de traitement") + m.kpis(c["k2"]) + im("trend", "Tendance des tickets") + m.note(s.get("delays"))
    if on("categories"):
        b += m.sec(3, f"Principales catégories de demandes en {c['ml']}") + m.table(["Catégorie"] + TOP_HEAD, c["top_rows"](c["cats"]), al) + im("cat", "Top catégories") + im("pareto", "Pareto des catégories") + m.note(s.get("categories"))
    if on("services"):
        b += m.sec(4, "Services les plus demandeurs") + m.table(["Service"] + TOP_HEAD, c["top_rows"](c["srv"]), al) + im("srv", "Top services demandeurs") + m.note(s.get("services"))
    if on("team"):
        b += m.sec(5, "Charge et mobilisation de l’équipe SI") + m.table(TECH_HEAD, c["tech_rows"], ["left"] + ["center"] * 5) + im("techvol_wide", "Volume de sollicitation") + im("tech", "Détail des sollicitations") + im("radar", "Radar de performance") + m.note(s.get("team"))
    if on("comparative"):
        b += m.sec(6, f"Analyse comparative {c['pl']} / {c['ml']}") + im("delay_wide", "Délais de traitement (heures)") + m.note(s.get("comparative"))
    if on("synth"):
        b += m.sec(7, "Synthèse finale et conclusion") + m.note(s.get("synthesis"))
    intro = f"Veuillez trouver ci-dessous le rapport mensuel du support Moussanada ({c['ml']})."
    html = app.strip_platform_mentions(m.shell("Rapport Support Moussanada", c["ml"], ms.get("greeting", "Bonjour,"), intro, b, signature(ms)))
    m.save("Moussanada_email_2026-09", app.build_subject("moussanada", c["ml"]), html)


if __name__ == "__main__":
    G.build_fonts()
    G.TARKHISS_BASE = {k: R[k] for k in ("title", "short", "period")}
    app, work = sandbox.load_app()
    print("copie isolée :", work)
    with app.app.app_context(), app.app.test_request_context():
        c = moussanada_ctx(app)
        pdf_moussanada(app, c)
        pdf_pchc(app)
        mail_tarkhiss(app)
        mail_moussanada(app, c)
