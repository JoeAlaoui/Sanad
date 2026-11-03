# -*- coding: utf-8 -*-
"""
ammps_charte.py — Charte graphique AMMPS des rapports (variante A validée) : couleurs, polices,
graphiques matplotlib, composants PDF (WeasyPrint), composants email (tableaux + styles en ligne)
et classe Deck (PPTX 16:9). Aucune donnée ici : les volets lui passent des valeurs déjà calculées.

Ressources : static/charte/ (logo, Montserrat et Inter statiques). PDF en Montserrat (titres, chiffres
clés) / Inter (texte, comme le dashboard) ; graphiques (images) en Inter partout ; texte des emails, PPTX
et Excel en Arial, seule police garantie sur les postes (ni Outlook ni PowerPoint n'embarquent les
autres). Les images d'un email restent en data: dans le HTML (aperçu) ; inline_images() les
transforme en pièces CID à la création du .eml.
"""
import base64
import hashlib
import html as html_module
import io
import math
import os
import re
import textwrap
from pathlib import Path

FOREST, GREEN, RED, INK, MUTED, LINE, PALE, LIGHT, DEEP, TINT, AMBER = (
    "#0d4823", "#058b44", "#df1e26", "#183928", "#5b6e62", "#dce6df", "#f3f7f4", "#eaf4ed", "#0b3420", "#8cc9a3", "#c98a1b")
FONT = "Arial"
AGENCY = "Agence Marocaine du Médicament et des Produits de Santé"
AGENCY_LINES = ["Agence Marocaine du Médicament", "et des Produits de Santé"]
CONTACT = "Rue Lamfadal Charkaoui, B.P. 6206 · Madinat Al Irfane, Rabat · Maroc"
ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "charte")
esc = html_module.escape


def hx(color):
    return color.lstrip("#").upper()


def data_uri(png):
    return "data:image/png;base64," + base64.b64encode(png).decode("ascii")


def uri_bytes(img):
    return img if isinstance(img, bytes) else base64.b64decode(img.split(",", 1)[1])


with open(os.path.join(ASSETS, "logo-ammps.png"), "rb") as _f:
    LOGO_PNG = _f.read()
LOGO = data_uri(LOGO_PNG)


def fr_pct(v, d=1):
    return f"{v:.{d}f}".replace(".", ",") + " %"


def delta_label(cur, prev, prev_label, fmt=str):
    """« −27 % vs septembre (283) », ou « nouveau » si le mois précédent est à zéro."""
    if not prev:
        return "nouveau" if cur else ""
    diff = (cur - prev) / prev * 100
    if diff >= 1000:   # mois de référence quasi vide : un pourcentage n'aurait pas de sens
        return f"×{cur / prev:.0f} vs {prev_label} ({fmt(prev)})"
    return f"{'+' if diff >= 0 else '−'}{abs(diff):.0f} % vs {prev_label} ({fmt(prev)})"


# ----------------------------------------------------------------------------- graphiques
def mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    known = {f.fname for f in font_manager.fontManager.ttflist}
    for w in (400, 500, 700):
        path = os.path.join(ASSETS, f"Inter-{w}.ttf")
        if path not in known:
            font_manager.fontManager.addfont(path)
    # Inter : police du dashboard, partagée par les graphiques des emails, PDF et PPTX (DejaVu : repli des symboles absents)
    plt.rcParams.update({"font.family": ["Inter", "DejaVu Sans"], "text.color": INK, "axes.edgecolor": LINE, "axes.labelcolor": MUTED,
                         "axes.facecolor": "white", "axes.grid": False, "xtick.color": INK, "ytick.color": MUTED, "font.size": 10})
    return plt


def png(fig, dpi=200):
    import matplotlib.pyplot as plt
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return data_uri(buf.getvalue())


def _flat(ax):
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0)


def round_bars(fig, ax, horizontal=False, radius_pt=4.3):
    """Arrondit l'extrémité libre des barres déjà tracées (coins du sommet, pas ceux de la base), comme
    le borderRadius des graphiques du dashboard. À appeler une fois axes et limites définitifs : le rayon
    est calculé en points, donc identique quelle que soit l'échelle des données."""
    from matplotlib.patches import FancyBboxPatch, Rectangle
    pos = ax.get_position()
    w_in, h_in = pos.width * fig.get_figwidth(), pos.height * fig.get_figheight()
    (x0, x1), (y0, y1) = ax.get_xlim(), ax.get_ylim()
    xu, yu = abs(x1 - x0) / w_in, abs(y1 - y0) / h_in     # unités de données par pouce
    ax.set_xlim(x0, x1)
    ax.set_ylim(y0, y1)
    for p in [p for p in ax.patches if isinstance(p, Rectangle)]:
        x, y, w, h = p.get_x(), p.get_y(), p.get_width(), p.get_height()
        along, across = (w, h) if horizontal else (h, w)
        u_along, u_across = (xu, yu) if horizontal else (yu, xu)
        if not along or not across:
            continue
        r = min(radius_pt / 72, abs(along) / u_along / 2, abs(across) / u_across / 2)
        color, z = p.get_facecolor(), p.get_zorder()
        rs = r * xu
        box = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={rs}", mutation_aspect=yu / xu,
                             facecolor=color, edgecolor="none", linewidth=0, zorder=z)
        cover = Rectangle((x, y), r * xu if horizontal else w, h if horizontal else r * yu * (1 if along > 0 else -1),
                          facecolor=color, edgecolor="none", linewidth=0, zorder=z)
        p.remove()
        ax.add_patch(box)
        ax.add_patch(cover)


TECH_VOL, TECH_OPEN, TECH_DONE = GREEN, GREEN, TINT     # verts de la charte : volume et ouverts en vert, résolus en vert clair
DASH_GRID, DASH_MUTED, TECH_HI, TECH_LO = "#E5E7EB", "#6B7280", "#B45309", "#0F766E"


def dash_bars(cats, series, colors, figsize=(8.2, 2.6), marks=None):
    """Barres au style du dashboard : axe Y et quadrillage clairs, valeurs en gras au-dessus des barres, coins
    arrondis, légende en bas (plusieurs séries), ▲ / ▼ sur le technicien le plus / le moins chargé
    (`marks` = volume de chaque catégorie ; aucun repère si tous sont égaux)."""
    import numpy as np
    from matplotlib.patches import Patch
    from matplotlib.ticker import MaxNLocator
    plt = mpl()
    k = len(series)
    per = figsize[0] / max(len(cats), 1)
    size = 9 if per >= 1.0 else 8 if per >= .85 else 6.5
    labels = [textwrap.wrap(str(c), 11) or [str(c)] for c in cats]
    i_max = i_min = -1
    if marks and len(marks) > 1 and max(marks) != min(marks):
        i_max, i_min = marks.index(max(marks)), len(marks) - 1 - marks[::-1].index(min(marks))
        labels[i_max][0] = "▲ " + labels[i_max][0]
        labels[i_min][0] = "▼ " + labels[i_min][0]
    legend = k > 1
    fig, ax = plt.subplots(figsize=figsize)
    fig.subplots_adjust(left=.055, right=.995, top=.93, bottom=(.22 + .15 * max(len(l) for l in labels) + (.3 if legend else 0)) / figsize[1])
    x = np.arange(len(cats))
    w, gap = (.42, 0) if k == 1 else (.3, .04)
    top = max([max(v, default=0) for v in series.values()] + [1])
    for i, (name, vals) in enumerate(series.items()):
        off = (i - (k - 1) / 2) * (w + gap)
        ax.bar(x + off, vals, w, color=colors[i % len(colors)], zorder=3)
        for xi, v in zip(x + off, vals):
            ax.text(xi, v + top * .025, f"{v:.0f}", ha="center", fontsize=9.5, fontweight="bold", color=INK, zorder=4)
    ax.set_xticks(x, ["\n".join(l) for l in labels], fontsize=size)
    for i, t in enumerate(ax.get_xticklabels()):
        t.set_color(TECH_HI if i == i_max else TECH_LO if i == i_min else DASH_MUTED)
        t.set_fontweight(500 if i in (i_max, i_min) else 400)
    ax.set_xlim(-.5, len(cats) - .5)
    ax.set_ylim(0, top * 1.2)
    ax.yaxis.set_major_locator(MaxNLocator(integer=True, nbins=4))
    ax.grid(axis="y", color=DASH_GRID, linewidth=.8)
    ax.set_axisbelow(True)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color(DASH_GRID)
    ax.tick_params(axis="x", length=0, pad=6)
    ax.tick_params(axis="y", length=0, labelsize=8.5, labelcolor=DASH_MUTED)
    if legend:
        fig.legend(handles=[Patch(facecolor=colors[i % len(colors)], label=n) for i, n in enumerate(series)], loc="lower center", ncol=k, frameon=False,
                   fontsize=9, labelcolor=DASH_MUTED, handlelength=.8, handleheight=.8, handletextpad=.5, columnspacing=1.6, bbox_to_anchor=(.5, 0))
    round_bars(fig, ax)
    return png(fig)


def bars(cats, series, figsize=(8.2, 2.4), colors=None, rotate=False):
    import numpy as np
    plt = mpl()
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
    # « Prénom Nom » sur deux lignes, jamais en diagonale : sinon les noms voisins se touchent et se lisent
    # comme une phrase. `rotate` (graphiques étroits) réduit seulement la police pour que tout tienne.
    per = figsize[0] / max(len(cats), 1)
    ax.set_xticks(x, ["\n".join(textwrap.wrap(str(c), 11)) or str(c) for c in cats], fontsize=9 if per >= 1.0 else 8 if per >= .85 else 6.5)
    _flat(ax)
    ax.tick_params(axis="x", pad=6)
    ax.set_ylim(0, top * (1.3 if len(series) > 1 else 1.15))
    if len(series) > 1:
        ax.legend(frameon=False, loc="upper left", ncol=len(series), fontsize=9)
    round_bars(fig, ax)
    return png(fig)


def hbars(labels, series, figsize=None, colors=(FOREST, TINT), maxlen=40):
    import numpy as np
    plt = mpl()
    n = len(labels)
    k = len(series)
    # hauteur par catégorie assez grande pour que les valeurs de deux barres voisines ne se chevauchent pas
    fig, ax = plt.subplots(figsize=figsize or (8.2, (.17 * k + .1) * n + .75))
    fig.subplots_adjust(top=1 - .3 / fig.get_figheight() if k > 1 else .98)   # place de la légende, au-dessus des barres
    y = np.arange(n)
    h = .78 / k
    top = max([max(v, default=0) for v in series.values()] + [1])
    bar_pt = fig.get_figheight() * ax.get_position().height / max(n, 1) * h * .92 * 72
    size = min(8.5, max(6.5, bar_pt * .85))
    for i, (name, vals) in enumerate(series.items()):
        bs = ax.barh(y + i * h - .39 + h / 2, vals, h * .92, color=colors[i % len(colors)], label=name)
        for b in bs:
            ax.text(b.get_width() + top * .014, b.get_y() + b.get_height() / 2, f"{b.get_width():.0f}", va="center", fontsize=size, fontweight="bold")
    ax.set_yticks(y, [(l[:maxlen] + "…") if len(l) > maxlen else l for l in labels])
    ax.invert_yaxis()
    ax.set_xticks([])
    ax.set_xlim(0, top * 1.12)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.tick_params(axis="y", length=0, labelcolor=INK, labelsize=9)
    if k > 1:   # légende au-dessus du graphique, hors des barres et de leurs valeurs
        ax.legend(frameon=False, loc="lower right", bbox_to_anchor=(1, 1), ncol=k, fontsize=9, borderaxespad=.15)
    ax.set_ylim(n - .5, -.5)
    round_bars(fig, ax, horizontal=True)
    return png(fig)


def lines(cats, series, figsize=(8.2, 2.3), colors=(FOREST, GREEN, TINT)):
    plt = mpl()
    fig, ax = plt.subplots(figsize=figsize)
    for (name, vals), col in zip(series.items(), colors):
        ax.plot(cats, vals, marker="o", linewidth=2.2, markersize=4.5, color=col, label=name)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color=LINE, linewidth=.6)
    ax.tick_params(length=0)
    ax.legend(frameon=False, loc="lower left", bbox_to_anchor=(0, 1), ncol=len(series), fontsize=9)
    return png(fig)


def pareto(cats, vals, cum, figsize=(8.2, 2.7)):
    plt = mpl()
    fig, ax = plt.subplots(figsize=figsize)
    ax.bar(range(len(cats)), vals, color=TINT, width=.7)
    ax.set_xticks(range(len(cats)), cats, rotation=30, ha="right", rotation_mode="anchor", fontsize=8)
    _flat(ax)
    ax2 = ax.twinx()
    ax2.plot(range(len(cats)), cum, color=FOREST, marker="o", markersize=4, linewidth=2)
    for i, c in enumerate(cum):
        ax2.text(i, c + 4, f"{c} %", ha="center", fontsize=8, fontweight="bold", color=FOREST)
    ax2.set_ylim(0, 118)
    ax2.set_yticks([])
    ax2.spines[["top", "right", "left"]].set_visible(False)
    round_bars(fig, ax)
    return png(fig)


def radar(techs, figsize=(4.4, 3.8)):
    import numpy as np
    plt = mpl()
    rt = techs[:6]
    mo = max([t.get("ouverts", 0) for t in rt] + [1])
    md = max([(t.get("delai_resolution_s") or 0) for t in rt] + [1])
    axes = ["Ouverts", "Résolus", "Taux\nrésolution", "Rapidité"]
    ang = np.linspace(0, 2 * np.pi, 4, endpoint=False).tolist()
    ang += ang[:1]
    fig, ax = plt.subplots(figsize=figsize, subplot_kw=dict(polar=True))
    for t, col, ls in zip(rt, (FOREST, GREEN, RED, MUTED, TINT, INK), ("-", "-", "-", "--", "--", "--")):
        o, r = t.get("ouverts", 0), t.get("resolus", 0)
        rate = min(100, round(r / o * 100)) if o else 0  # plafonné : résolus du mois pouvant inclure des tickets plus anciens
        vals = [round(o / mo * 100), rate, rate, round(100 - (t.get("delai_resolution_s") or 0) / md * 100) if t.get("delai_resolution_s") else 100]
        ax.plot(ang, vals + vals[:1], color=col, linewidth=1.8, linestyle=ls, label=t["label"])
    ax.set_xticks(ang[:-1], axes, fontsize=8.5)
    ax.tick_params(axis="x", pad=13)
    ax.set_yticklabels([])
    ax.set_ylim(0, 100)
    ax.spines["polar"].set_color(LINE)
    ax.grid(color=LINE)
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.17), ncol=2, frameon=False, fontsize=7.5)
    return png(fig)


def weekly(items, crit=None, figsize=(8.2, 3.4), short=False):
    """items = [(semaine, bugs, demandes, …)] ; la semaine critique ressort en rouge."""
    import numpy as np
    plt = mpl()
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(len(items))
    b_v, d_v = [w[1] for w in items], [w[2] for w in items]
    top = max(b_v + d_v + [1])
    b1 = ax.bar(x - .19, b_v, .36, color=FOREST, label="Bugs")
    b2 = ax.bar(x + .19, d_v, .36, color=TINT, label="Demandes")
    if crit is not None:
        b1[crit].set_color(RED)
    for group in (b1, b2):
        for b in group:
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + top * .02, int(b.get_height()), ha="center", fontsize=9.5, fontweight="bold")
    ax.set_xticks(x, [re.sub(r"\s+[^\d\s]+$", "", str(w[0])) if short else str(w[0]) for w in items])
    ax.set_yticks([])
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.tick_params(axis="x", length=0)
    ax.legend(frameon=False, loc="upper left", ncol=2, fontsize=9)
    ax.set_ylim(0, top * 1.18)
    round_bars(fig, ax)
    return png(fig)


def rank(items, figsize=(8.2, 3.3)):
    """items = [(libellé, nombre, pourcentage formaté, …)], du plus grand au plus petit."""
    plt = mpl()
    fig, ax = plt.subplots(figsize=figsize)
    rev = items[::-1]
    top = max([i[1] for i in items] + [1])
    colors = [FOREST if i >= len(rev) - 2 else TINT for i in range(len(rev))]
    bs = ax.barh([i[0] for i in rev], [i[1] for i in rev], color=colors, height=.66)
    for b, it in zip(bs, rev):
        ax.text(b.get_width() + top * .015, b.get_y() + b.get_height() / 2, f"{it[1]}  ·  {it[2]}", va="center", fontsize=9.5, fontweight="bold")
    ax.set_xticks([])
    ax.set_xlim(0, top * 1.24)
    ax.set_ylim(-.6, max(len(items), 4) - .4)   # une seule catégorie : barre de hauteur normale
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.tick_params(axis="y", length=0, labelcolor=INK, labelsize=10)
    round_bars(fig, ax, horizontal=True)
    return png(fig)


def heat(matrix, rows, cols, figsize=(8.2, 2.6)):
    import numpy as np
    from matplotlib.colors import LinearSegmentedColormap
    plt = mpl()
    cmap = LinearSegmentedColormap.from_list("ammps", [PALE, TINT, GREEN, FOREST])
    data = np.array(matrix, float)
    top = data.max() or 1
    fs = 10 if len(cols) <= 12 else 7
    fig, ax = plt.subplots(figsize=figsize)
    ax.imshow(data, cmap=cmap, aspect="auto", vmin=0, vmax=top)
    for y, row in enumerate(matrix):
        for x, v in enumerate(row):
            ax.text(x, y, f"{v:g}".replace(".", ",") if v else "–", ha="center", va="center", fontsize=fs, fontweight="bold" if v else "normal",
                    color="white" if v >= top * .55 else (INK if v else "#a9b8ae"))
    ax.set_xticks(range(len(cols)), cols)
    ax.set_yticks(range(len(rows)), rows)
    ax.xaxis.tick_top()
    ax.tick_params(length=0, labelsize=fs + 1)
    ax.set_xticks(np.arange(-.5, len(cols)), minor=True)
    ax.set_yticks(np.arange(-.5, len(rows)), minor=True)
    ax.grid(which="minor", color="white", linewidth=2.5)
    ax.tick_params(which="minor", length=0)
    ax.spines[:].set_visible(False)
    return png(fig)


# ----------------------------------------------------------------------------- PDF : composants
def kpi(label, value, sub="", tone=""):
    return (f'<div class="kpi {tone}"><div class="kl">{esc(str(label))}</div><div class="kv">{esc(str(value))}</div>'
            f'<div class="ks">{esc(str(sub)) or "&nbsp;"}</div></div>')


def kpi_grid(items, cols=4):
    return f'<div class="kgrid c{cols}">' + "".join(kpi(*i) for i in items) + "</div>"


def table(heads, rows, aligns, cls=""):
    """Une ligne peut être {"cells": […], "cls": "hl"|"tot"} ; les cellules sont du HTML déjà échappé."""
    th = "".join(f'<th class="{a}">{esc(h)}</th>' for h, a in zip(heads, aligns))
    body = ""
    for row in rows:
        rcls = row.get("cls", "hl") if isinstance(row, dict) else ""
        cells = row["cells"] if isinstance(row, dict) else row
        body += f'<tr class="{rcls}">' + "".join(f'<td class="{a}">{c}</td>' for c, a in zip(cells, aligns)) + "</tr>"
    return f'<table class="data {cls}"><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>'


def sec(n, title, eyebrow="", anchor=""):
    idx = f'<span class="idx">{n:02d}</span>' if n else ""
    eb = f'<div class="eyebrow">{esc(eyebrow)}</div>' if eyebrow else ""
    return f'<div class="sec" id="{anchor}">{idx}<div>{eb}<h2>{esc(title)}</h2></div></div>'


def fig(src, title, caption):
    if not src:
        return ""
    return f'<figure><div class="ft">{esc(title)}</div><img src="{src}"><figcaption>{esc(caption)}</figcaption></figure>'


def notebox(text, label="Analyse"):
    if not (text or "").strip():
        return ""
    paras = "".join(f"<p>{esc(p)}</p>" for p in text.split("\n") if p.strip())
    return f'<div class="notebox"><div class="nl">{esc(label)}</div>{paras}</div>'


def points(items):
    return '<ul class="pts">' + "".join(f"<li>{esc(str(i))}</li>" for i in items) + "</ul>"


def recos(items):
    rows = "".join(f'<li><span class="n">{i:02d}</span><span>{esc(str(r))}</span></li>' for i, r in enumerate(items, 1))
    return f'<ol class="recos">{rows}</ol>'


def toc(items):
    rows = "".join(f'<li><a href="#{a}"><span class="tn">{i:02d}</span><span class="tl">{esc(l)}</span><span class="tp" data-t="#{a}"></span></a></li>'
                   for i, (a, l) in enumerate(items, 1))
    return f'<ol class="toc">{rows}</ol>'


def head_html(title, period):
    return (f'<div class="head-wrap"><div class="head"><img src="{LOGO}"><div class="ag">{"<br>".join(AGENCY_LINES)}</div>'
            f'<div class="rt">{esc(title)}<br><b>{esc(period)}</b></div></div><div class="rule"></div></div>')


def cover(title, title2, subtitle, meta, eyebrow="Rapport mensuel", extra=""):
    rows = "".join(f"<div><span>{esc(k)}</span><b>{esc(v)}</b></div>" for k, v in meta)
    second = f"<br><em>{esc(title2)}</em>" if title2 else ""
    return f"""<section class="cover">
      <div class="cv-top"><img src="{LOGO}"></div><div class="rule cv-rule"></div>
      <div class="cv-main"><div class="eyebrow">{esc(eyebrow)}</div>
        <h1>{esc(title)}{second}</h1>{extra}
        <div class="cv-sub">{esc(subtitle)}</div>
        <div class="cv-meta">{rows}</div></div>
      <div class="cv-foot"><b>{esc(AGENCY)}</b>{esc(CONTACT)}</div></section>"""


def font_css():
    faces = []
    for fam, ws in (("Montserrat", (600, 700, 800)), ("Inter", (400, 500, 700))):
        faces += [f"@font-face{{font-family:{fam};src:url('{Path(ASSETS, f'{fam}-{w}.ttf').as_uri()}');font-weight:{w};}}" for w in ws]
    return "".join(faces)


def base_css(short, period):
    foot = f"AMMPS  ·  {short} — {period}".replace('"', "'")
    return font_css() + f"""
    @page {{ size:A4; margin:27mm 17mm 20mm;
      @top-left {{ content:element(head); width:100%; }}
      @bottom-left {{ content:"{foot}"; font:8pt Inter; color:{MUTED}; }}
      @bottom-right {{ content:"Page " counter(page) " / " counter(pages); font:700 8pt Inter; color:{FOREST}; }} }}
    @page cover {{ margin:0; @top-left {{ content:none; }} @bottom-left {{ content:none; }} @bottom-right {{ content:none; }} }}
    html {{ font-family:Inter,Arial,sans-serif; font-size:9.5pt; line-height:1.45; color:{INK}; }}
    body {{ margin:0; }}
    h1,h2,h3,h4 {{ font-family:Montserrat,Arial,sans-serif; margin:0; color:{FOREST}; break-after:avoid; }}
    h2 {{ font-size:15pt; font-weight:700; letter-spacing:-.2pt; line-height:1.2; }}
    h3 {{ font-size:10.5pt; font-weight:700; margin:9pt 0 5pt; }}
    h4 {{ font-size:8.5pt; font-weight:700; margin:9pt 0 4pt; }}
    p {{ margin:0; }}
    .head {{ display:flex; align-items:center; gap:9pt; padding-bottom:5pt; }}
    .head img {{ height:11mm; }}
    .head .ag {{ border-left:.6pt solid {LINE}; padding-left:9pt; font-size:7.2pt; line-height:1.3; color:{FOREST}; }}
    .head .rt {{ margin-left:auto; text-align:right; font-size:7.5pt; color:{MUTED}; }}
    .head-wrap {{ position:running(head); }}
    .rule {{ height:3pt; background:linear-gradient(90deg,{FOREST} 90%,{RED} 90%); }}
    .eyebrow {{ font:700 7pt Montserrat; letter-spacing:1.3pt; text-transform:uppercase; color:{FOREST}; margin-bottom:3pt; }}
    .eyebrow::before {{ content:''; display:inline-block; vertical-align:middle; width:15pt; height:1.6pt; margin:-1pt 6pt 0 0; background:{RED}; }}
    .sec {{ display:flex; gap:9pt; align-items:center; margin:0 0 9pt; break-after:avoid; }}
    .sec > div {{ flex:1; }}
    .idx {{ width:24pt; height:24pt; flex:none; border-radius:5pt; background:{LIGHT}; color:{GREEN}; font:700 9pt/24pt Montserrat; text-align:center; }}
    .newp {{ break-before:page; }}
    .gap {{ margin-top:16pt; }}
    .kgrid {{ display:grid; gap:6pt; margin-bottom:6pt; break-inside:avoid; }}
    .c4 {{ grid-template-columns:repeat(4,1fr); }} .c3 {{ grid-template-columns:repeat(3,1fr); }} .c2 {{ grid-template-columns:repeat(2,1fr); }}
    .kpi {{ border:.6pt solid {LINE}; border-radius:5pt; padding:7pt 8pt; background:#fff; }}
    .kl {{ font:600 6.8pt/1.25 Montserrat; color:{MUTED}; text-transform:uppercase; letter-spacing:.3pt; min-height:17pt; }}
    .mini .kl {{ min-height:0; }} .mini .ks {{ display:none; }} .mini .kpi {{ padding:5pt 8pt; }} .mini .kv {{ font-size:12pt; }}
    .two {{ display:grid; grid-template-columns:1fr 1fr; gap:12pt; break-inside:avoid; }}
    .kv {{ font:700 15pt/1.2 Montserrat; color:{FOREST}; letter-spacing:-.3pt; }}
    .ks {{ font-size:7.2pt; color:{MUTED}; margin-top:1pt; }}
    .kpi.alert {{ border-top:2.4pt solid {RED}; }} .kpi.alert .kv,.kpi.alert .ks {{ color:{RED}; }}
    .kpi.good {{ background:{LIGHT}; border-color:#cae4d3; }} .kpi.good .kv {{ color:{GREEN}; }}
    .kpi.warn {{ background:{FOREST}; border-color:{FOREST}; }} .kpi.warn .kv {{ color:#fff; }} .kpi.warn .kl {{ color:#b1cfba; }} .kpi.warn .ks {{ color:#c6dccd; }}
    table.data {{ width:100%; border-collapse:collapse; font-size:8.6pt; }}
    table.data th {{ background:{FOREST}; color:#fff; font:600 7.6pt Montserrat; padding:5pt 7pt; text-align:left; }}
    table.data th.c {{ text-align:center; }}
    table.data td {{ padding:4.5pt 7pt; border-bottom:.6pt solid {LINE}; vertical-align:middle; }}
    table.data tr {{ break-inside:avoid; }}
    table.data tbody tr:nth-child(even) td {{ background:{PALE}; }}
    table.data tbody tr.hl td {{ background:#fdeced; border-bottom-color:#f3caca; }}
    table.data tbody tr.tot td {{ background:{LIGHT}; font-weight:700; color:{FOREST}; }}
    table.data.tight td {{ padding:2.2pt 7pt; }}
    td.c {{ text-align:center; white-space:nowrap; }}
    .w30 {{ width:32%; }} .w6 {{ width:6%; }} .w9 {{ width:9%; }} .w18 {{ width:18%; }} .w8 {{ width:8%; }} .w11 {{ width:11%; }} .w14 {{ width:14%; }}
    .desc {{ color:{MUTED}; font-size:8pt; line-height:1.3; }} .keep {{ break-inside:avoid; }}
    figure {{ margin:0 0 12pt; break-inside:avoid; border:.6pt solid {LINE}; border-radius:6pt; padding:8pt 10pt 7pt; }}
    figure img {{ width:100%; display:block; }}
    .ft {{ font:700 9pt Montserrat; color:{FOREST}; margin-bottom:5pt; }}
    figcaption {{ font-size:7.8pt; color:{MUTED}; margin-top:4pt; }}
    .note {{ font-size:8pt; color:{MUTED}; margin:3pt 0 6pt; }}
    .words {{ font-size:8.5pt; }}
    .notebox {{ background:{LIGHT}; border-left:2.4pt solid {GREEN}; padding:7pt 11pt; margin:6pt 0 4pt; break-inside:avoid; font-size:9pt; }}
    .notebox p {{ margin:0 0 4pt; }} .notebox p:last-child {{ margin:0; }}
    .notebox .nl {{ font:700 7pt Montserrat; letter-spacing:.8pt; text-transform:uppercase; color:{GREEN}; margin-bottom:3pt; }}
    ul.pts {{ list-style:none; margin:0; padding:0; }}
    ul.pts li {{ padding:4.5pt 0 4.5pt 13pt; border-bottom:.6pt solid {LINE}; position:relative; break-inside:avoid; }}
    ul.pts li::before {{ content:''; position:absolute; left:1pt; top:9pt; width:5pt; height:5pt; background:{GREEN}; }}
    ol.recos {{ list-style:none; margin:0; padding:0; }}
    ol.recos li {{ display:flex; gap:9pt; align-items:center; background:{LIGHT}; border-left:2.4pt solid {GREEN}; padding:6pt 10pt; margin-bottom:5pt; break-inside:avoid; }}
    ol.recos .n {{ font:700 9pt Montserrat; color:{GREEN}; }}
    ol.toc {{ list-style:none; margin:0; padding:0; }}
    ol.toc li {{ border-bottom:.6pt solid {LINE}; }}
    ol.toc a {{ display:flex; gap:10pt; padding:5.5pt 0; color:{INK}; text-decoration:none; }}
    .tn {{ font:700 8.5pt Montserrat; color:{GREEN}; width:18pt; }} .tl {{ flex:1; }}
    .tp::after {{ content:target-counter(attr(data-t), page); }}
    .cover {{ page:cover; height:296.5mm; position:relative; background:#fff; }}
    .cv-top {{ padding:18mm 20mm 0; display:flex; align-items:center; gap:18pt; }}
    .cv-top img {{ height:32mm; }}
    .cv-rule {{ margin:12mm 0 0; }}
    .cv-main {{ padding:42mm 20mm 0; }}
    .cv-main .eyebrow {{ font-size:9pt; letter-spacing:2pt; margin-bottom:10pt; }}
    .cv-main h1 {{ font-size:36pt; line-height:1.12; font-weight:800; letter-spacing:-1pt; }}
    .cv-main h1 em {{ font-style:normal; color:{GREEN}; }}
    .cv-sub {{ font-size:13pt; color:{MUTED}; margin-top:9pt; }}
    .cv-meta {{ margin-top:22mm; border-top:.6pt solid {LINE}; width:115mm; }}
    .cv-meta div {{ display:flex; padding:6pt 0; border-bottom:.6pt solid {LINE}; }}
    .cv-meta span {{ width:40mm; color:{MUTED}; font-size:9pt; }} .cv-meta b {{ font-weight:500; font-size:10pt; }}
    .cv-foot {{ position:absolute; left:0; right:0; bottom:0; background:{DEEP}; color:#c2d3c8; padding:9mm 20mm; font-size:8.5pt; }}
    .cv-foot b {{ display:block; color:#fff; font:600 10pt Montserrat; margin-bottom:2pt; }}
    """


def render_pdf(title, short, period, body, css=""):
    """Octets du PDF. Lève une exception si WeasyPrint est indisponible : l'appelant bascule alors
    sur le rendu historique (repli xhtml2pdf)."""
    import weasyprint
    doc = (f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>{esc(title)} — {esc(period)}</title>'
           f'<style>{base_css(short, period)}{css}</style></head><body>{body}</body></html>')
    doc = re.sub(r"(\d)[ \u202f]%", "\\1\u00a0%", doc).replace("\u202f", "\u00a0")
    return weasyprint.HTML(string=doc, base_url=ASSETS).write_pdf()


# ----------------------------------------------------------------------------- emails
class Mail:
    """Corps d'email : tableaux et styles en ligne pour Outlook, Arial, images data: (voir inline_images)."""
    W = 720

    @staticmethod
    def img(uri, alt):
        if not uri:
            return ""
        return (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {LINE};margin:0 0 12px;"><tr><td style="padding:10px;">'
                f'<div style="font:bold 13px Arial,sans-serif;color:{FOREST};margin:0 0 6px;">{esc(alt)}</div>'
                f'<img src="{uri}" width="{Mail.W - 42}" alt="{esc(alt)}" style="display:block;width:100%;max-width:{Mail.W - 42}px;height:auto;border:0;"></td></tr></table>')

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
                    f'<div style="font:bold 10px Arial,sans-serif;color:{lc};text-transform:uppercase;">{esc(str(label))}</div>'
                    f'<div style="font:bold 21px Arial,sans-serif;color:{vc};padding:3px 0 2px;">{esc(str(val))}</div>'
                    f'<div style="font:11px Arial,sans-serif;color:{sc};">{esc(str(sub)) or "&nbsp;"}</div></td>')
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
                       f'<td style="font:13px/1.45 Arial,sans-serif;color:{INK};padding:6px 0;border-bottom:1px solid {LINE};">{esc(str(i))}</td></tr></table>' for i in items)

    @staticmethod
    def shell(title, period, greeting, intro, blocks, signature_parts, kind="Rapport mensuel", full_width=False):
        """Salutation et signature hors gabarit, tout à gauche ; le rapport est centré sur 720 px."""
        rule = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td width="90%" height="4" bgcolor="{FOREST}" style="background:{FOREST};font-size:0;line-height:0;">&nbsp;</td>'
                f'<td width="10%" height="4" bgcolor="{RED}" style="background:{RED};font-size:0;line-height:0;">&nbsp;</td></tr></table>')
        width = 'width="100%" style="width:100%;"' if full_width else f'width="{Mail.W}" style="width:{Mail.W}px;max-width:100%;"'
        signature = "<br>".join(esc(p) for p in signature_parts if p)
        html = f"""<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr><td align="left" style="padding:0 0 14px;font:14px/1.5 Arial,sans-serif;color:{INK};text-align:left;">
          <p style="margin:0 0 8px;">{esc(greeting)}</p><p style="margin:0;">{esc(intro)}</p></td></tr>
        <tr><td align="center">
        <table role="presentation" {width} cellpadding="0" cellspacing="0">
        <tr><td><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
          <td style="font:12px/1.4 Arial,sans-serif;color:{FOREST};padding:6px 0 10px;">{"<br>".join(AGENCY_LINES)}</td>
          <td align="right" style="font:bold 11px Arial,sans-serif;color:{MUTED};text-transform:uppercase;letter-spacing:1px;">{esc(kind)}<br><span style="font-size:14px;color:{FOREST};text-transform:none;letter-spacing:0;">{esc(period)}</span></td>
        </tr></table>{rule}</td></tr>
        <tr><td style="padding:16px 0 2px;font:bold 25px/1.2 Arial,sans-serif;color:{FOREST};text-align:left;">{esc(title)}</td></tr>
        <tr><td style="text-align:left;">{blocks}</td></tr>
        </table></td></tr>
        <tr><td align="left" style="padding:22px 0 0;font:14px/1.5 Arial,sans-serif;color:{INK};text-align:left;">Cordialement,{'<br>' + signature if signature else ''}</td></tr>
        </table>"""
        return re.sub(r"(\d)[ \u202f]%", "\\1&nbsp;%", html)


_INLINE_IMG = re.compile(r'src="data:image/(png|jpeg|gif);base64,([A-Za-z0-9+/=]+)"')


def inline_images(msg, html):
    """Ajoute le HTML à un EmailMessage en remplaçant chaque image data: par une pièce CID (Outlook
    n'affiche pas les data:). Une image répétée n'est jointe qu'une fois."""
    images = {}

    def to_cid(m):
        data = base64.b64decode(m.group(2))
        cid = "img-" + hashlib.sha1(data).hexdigest()[:16]
        images[cid] = (data, m.group(1))
        return f'src="cid:{cid}"'
    body = _INLINE_IMG.sub(to_cid, html)
    msg.add_alternative(body, subtype="html")
    part = msg.get_payload()[-1]
    for cid, (data, sub) in images.items():
        part.add_related(data, maintype="image", subtype=sub, cid=f"<{cid}>", filename=f"{cid}.{sub}", disposition="inline")
    return msg


# ----------------------------------------------------------------------------- PPTX
def plain(v):
    return html_module.unescape(re.sub(r"<[^>]+>", "", str(v)))


def fit_size(value, w, h, size, floor=8):
    """Taille de police qui fait tenir un texte libre (analyses) dans sa boîte."""
    paras = value if isinstance(value, list) else [value]
    while size > floor:
        per_line = max(8, w * 72 / (size * .5))
        n = sum(max(1, math.ceil(len(str(p)) / per_line)) for p in paras)
        if n * size * 1.22 / 72 + (len(paras) - 1) * 3 / 72 <= h:
            break
        size -= .5
    return size


class Deck:
    def __init__(self, title, period, footer, clean=None):
        from pptx import Presentation
        from pptx.util import Inches
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(13.333), Inches(7.5)
        self.title, self.period, self.footer = title, period, footer
        self.clean = clean or (lambda s: s)
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
        sh.fill.fore_color.rgb = RGBColor.from_string(hx(color))
        if line:
            sh.line.color.rgb = RGBColor.from_string(hx(line))
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
            r.text = re.sub(r"(\d)[ \u202f]%", "\\1\u00a0%", self.clean(str("" if line is None else line)))
            r.font.size, r.font.bold, r.font.italic, r.font.name = Pt(size), bold, italic, FONT
            r.font.color.rgb = RGBColor.from_string(hx(color))
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
        if img:
            self.pic(s, img, x + .1, y + .36, w - .2, h - .44)
        else:
            self.text(s, x + .1, y + .36, w - .2, h - .44, "Aucune donnée", 11, False, MUTED, "c", "m")

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
                run.text = self.clean(plain(row[c])) if c < len(row) else ""
                run.font.size, run.font.name = Pt(font), FONT
                run.font.bold = r == 0 or bool(bolds and bolds(r - 1, c))
                run.font.color.rgb = RGBColor.from_string("FFFFFF" if r == 0 else hx(INK))
                cell.fill.solid()
                f = theme if r == 0 else ((fills(r - 1, c) if fills else None) or ("FFFFFF" if r % 2 else PALE))
                cell.fill.fore_color.rgb = RGBColor.from_string(hx(f))

    def note(self, s, x, y, w, h, label, value, size=11, theme=GREEN, bg=LIGHT, color=INK):
        self.rect(s, x, y, w, h, bg)
        self.rect(s, x, y, .05, h, theme)
        self.text(s, x + .14, y + .05, w - .22, .25, label.upper(), 8.5, True, theme)
        empty = value is None or value == "" or value == []
        self.text(s, x + .14, y + .29, w - .22, h - .32, "Analyse non renseignée." if empty else value,
                  fit_size(value or "", w - .3, h - .36, size), color=MUTED if empty and color == INK else color, italic=empty)

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
            self.text(s, cx + .08, y + .07, cw - .16, .25, str(lab).upper(), 8.5, True, lc)
            self.text(s, cx + .08, y + .31, cw - .16, .45, val, vs, True, vc)
            if sub:
                self.text(s, cx + .08, y + h - .3, cw - .16, .25, sub, 9.5, False, sc)

    def cover(self, eyebrow, title, title2, subtitle, meta, stripe=()):
        s = self.slide()
        self.pic(s, LOGO_PNG, .7, .35, 2.5, 1.2)
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
        self.text(s, .7, 6.72, 12, .3, AGENCY, 13, True, "FFFFFF")
        self.text(s, .7, 7.03, 12, .3, CONTACT, 10.5, False, "C2D3C8")

    def page(self, n, title, eyebrow):
        """Diapositive charte : en-tête logo + filet, carré d'index, eyebrow à tiret rouge, titre."""
        s = self.slide()
        self.pic(s, LOGO_PNG, .45, .12, 1, .66)
        self.rect(s, 1.52, .22, .01, .46, LINE)
        self.text(s, 1.6, .2, 4.5, .5, AGENCY_LINES, 9, False, FOREST, anchor="m")
        self.text(s, 8.3, .2, 4.6, .5, [self.title, self.period], 9.5, False, MUTED, "r", "m")
        self.filet(s, .85)
        self.rect(s, .45, 1.03, .5, .5, LIGHT)
        self.text(s, .45, 1.03, .5, .5, f"{n:02d}", 13, True, GREEN, "c", "m")
        self.rect(s, 1.07, 1.11, .22, .025, RED)
        self.text(s, 1.31, 1.0, 8, .24, eyebrow.upper(), 8.5, True, FOREST)
        self.text(s, 1.03, 1.17, 11.8, .42, title, 22, True, FOREST)
        return s

    def band(self, title, subtitle, theme):
        """Diapositive PCHC : bandeau du volet, logo sur pastille blanche, filet AMMPS."""
        s = self.slide()
        self.rect(s, 0, 0, 13.333, .95, theme)
        self.rect(s, 11.75, .09, 1.2, .77, "FFFFFF", rounded=True)
        self.pic(s, LOGO_PNG, 11.82, .12, 1.06, .71)
        self.filet(s, .95)
        self.text(s, .45, .1, 11, .5, title, 24, True, "FFFFFF")
        self.text(s, .45, .57, 11, .3, subtitle, 11.5, False, "FFFFFF")
        return s

    def finish(self):
        n = len(self.prs.slides)
        for i, s in enumerate(self.prs.slides, 1):
            if i == 1:
                continue
            self.rect(s, .45, 7.02, 12.43, .01, LINE)
            self.text(s, .45, 7.06, 10.5, .3, self.footer, 9, False, MUTED)
            self.text(s, 11.3, 7.06, 1.58, .3, f"{i} / {n}", 9.5, True, FOREST, "r")
        buf = io.BytesIO()
        self.prs.save(buf)
        return buf.getvalue()
