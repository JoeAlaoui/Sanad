# -*- coding: utf-8 -*-
"""
pchc_deliverables.py — Livrables du Reporting Métier PCHC : PPTX (16:9), PDF (A4 paysage, serveur),
corps de l'e-mail et message .eml avec les 3 livrables en pièces jointes. Modèle hybride : cadre de la
charte AMMPS (ammps_charte) + paysage, tableau à gauche, graphique et constats à droite, couleur par volet.

Toutes les valeurs viennent de pchc_raw.pchc_compute() (aucun calcul ici) : l'Excel, le PPTX, le PDF,
l'e-mail et l'UI affichent donc exactement les mêmes chiffres.
"""
import base64
import io
import os
import re
from datetime import datetime
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from html import escape as esc

import ammps_charte as C
import pchc_raw as R
from pchc_excel import CAT_FILL, CAT_PT, TYPE_PAL, pt_colors

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF_MIME = "application/pdf"
NAVY = C.hx(C.FOREST)
RED = C.hx(C.RED)
INK, MUTED, LINE, SOFT = C.hx(C.INK), C.hx(C.MUTED), C.hx(C.LINE), C.hx(C.PALE)
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
    plt = C.mpl()
    plt.rcParams.update({"font.size": 9, "axes.edgecolor": "#B8C2CF", "axes.spines.top": False, "axes.spines.right": False})
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
        ax.text(i, bottom[i] + top * 0.015, f"{s['taux'] * 100:.1f} %".replace(".", ","), ha="center", va="bottom", fontsize=8.5, color=C.FOREST, fontweight="bold")
    ax.set_ylim(0, top * 1.12)
    ax.set_yticks([])
    ax.spines["left"].set_visible(False)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3, frameon=False, fontsize=8.5)
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


def keep_lines(comp):
    t = comp["totals"]
    return [f"{R.fmt_int(t['total'])} dossiers, {R.fmt_pct(t['taux'])} délivrés/validés.", f"{R.fmt_int(t['encours'])} dossiers en cours de traitement.",
            f"{R.fmt_int(t['bloques'])} bloqués ou refusés.", f"{sum(1 for a in comp['alerts'] if 'CRITIQUE' in a['status'])} alerte(s) critique(s) sur {len(comp['alerts'])}."]


def synth_rows(comp):
    rows = [[r["module"], R.fmt_int(r["total"]), R.fmt_int(r["delivered"]), R.fmt_pct(r["taux"]), R.fmt_int(r["encours"]), "–" if r["bloques"] is None else R.fmt_int(r["bloques"]),
             f"≥ {r['seuil_vert'] * 100:.0f} % / ≥ {r['seuil_orange'] * 100:.0f} %", "● " + voyant_info(r["voyant"])[1]] for r in comp["synthesis"]]
    t = comp["totals"]
    rows.append(["TOTAL", R.fmt_int(t["total"]), R.fmt_int(t["delivered"]), R.fmt_pct(t["taux"]), R.fmt_int(t["encours"]), R.fmt_int(t["bloques"]), "", ""])
    return rows


def alert_rows(comp):
    out = []
    for a in comp["alerts"]:
        pf = a["type"] == "taux"
        fmt = (lambda v: f"{v * 100:.0f} %") if pf else R.fmt_int
        out.append([a["indicator"], a["module"], R.fmt_pct(a["value"]) if pf else R.fmt_int(a["value"]), fmt(a["warn"]), fmt(a["crit"]), "● " + ALERT_LABEL[alert_key(a["status"])], a["priority"].capitalize()])
    return out


SYN_HEAD = ["Module", "Total", "Délivrés / Validés", "Taux", "En cours*", "Bloqués / Refusés", "Seuils vert / orange", "Voyant"]
ALERT_HEAD = ["Indicateur", "Module", "Valeur actuelle", "Seuil attention", "Seuil critique", "Statut", "Priorité"]
NOTE_SYN = "* En cours = statuts de traitement + statuts non listés. Bloqués/Refusés = Non conforme + Irrecevable + Rejeté."
NOTE_ALERT = "Seuils : « taux » → alerte si la valeur est inférieure au seuil ; « nombre » → supérieure. Seuils = paramètres du classeur."
CARD_BG = {"ok": "C6EFCE", "info": "E4EEF9", "bad": "FFCCCC"}
SUBTITLE = "Plateforme Tarkhiss — Produits Cosmétiques et d’Hygiène Corporelle"


def no_emoji(line):
    return line.replace("🟢", "vert").replace("🟡", "orange").replace("🔴", "rouge").replace("⚪ ", "")


def module_subtitle(m, mod):
    unit = "opérateurs" if m["key"] == "identification" else "dossiers"
    return f"Feuille « {m['sheet']} » — {R.fmt_int(mod['total'])} {unit} · voyant {voyant_info(mod['voyant'])[1]}"


# ----------------------------------------------------------------------------- PPTX
def build_pptx(comp, meta, ctx):
    """Présentation 16:9 : garde, synthèse, chiffres clés & alertes, une section par KPI, méthodologie."""
    gen = ctx["generated"]
    label = ctx["period_label"].capitalize()
    d = C.Deck("Reporting Métier PCHC", label, f"AMMPS · Reporting Métier PCHC — {label}", ctx.get("clean"))
    d.cover("Revue de direction", "Reporting Métier", "PCHC", SUBTITLE, [("Période", label), ("Situation au", fmt_dt(meta.get("imported_at")) if meta else "—")],
            stripe=[m["theme"] for m in R.MODULES])

    def cards(s, items, y, h, vs, theme):
        gap, w = .12, 12.43
        cw = (w - gap * (len(items) - 1)) / len(items)
        for i, (lab, val, bg) in enumerate(items):
            cx = .45 + i * (cw + gap)
            d.rect(s, cx, y, cw, h, bg, LINE, True)
            d.text(s, cx, y + .06, cw, .25, lab.upper(), 8.5, True, MUTED, "c")
            d.text(s, cx, y + .3, cw, h - .34, val, vs, True, theme, "c", "m")

    syn = comp["synthesis"]
    p = d.band("Synthèse exécutive", f"Situation par volet — {label}", NAVY)
    d.table(p, .45, 1.18, 12.43, [3.3, 1.2, 1.4, 1.2, 1.2, 1.4, 1.8, 1.3], SYN_HEAD, synth_rows(comp), font=11, row_h=.36, bolds=lambda r, c: r == len(syn),
            fills=lambda r, c: "D9E2F3" if r == len(syn) else (light_hex(voyant_info(syn[r]["voyant"])[0], .7) if c == 7 else None))
    d.text(p, .45, 3.72, 12.4, .28, NOTE_SYN, 8.5, False, MUTED, italic=True)
    d.pic(p, chart_modules(comp, (8.6, 2.9)), .45, 4.02, 8.2, 2.9)
    d.bullets(p, 8.95, 4.1, 3.93, 2.3, "À retenir", keep_lines(comp), 11.5, NAVY)

    p = d.band("Chiffres clés & points d’alerte", "Résumé décisionnel — actions requises", NAVY)
    cards(p, [(c["label"].title(), R.fmt_int(c["value"]), CARD_BG[c["kind"]]) for c in comp["cards"]], 1.18, 1.15, 24, NAVY)
    al = comp["alerts"]
    d.table(p, .45, 2.5, 12.43, [3.6, 1.4, 1.5, 1.5, 1.5, 1.6, 1.3], ALERT_HEAD, alert_rows(comp), RED, 10.5, .36, fills=lambda r, c: ALERT_FILL[alert_key(al[r]["status"])] if c >= 5 else None)
    d.text(p, .45, 6.2, 12.4, .3, NOTE_ALERT, 9, False, MUTED, italic=True)

    for m in R.MODULES:
        mod = comp["modules"][m["key"]]
        th = m["theme"]
        p = d.band(f"KPI {m['label']}", module_subtitle(m, mod), th)
        cards(p, key_indicators(mod), 1.18, .95, 18, th)
        rows = status_table_rows(mod)
        n = len(rows)
        last = lambda r, c, n=n: r == n - 1
        fills = lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None
        head = ["Statut", "Nombre", "%", "Catégorie"]
        if m["key"] == "identification":
            ty = [[x["label"], R.fmt_int(x["count"]), R.fmt_pct(x["pct"])] for x in mod["types"]]
            ty.append(["TOTAL", R.fmt_int(mod["types_total"]), R.fmt_pct(1 if mod["types_total"] else 0)])
            d.table(p, .45, 2.35, 6.1, [3.2, 1, 1, 1.4], head, [r[:4] for r in rows], th, 10.5, .32, fills=fills, bolds=last)
            d.table(p, .45, 3.85, 6.1, [3.8, 1, 1], ["Type d'opérateur", "Nombre", "%"], ty, th, 10.5, .32, fills=lambda r, c, k=len(ty): "D9E2F3" if r == k - 1 else None, bolds=lambda r, c, k=len(ty): r == k - 1)
        else:
            d.table(p, .45, 2.35, 6.1, [3, 1, 1, 1.3], head, [r[:4] for r in rows], th, 10 if n > 12 else 10.5, min(.3, 4.5 / (n + 1)), fills=fills, bolds=last)
        d.pic(p, module_chart(mod), 6.8, 2.3, 6.08, 3.15)
        d.bullets(p, 6.8, 5.5, 6.08, 1.42, "Constats", comp["insights"][m["key"]], 10.5, th, light_hex(th, .9))

    md = methodology(comp, meta, gen)
    p = d.band("Méthodologie & qualité des données", "Sources, contrôles, coquilles normalisées, hypothèses des seuils", NAVY)
    d.bullets(p, .45, 1.18, 6.15, 3.15, "Sources", md["sources"], 9, NAVY)
    d.bullets(p, 6.73, 1.18, 6.15, 2.25, "Contrôles", md["controls"][:8], 9, NAVY)
    d.bullets(p, .45, 4.43, 6.15, 2.48, "Coquilles normalisées & statuts non référencés", md["typos"] + md["unknown"], 9, NAVY)
    d.bullets(p, 6.73, 3.53, 6.15, 3.38, "Hypothèses des seuils", [no_emoji(h) for h in md["hypotheses"]], 8.5, NAVY)
    return d.finish()


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
                  rows, RED, [30, 12, 13, 12, 12, 11, 10], font=font,
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
.nt {{ font-size:7.2pt; color:{C.MUTED}; font-style:italic; margin-top:3pt; }}
table {{ break-inside:avoid; }}
.rh td, .rh th {{ height:var(--rh); vertical-align:middle; }}
.cw {{ margin:0 -4px; }}
"""


def charte_pdf(comp, meta, ctx):
    gen = ctx["generated"]
    label = ctx["period_label"].capitalize()
    uri = C.data_uri
    cards = lambda items, th, size: '<div class="cw">' + _cards_html(items, th, size).replace("margin:0 -4px 10px", "margin:0 0 10px") + "</div>"
    band = lambda n, t, sub, th: f'<div class="band" style="background:#{th}"><span class="bi">{n:02d}</span><div><div class="bt">{esc(t)}</div><div class="bs">{esc(sub)}</div></div></div>'
    stripe = '<div class="stripe">' + "".join(f'<span style="background:#{m["theme"]}"></span>' for m in R.MODULES) + "</div>"
    body = C.cover("Reporting Métier", "PCHC", SUBTITLE, [("Période", label), ("Situation au", fmt_dt(meta.get("imported_at")) if meta else "—"), ("Édition", gen.strftime("%d/%m/%Y"))],
                   "Revue de direction", stripe) + C.head_html("Reporting Métier PCHC", label)
    body += (f'<div class="pg">{band(1, "Synthèse exécutive", f"Situation par volet — {label}", NAVY)}<div class="rh" style="--rh:25px">{synthesis_table_html(comp, 12)}</div><div class="nt">{esc(NOTE_SYN)}</div>'
             f'<div class="syn"><div><img src="{uri(chart_modules(comp, (8.6, 3.2)))}"></div><div>{_box("À retenir", keep_lines(comp), NAVY, 12.5)}</div></div></div>')
    alerts = re.sub(r">(HAUTE|MOYENNE|BASSE)<", lambda x: f">{x.group(1).capitalize()}<", alerts_table_html(comp, 13))
    body += (f'<div class="pg">{band(2, "Chiffres clés & points d’alerte", "Résumé décisionnel — actions requises", NAVY)}'
             + cards([(c["label"], R.fmt_int(c["value"]), CARD_BG[c["kind"]]) for c in comp["cards"]], NAVY, 22)
             + f'<div class="rh" style="--rh:32px">{alerts}</div><div class="nt">{esc(NOTE_ALERT)}</div></div>')
    for i, m in enumerate(R.MODULES, 3):
        mod = comp["modules"][m["key"]]
        th = m["theme"]
        rows = status_table_rows(mod)
        tr = [[esc(r[0]), r[1], r[2], esc(r[3])] for r in rows]
        tbl = _table(["Statut", "Nombre", "%", "Catégorie"], tr, th, [46, 16, 16, 22], font=12, fills=lambda r, c, rows=rows: rows[r][4] if c in (0, 3) else None, bold_last=True)
        ident = m["key"] == "identification"
        if ident:
            ty = [[esc(x["label"]), R.fmt_int(x["count"]), R.fmt_pct(x["pct"])] for x in mod["types"]]
            ty.append(["TOTAL", R.fmt_int(mod["types_total"]), R.fmt_pct(1 if mod["types_total"] else 0)])
            tbl += '<div style="height:8px"></div>' + _table(["Type d'opérateur", "Nombre", "%"], ty, th, [62, 19, 19], font=12, bold_last=True, fills=lambda r, c, n=len(ty): "D9E2F3" if r == n - 1 else None)
        nb = len(rows) + 1 + (len(mod["types"]) + 2 if ident else 0)
        chart = chart_types(mod, (6.4, 4.1)) if ident else chart_status(mod, (6.4, 4.5))
        body += (f'<div class="pg">{band(i, "KPI " + m["label"], module_subtitle(m, mod), th)}' + cards(key_indicators(mod), th, 17)
                 + f'<div class="split"><div class="rh" style="--rh:{min(27, 440 // nb - 9)}px">{tbl}</div><div><img src="{uri(chart)}">'
                 + _box("Constats", comp["insights"][m["key"]], th, 11.5) + "</div></div></div>")
    md = methodology(comp, meta, gen)
    n = len(R.MODULES) + 3
    body += (f'<div class="pg">{band(n, "Méthodologie & qualité des données", "Sources, contrôles, coquilles normalisées, hypothèses des seuils", NAVY)}<div class="split" style="grid-template-columns:1fr 1fr">'
             f'<div>{_box("Sources", md["sources"], NAVY, 10.5)}{_box("Coquilles normalisées & statuts non référencés", md["typos"] + md["unknown"], NAVY, 10.5)}</div>'
             f'<div>{_box("Contrôles", md["controls"], NAVY, 10.5)}{_box("Hypothèses des seuils", [no_emoji(h) for h in md["hypotheses"]], NAVY, 10.5)}</div></div></div>')
    return C.render_pdf("Reporting Métier PCHC", "Reporting Métier PCHC", label, ctx.get("clean", str)(body), MIX_CSS)


def build_pdf(comp, meta, ctx):
    """PDF A4 paysage côté serveur : modèle charte (WeasyPrint), repli xhtml2pdf sur le rapport simple."""
    try:
        return charte_pdf(comp, meta, ctx), "weasyprint"
    except Exception:
        pass
    pages = report_html(comp, meta, ctx)
    gen = ctx["generated"]
    head = f"{ctx.get('agency', 'AMMPS — DSID')} · Reporting Métier PCHC"
    foot = f"Généré le {gen.strftime('%d/%m/%Y à %H:%M')} · {ctx['period_label']}"
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
def mail_html(comp, meta, ctx, greeting, signature, full_width=False):
    """Corps de l'e-mail : synthèse exécutive et points d'alerte, salutation et signature hors gabarit."""
    label = ctx["period_label"].capitalize()
    t = comp["totals"]
    m = C.Mail
    tone = {"ok": "good", "bad": "alert"}
    cards = [(c["label"].capitalize(), R.fmt_int(c["value"]), "", tone.get(c["kind"], "")) for c in comp["cards"]]
    crit = sum(1 for a in comp["alerts"] if "CRITIQUE" in a["status"])
    dot = lambda col, txt: f'<span style="color:#{col};">&#9632;</span> {esc(txt)}'
    rows = synth_rows(comp)
    syn = []
    for r, row in zip(comp["synthesis"], rows):
        col, name = voyant_info(r["voyant"])
        syn.append([f"<b>{esc(row[0])}</b>", row[1], row[2], f"<b>{row[3]}</b>", row[4], row[5], dot(col, name)])
    syn.append({"cls": "tot", "cells": [f"<b>{x}</b>" for x in rows[-1][:6]] + [""]})
    colors = {"Critique": RED, "Attention": C.hx(C.AMBER), "OK": "2E8B57"}
    al = [[esc(r[0]), esc(r[1]), f"<b>{r[2]}</b>", r[3], r[4], dot(colors.get(r[5][2:], MUTED), r[5][2:])] for r in alert_rows(comp)]
    blocks = (m.sec(1, "Synthèse exécutive")
              + m.note(f"Situation des dossiers au {fmt_dt(meta.get('imported_at')) if meta else '—'} : {R.fmt_int(t['total'])} dossiers, {R.fmt_pct(t['taux'])} délivrés/validés, "
                       f"{R.fmt_int(t['encours'])} en cours, {R.fmt_int(t['bloques'])} bloqués ou refusés — {crit} alerte(s) critique(s).")
              + m.kpis(cards[:3], 3) + m.kpis(cards[3:], 3)
              + m.table(["Module", "Total", "Délivrés", "Taux", "En cours", "Bloqués", "Voyant"], syn, ["left", "center", "center", "center", "center", "center", "left"])
              + m.img(C.data_uri(chart_modules(comp, (8.2, 2.6))), "Dossiers par volet")
              + m.sec(2, "Points d’alerte")
              + m.table(["Indicateur", "Module", "Valeur", "Attention", "Critique", "Statut"], al, ["left", "left", "center", "center", "center", "left"])
              + m.note("Détail par volet, graphiques et méthodologie : voir les pièces jointes."))
    intro = f"Veuillez trouver ci-dessous le reporting métier PCHC ({label})."
    return m.shell("Reporting Métier PCHC", label, greeting, intro, blocks, signature, "Revue de direction", full_width)


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
    C.inline_images(msg, body_html)
    for data, mime, name in attachments:
        _attach(msg, data, mime, name)
    return bytes(msg)
