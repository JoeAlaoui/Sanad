"""Prototype charte AMMPS — rapport mensuel Tarkhiss (octobre 2026), 3 variantes.
Données : recopiées du PDF de référence (reference/rapport_tarkhiss_2026-10.pdf), aucun accès à data/.
Lancer depuis la racine du projet : .venv/bin/python mockups/charte_ammps/generate.py
"""
import base64
import html
import io
import os
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
CHARTE = os.path.join(ROOT, "reference", "charteGraphiqueAMMPS", "assets")
FONTS = os.path.join(HERE, "fonts")

FOREST, GREEN, RED, INK, MUTED, LINE, PALE, LIGHT, DEEP = (
    "#0d4823", "#058b44", "#df1e26", "#183928", "#5b6e62", "#dce6df", "#f3f7f4", "#eaf4ed", "#0b3420")
TINT = "#8cc9a3"
esc = html.escape

# ----------------------------------------------------------------------------- Données (PDF de référence)
R = dict(
    title="Rapport Support Tarkhiss", short="Tarkhiss", subtitle="Support Email & Hotline", period="Octobre 2026",
    agency="Agence Marocaine du Médicament et des Produits de Santé",
    prev="septembre",
    kpi_volume=[("Emails reçus", "207", "−27 % vs septembre (283)"), ("Emails envoyés", "160", "−37 % vs septembre (255)"),
                ("Appels", "179", "−21 % vs septembre (226)"), ("Temps de communication", "07:41:46", "hh:mm:ss")],
    kpi_quali=[("Problèmes techniques", "166", "+48 % vs septembre (112)", "alert"), ("Demandes d’information", "91", "−47 % vs septembre (171)", "good"),
               ("Ratio bugs / demandes", "1,82", "", ""), ("Semaine critique", "19–25 oct.", "80 cas qualifiés", "warn")],
    kpi_charge=[("Durée moyenne par appel", "00:02:34", ""), ("Canal dominant", "54 % emails", "46 % appels")],
    problems=[("Accès, connexion et OTP/MFA", 46, "27,7 %", "Connexion impossible, codes OTP non reçus et réinitialisation Karaz/MFA."),
              ("Dépôt, soumission et pièces jointes", 40, "24,1 %", "Blocage du dépôt, ajout de documents ou absence du menu « Nouveau dossier »."),
              ("Autres problèmes techniques", 26, "15,7 %", "Dysfonctionnements spécifiques ne relevant pas clairement des catégories principales."),
              ("Paiement et références Fatourati", 19, "11,4 %", "Paiements refusés, références invalides ou à régénérer."),
              ("Inscription et identification opérateur", 19, "11,4 %", "Difficultés de création de compte, d’inscription ou d’identification d’établissement."),
              ("Visibilité, statut et téléchargement", 11, "6,6 %", "Dossiers invisibles, statuts non actualisés et documents impossibles à télécharger."),
              ("Erreurs applicatives générales", 5, "3,0 %", "Erreurs JavaScript, blocages généraux ou indisponibilités de la plateforme.")],
    demandes=[("Suivi et traitement des dossiers", 28, "30,8 %", "État d’avancement, délais, validation et relances de dossiers."),
              ("Procédures de dépôt et exigences réglementaires", 17, "18,7 %", "Canal de dépôt, autorisations et pièces ou étapes réglementaires requises."),
              ("Gestion administrative des comptes", 14, "15,4 %", "Transfert, annulation ou changement d’identifiant, d’email ou de représentant."),
              ("Assistance et clarification générales", 12, "13,2 %", "Explications générales sur le fonctionnement ou le traitement d’une demande."),
              ("Rectification et mise à jour de dossiers", 12, "13,2 %", "Correction d’informations, d’attestations ou de données déjà déposées."),
              ("Paiement, facturation et références", 6, "6,6 %", "Informations sur factures, références Fatourati et justificatifs."),
              ("Autres demandes", 2, "2,2 %", "Demandes ponctuelles insuffisamment détaillées.")],
    weekly=[("1–4 octobre", 11, 8, "Début de mois modéré, avec 19 cas qualifiés et de nombreux messages administratifs."),
            ("5–11 octobre", 23, 19, "Hausse équilibrée : les demandes représentent 45 % des cas qualifiés."),
            ("12–18 octobre", 48, 20, "Pic de trafic total, porté par une forte progression des incidents techniques."),
            ("19–25 octobre", 55, 25, "Semaine la plus lourde opérationnellement, avec 80 bugs ou demandes."),
            ("26–31 octobre", 29, 19, "Malgré une semaine incomplète, 48 cas qualifiés indiquent une forte intensité quotidienne.")],
    crit=3,
    heat=[[0, 0, 0, 16, 8, 0, 0], [7, 9, 11, 11, 4, 0, 0], [6, 10, 7, 8, 7, 0, 0], [7, 7, 9, 9, 13, 0, 0], [9, 14, 7, 0, 0, 0, 0]],
    calls_kpi=[("Taux de décroché", "74,3 %"), ("Appels manqués", "62"), ("Durée médiane", "00:02:07"), ("Durée maximale", "00:17:54")],
    callback="Rappel des appels manqués (62) : 3 rappelés par nous, 39 le client a rerappelé, 20 sans suite sous 24 h.",
    callers=[("•••• 475", 12, 1, "00:24:42"), ("•••• 202", 11, 7, "00:23:46"), ("•••• (masqué)", 9, 8, "00:26:52"), ("•••• (masqué)", 6, 1, "00:11:14"),
             ("•••• 951", 5, 2, "00:05:15"), ("•••• 428", 5, 2, "00:07:21"), ("•••• (masqué)", 5, 1, "00:33:46"), ("•••• (masqué)", 5, 0, "00:11:26")],
    emails_kpi=[("Délai médian (h. ouvrées)", "11 h 12"), ("P90", "19 h 16"), ("Sans réponse sous 7 j", "53"), ("Non lus", "19,3 %")],
    emails_note="Indicatif : appariement par sujet (pas d’ID de conversation dans l’export). Base : 8h-17h, lun-ven. Correspondance par sujet : 72,6 % des échanges appariés.",
    subjects=[("AKSAL BEAUTY / REGENERATION REFERENCE FATOURATI", 6), ("Demande de validation du compte E-services – MEDIPRO SOLUTION (medipro", 6),
              ("dysfonctionnement technique plateforme Tarkhiss", 6), ("Demande de transfert d’accès OTP - Plateforme Tarkhiss", 5),
              ("Problème de validité des références de paiement", 5), ("ECOLAB MAROC - URGENT SVP- Blocage technique sur l’application Tarkhis", 4)],
    words="plateforme (36) · dossier (35) · problème (29) · tarkhis (21) · technique (21) · réclamation (21) · compte (20) · déclaration (18) · sarl (16) · accès (16)",
    recipients=[("•••• (masqué)", 12), ("•••• (masqué)", 7), ("•••• (masqué)", 6), ("•••• (masqué)", 6), ("•••• (masqué)", 5), ("•••• (masqué)", 5)],
    constats=["Les problèmes techniques sont 1,82 fois plus nombreux que les demandes d’information : 166 contre 91.",
              "L’accès/OTP et le dépôt de documents concentrent 86 incidents, soit 51,8 % de tous les bugs.",
              "Les demandes de suivi de dossiers dominent les demandes d’information avec 28 cas (30,8 %).",
              "Les deux semaines du 12 au 25 octobre totalisent 211 emails, soit 54,1 % du trafic mensuel.",
              "La semaine du 19–25 octobre concentre 55 bugs, soit 33,1 % des incidents du mois.",
              "La période du 26–31 octobre reste très tendue : 48 cas qualifiés sur 54 emails, malgré seulement six jours."],
    recos=["Mettre en place un parcours automatisé de diagnostic et réinitialisation OTP/MFA.",
           "Créer un suivi en libre-service des dossiers avec statut, prochaine étape et délai indicatif.",
           "Publier une base de connaissances courte sur Fatourati, l’inscription et les procédures de dépôt."],
)


# ----------------------------------------------------------------------------- Polices statiques
def build_fonts():
    """Instancie les polices variables de la charte : le PDF embarque alors de vraies graisses."""
    from fontTools.ttLib import TTFont
    from fontTools.varLib import instancer
    os.makedirs(FONTS, exist_ok=True)
    for name, src, weights in (("Montserrat", "montserrat.ttf", (600, 700, 800)), ("Inter", "inter.ttf", (400, 500, 700))):
        for w in weights:
            out = os.path.join(FONTS, f"{name}-{w}.ttf")
            if os.path.exists(out):
                continue
            font = TTFont(os.path.join(CHARTE, src))
            axes = {a.axisTag: (w if a.axisTag == "wght" else a.defaultValue) for a in font["fvar"].axes}
            inst = instancer.instantiateVariableFont(font, axes)
            inst["OS/2"].usWeightClass = w
            for rec in inst["name"].names:    # « Inter Variable » → « Inter » : matplotlib retrouve la police par ce nom
                if rec.nameID in (1, 16):
                    rec.string = name
            inst.save(out)


def font_css():
    faces = []
    for fam, ws in (("Montserrat", (600, 700, 800)), ("Inter", (400, 500, 700))):
        faces += [f"@font-face{{font-family:{fam};src:url('file://{FONTS}/{fam}-{w}.ttf');font-weight:{w};}}" for w in ws]
    return "".join(faces)


# ----------------------------------------------------------------------------- Graphiques (même charte que le PDF)
def mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    for w in (400, 500, 700):
        font_manager.fontManager.addfont(os.path.join(FONTS, f"Inter-{w}.ttf"))
    plt.rcParams.update({"font.family": "Inter", "text.color": INK, "axes.edgecolor": LINE, "axes.labelcolor": MUTED,
                         "xtick.color": INK, "ytick.color": MUTED, "font.size": 10})
    return plt


def plt_close(fig):
    import matplotlib.pyplot as plt
    plt.close(fig)


def png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=240, bbox_inches="tight", facecolor="white")
    plt_close(fig)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def chart_weekly(figsize=(8.2, 3.4), short=False):
    import numpy as np
    plt = mpl()
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(len(R["weekly"]))
    bugs = [w[1] for w in R["weekly"]]
    dem = [w[2] for w in R["weekly"]]
    b1 = ax.bar(x - .19, bugs, .36, color=FOREST, label="Bugs")
    b2 = ax.bar(x + .19, dem, .36, color=TINT, label="Demandes")
    b1[R["crit"]].set_color(RED)
    for bars in (b1, b2):
        for b in bars:
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 1, int(b.get_height()), ha="center", fontsize=9.5, fontweight="bold")
    ax.set_xticks(x, [w[0].replace(" octobre", "") if short else w[0] for w in R["weekly"]])
    ax.set_yticks([])
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.tick_params(axis="x", length=0)
    ax.legend(frameon=False, loc="upper left", ncol=2, fontsize=9)
    ax.set_ylim(0, max(bugs) * 1.18)
    return png(fig)


def chart_rank(figsize=(8.2, 3.3), compact=False):
    plt = mpl()
    fig, ax = plt.subplots(figsize=figsize)
    items = R["problems"][::-1]
    colors = [FOREST if i >= len(items) - 2 else TINT for i in range(len(items))]
    bars = ax.barh([i[0] for i in items], [i[1] for i in items], color=colors, height=.66)
    for b, it in zip(bars, items):
        ax.text(b.get_width() + .7, b.get_y() + b.get_height() / 2, f"{it[1]}  ·  {it[2]}", va="center", fontsize=9.5, fontweight="bold")
    ax.set_xticks([])
    ax.set_xlim(0, 74 if compact else 56)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.tick_params(axis="y", length=0, labelcolor=INK, labelsize=8 if compact else 10)
    return png(fig)


def chart_heat(figsize=(8.2, 2.6)):
    import numpy as np
    from matplotlib.colors import LinearSegmentedColormap
    plt = mpl()
    cmap = LinearSegmentedColormap.from_list("ammps", [PALE, TINT, GREEN, FOREST])
    data = np.array(R["heat"], float)
    fig, ax = plt.subplots(figsize=figsize)
    ax.imshow(data, cmap=cmap, aspect="auto", vmin=0, vmax=data.max())
    for y, row in enumerate(R["heat"]):
        for x, v in enumerate(row):
            ax.text(x, y, v if v else "–", ha="center", va="center", fontsize=10, fontweight="bold" if v else "normal",
                    color="white" if v >= 9 else (INK if v else "#a9b8ae"))
    ax.set_xticks(range(7), ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"])
    ax.set_yticks(range(5), [f"Sem. {i}" for i in range(1, 6)])
    ax.xaxis.tick_top()
    ax.tick_params(length=0)
    ax.set_xticks(np.arange(-.5, 7), minor=True)
    ax.set_yticks(np.arange(-.5, 5), minor=True)
    ax.grid(which="minor", color="white", linewidth=2.5)
    ax.tick_params(which="minor", length=0)
    ax.spines[:].set_visible(False)
    return png(fig)


# ----------------------------------------------------------------------------- Composants
def b64(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


LOGO = b64(os.path.join(CHARTE, "logo-ammps.png"))


def kpi(label, value, sub="", tone=""):
    return (f'<div class="kpi {tone}"><div class="kl">{esc(label)}</div><div class="kv">{esc(value)}</div>'
            f'<div class="ks">{esc(sub) or "&nbsp;"}</div></div>')


def kpi_grid(items, cols=4):
    return f'<div class="kgrid c{cols}">' + "".join(kpi(*i) for i in items) + "</div>"


def table(heads, rows, aligns, cls=""):
    th = "".join(f'<th class="{a}">{esc(h)}</th>' for h, a in zip(heads, aligns))
    body = ""
    for row in rows:
        rcls = row.get("cls", "hl") if isinstance(row, dict) else ""
        cells = row["cells"] if isinstance(row, dict) else row
        body += f'<tr class="{rcls}">' +"".join(f'<td class="{a}">{c}</td>' for c, a in zip(cells, aligns)) + "</tr>"
    return f'<table class="data {cls}"><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table>'


def sec(n, title, eyebrow="", anchor=""):
    idx = f'<span class="idx">{n:02d}</span>' if n else ""
    eb = f'<div class="eyebrow">{esc(eyebrow)}</div>' if eyebrow else ""
    return f'<div class="sec" id="{anchor}">{idx}<div>{eb}<h2>{esc(title)}</h2></div></div>'


def fig(src, title, caption):
    return (f'<figure><div class="ft">{esc(title)}</div><img src="{src}"><figcaption>{esc(caption)}</figcaption></figure>')


def cat_table(items):
    rows = [[f"<b>{esc(l)}</b>", f"<b>{n}</b>", p, f'<span class="desc">{esc(d)}</span>'] for l, n, p, d in items]
    return table(["Type", "Nb", "%", "Description"], rows, ["l w30", "c w6", "c w9", "l"])


def weekly_table():
    rows = []
    for i, (w, b, d, o) in enumerate(R["weekly"]):
        cells = [f"<b>{w}</b>", f"<b>{b}</b>", f"<b>{d}</b>", f'<span class="desc">{esc(o)}</span>']
        rows.append({"cells": cells} if i == R["crit"] else cells)
    return table(["Semaine", "Bugs", "Demandes", "Observations"], rows, ["l w18", "c w8", "c w11", "l"])


def advanced():
    callers = table(["Contact", "Appels", "Manqués", "Durée cumulée"], [[esc(a), b, c, d] for a, b, c, d in R["callers"]], ["l", "c", "c", "c"], "tight")
    subjects = table(["Sujet", "Occurrences"], [[esc(s), f"<b>{n}</b>"] for s, n in R["subjects"]], ["l", "c w14"], "tight")
    recipients = table(["Destinataire", "Emails envoyés"], [[esc(s), n] for s, n in R["recipients"]], ["l", "c"], "tight")
    return f"""
    <h3>Journal d’appels</h3>
    <div class="mini">{kpi_grid(R['calls_kpi'])}</div>
    <p class="note">{esc(R['callback'])}</p>
    <h4>Top appelants (anonymisé)</h4>{callers}
    <h3>Emails</h3>
    <div class="mini">{kpi_grid(R['emails_kpi'])}</div>
    <p class="note">{esc(R['emails_note'])}</p>
    <h4>Sujets les plus fréquents</h4>{subjects}
    <div class="two"><div><h4>Mots-clés fréquents</h4><p class="words">{esc(R['words'])}</p></div>
    <div><h4>Top destinataires (anonymisé)</h4>{recipients}</div></div>"""


def synthesis():
    cs = "".join(f"<li>{esc(c)}</li>" for c in R["constats"])
    rs = "".join(f'<li><span class="n">{i:02d}</span><span>{esc(r)}</span></li>' for i, r in enumerate(R["recos"], 1))
    return f'<h3>Constats clés</h3><ul class="pts">{cs}</ul><h3>Recommandations</h3><ol class="recos">{rs}</ol>'


TOC = [("s1", "Indicateurs clés"), ("s2", "Analyse graphique"), ("s3", "Problèmes techniques"), ("s4", "Demandes d’information"),
       ("s5", "Évolution hebdomadaire"), ("s6", "Heatmap de charge"), ("s7", "Statistiques avancées — appels & emails"), ("s8", "Constats & recommandations")]
HEAT_NOTE = "Volume d’appels par jour de semaine et semaine du mois (granularité journalière : Tarkhiss ne journalise pas l’heure des appels)."
ADV_NOTE = "Imports bruts : journal d’appels et export Outlook"
W_CAP = "Comparaison du nombre de bugs signalés et de demandes d’information, semaine par semaine. Semaine critique en rouge."
P_CAP = "Répartition des problèmes techniques par type, en proportion du total du mois."


# ----------------------------------------------------------------------------- CSS commun (tokens + composants)
def base_css():
    return font_css() + f"""
    @page {{ size:A4; margin:27mm 17mm 20mm;
      @top-left {{ content:element(head); width:100%; }}
      @bottom-left {{ content:"AMMPS  ·  {R['short']} — {R['period']}"; font:8pt Inter; color:{MUTED}; }}
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
    .c4 {{ grid-template-columns:repeat(4,1fr); }} .c2 {{ grid-template-columns:repeat(2,1fr); }}
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
    table.data.tight td {{ padding:2.2pt 7pt; }}
    td.c {{ text-align:center; }} .w30 {{ width:32%; }} .w6 {{ width:6%; }} .w9 {{ width:9%; }} .w18 {{ width:18%; }} .w8 {{ width:8%; }} .w11 {{ width:11%; }} .w14 {{ width:14%; }}
    .desc {{ color:{MUTED}; font-size:8pt; line-height:1.3; }}
    figure {{ margin:0 0 12pt; break-inside:avoid; border:.6pt solid {LINE}; border-radius:6pt; padding:8pt 10pt 7pt; }}
    figure img {{ width:100%; display:block; }}
    .ft {{ font:700 9pt Montserrat; color:{FOREST}; margin-bottom:5pt; }}
    figcaption {{ font-size:7.8pt; color:{MUTED}; margin-top:4pt; }}
    .note {{ font-size:8pt; color:{MUTED}; margin:3pt 0 6pt; }}
    .words {{ font-size:8.5pt; }}
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
    """


def head_html(label=True):
    return (f'<div class="head-wrap"><div class="head"><img src="{LOGO}"><div class="ag">Agence Marocaine du Médicament<br>et des Produits de Santé</div>'
            f'<div class="rt">{esc(R["title"])}<br><b>{R["period"]}</b></div></div><div class="rule"></div></div>')


CONTACT = "Rue Lamfadal Charkaoui, B.P. 6206 · Madinat Al Irfane, Rabat · Maroc"


def body_sections():
    """Sections 2 à 8, identiques pour les variantes A et B."""
    return f"""
    <div class="newp">{sec(2, 'Analyse graphique', 'Tendances', 's2')}
      {fig(chart_weekly(), 'Évolution hebdomadaire — Bugs vs Demandes', W_CAP)}
      {fig(chart_rank(), 'Répartition des problèmes techniques', P_CAP)}</div>
    <div class="newp">{sec(3, 'Problèmes techniques signalés (166)', 'Qualification', 's3')}{cat_table(R['problems'])}</div>
    <div class="gap">{sec(4, 'Demandes d’information reçues (91)', 'Qualification', 's4')}{cat_table(R['demandes'])}</div>
    <div class="newp">{sec(5, 'Évolution hebdomadaire', 'Rythme du mois', 's5')}{weekly_table()}</div>
    <div class="gap">{sec(6, 'Heatmap de charge', 'Rythme du mois', 's6')}{fig(chart_heat(), 'Appels par jour et semaine du mois', HEAT_NOTE)}</div>
    <div class="newp">{sec(7, 'Statistiques avancées — appels & emails', ADV_NOTE, 's7')}{advanced()}</div>
    <div class="gap">{sec(8, 'Constats & recommandations', 'Synthèse', 's8')}{synthesis()}</div>"""


def toc_page():
    rows = ""
    for i, (a, l) in enumerate(TOC, 1):
        rows += f'<li><a href="#{a}"><span class="tn">{i:02d}</span><span class="tl">{esc(l)}</span><span class="tp" data-t="#{a}"></span></a></li>'
    return f'<ol class="toc">{rows}</ol>'


def kpi_section():
    return (f'{sec(1, "Indicateurs clés", "Vue d’ensemble", "s1")}'
            f'<h4>Volumes</h4>{kpi_grid(R["kpi_volume"])}<h4>Qualification</h4>{kpi_grid(R["kpi_quali"])}'
            f'<h4>Charge</h4>{kpi_grid(R["kpi_charge"], 2)}')


# ----------------------------------------------------------------------------- Variante A : Institutionnelle
def cover_a(title, title2, subtitle, meta, eyebrow="Rapport mensuel"):
    rows = "".join(f"<div><span>{esc(k)}</span><b>{esc(v)}</b></div>" for k, v in meta)
    return f"""<section class="cover">
      <div class="cv-top"><img src="{LOGO}"></div><div class="rule cv-rule"></div>
      <div class="cv-main"><div class="eyebrow">{esc(eyebrow)}</div>
        <h1>{esc(title)}<br><em>{esc(title2)}</em></h1>
        <div class="cv-sub">{esc(subtitle)}</div>
        <div class="cv-meta">{rows}</div></div>
      <div class="cv-foot"><b>{esc(R['agency'])}</b>{esc(CONTACT)}</div></section>"""


def cover_a_css():
    return f"""
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


def variant_a():
    css = cover_a_css()
    cover = cover_a("Rapport Support", "Tarkhiss", R["subtitle"],
                    [("Période", R["period"]), ("Volet", "Tarkhiss"), ("Édition", date(2026, 10, 30).strftime("%d/%m/%Y"))])
    body = f"""{head_html()}<div>{sec(0, 'Sommaire', 'Rapport mensuel')}{toc_page()}</div><div class="gap">{kpi_section()}</div>{body_sections()}"""
    return css, cover + body


# ----------------------------------------------------------------------------- Variante B : Direction
def variant_b():
    css = f"""
    .cover {{ page:cover; height:296.5mm; position:relative; background:{FOREST}; color:#fff; }}
    .cv-plate {{ position:absolute; left:0; top:0; background:#fff; padding:12mm 12mm 10mm 20mm; border-bottom-right-radius:8mm; }}
    .cv-plate img {{ height:30mm; display:block; }}
    .cv-main {{ position:absolute; left:20mm; right:20mm; top:92mm; }}
    .cv-main .eyebrow {{ color:#9cd9b1; font-size:9pt; letter-spacing:2pt; margin-bottom:12pt; }}
    .cv-main h1 {{ color:#fff; font-size:40pt; line-height:1.1; font-weight:800; letter-spacing:-1.2pt; }}
    .cv-sub {{ color:#c6dccd; font-size:14pt; margin-top:10pt; }}
    .cv-stats {{ position:absolute; left:20mm; right:20mm; top:180mm; display:grid; grid-template-columns:repeat(3,1fr); gap:0; }}
    .cv-stats div {{ border-top:.6pt solid #ffffff55; padding-top:8pt; }}
    .cv-stats b {{ display:block; font:700 24pt Montserrat; }} .cv-stats span {{ font-size:9pt; color:#b1cfba; }}
    .cv-foot {{ position:absolute; left:20mm; right:20mm; bottom:16mm; font-size:8.5pt; color:#c2d3c8; }}
    .cv-foot b {{ display:block; color:#fff; font:600 10pt Montserrat; margin-bottom:2pt; }}
    .cv-bar {{ position:absolute; left:0; right:0; bottom:0; height:5pt; background:linear-gradient(90deg,{GREEN} 90%,{RED} 90%); }}
    .exec {{ background:{FOREST}; color:#fff; border-radius:7pt; padding:14pt 16pt; margin-bottom:12pt; break-inside:avoid; }}
    .exec .eyebrow {{ color:#9cd9b1; }} .exec .eyebrow::before {{ background:{RED}; }}
    .exec h2 {{ color:#fff; margin-bottom:6pt; }}
    .exec li {{ padding:4pt 0; border-top:.6pt solid #ffffff24; }}
    .exec ul {{ list-style:none; margin:0; padding:0; font-size:9.5pt; color:#e6f0e9; }}
    """
    cover = f"""<section class="cover"><div class="cv-plate"><img src="{LOGO}"></div>
      <div class="cv-main"><div class="eyebrow">Rapport mensuel · {R['period']}</div><h1>Rapport Support<br>Tarkhiss</h1><div class="cv-sub">{esc(R['subtitle'])}</div></div>
      <div class="cv-stats"><div><b>207</b><span>Emails reçus</span></div><div><b>179</b><span>Appels</span></div><div><b>166</b><span>Problèmes techniques</span></div></div>
      <div class="cv-foot"><b>{esc(R['agency'])}</b>{esc(CONTACT)}</div><div class="cv-bar"></div></section>"""
    takeaways = "".join(f"<li>{esc(c)}</li>" for c in R["constats"][:3])
    body = f"""{head_html()}<div class="exec"><div class="eyebrow">À retenir</div><h2>Synthèse d’{R['period'].lower()}</h2><ul>{takeaways}</ul></div>
      {kpi_section()}<div class="gap">{sec(0, 'Sommaire', '')}{toc_page()}</div>{body_sections()}"""
    return css, cover + body


# ----------------------------------------------------------------------------- Variante C : Bulletin compact
def variant_c():
    css = f"""
    @page {{ margin:18mm 15mm 16mm; @top-left {{ content:none; }} }}
    html {{ font-size:9pt; }}
    .lh {{ display:flex; align-items:center; gap:12pt; }}
    .lh img {{ height:21mm; }}
    .lh .ag {{ border-left:.6pt solid {LINE}; padding-left:12pt; font-size:9pt; line-height:1.35; color:{FOREST}; }}
    .lh .rt {{ margin-left:auto; text-align:right; }}
    .lh .rt .eyebrow {{ justify-content:flex-end; }}
    .lh .rt b {{ font:700 11pt Montserrat; color:{FOREST}; }}
    .ttl {{ margin:9pt 0 10pt; }}
    .ttl h1 {{ font-size:21pt; font-weight:800; letter-spacing:-.6pt; line-height:1.15; }}
    .ttl h1 em {{ font-style:normal; color:{GREEN}; }}
    .ttl p {{ color:{MUTED}; font-size:10pt; }}
    .sec {{ margin-top:10pt; margin-bottom:6pt; }}
    h2 {{ font-size:12.5pt; }} .idx {{ width:20pt; height:20pt; line-height:20pt; font-size:8pt; }}
    h4 {{ margin:6pt 0 3pt; }}
    .kpi {{ padding:5pt 7pt; }} .kv {{ font-size:13pt; }} .kl {{ min-height:0; }}
    figure {{ margin-bottom:7pt; padding:6pt 7pt; }}
    table.data td {{ padding:2.6pt 6pt; }} .desc {{ font-size:7.5pt; line-height:1.2; }}
    """
    head = f"""<div class="lh"><img src="{LOGO}"><div class="ag">Agence Marocaine du Médicament<br>et des Produits de Santé</div>
      <div class="rt"><div class="eyebrow">Rapport mensuel</div><b>{R['period']}</b></div></div><div class="rule" style="margin-top:7pt"></div>
      <div class="ttl"><h1>Rapport Support <em>Tarkhiss</em></h1><p>{esc(R['subtitle'])}</p></div>"""
    body = f"""{head}{kpi_section()}
      {sec(0, 'Constats clés')}<ul class="pts">{''.join(f'<li>{esc(c)}</li>' for c in R['constats'])}</ul>
      <h3>Recommandations</h3><ol class="recos">{''.join(f'<li><span class="n">{i:02d}</span><span>{esc(r)}</span></li>' for i, r in enumerate(R['recos'], 1))}</ol>
      <div class="newp">{sec(2, 'Analyse graphique', '', 's2')}<div class="two">
        {fig(chart_weekly((4.3, 2.5), short=True), 'Évolution hebdomadaire', W_CAP)}{fig(chart_rank((4.3, 2.5), compact=True), 'Problèmes techniques', P_CAP)}</div>
      {sec(3, 'Problèmes techniques signalés (166)', '', 's3')}{cat_table(R['problems'])}
      {sec(4, 'Demandes d’information reçues (91)', '', 's4')}{cat_table(R['demandes'])}</div>
      <div class="newp">{sec(5, 'Évolution hebdomadaire', '', 's5')}{weekly_table()}
      {sec(6, 'Heatmap de charge', '', 's6')}{fig(chart_heat((8.6, 2.2)), 'Appels par jour et semaine du mois', HEAT_NOTE)}
      {sec(7, 'Statistiques avancées — appels & emails', '', 's7')}<p class="note">{ADV_NOTE}</p>{advanced()}</div>"""
    return css, body


def render(name, fn):
    import weasyprint
    css, body = fn()
    doc = f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>{esc(R["title"])} — {R["period"]}</title><style>{base_css()}{css}</style></head><body>{body}</body></html>'
    out = os.path.join(HERE, name)
    weasyprint.HTML(string=doc, base_url=HERE).write_pdf(out)
    from pypdf import PdfReader
    print(name, len(PdfReader(out).pages), "pages", os.path.getsize(out) // 1024, "Ko")


if __name__ == "__main__":
    build_fonts()
    render("variante_A_institutionnelle.pdf", variant_a)
    render("variante_B_direction.pdf", variant_b)
    render("variante_C_bulletin.pdf", variant_c)
