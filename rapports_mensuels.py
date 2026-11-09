# -*- coding: utf-8 -*-
"""
rapports_mensuels.py — Rapports mensuels Tarkhiss et Moussanada selon la charte AMMPS : PDF, PPTX et
corps d'email. tarkhiss_ctx() / moussanada_ctx() rassemblent les données une seule fois ; les trois
sorties affichent donc les mêmes chiffres. `a` = module app (helpers de lecture et de format), passé
en paramètre pour éviter un import circulaire.

Règles : le PPTX Tarkhiss exclut les statistiques de délai de réponse ; les emails respectent les
sections cochées en Administration et n'ont ni heatmap ni statistiques avancées.
"""
import re
from datetime import datetime

import ammps_charte as C
import retards
from ammps_charte import FOREST, GREEN, LIGHT, MUTED, RED, TINT, Deck, Mail, esc, fr_pct

CATEGORY_HEAD = ["Type", "Nb", "%", "Description"]
WEEK_HEAD = ["Semaine", "Bugs", "Demandes", "Observations"]
TOP_HEAD = ["Ouverts", "Résolus", "%"]
TECH_HEAD = ["Technicien", "Sollicitations", "Ouverts", "Résolus", "En instance", "vs équipe"]
LATE_HEAD = ["N°", "Titre", "Technicien", "Priorité", "Échéance", "Retard"]
ALERT_BG = "#FDECED"


class Charts:
    """Graphiques calculés à la demande et mémorisés : chaque sortie n'en utilise qu'une partie."""

    def __init__(self, makers):
        self.makers, self.done = makers, {}

    def __getitem__(self, key):
        if key not in self.done:
            maker = self.makers.get(key)
            self.done[key] = maker() if maker else None
        return self.done[key]


def titles(a, module, fallback_sub):
    """(titre, 2e ligne de garde, titre complet, sous-titre) : respecte les libellés de l'Administration."""
    gs = a.load_global_settings()
    label = a.MODULES[module]["label"]
    custom = (gs.get(f"report_title_{module}") or "").strip()
    sub = (gs.get(f"report_subtitle_{module}") or "").strip()
    return (custom or "Rapport Support", "" if custom else label, custom or f"Rapport Support {label}",
            a.report_subtitle(gs, module, "") if sub else fallback_sub)


def signature_parts(ms):
    return [ms.get("signature_name"), ms.get("signature_function"), ms.get("signature_phone")]


def full_width(a):
    return bool(a.load_global_settings().get("email_full_width"))


def edition():
    return datetime.now().strftime("%d/%m/%Y")


# ----------------------------------------------------------------------------- Tarkhiss
def tarkhiss_ctx(a, ym):
    calls, emails = a._load("tarkhiss", "calls", ym, {}), a._load("tarkhiss", "emails", ym, {})
    analysis = a._load("tarkhiss", "analysis", ym, a.empty_analysis(ym))
    py = a.prev_ym(ym)
    p_calls, p_emails = a._load("tarkhiss", "calls", py, {}), a._load("tarkhiss", "emails", py, {})
    p_analysis = a._load("tarkhiss", "analysis", py, a.empty_analysis(py))
    period = analysis.get("month_label") or a.month_label_fr(ym)
    prev = a.month_label_fr(py).split()[0].lower()
    lbl = a.get_kpi_labels("tarkhiss")
    title, title2, full, sub = titles(a, "tarkhiss", "Support Email & Hotline")

    def total(rows, key):
        return sum((d.get(key) or 0) for d in rows.values())

    received, sent, nb_calls = total(emails, "received"), total(emails, "sent"), total(calls, "calls")
    seconds = sum(a.duration_seconds(d) for d in calls.values())

    def category(items):
        tot = sum(i.get("count", 0) for i in items)
        return [(i.get("label", ""), i.get("count", 0), fr_pct(a.pct(i.get("count", 0), tot)), i.get("desc", "")) for i in items], tot

    problems, nb_prob = category(analysis.get("problems", []))
    demandes, nb_dem = category(analysis.get("demandes", []))
    p_prob = sum(i.get("count", 0) for i in p_analysis.get("problems", []))
    p_dem = sum(i.get("count", 0) for i in p_analysis.get("demandes", []))
    weekly = [(w.get("week", ""), w.get("bugs", 0) or 0, w.get("demandes", 0) or 0, w.get("obs", "")) for w in analysis.get("weekly", [])]
    crit = max(range(len(weekly)), key=lambda i: weekly[i][1] + weekly[i][2]) if weekly else None
    channels = nb_calls + received
    pc_mail = round(received / channels * 100) if channels else 0
    if not channels:
        channel = ("—", "")
    elif pc_mail >= 50:
        channel = (f"{pc_mail} % emails", f"{100 - pc_mail} % appels")
    else:
        channel = (f"{100 - pc_mail} % appels", f"{pc_mail} % emails")
    d = lambda cur, pv: C.delta_label(cur, pv, prev)

    matrix, weeks = a.compute_tarkhiss_heatmap(calls, ym)
    has_heat = any(any(r) for r in matrix)
    show_names = a.load_global_settings().get("report_show_contact_names", False)
    m_calls, m_emails = a._load("tarkhiss", "meta_calls", ym, {}), a._load("tarkhiss", "meta_emails", ym, {})
    adv = {}
    makers = {
        "weekly": lambda: C.weekly(weekly, crit) if weekly else None,
        "weekly_slide": lambda: C.weekly(weekly, crit, (6.1, 4.3), True) if weekly else None,
        "problems": lambda: C.rank(problems) if problems else None,
        "problems_slide": lambda: C.rank(problems, (6.1, 5.0)) if problems else None,
        "weekly_month": lambda: C.weekly(weekly, crit, (6.1, 3.0), True) if weekly else None,
        "problems_month": lambda: C.rank(problems, (6.1, 3.0)) if problems else None,
        "heat": lambda: C.heat(matrix, weeks, a_days()) if has_heat else None,
        "heat_slide": lambda: C.heat(matrix, weeks, a_days(), (8.0, 3.3)) if has_heat else None,
    }
    if m_calls:
        cn, du, cb = m_calls.get("counts", {}), m_calls.get("duration", {}), m_calls.get("callback", {})
        rate = m_calls.get("answer_rate")
        adv["calls_kpi"] = [("Taux de décroché", fr_pct(rate) if rate is not None else "—"), ("Appels manqués", cn.get("missed", 0)),
                            ("Durée médiane", a.format_hms(du.get("median_sec") or 0)), ("Durée maximale", a.format_hms(du.get("max_sec") or 0))]
        adv["callback"] = (f"Rappel des appels manqués ({cb.get('missed_total', 0)}) : {cb.get('called_back_by_us', 0)} rappelés par nous, "
                           f"{cb.get('client_called_again', 0)} le client a rerappelé, {cb.get('no_callback_24h', 0)} sans suite sous 24 h.")
        adv["callers"] = [(e["label"] if show_names else a.mask_call_label(e), e["calls"], e["missed"], a.format_hms(e["duration_sec"])) for e in m_calls.get("top_callers", [])[:8]]
        hc, hh = a.tm.business_hours_view(m_calls["hourly"]["in"])
        makers["calls_hourly"] = lambda: C.heat(hc, ["Lun", "Mar", "Mer", "Jeu", "Ven"], hh, (8.2, 2.5))
    if m_emails:
        rs = m_emails.get("response", {})
        rate = rs.get("match_rate")
        adv["emails_kpi"] = [("Délai médian (h. ouvrées)", fmt_min(rs.get("median_biz_min"))), ("P90", fmt_min(rs.get("p90_biz_min"))),
                             ("Sans réponse sous 7 j", m_emails.get("unanswered", 0)), ("Non lus", fr_pct(m_emails.get("pct_unread", 0)))]
        adv["emails_note"] = (f"{rs.get('note', '')} Base : {rs.get('business_hours', '')}. Correspondance par sujet : "
                              f"{fr_pct(rate) if rate is not None else '—'} des échanges appariés.")
        adv["subjects"] = [(s["subject"][:70], s["count"]) for s in m_emails.get("top_subjects", [])[:6]]
        adv["words"] = " · ".join(f"{w['word']} ({w['count']})" for w in m_emails.get("top_words", [])[:10])
        adv["recipients"] = [(r["name"] if show_names else "•••• (masqué)", r["count"]) for r in m_emails.get("top_recipients", [])[:6]]
        he, hh2 = a.tm.business_hours_view(m_emails["hourly"]["received"])
        makers["emails_hourly"] = lambda: C.heat(he, ["Lun", "Mar", "Mer", "Jeu", "Ven"], hh2, (8.2, 2.5))

    week_kpi = ("-", "")
    if crit is not None:
        week_kpi = (weekly[crit][0], f"{weekly[crit][1] + weekly[crit][2]} cas qualifiés")
    return dict(
        a=a, ym=ym, period=period, title=title, title2=title2, full=full, subtitle=sub, charts=Charts(makers), adv=adv,
        problems=problems, demandes=demandes, nb_prob=nb_prob, nb_dem=nb_dem, weekly=weekly, crit=crit, has_heat=has_heat,
        constats=analysis.get("constats", []), recos=analysis.get("recommandations", []),
        kpi_volume=[(lbl["emails_received"], received, d(received, total(p_emails, "received"))), (lbl["emails_sent"], sent, d(sent, total(p_emails, "sent"))),
                    (lbl["calls"], nb_calls, d(nb_calls, total(p_calls, "calls"))), (lbl["talk_time"], a.format_hms(seconds), "hh:mm:ss")],
        kpi_quali=[(lbl["problems"], nb_prob, d(nb_prob, p_prob), "alert"), (lbl["demandes"], nb_dem, d(nb_dem, p_dem), "good"),
                   (lbl["ratio_bd"], f"{(nb_prob / nb_dem if nb_dem else 0):.2f}".replace(".", ","), ""), (lbl["critical_week"], *week_kpi, "warn")],
        kpi_charge=[(lbl["avg_call"], a.format_hms(seconds / nb_calls if nb_calls else 0), ""), (lbl["dominant_channel"], *channel)],
    )


def a_days():
    return ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]


def fmt_min(v):
    if v is None:
        return "—"
    h, m = divmod(int(round(v)), 60)
    return f"{h} h {m:02d}" if h else f"{m} min"


def category_rows(items):
    return [[f"<b>{esc(l)}</b>", f"<b>{n}</b>", p, f'<span class="desc">{esc(d)}</span>'] for l, n, p, d in items]


def category_table(items):
    return C.table(CATEGORY_HEAD, category_rows(items), ["l w30", "c w6", "c w9", "l"])


def tarkhiss_pdf(c):
    ch, adv = c["charts"], c["adv"]
    entries = [("s1", "Indicateurs clés"), ("s2", "Analyse graphique"), ("s3", "Problèmes techniques"), ("s4", "Demandes d’information"), ("s5", "Évolution hebdomadaire")]
    if c["has_heat"]:
        entries.append(("s6", "Heatmap de charge"))
    if adv:
        entries.append(("s7", "Statistiques avancées — appels & emails"))
    entries.append(("s8", "Constats & recommandations"))
    num = {k: i for i, (k, _) in enumerate(entries, 1)}
    rows = []
    for i, (w, b, d, o) in enumerate(c["weekly"]):
        cells = [f"<b>{esc(w)}</b>", f"<b>{b}</b>", f"<b>{d}</b>", f'<span class="desc">{esc(o)}</span>']
        rows.append({"cells": cells} if i == c["crit"] else cells)
    weekly = C.table(WEEK_HEAD, rows, ["l w18", "c w8", "c w11", "l"])
    body = C.cover(c["title"], c["title2"], c["subtitle"], [("Période", c["period"]), ("Volet", "Tarkhiss"), ("Édition", edition())])
    body += f"""{C.head_html(c['full'], c['period'])}<div>{C.sec(0, 'Sommaire', 'Rapport mensuel')}{C.toc(entries)}</div>
    <div class="gap">{C.sec(num['s1'], 'Indicateurs clés', 'Vue d’ensemble', 's1')}
      <h4>Volumes</h4>{C.kpi_grid(c['kpi_volume'])}<h4>Qualification</h4>{C.kpi_grid(c['kpi_quali'])}<h4>Charge</h4>{C.kpi_grid(c['kpi_charge'], 2)}</div>
    <div class="newp">{C.sec(num['s2'], 'Analyse graphique', 'Tendances', 's2')}
      {C.fig(ch['weekly'], 'Évolution hebdomadaire — Bugs vs Demandes', 'Comparaison du nombre de bugs signalés et de demandes d’information, semaine par semaine. Semaine critique en rouge.')}
      {C.fig(ch['problems'], 'Répartition des problèmes techniques', 'Répartition des problèmes techniques par type, en proportion du total du mois.')}</div>
    <div class="newp">{C.sec(num['s3'], f"Problèmes techniques signalés ({c['nb_prob']})", 'Qualification', 's3')}{category_table(c['problems'])}</div>
    <div class="gap">{C.sec(num['s4'], f"Demandes d’information reçues ({c['nb_dem']})", 'Qualification', 's4')}{category_table(c['demandes'])}</div>
    <div class="newp">{C.sec(num['s5'], 'Évolution hebdomadaire', 'Rythme du mois', 's5')}{weekly}</div>"""
    if c["has_heat"]:
        body += f"""<div class="gap">{C.sec(num['s6'], 'Heatmap de charge', 'Rythme du mois', 's6')}{C.fig(ch['heat'], 'Appels par jour et semaine du mois',
                                                                                                           'Volume d’appels par jour de semaine et semaine du mois (granularité journalière : Tarkhiss ne journalise pas l’heure des appels).')}</div>"""
    if adv:
        body += f"""<div class="newp">{C.sec(num['s7'], 'Statistiques avancées — appels & emails', 'Imports bruts : journal d’appels et export Outlook', 's7')}{advanced_html(c)}</div>"""
    body += f"""<div class="{'newp' if adv else 'gap'}">{C.sec(num['s8'], 'Constats & recommandations', 'Synthèse', 's8')}
      <h3>Constats clés</h3>{C.points(c['constats'])}<h3>Recommandations</h3>{C.recos(c['recos'])}</div>"""
    return C.render_pdf(c["full"], "Tarkhiss", c["period"], c["a"].strip_platform_mentions(body))


def advanced_html(c):
    adv, ch = c["adv"], c["charts"]
    out = ""
    if "calls_kpi" in adv:
        callers = C.table(["Contact", "Appels", "Manqués", "Durée cumulée"], [[esc(str(x)), y, z, t] for x, y, z, t in adv["callers"]], ["l", "c", "c", "c"], "tight")
        out += f"""<h3>Journal d’appels</h3><div class="mini">{C.kpi_grid(adv['calls_kpi'])}</div><p class="note">{esc(adv['callback'])}</p>
        {C.fig(ch['calls_hourly'], 'Appels reçus par jour et par heure', 'Jours ouvrés, 08h-17h, heures locales.')}<h4>Top appelants (anonymisé)</h4>{callers}"""
    if "emails_kpi" in adv:
        subjects = C.table(["Sujet", "Occurrences"], [[esc(s), f"<b>{n}</b>"] for s, n in adv["subjects"]], ["l", "c w14"], "tight")
        recipients = C.table(["Destinataire", "Emails envoyés"], [[esc(s), n] for s, n in adv["recipients"]], ["l", "c"], "tight")
        out += f"""<h3>Emails</h3><div class="mini">{C.kpi_grid(adv['emails_kpi'])}</div><p class="note">{esc(adv['emails_note'])}</p>
        {C.fig(ch['emails_hourly'], 'Emails reçus par jour et par heure', 'Jours ouvrés, 08h-17h, heures locales.')}<h4>Sujets les plus fréquents</h4>{subjects}
        <div class="two"><div><h4>Mots-clés fréquents</h4><p class="words">{esc(adv['words'])}</p></div><div><h4>Top destinataires (anonymisé)</h4>{recipients}</div></div>"""
    return out


def tarkhiss_pptx(c):
    a, ch, adv = c["a"], c["charts"], c["adv"]
    per = c["period"]
    d = Deck(c["full"], per, f"AMMPS · Tarkhiss — {per}", a.strip_platform_mentions)
    d.cover("Rapport mensuel", c["title"], c["title2"], c["subtitle"], [("Période", per), ("Édition", edition())])
    n = 1
    p = d.page(n, "Indicateurs clés", "Vue d’ensemble")
    d.kcards(p, c["kpi_volume"], .45, 1.72, 12.43, 1.5)
    d.kcards(p, c["kpi_quali"], .45, 3.34, 12.43, 1.5)
    d.kcards(p, c["kpi_charge"], .45, 4.96, 12.43, 1.5)
    n += 1
    p = d.page(n, "Analyse graphique", "Tendances")
    d.box_pic(p, ch["weekly_slide"], .45, 1.72, 6.13, 5.17, "Évolution hebdomadaire — Bugs vs Demandes")
    d.box_pic(p, ch["problems_slide"], 6.75, 1.72, 6.13, 5.17, "Répartition des problèmes techniques")
    wt, al = [3.1, .6, .8, 6.5], ["l", "c", "c", "l"]
    bold = lambda r, k: k < 2
    for title, items in ((f"Problèmes techniques signalés ({c['nb_prob']})", c["problems"]), (f"Demandes d’information reçues ({c['nb_dem']})", c["demandes"])):
        n += 1
        p = d.page(n, title, "Qualification")
        d.table(p, .45, 1.72, 12.43, wt, CATEGORY_HEAD, [list(i) for i in items], font=12, row_h=min(.58, 5.1 / (len(items) + 1)), aligns=al, bolds=bold)
    n += 1
    p = d.page(n, "Évolution hebdomadaire", "Rythme du mois")
    w, crit = c["weekly"], c["crit"]
    d.table(p, .45, 1.72, 12.43, [1.7, .8, 1, 7.5], WEEK_HEAD, [list(x) for x in w], font=12, row_h=min(.62, 3.7 / (len(w) + 1)), aligns=al,
            fills=lambda r, k: ALERT_BG if r == crit else None, bolds=lambda r, k: k < 3)
    if crit is not None:
        cw = w[crit]
        d.note(p, .45, 5.6, 12.43, 1.1, "Semaine critique", f"{cw[0]} : {cw[1] + cw[2]} cas qualifiés, dont {cw[1]} bugs ({fr_pct(a.pct(cw[1], c['nb_prob']))} des incidents du mois).", 13, RED, ALERT_BG)
    if c["has_heat"] or "calls_kpi" in adv:
        n += 1
        p = d.page(n, "Activité hotline", "Appels")
        top = 1.72
        if "calls_kpi" in adv:
            d.kcards(p, adv["calls_kpi"], .45, 1.72, 12.43, 1.05)
            top = 2.9
            d.box_pic(p, ch["heat_slide"], .45, top, 8.2, 6.9 - top, "Répartition des appels par jour et par semaine")
            d.note(p, 8.8, top, 4.08, 6.9 - top, "Rappel des appels manqués", adv["callback"], 13)
        else:
            d.box_pic(p, ch["heat_slide"], .45, top, 12.43, 6.9 - top, "Répartition des appels par jour et par semaine")
    n += 1
    p = d.page(n, "Constats clés et recommandations", "Synthèse")
    d.bullets(p, .45, 1.72, 7.4, 5.18, "Constats clés", c["constats"], 14)
    recos = c["recos"][:4]
    h = 5.18 / max(len(recos), 1)
    for i, r in enumerate(recos):
        y = 1.72 + i * h
        d.rect(p, 8.0, y, 4.88, h - .07, LIGHT)
        d.rect(p, 8.0, y, .05, h - .07, GREEN)
        d.text(p, 8.14, y + .06, 1, .4, f"{i + 1:02d}", 15, True, GREEN)
        d.text(p, 8.14, y + .45, 4.65, h - .55, r, C.fit_size(r, 4.6, h - .6, 12.5))
    return d.finish()


def tarkhiss_mail(c, greeting, ms):
    a, ch = c["a"], c["charts"]
    es = ms.get("email_sections", {})
    on = lambda k: es.get(k, True)
    m = Mail
    b = m.sec(1, "Indicateurs clés") + m.kpis(c["kpi_volume"]) + m.kpis(c["kpi_quali"]) + m.kpis(c["kpi_charge"], 2)
    if on("graphs"):
        b += m.sec(2, "Analyse graphique")
        if on("chart_weekly"):
            b += m.img(ch["weekly"], "Évolution hebdomadaire — Bugs vs Demandes")
        if on("chart_problems"):
            b += m.img(ch["problems"], "Répartition des problèmes techniques")
    cat = lambda items: [[f"<b>{esc(l)}</b>", f"<b>{n}</b>", p, f'<span style="color:{MUTED};font-size:12px;">{esc(d)}</span>'] for l, n, p, d in items]
    al = ["left", "center", "center", "left"]
    if on("problems"):
        b += m.sec(3, f"Problèmes techniques signalés ({c['nb_prob']})") + m.table(CATEGORY_HEAD, cat(c["problems"]), al)
    if on("demandes"):
        b += m.sec(4, f"Demandes d’information reçues ({c['nb_dem']})") + m.table(CATEGORY_HEAD, cat(c["demandes"]), al)
    if on("weekly"):
        rows = []
        for i, (w, x, y, o) in enumerate(c["weekly"]):
            cells = [f'<b style="white-space:nowrap;">{esc(w)}</b>', f"<b>{x}</b>", f"<b>{y}</b>", f'<span style="color:{MUTED};font-size:12px;">{esc(o)}</span>']
            rows.append({"cells": cells} if i == c["crit"] else cells)
        b += m.sec(5, "Évolution hebdomadaire") + m.table(WEEK_HEAD, rows, al)
    if on("synth"):
        b += m.sec(6, "Constats clés") + m.bullets(c["constats"]) + m.sec(0, "Recommandations") + "".join(m.note(f"{i:02d}   {r}") for i, r in enumerate(c["recos"], 1))
    intro = f"Veuillez trouver ci-dessous le rapport mensuel du support Tarkhiss ({c['period']})."
    return a.strip_platform_mentions(m.shell(c["full"], c["period"], greeting, intro, b, signature_parts(ms), full_width=full_width(a)))


# ----------------------------------------------------------------------------- Moussanada
def tail_labels(items):
    """Libellés de graphique : dernier niveau de l'arborescence, deux niveaux si homonymes."""
    parts = [[p.strip() for p in i.get("label", "").split(">")] for i in items]
    last = [p[-1] for p in parts]
    return [" > ".join(p[-2:]) if last.count(p[-1]) > 1 else p[-1] for p in parts]


def moussanada_ctx(a, ym):
    d = a.moussanada_data(ym).get_json()
    m, ts, notes = d["month"], d["timeseries"], d["notes"]
    tk, du = m["tickets"], m["durations"]
    py = a.prev_ym(ym)
    prev = ts.get(py, {})
    ml, pl = a.month_label_fr(ym), a.month_label_fr(py)
    plow = pl.split()[0].lower()
    lbl = a.get_kpi_labels("moussanada")
    inst = lbl["tickets_retard"] if lbl["tickets_retard"] != a.MOUSSANADA_DEFAULT_KPI_LABELS["tickets_retard"] else "Tickets en instance"
    title, title2, full, sub = titles(a, "moussanada", "Support Helpdesk")

    def k(label, val, cur, pv, pol, fmt=str):
        s = C.delta_label(cur, pv, plow, fmt)
        tone = ""
        if pv and pol != "neutral" and abs(cur - pv) / pv >= .01:
            tone = "good" if (cur > pv) == (pol == "positive") else "alert"
        return (label, str(val), s, tone)

    pc, dh = a.pct, a.format_dh
    tr, tc = pc(tk["resolus"], tk["ouverts"]), pc(tk["clos"], tk["ouverts"])
    ptr, ptc = pc(prev.get("resolus", 0), prev.get("ouverts", 0)), pc(prev.get("clos", 0), prev.get("ouverts", 0))
    months = a.real_months(ts, require_data=True)
    labs = [a.month_label_fr(x).split()[0][:4] + "." for x in months]
    cats = sorted(m["categories"], key=lambda x: x.get("ouverts", 0), reverse=True)
    srv = sorted(m["services"], key=lambda x: x.get("ouverts", 0), reverse=True)
    techs = sorted(m["techniciens"], key=a.tech_volume, reverse=True)
    tot, cum, cp = sum(c.get("ouverts", 0) for c in cats) or 1, 0, []
    for c in cats[:12]:
        cum += c.get("ouverts", 0)
        cp.append(round(cum / tot * 100))
    keys = ("ouverts", "resolus", "en_retard", "clos")
    names = ["Ouverts", "Résolus", "En instance", "Clos"]
    vs = {pl: [prev.get(x, 0) or 0 for x in keys], ml: [tk[x] for x in keys]}
    delays = {pl: [round(prev.get("resolution_h", 0) or 0, 1), round(prev.get("cloture_h", 0) or 0, 1)], ml: [round(du["resolution_h"] or 0, 1), round(du["cloture_h"] or 0, 1)]}
    delay_names = ["Délai résolution", "Délai clôture"]
    two = lambda rows: {"Ouverts": [r.get("ouverts", 0) for r in rows], "Résolus": [r.get("resolus", 0) for r in rows]}
    ts_of = lambda key: [ts[x].get(key, 0) or 0 for x in months]
    tvol = lambda: {"Sollicitations (ouverts + résolus)": [a.tech_volume(t) for t in techs]}
    # graphiques techniciens : rendu du dashboard (couleurs, valeurs, quadrillage), sans les repères ▲ / ▼ (marks)
    dash = lambda series, size: C.dash_bars(tnames, series, [C.TECH_VOL] if len(series) == 1 else [C.TECH_OPEN, C.TECH_DONE], size)
    tnames = [t["label"] for t in techs]
    makers = {
        "volume": lambda: C.bars(names, vs),
        "delay": lambda: C.bars(delay_names, delays, figsize=(4.4, 2.8)),
        "volume_half": lambda: C.bars(names, vs, figsize=(4.4, 2.8)),
        "delay_wide": lambda: C.bars(delay_names, delays, figsize=(8.2, 2.2)),
        "s_volume": lambda: C.bars(names, vs, figsize=(6.0, 2.7)),
        "s_volume2": lambda: C.bars(names, vs, figsize=(6.0, 2.6)),
        "s_delay": lambda: C.bars(delay_names, delays, figsize=(6.0, 2.6)),
    }
    if months:
        trend = lambda size: C.lines(labs, {"Ouverts": ts_of("ouverts"), "Résolus": ts_of("resolus"), "Clos": ts_of("clos")}, figsize=size)
        makers.update(backlog=lambda: C.bars(labs, {"En instance": ts_of("en_retard")}, figsize=(8.2, 1.9)), trend=lambda: trend((8.2, 2.3)),
                      s_backlog=lambda: C.bars(labs, {"En instance": ts_of("en_retard")}, figsize=(6.0, 2.7)), s_trend=lambda: trend((7.6, 3.9)))
    if cats:
        cl, pl_c = tail_labels(cats[:10]), tail_labels(cats[:12])
        pv = [c.get("ouverts", 0) for c in cats[:12]]
        makers.update(cat=lambda: C.hbars(cl, two(cats[:10])), pareto=lambda: C.pareto(pl_c, pv, cp),
                      s_cat=lambda: C.hbars(cl, two(cats[:10]), figsize=(6.2, 5.1), maxlen=34), s_pareto=lambda: C.pareto(pl_c, pv, cp, figsize=(12.2, 3.3)))
    if srv:
        sl = tail_labels(srv[:10])
        makers.update(srv=lambda: C.hbars(sl, two(srv[:10])), s_srv=lambda: C.hbars(sl, two(srv[:10]), figsize=(6.2, 3.7), maxlen=46))
    if techs:
        makers.update(
            techvol=lambda: dash(tvol(), (4.4, 3.8)), tech=lambda: dash(two(techs), (8.2, 2.9)),
            techvol_wide=lambda: dash(tvol(), (8.2, 2.5)), radar=lambda: C.radar(techs),
            s_techvol=lambda: dash(tvol(), (4.4, 2.8)), s_tech=lambda: dash(two(techs), (4.4, 2.8)),
            s_radar=lambda: C.radar(techs, figsize=(3.6, 2.9)))
    raw = a.load_raw_tickets(ym, ym)
    matrix = a.compute_heatmap_matrix(raw) if raw else None
    if matrix:
        makers["heat"] = lambda: C.heat(matrix, a_days(), [f"{h}h" for h in range(24)], (8.2, 2.8))
        makers["s_heat"] = lambda: C.heat(matrix, a_days(), [f"{h}h" for h in range(24)], (12.2, 4.3))

    n = len(techs) or 1

    def top_rows(items):
        total = sum(i.get("ouverts", 0) for i in items) or 1
        return [[esc(i.get("label", "")), f"<b>{i.get('ouverts', 0)}</b>", i.get("resolus", 0), fr_pct(pc(i.get("ouverts", 0), total))] for i in items]

    avg_v = sum(a.tech_volume(t) for t in techs) / n
    avg = lambda key: f"{sum(t.get(key, 0) for t in techs) / n:.1f}".replace(".", ",")
    tech_rows = [[esc(t.get("label", "")), f"<b>{a.tech_volume(t)}</b>", t.get("ouverts", 0), t.get("resolus", 0), t.get("en_retard", 0),
                  "≥ moy." if a.tech_volume(t) >= avg_v else "&lt; moy."] for t in techs]
    if techs:
        tech_rows.append({"cls": "tot", "cells": ["Moyenne équipe", f"{avg_v:.1f}".replace(".", ","), avg("ouverts"), avg("resolus"), "", ""]})
    return dict(
        a=a, ym=ym, ml=ml, pl=pl, title=title, title2=title2, full=full, subtitle=sub, sections=notes.get("sections", {}), charts=Charts(makers),
        constats=notes.get("constats", []) or [], recos=notes.get("recommandations", []) or [], cats=cats[:10], srv=srv[:10], top_rows=top_rows, tech_rows=tech_rows, has_heat=bool(matrix),
        k1=[k(lbl["tickets_ouverts"], tk["ouverts"], tk["ouverts"], prev.get("ouverts", 0), "neutral"), k(lbl["tickets_resolus"], tk["resolus"], tk["resolus"], prev.get("resolus", 0), "positive"),
            k(lbl["tickets_clos"], tk["clos"], tk["clos"], prev.get("clos", 0), "positive"), k(inst, tk["en_retard"], tk["en_retard"], prev.get("en_retard", 0), "negative")],
        k2=[k(lbl["taux_resolution"], fr_pct(tr), tr, ptr, "positive", fr_pct), k(lbl["taux_cloture"], fr_pct(tc), tc, ptc, "positive", fr_pct),
            k(lbl["delai_resolution"], dh(du["resolution_h"]), du["resolution_h"], prev.get("resolution_h", 0), "negative", dh),
            k(lbl["delai_cloture"], dh(du["cloture_h"]), du["cloture_h"], prev.get("cloture_h", 0), "negative", dh)],
    )


def moussanada_pdf(c):
    s, ch = c["sections"], c["charts"]
    title3, title6 = f"Principales catégories de demandes en {c['ml']}", f"Analyse comparative {c['pl']} / {c['ml']}"
    items = [("m1", "Volume global des tickets"), ("m2", "Évolution des délais de traitement"), ("m3", title3), ("m4", "Services les plus demandeurs"),
             ("m5", "Charge et mobilisation de l’équipe SI"), ("m6", title6)]
    if c["has_heat"]:
        items.append(("m7", "Heatmap de charge (jour x heure)"))
    items.append(("m8", "Synthèse finale et conclusion"))
    num = {k: i for i, (k, _) in enumerate(items, 1)}
    A = ["l", "c w9", "c w9", "c w9"]
    tb = lambda rows, first: C.table([first] + TOP_HEAD, rows, A, "tight")
    body = C.cover(c["title"], c["title2"], c["subtitle"], [("Période", c["ml"]), ("Volet", "Moussanada"), ("Édition", edition())])
    body += f"""{C.head_html(c['full'], c['ml'])}<div>{C.sec(0, 'Sommaire', 'Rapport mensuel')}{C.toc(items)}</div>
    <div class="gap">{C.sec(1, items[0][1], 'Volumes', 'm1')}{C.kpi_grid(c['k1'])}
      {C.fig(ch['volume'], 'Volumes de tickets — comparatif', f"Comparatif {c['pl']} / {c['ml']}.")}
      {C.fig(ch['backlog'], 'Backlog — tickets en instance', 'Tickets en instance, mois par mois.')}{C.notebox(s.get('volume'))}</div>
    <div class="gap">{C.sec(2, items[1][1], 'Délais', 'm2')}{C.kpi_grid(c['k2'])}
      {C.fig(ch['trend'], 'Tendance des tickets', 'Tickets ouverts, résolus et clos, mois par mois.')}{C.notebox(s.get('delays'))}</div>
    <div class="gap"><div class="keep">{C.sec(3, title3, 'Demandes', 'm3')}{tb(c['top_rows'](c['cats']), 'Catégorie')}</div>
      <div class="gap">{C.fig(ch['cat'], 'Top catégories', 'Dix premières catégories : tickets ouverts et résolus.')}</div>
      {C.fig(ch['pareto'], 'Pareto des catégories', 'Part cumulée des tickets ouverts.')}{C.notebox(s.get('categories'))}</div>
    <div class="gap"><div class="keep">{C.sec(4, items[3][1], 'Demandes', 'm4')}{tb(c['top_rows'](c['srv']), 'Service')}</div>
      <div class="gap">{C.fig(ch['srv'], 'Top services demandeurs', 'Dix premiers services : tickets ouverts et résolus.')}</div>{C.notebox(s.get('services'))}</div>
    <div class="gap"><div class="keep">{C.sec(5, items[4][1], 'Équipe', 'm5')}{C.table(TECH_HEAD, c['tech_rows'], ['l', 'c', 'c', 'c', 'c', 'c'], 'tight')}</div>
      <div class="gap">{C.fig(ch['tech'], 'Détail des sollicitations', 'Tickets ouverts et résolus par technicien.')}</div>
      <div class="two">{C.fig(ch['techvol'], 'Volume de sollicitation', 'Ouverts + résolus.')}{C.fig(ch['radar'], 'Radar de performance', 'Principaux techniciens.')}</div>{C.notebox(s.get('team'))}</div>
    <div class="gap"><div class="keep">{C.sec(6, title6, 'Comparatif', 'm6')}
      <div class="two">{C.fig(ch['volume_half'], 'Volumes', 'Comparatif des volumes.')}{C.fig(ch['delay'], 'Délais (heures)', 'Délais moyens de traitement.')}</div></div>{C.notebox(s.get('comparative'))}</div>"""
    if c["has_heat"]:
        body += f"""<div class="gap keep">{C.sec(num['m7'], 'Heatmap de charge (jour x heure)', 'Rythme du mois', 'm7')}{C.fig(ch['heat'], 'Tickets ouverts par jour et heure',
                                                                                                                           'Basée sur les tickets bruts importés. Utile pour dimensionner les plages de présence hotline.')}</div>"""
    constats = f"<h3>Constats clés</h3>{C.points(c['constats'])}" if c["constats"] else ""
    recos = f"<h3>Recommandations</h3>{C.recos(c['recos'])}" if c["recos"] else ""
    body += f"""<div class="gap">{C.sec(num['m8'], 'Synthèse finale et conclusion', 'Synthèse', 'm8')}{constats}{recos}{C.notebox(s.get('synthesis'), 'Conclusion')}</div>"""
    return C.render_pdf(c["full"], "Moussanada", c["ml"], c["a"].strip_platform_mentions(body))


def moussanada_pptx(c):
    a, ch, s = c["a"], c["charts"], c["sections"]
    d = Deck(c["full"], c["ml"], f"AMMPS · Moussanada — {c['ml']}", a.strip_platform_mentions)
    d.cover("Rapport mensuel", c["title"], c["title2"], c["subtitle"], [("Période", c["ml"]), ("Édition", edition())])

    p = d.page(1, "Volume global des tickets", "Volumes")
    d.kcards(p, c["k1"], .45, 1.72, 12.43, 1.1)
    d.box_pic(p, ch["s_volume"], .45, 2.93, 6.13, 2.85, "Volumes de tickets — comparatif")
    d.box_pic(p, ch["s_backlog"], 6.75, 2.93, 6.13, 2.85, "Backlog — tickets en instance")
    d.note(p, .45, 5.88, 12.43, 1.06, "Analyse", s.get("volume"), 11)

    p = d.page(2, "Évolution des délais de traitement", "Délais")
    d.kcards(p, c["k2"], .45, 1.72, 12.43, 1.1)
    d.box_pic(p, ch["s_trend"], .45, 2.93, 7.6, 3.98, "Tendance des tickets")
    d.note(p, 8.2, 2.93, 4.68, 3.98, "Analyse", s.get("delays"), 12)

    rows = lambda items: [[t] + r[1:] for t, r in zip(tail_labels(items), c["top_rows"](items))]
    bold = lambda r, k: k == 1
    p = d.page(3, f"Principales catégories de demandes en {c['ml']}", "Demandes")
    d.table(p, .45, 1.72, 6.05, [3.5, .9, .9, .9], ["Catégorie"] + TOP_HEAD, rows(c["cats"]), row_h=.46, bolds=bold, font=11)
    d.box_pic(p, ch["s_cat"], 6.65, 1.72, 6.23, 5.17, "Top catégories")

    p = d.page(4, "Pareto des catégories", "Demandes")
    d.box_pic(p, ch["s_pareto"], .45, 1.72, 12.43, 3.75, "Part cumulée des tickets ouverts")
    d.note(p, .45, 5.58, 12.43, 1.36, "Analyse", s.get("categories"), 11.5)

    p = d.page(5, "Services les plus demandeurs", "Demandes")
    d.table(p, .45, 1.72, 6.05, [3.5, .9, .9, .9], ["Service"] + TOP_HEAD, rows(c["srv"]), row_h=.35, bolds=bold, font=10.5)
    d.box_pic(p, ch["s_srv"], 6.65, 1.72, 6.23, 3.85, "Top services demandeurs")
    d.note(p, .45, 5.68, 12.43, 1.26, "Analyse", s.get("services"), 11)

    p = d.page(6, "Charge et mobilisation de l’équipe SI", "Équipe")
    rows_t = [r["cells"] if isinstance(r, dict) else r for r in c["tech_rows"]]
    nt = len(rows_t)
    d.table(p, .45, 1.72, 7.0, [2.4, 1.2, .9, .9, 1, .9], TECH_HEAD, rows_t, row_h=min(.31, 2.2 / (nt + 1)), font=10.5 if nt < 7 else 9,
            fills=lambda r, k: LIGHT if r == nt - 1 else None, bolds=lambda r, k: r == nt - 1 or k == 1)
    d.note(p, 7.6, 1.72, 5.28, 2.2, "Analyse", s.get("team"), 10.5)
    d.box_pic(p, ch["s_tech"], .45, 4.02, 4.45, 2.88, "Détail des sollicitations")
    d.box_pic(p, ch["s_techvol"], 5.0, 4.02, 4.45, 2.88, "Volume de sollicitation")
    d.box_pic(p, ch["s_radar"], 9.55, 4.02, 3.33, 2.88, "Radar de performance")

    n = 7
    if c["has_heat"]:
        p = d.page(n, "Heatmap de charge (jour x heure)", "Rythme du mois")
        d.box_pic(p, ch["s_heat"], .45, 1.72, 12.43, 5.17, "Tickets ouverts par jour et heure")
        n += 1
    p = d.page(n, f"Analyse comparative {c['pl']} / {c['ml']}", "Comparatif et synthèse")
    d.box_pic(p, ch["s_volume2"], .45, 1.72, 6.13, 2.85, "Volumes")
    d.box_pic(p, ch["s_delay"], 6.75, 1.72, 6.13, 2.85, "Délais (heures)")
    d.note(p, .45, 4.68, 12.43, 1.1, "Analyse", s.get("comparative"), 10.5)
    d.note(p, .45, 5.86, 12.43, 1.08, "Synthèse finale et conclusion", s.get("synthesis"), 10.5, "#FFFFFF", FOREST, "#FFFFFF")
    if c["constats"] or c["recos"]:
        n += 1
        p = d.page(n, "Constats clés et recommandations", "Synthèse")
        d.bullets(p, .45, 1.72, 6.15, 5.18, "Constats clés", c["constats"], 13)
        d.bullets(p, 6.73, 1.72, 6.15, 5.18, "Recommandations", c["recos"], 13)
    return d.finish()


def late_section(a, ms, number, attached, now=None):
    """Section « Tickets en retard » de l'email : cartes KPI + TOP N (techniciens exclus écartés).
    Vide tant qu'aucun export de tickets bruts n'a été importé. Les exclusions ne sont jamais nommées."""
    if not a.raw_tickets_count():
        return ""
    v = retards.view(a.load_raw_tickets(open_only=True), ms, now)
    m = Mail
    n = v["top_n"]
    title = f"Tickets en retard — TOP {n}" if v["shown_total"] > n else "Tickets en retard"
    if not v["shown"]:
        return m.sec(number, title) + m.note("Aucun ticket en retard à signaler à ce jour.")
    total, avg_days, max_days, high = retards.stats(v["shown"])   # techniciens écartés exclus des KPI comme du TOP et de l'Excel
    cards = [("Tickets en retard", total, "échéance dépassée", "alert"),
             ("Retard moyen", f"{retards_days(avg_days)} j", "écart à l’échéance", ""),
             ("Retard maximal", f"{retards_days(max_days)} j", "plus ancien retard", "alert" if max_days >= 30 else ""),
             ("Priorité haute", high, "en retard", "alert" if high else "")]
    cut = lambda s, k=88: s if len(s) <= k else s[:k - 1].rstrip() + "…"
    rows = [{"cls": "hl" if t["priorite"].lower() in ("haute", "très haute", "majeure") else "", "cells": [
        esc(t["id"]), esc(cut(t["titre"])), esc(t["technicien"]), esc(t["priorite"]), t["echeance"].strftime("%d/%m/%Y"), f"<b>{esc(t['retard'])}</b>"]}
        for t in v["top"]]
    note = f"Les {len(rows)} tickets les plus en retard, sur {total} au total. " if len(rows) < total else ""
    note += f"Situation au {v['as_of']:%d/%m/%Y}."
    if attached:
        note += f"\nLa liste complète des {total} tickets en retard figure dans le fichier Excel joint."
    return (m.sec(number, title) + m.kpis(cards) + m.table(LATE_HEAD, rows, ["center", "left", "left", "center", "center", "center"])
            + m.note(note))


def retards_days(v):
    return f"{v:.1f}".replace(".", ",")


def moussanada_mail(c, greeting, ms):
    a, ch, s = c["a"], c["charts"], c["sections"]
    es = ms.get("email_sections", {})
    on = lambda k: es.get(k, True)
    m = Mail
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
    if on("late"):
        b += late_section(a, ms, 6, on("late_xlsx"))
    if on("comparative"):
        b += m.sec(7, f"Analyse comparative {c['pl']} / {c['ml']}") + im("delay_wide", "Délais de traitement (heures)") + m.note(s.get("comparative"))
    if on("synth"):
        b += m.sec(8, "Synthèse finale et conclusion")
        if c["constats"]:
            b += m.bullets(c["constats"])
        if c["recos"]:
            b += m.sec(0, "Recommandations") + "".join(m.note(f"{i:02d}   {r}") for i, r in enumerate(c["recos"], 1))
        b += m.note(s.get("synthesis"))
    intro = f"Veuillez trouver ci-dessous le rapport mensuel du support Moussanada ({c['ml']})."
    return a.strip_platform_mentions(m.shell(c["full"], c["ml"], greeting, intro, b, signature_parts(ms), full_width=full_width(a)))


# ----------------------------------------------------------------------------- points d'entrée de app.py
def monthly_pdf(a, module, ym):
    return tarkhiss_pdf(tarkhiss_ctx(a, ym)) if module == "tarkhiss" else moussanada_pdf(moussanada_ctx(a, ym))


def monthly_pptx(a, module, ym):
    return tarkhiss_pptx(tarkhiss_ctx(a, ym)) if module == "tarkhiss" else moussanada_pptx(moussanada_ctx(a, ym))


def monthly_mail(a, module, ym, greeting, ms):
    """(HTML du corps, libellé du mois)."""
    if module == "tarkhiss":
        c = tarkhiss_ctx(a, ym)
        return tarkhiss_mail(c, greeting, ms), c["period"]
    c = moussanada_ctx(a, ym)
    return moussanada_mail(c, greeting, ms), c["ml"]
