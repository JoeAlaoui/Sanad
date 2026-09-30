"""SANAD — Tarkhiss : agrégations détaillées issues des exports bruts.

Deux sources (aucune BDD, JSON à plat, stockage géré par app.py) :
  - export Outlook (emails)      -> aggregate_emails()
  - journal d'appels du mobile   -> parse_calllog_csv() + aggregate_calls()

Fonctions pures, sans dépendance Flask : testables seules.
Heures : celles du fichier, supposées en heure locale (Maroc). Jours : 0 = lundi.
"""
import csv
import io
import re
import statistics
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, timedelta

# Plage "ouvrée" utilisée pour le délai de première réponse (hypothèse, à ajuster ici).
BIZ_START_H, BIZ_END_H = 8, 17
BIZ_DAY_MIN = (BIZ_END_H - BIZ_START_H) * 60
REPLY_MAX_DAYS = 7  # au-delà, l'échange est considéré sans réponse (sujet réutilisé)

STOPWORDS = set("""
de du des la le les un une et en au aux pour par sur dans avec sans que qui quoi est sont pas plus
ce cet cette ces mon ma mes ton ta tes son sa ses nos vos leur leurs the and for with from your
this that you are not have has was will please merci bonjour bonsoir cordialement demande
svp urgent objet suite concernant sujet tarkhiss ammps
""".split())
_PREFIX_RE = re.compile(r"^\s*((re|tr|fw|fwd|rép|rep|réf|ref)\s*:\s*)+", re.I)
_TAG_RE = re.compile(r"\[[^\]]*\]")
_TOKEN_RE = re.compile(r"[A-Za-zÀ-ÿ']{4,}")


def _empty_matrix():
    return [[0] * 24 for _ in range(7)]


def _percentile(sorted_vals, p):
    if not sorted_vals:
        return None
    k = (len(sorted_vals) - 1) * p
    lo, hi = int(k), min(int(k) + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (k - lo)


def _round(v, n=1):
    return None if v is None else round(v, n)


# Horaires d'ouverture Tarkhiss (hotline + support email) — utilisés pour restreindre l'AFFICHAGE
# des heatmaps jour×heure (appels/emails) aux jours ouvrés et à la plage réellement couverte par
# le service. Les autres statistiques (taux de décroché, délai de réponse...) restent, elles,
# calculées sur la totalité des données importées — seule la heatmap est recadrée à l'affichage.
HEATMAP_HOUR_START, HEATMAP_HOUR_END = 8, 16  # colonnes affichées : 8h à 16h inclus (≈ 08h30-16h30)
HEATMAP_DAYS = 5  # lignes affichées : Lun-Ven uniquement


def business_hours_view(matrix):
    """Découpe une matrice 7×24 (jour×heure, Lun=0..Dim=6) sur la plage d'ouverture Tarkhiss :
    jours ouvrés (Lun-Ven) et heures HEATMAP_HOUR_START à HEATMAP_HOUR_END inclus. Retourne
    (sous_matrice, labels_heures) — à utiliser uniquement pour l'affichage de la heatmap."""
    hours = list(range(HEATMAP_HOUR_START, HEATMAP_HOUR_END + 1))
    sub = [[row[h] for h in hours] for row in matrix[:HEATMAP_DAYS]]
    return sub, [str(h) for h in hours]


def ym_of(dt):
    return f"{dt.year}-{dt.month:02d}"


# ---------------------------------------------------------------------------
# EMAILS
# ---------------------------------------------------------------------------
def normalize_subject(subject):
    s = _TAG_RE.sub(" ", str(subject or ""))
    s = _PREFIX_RE.sub("", s)
    return re.sub(r"\s+", " ", s).strip()


def business_minutes(a, b):
    """Minutes ouvrées (lun-ven, BIZ_START_H-BIZ_END_H) entre a et b."""
    if b <= a:
        return 0.0
    total = 0.0
    day = a.replace(hour=0, minute=0, second=0, microsecond=0)
    while day <= b:
        if day.weekday() < 5:
            s = max(a, day.replace(hour=BIZ_START_H))
            e = min(b, day.replace(hour=BIZ_END_H))
            if e > s:
                total += (e - s).total_seconds() / 60
        day += timedelta(days=1)
    return total


def _split_recipients(raw):
    return [p.strip() for p in str(raw or "").split(";") if p.strip()]


def aggregate_emails(events):
    """events : liste de dicts {kind: 'received'|'sent', dt, subject, to, importance, attachment, is_read}.
    Retourne {ym: meta} (un meta par mois présent)."""
    metas = {}

    def month(ym):
        return metas.setdefault(ym, {
            "ym": ym,
            "hourly": {"received": _empty_matrix(), "sent": _empty_matrix()},
            "weekday": {"received": [0] * 7, "sent": [0] * 7},
            "received": 0, "sent": 0, "attachments": 0, "high_importance": 0, "unread": 0,
            "_delays_cal": [], "_delays_biz": [],
            "threads": 0, "unanswered": 0,
            "_subjects": Counter(), "_subject_label": {}, "_words": Counter(), "_rcpt": Counter(),
        })

    for ev in events:
        dt = ev["dt"]
        m = month(ym_of(dt))
        k = ev["kind"]
        m["hourly"][k][dt.weekday()][dt.hour] += 1
        m["weekday"][k][dt.weekday()] += 1
        m[k] += 1
        if k == "received":
            if ev.get("attachment"):
                m["attachments"] += 1
            if ev.get("importance") == "High":
                m["high_importance"] += 1
            if ev.get("is_read") in (False, 0, "False", "false"):
                m["unread"] += 1
            norm = normalize_subject(ev.get("subject"))
            if norm:
                key = norm.lower()
                m["_subjects"][key] += 1
                m["_subject_label"].setdefault(key, norm)
                for w in _TOKEN_RE.findall(norm.lower()):
                    if w not in STOPWORDS:
                        m["_words"][w] += 1
        else:
            for r in _split_recipients(ev.get("to")):
                m["_rcpt"][r] += 1

    # Délai de première réponse : appariement par sujet normalisé (pas d'ID de conversation
    # dans l'export -> INDICATIF). Pour chaque sujet : un reçu ouvre une attente, le premier
    # envoyé suivant (<= REPLY_MAX_DAYS) la clôture.
    by_subject = defaultdict(list)
    for ev in events:
        norm = normalize_subject(ev.get("subject")).lower()
        if norm:
            by_subject[norm].append(ev)
    for evs in by_subject.values():
        evs.sort(key=lambda e: e["dt"])
        pending = None
        for ev in evs:
            if ev["kind"] == "received":
                if pending is not None and (ev["dt"] - pending).days >= REPLY_MAX_DAYS:
                    month(ym_of(pending))["unanswered"] += 1
                    pending = None
                if pending is None:
                    pending = ev["dt"]
                    month(ym_of(pending))["threads"] += 1
            elif pending is not None:
                if (ev["dt"] - pending).days < REPLY_MAX_DAYS:
                    mm = month(ym_of(ev["dt"]))
                    mm["_delays_cal"].append((ev["dt"] - pending).total_seconds() / 60)
                    mm["_delays_biz"].append(business_minutes(pending, ev["dt"]))
                else:
                    month(ym_of(pending))["unanswered"] += 1
                pending = None
        if pending is not None:
            month(ym_of(pending))["unanswered"] += 1

    for ym, m in metas.items():
        cal, biz = sorted(m.pop("_delays_cal")), sorted(m.pop("_delays_biz"))
        buckets = [("≤ 30 min", 30), ("≤ 2 h", 120), ("≤ 4 h", 240), ("≤ 1 j ouvré", BIZ_DAY_MIN)]
        dist = []
        prev = -1
        for label, lim in buckets:
            dist.append({"label": label, "count": sum(1 for v in biz if prev < v <= lim or (prev < 0 and v <= lim))})
            prev = lim
        dist.append({"label": "> 1 j ouvré", "count": sum(1 for v in biz if v > BIZ_DAY_MIN)})
        m["response"] = {
            "replied": len(biz),
            "median_biz_min": _round(_percentile(biz, 0.5)),
            "p90_biz_min": _round(_percentile(biz, 0.9)),
            "median_cal_min": _round(_percentile(cal, 0.5)),
            "p90_cal_min": _round(_percentile(cal, 0.9)),
            "distribution": dist,
            "match_rate": round(len(biz) / m["threads"] * 100, 1) if m["threads"] else None,
            "business_hours": f"{BIZ_START_H}h-{BIZ_END_H}h, lun-ven",
            "note": "Indicatif : appariement par sujet (pas d'ID de conversation dans l'export).",
        }
        subj, labels = m.pop("_subjects"), m.pop("_subject_label")
        m["top_subjects"] = [{"subject": labels[k], "count": c} for k, c in subj.most_common(10)]
        m["top_words"] = [{"word": w, "count": c} for w, c in m.pop("_words").most_common(15)]
        m["top_recipients"] = [{"name": n, "count": c} for n, c in m.pop("_rcpt").most_common(10)]
        m["pct_unread"] = round(m["unread"] / m["received"] * 100, 1) if m["received"] else 0
    return metas


# ---------------------------------------------------------------------------
# JOURNAL D'APPELS
# ---------------------------------------------------------------------------
def _strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def normalize_call_type(raw):
    t = _strip_accents(str(raw or "")).strip().lower()
    if t.startswith(("entrant", "incoming", "recu")):
        return "in"
    if t.startswith(("manqu", "missed")):
        return "missed"
    if t.startswith(("sortant", "outgoing", "emis")):
        return "out"
    if t.startswith(("bloqu", "blocked")):
        return "blocked"
    if t.startswith(("refus", "reject", "declin")):
        return "rejected"
    return None


def _parse_dt(s):
    s = str(s or "").strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _parse_hms(s):
    m = re.match(r"^\s*(\d+):(\d{2}):(\d{2})\s*$", str(s or ""))
    return int(m[1]) * 3600 + int(m[2]) * 60 + int(m[3]) if m else 0


def mask_phone(phone):
    digits = re.sub(r"\D", "", str(phone or ""))
    return f"•••• {digits[-3:]}" if len(digits) >= 3 else "•••• ???"


def phone_key(phone):
    d = re.sub(r"\D", "", str(phone or ""))
    return d[-9:] if d else ""


def parse_calllog_csv(raw_bytes):
    """CSV du journal d'appels (colonnes : Name, Phone, Date, Type, Duration(HH:MM:SS),
    Duration(secs), SIM). Retourne (calls, skipped)."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            text = raw_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    first = text.splitlines()[0] if text.strip() else ""
    delim = ";" if first.count(";") > first.count(",") else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    if not reader.fieldnames or "Date" not in reader.fieldnames or "Type" not in reader.fieldnames:
        raise ValueError("Colonnes attendues : Name, Phone, Date, Type, Duration(secs)… (export du journal d'appels)")
    calls, skipped = [], 0
    for r in reader:
        dt, ctype = _parse_dt(r.get("Date")), normalize_call_type(r.get("Type"))
        if not dt or not ctype:
            skipped += 1
            continue
        try:
            dur = int(float(r.get("Duration(secs)") or 0))
        except ValueError:
            dur = _parse_hms(r.get("Duration(HH:MM:SS)"))
        calls.append({"dt": dt, "type": ctype, "duration": dur,
                      "name": (r.get("Name") or "").strip(), "phone": (r.get("Phone") or "").strip()})
    return calls, skipped


def aggregate_calls(calls):
    """Retourne {ym: {'days': {'AAAA-MM-JJ': {calls, duration_sec}}, 'meta': {...}}}.
    Clé jour en pleine date (AAAA-MM-JJ), alignée sur le format de la saisie manuelle Tarkhiss."""
    by_month = defaultdict(list)
    for c in calls:
        by_month[ym_of(c["dt"])].append(c)

    # Index par numéro pour l'analyse des rappels (tous mois confondus)
    by_phone = defaultdict(list)
    for c in sorted(calls, key=lambda x: x["dt"]):
        k = phone_key(c["phone"])
        if k:
            by_phone[k].append(c)

    out = {}
    for ym, items in by_month.items():
        days = defaultdict(lambda: {"calls": 0, "duration_sec": 0})
        hourly = {"in": _empty_matrix(), "missed": _empty_matrix(), "out": _empty_matrix()}
        weekday = {"in": [0] * 7, "missed": [0] * 7, "out": [0] * 7}
        counts = Counter()
        in_durs = []
        callers = defaultdict(lambda: {"label": "", "named": False, "calls": 0, "duration_sec": 0, "missed": 0})
        for c in items:
            t, dt = c["type"], c["dt"]
            counts[t] += 1
            d = days[dt.strftime("%Y-%m-%d")]  # clé pleine date — alignée sur la saisie manuelle
            if t in hourly:
                hourly[t][dt.weekday()][dt.hour] += 1
                weekday[t][dt.weekday()] += 1
            if t == "in":
                d["calls"] += 1
                d["duration_sec"] += c["duration"]
                if c["duration"] > 0:
                    in_durs.append(c["duration"])
                else:
                    counts["in_zero"] += 1
            k = phone_key(c["phone"]) or c["name"] or "?"
            if t in ("in", "missed"):
                e = callers[k]
                if not e["label"]:
                    e["named"] = bool(c["name"])
                    e["label"] = c["name"] or mask_phone(c["phone"])
                e["missed" if t == "missed" else "calls"] += 1
                if t == "in":
                    e["duration_sec"] += c["duration"]

        # Rappel des appels manqués : premier appel sortant (par nous) ou entrant (le client
        # rappelle) sur le même numéro dans les 24 h.
        cb_out = cb_in = no_cb = 0
        cb_delays = []
        for c in items:
            if c["type"] != "missed":
                continue
            k = phone_key(c["phone"])
            nxt = next((x for x in by_phone.get(k, [])
                        if x["dt"] > c["dt"] and x["dt"] - c["dt"] <= timedelta(hours=24)
                        and x["type"] in ("out", "in")), None) if k else None
            if nxt is None:
                no_cb += 1
            elif nxt["type"] == "out":
                cb_out += 1
                cb_delays.append((nxt["dt"] - c["dt"]).total_seconds() / 60)
            else:
                cb_in += 1

        durs = sorted(in_durs)
        buckets = [("< 30 s", 30), ("30 s – 2 min", 120), ("2 – 5 min", 300), ("5 – 10 min", 600)]
        dist, prev = [], 0
        for label, lim in buckets:
            dist.append({"label": label, "count": sum(1 for v in durs if prev <= v < lim)})
            prev = lim
        dist.append({"label": "> 10 min", "count": sum(1 for v in durs if v >= 600)})

        answered, missed = counts["in"], counts["missed"]
        miss_by_hour = Counter()
        for c in items:
            if c["type"] == "missed":
                miss_by_hour[c["dt"].hour] += 1
        top = sorted(callers.values(), key=lambda e: (-e["calls"], -e["missed"]))[:10]
        out[ym] = {
            "days": dict(days),
            "meta": {
                "ym": ym, "hourly": hourly, "weekday": weekday,
                "counts": {"in": answered, "missed": missed, "out": counts["out"],
                           "blocked": counts["blocked"], "rejected": counts["rejected"],
                           "in_zero": counts["in_zero"]},
                "answer_rate": round(answered / (answered + missed) * 100, 1) if (answered + missed) else None,
                "duration": {
                    "total_sec": sum(durs), "avg_sec": _round(statistics.mean(durs)) if durs else None,
                    "median_sec": _round(_percentile(durs, 0.5)), "p90_sec": _round(_percentile(durs, 0.9)),
                    "max_sec": max(durs) if durs else None, "distribution": dist,
                },
                "callback": {"missed_total": missed, "called_back_by_us": cb_out,
                             "client_called_again": cb_in, "no_callback_24h": no_cb,
                             "median_callback_min": _round(_percentile(sorted(cb_delays), 0.5)),
                             "note": "Rappel = appel sortant (ou nouvel appel entrant) sur le même numéro sous 24 h."},
                "peak_missed_hour": (miss_by_hour.most_common(1)[0][0] if miss_by_hour else None),
                "unique_callers": len(callers),
                "recurrent_callers": sum(1 for e in callers.values() if e["calls"] >= 5),
                "top_callers": top,
            },
        }
    return out
