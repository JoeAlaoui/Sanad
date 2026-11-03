# -*- coding: utf-8 -*-
"""
retards.py — Tickets en retard (Moussanada), calculés à partir des tickets bruts importés.

Un ticket est « en retard » quand il n'est ni résolu ni clos ET que son échéance (colonne TTR de
l'export, « temps de résolution » = date limite) est dépassée. Un ticket sans échéance renseignée
n'est pas compté en retard (il est signalé à part). Le TOP de l'email peut écarter des techniciens
(réglage `late_tickets.excluded_technicians`) ; l'Excel (joint ou téléchargé) les écarte aussi (`shown`).
"""
import unicodedata
from datetime import datetime

from glpi_import import parse_glpi_datetime

DEFAULT_TOP_N = 10
TOP_N_CHOICES = (5, 10, 15, 20)
UNASSIGNED = "Non assigné"
_CLOSED = {"resolu", "clos", "closed", "solved"}


def _norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


def is_open(ticket):
    status = _norm(ticket.get("statut"))
    return bool(status) and status not in _CLOSED


def technician(ticket):
    return (ticket.get("technicien") or "").strip() or UNASSIGNED


def settings_of(module_settings):
    """(top_n, techniciens exclus) lus dans les paramètres du volet, avec valeurs par défaut sûres."""
    cfg = (module_settings or {}).get("late_tickets") or {}
    try:
        top_n = int(cfg.get("top_n", DEFAULT_TOP_N))
    except (TypeError, ValueError):
        top_n = DEFAULT_TOP_N
    top_n = min(max(top_n, 1), 50)
    return top_n, [str(x) for x in (cfg.get("excluded_technicians") or []) if str(x).strip()]


def late_label(hours):
    """« 5 h » sous un jour, « 34 j » au-delà."""
    return f"{int(hours)} h" if hours < 24 else f"{int(hours // 24)} j"


def compute(tickets, now=None):
    """(en retard, sans échéance, nombre de tickets ouverts). Les retards sont triés du plus ancien au
    plus récent par échéance dépassée (le plus en retard d'abord)."""
    now = now or datetime.now()
    late, no_due, open_total = [], 0, 0
    for t in tickets:
        if not is_open(t):
            continue
        open_total += 1
        due = parse_glpi_datetime(t.get("ttr"))
        if not due:
            no_due += 1
            continue
        if due >= now:
            continue
        opened = parse_glpi_datetime(t.get("date_ouverture"))
        hours = (now - due).total_seconds() / 3600
        late.append({
            "id": t.get("id", ""), "titre": t.get("titre", ""), "statut": t.get("statut", ""),
            "priorite": t.get("priorite", ""), "technicien": technician(t), "demandeur": t.get("demandeur", ""),
            "categorie": t.get("categorie", ""), "type": t.get("type", ""),
            "date_ouverture": opened, "echeance": due,
            "retard_h": hours, "retard_j": int(hours // 24), "retard": late_label(hours),
        })
    late.sort(key=lambda x: (-x["retard_h"], str(x["id"])))
    return late, no_due, open_total


def stats(items):
    """(nombre, retard moyen en jours, retard maximal en jours, priorité haute) d'une liste de retards."""
    hours = [t["retard_h"] for t in items]
    high = sum(1 for t in items if _norm(t["priorite"]) in ("haute", "tres haute", "majeure"))
    return len(items), (sum(hours) / len(hours) / 24) if hours else 0, (max(hours) / 24) if hours else 0, high


def view(tickets, module_settings=None, now=None):
    """Tout ce dont le dashboard, l'email et l'Excel ont besoin, calculé une seule fois."""
    now = now or datetime.now()
    top_n, excluded = settings_of(module_settings)
    late, no_due, open_total = compute(tickets, now)
    skip = {_norm(x) for x in excluded}
    shown = [t for t in late if _norm(t["technicien"]) not in skip]
    by_tech = {}
    for t in late:
        by_tech.setdefault(t["technicien"], []).append(t)
    technicians = sorted(({"name": name, "late": len(items), "excluded": _norm(name) in skip}
                          for name, items in by_tech.items()), key=lambda x: (-x["late"], x["name"].lower()))
    total, avg_days, max_days, high = stats(late)
    return {
        "as_of": now, "top_n": top_n, "excluded": excluded, "late": late, "shown": shown, "total": len(late),
        "top": shown[:top_n], "hidden": len(late) - len(shown), "shown_total": len(shown),
        "no_due": no_due, "open_total": open_total, "technicians": technicians,
        "avg_days": avg_days, "max_days": max_days, "high": high,
    }


def to_json(v):
    """Version sérialisable de view() pour l'API du dashboard."""
    def row(t):
        return {"id": t["id"], "titre": t["titre"], "statut": t["statut"], "priorite": t["priorite"],
                "technicien": t["technicien"], "categorie": t["categorie"],
                "date_ouverture": t["date_ouverture"].strftime("%d/%m/%Y %H:%M") if t["date_ouverture"] else "",
                "echeance": t["echeance"].strftime("%d/%m/%Y %H:%M"), "retard": t["retard"], "retard_j": t["retard_j"]}
    skip = {_norm(x) for x in v["excluded"]}
    return {
        "as_of": v["as_of"].strftime("%d/%m/%Y %H:%M"), "top_n": v["top_n"], "top_n_choices": list(TOP_N_CHOICES),
        "excluded": v["excluded"], "total": v["total"], "hidden": v["hidden"], "no_due": v["no_due"],
        "open_total": v["open_total"], "technicians": v["technicians"],
        "avg_days": round(v["avg_days"], 1), "max_days": round(v["max_days"], 1), "high": v["high"],
        "top": [row(t) for t in v["top"]],
        "all": [dict(row(t), excluded=_norm(t["technicien"]) in skip) for t in v["late"]],
    }
