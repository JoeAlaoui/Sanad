# -*- coding: utf-8 -*-
"""
glpi_profiles.py — Profils d'import GLPI (Lot C).

Chaque profil décrit comment obtenir et transformer des données GLPI (via glpi_client.py) en
enregistrements normalisés compatibles avec les fonctions de stockage déjà existantes dans
glpi_import.py (store_ticket_series, store_source_items). Ajouter une nouvelle section GLPI
(groupes, SLA, etc.) = ajouter une entrée dans PROFILES + une fonction d'agrégation dédiée si
besoin — jamais une réécriture du moteur d'import ou des routes Flask.

IMPORTANT — IDs de champs GLPI : les identifiants numériques ci-dessous (TICKET_FIELDS)
correspondent à une installation GLPI 11.x standard, non personnalisée. Si des champs custom ou
une configuration spécifique existent chez vous, ces IDs doivent être vérifiés (Configuration >
Listes > "Rechercher" sur l'itemtype Ticket affiche les IDs réels) avant mise en production —
c'est le seul point qui nécessite une validation manuelle contre votre GLPI réel, le reste de
l'architecture est indépendant de cette numérotation.
"""
import re
from datetime import datetime

# ---------------------------------------------------------------------------
# Référentiel des champs de recherche GLPI standards pour l'itemtype Ticket
# ---------------------------------------------------------------------------
TICKET_FIELDS = {
    "id": 2,
    "name": 1,
    "status": 12,
    "date_creation": 15,
    "date_mod": 19,
    "closedate": 17,
    "solvedate": 18,
    "category": 7,
    "priority": 3,
    "type": 14,
    "requester": 4,        # "Demandeur - Demandeur" dans les exports bruts GLPI
    "requester_group": 71,
    "assigned_tech": 5,    # "Attribué à - Technicien"
    "time_to_resolve": 82,
}

# Statuts GLPI standards (id -> libellé). 5=Résolu, 6=Clos sont considérés "terminés".
GLPI_STATUS_LABELS = {
    1: "Nouveau", 2: "En cours (Attribué)", 3: "En cours (Planifié)",
    4: "En attente", 5: "Résolu", 6: "Clos",
}
GLPI_RESOLVED_STATUSES = {5, 6}
GLPI_CLOSED_STATUSES = {6}
GLPI_OPEN_STATUSES = {1, 2, 3, 4}


def glpi_status_label(status_id):
    try:
        return GLPI_STATUS_LABELS.get(int(status_id), f"Statut {status_id}")
    except (TypeError, ValueError):
        return "Statut inconnu"


# ---------------------------------------------------------------------------
# Profils — la config déclarative. "target" indique la fonction de stockage de glpi_import.py
# à appeler : "ticket_series" -> store_ticket_series ; "source_items:<kind>" -> store_source_items.
# ---------------------------------------------------------------------------
PROFILES = {
    "tickets": {
        "label": "Tickets (série mensuelle)",
        "itemtype": "Ticket",
        "date_field": "date_creation",
        "forcedisplay": [TICKET_FIELDS["id"], TICKET_FIELDS["status"], TICKET_FIELDS["date_creation"],
                          TICKET_FIELDS["closedate"], TICKET_FIELDS["solvedate"]],
        "target": "ticket_series",
        "csv_equivalent_headers": ["mois", "ouverts", "resolus", "en_retard", "clos"],
        "aggregator": "aggregate_tickets_by_month",
    },
    "categories": {
        "label": "Catégories de tickets",
        "itemtype": "Ticket",
        "date_field": "date_creation",
        "group_field": "category",
        "forcedisplay": [TICKET_FIELDS["id"], TICKET_FIELDS["status"], TICKET_FIELDS["date_creation"],
                          TICKET_FIELDS["category"]],
        "target": "source_items:categories",
        "csv_equivalent_headers": ["categorie", "ouverts", "resolus", "en_retard", "fermes"],
        "aggregator": "aggregate_tickets_by_label",
    },
    "demandeurs": {
        "label": "Demandeurs",
        "itemtype": "Ticket",
        "date_field": "date_creation",
        "group_field": "requester",
        "forcedisplay": [TICKET_FIELDS["id"], TICKET_FIELDS["status"], TICKET_FIELDS["date_creation"],
                          TICKET_FIELDS["requester"]],
        "target": "source_items:demandeurs",
        "csv_equivalent_headers": ["demandeur", "ouverts", "resolus", "en_retard", "fermes"],
        "aggregator": "aggregate_tickets_by_label",
    },
    "techniciens": {
        "label": "Statistiques techniciens (mensuel)",
        "itemtype": "Ticket",
        "date_field": "date_creation",
        "group_field": "assigned_tech",
        "forcedisplay": [TICKET_FIELDS["id"], TICKET_FIELDS["status"], TICKET_FIELDS["date_creation"],
                          TICKET_FIELDS["assigned_tech"]],
        "target": "source_items:techniciens",
        "csv_equivalent_headers": ["technicien", "ouverts", "resolus", "en_retard", "fermes"],
        "aggregator": "aggregate_tickets_by_label",
        # Seul profil soumis au filtre exclusion/whitelist (apply_technicien_filter côté
        # app.py) — appliqué après agrégation, à l'import ET à chaque lecture pour rapport.
        "filterable": True,
    },
}


def list_profiles():
    """Pour le panneau Admin : liste des profils disponibles avec leur libellé."""
    return [{"key": k, "label": v["label"], "itemtype": v["itemtype"]} for k, v in PROFILES.items()]


# ---------------------------------------------------------------------------
# Normalisation des résultats bruts de l'API GLPI (dicts indexés par ID de champ, ex: "12")
# ---------------------------------------------------------------------------
def _field(raw_ticket, field_id):
    return raw_ticket.get(str(field_id))


def _parse_glpi_date(raw):
    if not raw:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", str(raw))
    if not m:
        return None
    try:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def aggregate_tickets_by_month(raw_tickets, profile):
    """Agrège une liste de tickets bruts GLPI (format search : dict indexé par ID de champ) en
    série mensuelle {"mois": "YYYY-MM", "ouverts", "resolus", "en_retard", "clos"} — équivalent
    du contenu d'un export CSV 'tickets' consolidé par mois.

    Règle "en_retard" : ticket encore ouvert (statut 1-4) dont la date de création dépasse 30
    jours — seuil par défaut documenté ici, à ajuster si un SLA différent doit être appliqué
    (aucun champ SLA n'est actuellement mappé dans TICKET_FIELDS)."""
    by_month = {}
    status_f = str(TICKET_FIELDS["status"])
    date_f = str(TICKET_FIELDS[profile["date_field"]])
    close_f = str(TICKET_FIELDS["closedate"])
    now = datetime.now()

    for t in raw_tickets:
        d = _parse_glpi_date(t.get(date_f))
        if not d:
            continue
        month = f"{d.year}-{d.month:02d}"
        by_month.setdefault(month, {"mois": month, "ouverts": 0, "resolus": 0, "en_retard": 0, "clos": 0})
        row = by_month[month]
        row["ouverts"] += 1
        try:
            status = int(t.get(status_f) or 0)
        except (TypeError, ValueError):
            status = 0
        if status in GLPI_RESOLVED_STATUSES:
            row["resolus"] += 1
        if status in GLPI_CLOSED_STATUSES:
            row["clos"] += 1
        if status in GLPI_OPEN_STATUSES and (now - d).days > 30:
            row["en_retard"] += 1

    return list(by_month.values())


def aggregate_tickets_by_label(raw_tickets, profile):
    """Agrège une liste de tickets bruts GLPI par étiquette (catégorie ou demandeur, selon
    profile["group_field"]) en {"label", "ouverts", "resolus", "en_retard", "fermes"} —
    équivalent du contenu d'un export CSV 'categories'/'demandeurs'."""
    group_field_id = str(TICKET_FIELDS[profile["group_field"]])
    status_f = str(TICKET_FIELDS["status"])
    date_f = str(TICKET_FIELDS[profile["date_field"]])
    now = datetime.now()
    by_label = {}

    for t in raw_tickets:
        label = (t.get(group_field_id) or "Non catégorisé").strip() or "Non catégorisé"
        by_label.setdefault(label, {"label": label, "ouverts": 0, "resolus": 0, "en_retard": 0, "fermes": 0})
        row = by_label[label]
        row["ouverts"] += 1
        try:
            status = int(t.get(status_f) or 0)
        except (TypeError, ValueError):
            status = 0
        if status in GLPI_RESOLVED_STATUSES:
            row["resolus"] += 1
        if status in GLPI_CLOSED_STATUSES:
            row["fermes"] += 1
        d = _parse_glpi_date(t.get(date_f))
        if status in GLPI_OPEN_STATUSES and d and (now - d).days > 30:
            row["en_retard"] += 1

    return list(by_label.values())


AGGREGATORS = {
    "aggregate_tickets_by_month": aggregate_tickets_by_month,
    "aggregate_tickets_by_label": aggregate_tickets_by_label,
}


def run_profile(client, profile_key, date_start=None, date_end=None):
    """Exécute un profil : recherche GLPI (avec filtre de date optionnel) puis agrégation.
    Retourne une liste d'enregistrements normalisés, prête pour store_ticket_series() ou
    store_source_items() selon profile["target"]. `client` est un GLPIClient déjà connecté
    (voir glpi_client.py)."""
    profile = PROFILES[profile_key]
    criteria = []
    if date_start:
        criteria.append({"field": TICKET_FIELDS[profile["date_field"]], "searchtype": "morethan", "value": date_start})
    if date_end:
        criteria.append({"field": TICKET_FIELDS[profile["date_field"]], "searchtype": "lessthan", "value": date_end})

    raw_tickets = client.get_all(profile["itemtype"], criteria=criteria, forcedisplay=profile["forcedisplay"])
    aggregator = AGGREGATORS[profile["aggregator"]]
    return aggregator(raw_tickets, profile), len(raw_tickets)
