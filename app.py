# -*- coding: utf-8 -*-
"""
Helpdesk Dashboard - Application locale multi-volets (Tarkhiss, Moussanada...)
Stockage : fichiers JSON (data/<module>/). Aucune base de données.
Lancer : python app.py  -> http://127.0.0.1:5050
"""
import csv
import io
import json
import os
import platform
import re
import html as html_module
from glpi_import import (
    glpi_cell_to_seconds, decimal_hours_cell, format_dh, normalize_header,
    match_header, parse_glpi_row_sheet, import_moussanada_xlsx as _parse_moussanada_xlsx,
    import_moussanada_csv as _parse_moussanada_csv,
)
import base64
import subprocess
import zipfile
from datetime import datetime
from email.message import EmailMessage

from functools import wraps
from flask import Flask, jsonify, request, render_template, send_file, abort, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = os.path.join(BASE_DIR, "static", "uploads")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

GLOBAL_SETTINGS_FILE = os.path.join(DATA_DIR, "global_settings.json")

app = Flask(__name__)

SECRET_KEY_FILE = os.path.join(DATA_DIR, ".secret_key")
if os.path.exists(SECRET_KEY_FILE):
    with open(SECRET_KEY_FILE, "r") as f:
        app.secret_key = f.read().strip()
else:
    app.secret_key = os.urandom(32).hex()
    with open(SECRET_KEY_FILE, "w") as f:
        f.write(app.secret_key)

MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin",
           "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# ---------------------------------------------------------------------------
# Volets (modules) de l'application
# ---------------------------------------------------------------------------
MODULES = {
    "tarkhiss": {"label": "Tarkhiss", "subtitle": "Support Email & Hotline", "ready": True},
    "moussanada": {"label": "Moussanada", "subtitle": "Helpdesk GLPI", "ready": True},
}

# Catégories de la feuille "rawData - Ticket - X" : colonnes fixes GLPI
GLPI_ROW_FIELDS = [
    "ouverts", "resolus", "en_retard", "fermes",
    "enquetes_ouvertes", "enquetes_reponses", "satisfaction",
    "delai_prise_en_compte_s", "delai_resolution_s", "delai_fermeture_s",
    "duree_reelle_moy_s", "duree_reelle_s",
]

# ---------------------------------------------------------------------------
# AUTHENTIFICATION & RÔLES
# ---------------------------------------------------------------------------
USERS_FILE = os.path.join(DATA_DIR, "users.json")
ROLES = ("admin", "hotliner", "superviseur")


def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def public_user(u):
    return {"id": u["id"], "username": u["username"], "name": u.get("name", ""),
            "role": u["role"], "modules": u.get("modules", list(MODULES.keys())), "active": u.get("active", True)}


def find_user(username=None, user_id=None):
    users = load_users()
    for u in users:
        if username is not None and u["username"] == username:
            return u
        if user_id is not None and u["id"] == user_id:
            return u
    return None


def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    u = find_user(user_id=uid)
    if not u or not u.get("active", True):
        return None
    return u


@app.before_request
def auth_gate():
    path = request.path
    if path == "/" or path.startswith("/static/") or path.startswith("/api/auth/"):
        return
    if path == "/api/modules":
        return
    user = current_user()
    if not user:
        return jsonify({"error": "auth_required"}), 401
    role = user["role"]
    if role == "admin":
        return
    if path.startswith("/api/users"):
        return jsonify({"error": "forbidden"}), 403
    if role == "hotliner":
        if path.startswith("/api/notes") or path.startswith("/api/knowledge") or path == "/api/modules":
            return
        if request.method == "GET" and (path == "/api/global-settings" or re.match(r"^/api/[^/]+/settings$", path)):
            return
        return jsonify({"error": "forbidden"}), 403
    if role == "superviseur":
        if path.startswith("/api/notes"):
            return jsonify({"error": "forbidden"}), 403
        if request.method != "GET":
            return jsonify({"error": "forbidden"}), 403
        return
    return jsonify({"error": "forbidden"}), 403


@app.route("/api/auth/me")
def auth_me():
    users = load_users()
    if not users:
        return jsonify({"authenticated": False, "setup_required": True})
    user = current_user()
    if not user:
        return jsonify({"authenticated": False, "setup_required": False})
    return jsonify({"authenticated": True, "setup_required": False, "user": public_user(user)})


@app.route("/api/auth/setup", methods=["POST"])
def auth_setup():
    users = load_users()
    if users:
        return jsonify({"ok": False, "error": "Un compte existe déjà"}), 400
    payload = request.json or {}
    username = (payload.get("username") or "").strip()
    password = payload.get("password") or ""
    name = (payload.get("name") or "").strip()
    if not username or len(password) < 4:
        return jsonify({"ok": False, "error": "Identifiant requis, mot de passe 4 caractères min."}), 400
    user = {"id": 1, "username": username, "password_hash": generate_password_hash(password),
            "name": name or username, "role": "admin", "modules": list(MODULES.keys()), "active": True}
    save_users([user])
    session["user_id"] = user["id"]
    return jsonify({"ok": True, "user": public_user(user)})


@app.route("/api/auth/login", methods=["POST"])
def auth_login():
    payload = request.json or {}
    u = find_user(username=(payload.get("username") or "").strip())
    if not u or not u.get("active", True) or not check_password_hash(u["password_hash"], payload.get("password") or ""):
        return jsonify({"ok": False, "error": "Identifiants incorrects"}), 401
    session["user_id"] = u["id"]
    return jsonify({"ok": True, "user": public_user(u)})


@app.route("/api/auth/logout", methods=["POST"])
def auth_logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/auth/change-password", methods=["POST"])
def auth_change_password():
    user = current_user()
    payload = request.json or {}
    if not check_password_hash(user["password_hash"], payload.get("old_password") or ""):
        return jsonify({"ok": False, "error": "Ancien mot de passe incorrect"}), 400
    new_pw = payload.get("new_password") or ""
    if len(new_pw) < 4:
        return jsonify({"ok": False, "error": "Mot de passe trop court (4 caractères min.)"}), 400
    users = load_users()
    for u in users:
        if u["id"] == user["id"]:
            u["password_hash"] = generate_password_hash(new_pw)
    save_users(users)
    return jsonify({"ok": True})


# --------- Gestion des utilisateurs (admin) ---------
@app.route("/api/users", methods=["GET"])
def list_users():
    return jsonify([public_user(u) for u in load_users()])


@app.route("/api/users", methods=["POST"])
def create_user():
    payload = request.json or {}
    username = (payload.get("username") or "").strip()
    password = payload.get("password") or ""
    role = payload.get("role")
    if role not in ROLES:
        return jsonify({"ok": False, "error": "Rôle invalide"}), 400
    if not username or len(password) < 4:
        return jsonify({"ok": False, "error": "Identifiant requis, mot de passe 4 caractères min."}), 400
    users = load_users()
    if any(u["username"] == username for u in users):
        return jsonify({"ok": False, "error": "Identifiant déjà utilisé"}), 400
    new_id = max([u["id"] for u in users], default=0) + 1
    user = {"id": new_id, "username": username, "password_hash": generate_password_hash(password),
            "name": (payload.get("name") or "").strip() or username, "role": role,
            "modules": payload.get("modules") or list(MODULES.keys()), "active": True}
    users.append(user)
    save_users(users)
    return jsonify({"ok": True, "user": public_user(user)})


@app.route("/api/users/<int:uid>", methods=["POST"])
def update_user(uid):
    payload = request.json or {}
    users = load_users()
    found = None
    for u in users:
        if u["id"] == uid:
            found = u
            if "name" in payload:
                u["name"] = payload["name"]
            if "role" in payload and payload["role"] in ROLES:
                u["role"] = payload["role"]
            if "modules" in payload:
                u["modules"] = payload["modules"]
            if "active" in payload:
                u["active"] = bool(payload["active"])
            if payload.get("new_password"):
                u["password_hash"] = generate_password_hash(payload["new_password"])
    if not found:
        return jsonify({"ok": False, "error": "Utilisateur introuvable"}), 404
    save_users(users)
    return jsonify({"ok": True, "user": public_user(found)})


@app.route("/api/users/<int:uid>", methods=["DELETE"])
def delete_user(uid):
    users = load_users()
    users = [u for u in users if u["id"] != uid]
    save_users(users)
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# NOTES (échange Hotliner <-> Admin)
# ---------------------------------------------------------------------------
NOTES_DIR = os.path.join(DATA_DIR, "notes")
os.makedirs(NOTES_DIR, exist_ok=True)


def notes_thread_path(hotliner_id):
    return os.path.join(NOTES_DIR, f"thread_{hotliner_id}.json")


def load_notes_thread(hotliner_id):
    p = notes_thread_path(hotliner_id)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_notes_thread(hotliner_id, thread):
    with open(notes_thread_path(hotliner_id), "w", encoding="utf-8") as f:
        json.dump(thread, f, ensure_ascii=False, indent=2)


def append_note(hotliner_id, sender_role, text):
    thread = load_notes_thread(hotliner_id)
    thread.append({"id": len(thread) + 1, "from": sender_role, "text": text,
                    "date": datetime.now().isoformat(timespec="seconds")})
    save_notes_thread(hotliner_id, thread)
    return thread


@app.route("/api/notes/mine", methods=["GET"])
def notes_mine():
    user = current_user()
    return jsonify(load_notes_thread(user["id"]))


@app.route("/api/notes/mine", methods=["POST"])
def notes_mine_post():
    user = current_user()
    text = (request.json or {}).get("text", "").strip()
    if not text:
        return jsonify({"ok": False, "error": "Message vide"}), 400
    thread = append_note(user["id"], "hotliner", text)
    return jsonify({"ok": True, "thread": thread})


@app.route("/api/notes/hotliners", methods=["GET"])
def notes_hotliners():
    users = [u for u in load_users() if u["role"] == "hotliner"]
    return jsonify([public_user(u) for u in users])


@app.route("/api/notes/user/<int:uid>", methods=["GET"])
def notes_user_get(uid):
    return jsonify(load_notes_thread(uid))


@app.route("/api/notes/user/<int:uid>", methods=["POST"])
def notes_user_post(uid):
    text = (request.json or {}).get("text", "").strip()
    if not text:
        return jsonify({"ok": False, "error": "Message vide"}), 400
    thread = append_note(uid, "admin", text)
    return jsonify({"ok": True, "thread": thread})


# ---------------------------------------------------------------------------
# BASE DE CONNAISSANCES
# ---------------------------------------------------------------------------
KNOWLEDGE_FILE = os.path.join(DATA_DIR, "knowledge.json")


def load_knowledge():
    if os.path.exists(KNOWLEDGE_FILE):
        with open(KNOWLEDGE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_knowledge(items):
    with open(KNOWLEDGE_FILE, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


@app.route("/api/knowledge")
def knowledge_list():
    module = request.args.get("module")
    items = load_knowledge()
    if module:
        items = [i for i in items if i.get("module") == module]
    return jsonify(items)


@app.route("/api/knowledge", methods=["POST"])
def knowledge_upsert():
    payload = request.json or {}
    items = load_knowledge()
    if payload.get("id"):
        for it in items:
            if it["id"] == payload["id"]:
                it.update({"title": payload.get("title", it["title"]),
                            "content": payload.get("content", it["content"]),
                            "module": payload.get("module", it["module"]),
                            "updated": datetime.now().isoformat(timespec="seconds")})
                save_knowledge(items)
                return jsonify({"ok": True, "item": it})
        return jsonify({"ok": False, "error": "Article introuvable"}), 404
    new_id = max([i["id"] for i in items], default=0) + 1
    item = {"id": new_id, "module": payload.get("module", "tarkhiss"), "title": payload.get("title", ""),
            "content": payload.get("content", ""), "updated": datetime.now().isoformat(timespec="seconds")}
    items.append(item)
    save_knowledge(items)
    return jsonify({"ok": True, "item": item})


@app.route("/api/knowledge/<int:kid>", methods=["DELETE"])
def knowledge_delete(kid):
    items = [i for i in load_knowledge() if i["id"] != kid]
    save_knowledge(items)
    return jsonify({"ok": True})


def check_module(module):
    if module not in MODULES:
        abort(404, description="Volet inconnu")


def module_dir(module):
    d = os.path.join(DATA_DIR, module)
    os.makedirs(os.path.join(d, "exports"), exist_ok=True)
    return d


# ---------------------------------------------------------------------------
# Utilitaires stockage JSON (par volet)
# ---------------------------------------------------------------------------
def _path(module, kind, ym):
    return os.path.join(module_dir(module), f"{kind}_{ym}.json")


def _load(module, kind, ym, default):
    p = _path(module, kind, ym)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return default


def _save(module, kind, ym, payload):
    with open(_path(module, kind, ym), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def month_label_fr(ym):
    y, m = ym.split("-")
    return f"{MOIS_FR[int(m) - 1].capitalize()} {y}"


def prev_ym(ym):
    y, m = ym.split("-")
    y, m = int(y), int(m)
    if m == 1:
        return f"{y-1}-12"
    return f"{y}-{m-1:02d}"


def duration_seconds(rec):
    """Compatibilité : ancien format duration_min, nouveau duration_sec."""
    if rec.get("duration_sec") is not None:
        return rec.get("duration_sec") or 0
    if rec.get("duration_min") is not None:
        return (rec.get("duration_min") or 0) * 60
    return 0


def format_hms(total_seconds):
    total_seconds = int(total_seconds or 0)
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


# ---------------------------------------------------------------------------
# Paramètres globaux (agence, logo) - communs à tous les volets
# ---------------------------------------------------------------------------
def load_global_settings():
    defaults = {
        "agency_name": "AMMPS — DSID",
        "logo_filename": None,          # logo AMMPS
        "sanad_logo_filename": None,    # logo SANAD (identité plateforme)
        "app_name": "SANAD",
        "app_subtitle": "Plateforme de pilotage du Helpdesk SI",
    }
    if os.path.exists(GLOBAL_SETTINGS_FILE):
        with open(GLOBAL_SETTINGS_FILE, "r", encoding="utf-8") as f:
            defaults.update(json.load(f))
    return defaults


def save_global_settings(payload):
    with open(GLOBAL_SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Paramètres par volet (contacts, salutation, signature)
# ---------------------------------------------------------------------------
def module_settings_file(module):
    return os.path.join(module_dir(module), "settings.json")


def load_module_settings(module):
    defaults = {
        "contacts": [],           # [{email, name, role: "to"|"cc"}]
        "greeting": "Bonjour,",
        "signature_name": "",
        "signature_function": "",
        "signature_phone": "",
        "alert_thresholds": {"resolution_rate_min": None, "backlog_max": None, "delay_max_h": None, "volume_variation_max": None},
        "reminder_day": 5,
    }
    p = module_settings_file(module)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        defaults.update(data)
        if data.get("recipients") and not data.get("contacts"):
            defaults["contacts"] = [{"email": e, "name": "", "role": "to"} for e in data["recipients"]]
    return defaults


def save_module_settings(module, payload):
    with open(module_settings_file(module), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def send_log_file(module):
    return os.path.join(module_dir(module), "send_log.json")


def alert_history_file(module):
    return os.path.join(module_dir(module), "alert_history.json")


def load_alert_history(module):
    p = alert_history_file(module)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_alert_history(module, history):
    with open(alert_history_file(module), "w", encoding="utf-8") as f:
        json.dump(history[:300], f, ensure_ascii=False, indent=2)


@app.route("/api/<module>/alert-history", methods=["GET"])
def get_alert_history(module):
    check_module(module)
    return jsonify(load_alert_history(module))


@app.route("/api/<module>/alert-history", methods=["POST"])
def post_alert_history(module):
    check_module(module)
    payload = request.json or {}
    ym = payload.get("ym")
    alerts = payload.get("alerts") or []
    if not ym or not alerts:
        return jsonify({"ok": True, "logged": False})
    today = datetime.now().strftime("%Y-%m-%d")
    history = load_alert_history(module)
    already = any(h["ym"] == ym and h["date"] == today for h in history)
    if not already:
        history.insert(0, {"date": today, "ym": ym, "alerts": alerts})
        save_alert_history(module, history)
    return jsonify({"ok": True, "logged": not already})


def load_send_log(module):
    p = send_log_file(module)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def append_send_log(module, entry):
    log = load_send_log(module)
    log.insert(0, entry)
    with open(send_log_file(module), "w", encoding="utf-8") as f:
        json.dump(log[:200], f, ensure_ascii=False, indent=2)


def empty_analysis(ym):
    return {
        "month_label": month_label_fr(ym),
        "problems": [], "demandes": [], "weekly": [],
        "constats": [], "recommandations": [],
        "keywords": [], "response_delay": "", "internal_notes": ""
    }


# ---------------------------------------------------------------------------
# MOUSSANADA — stockage, import GLPI, KPI, rapports
# ---------------------------------------------------------------------------
def timeseries_file():
    return os.path.join(module_dir("moussanada"), "timeseries.json")


def load_timeseries():
    p = timeseries_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_timeseries(data):
    with open(timeseries_file(), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def sources_path(ym):
    return os.path.join(module_dir("moussanada"), f"sources_{ym}.json")


def load_sources(ym):
    p = sources_path(ym)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"categories": [], "services": [], "techniciens": [], "demandeurs": []}


def save_sources(ym, data):
    with open(sources_path(ym), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def notes_path(ym):
    return os.path.join(module_dir("moussanada"), f"notes_{ym}.json")


MOUSSANADA_SECTIONS = [
    ("volume", "Volume global des tickets"),
    ("delays", "Évolution des délais de traitement"),
    ("categories", "Principales catégories de demandes"),
    ("services", "Services les plus demandeurs"),
    ("team", "Charge et mobilisation de l'équipe SI"),
    ("comparative", "Analyse comparative M-1 / M"),
    ("synthesis", "Synthèse finale et conclusion"),
]


def empty_moussanada_notes():
    return {
        "sections": {key: "" for key, _ in MOUSSANADA_SECTIONS},
        "constats": [], "recommandations": [],
    }


def load_notes(ym):
    p = notes_path(ym)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        base = empty_moussanada_notes()
        sections = data.pop("sections", None)
        base.update(data)
        if isinstance(sections, dict):
            base["sections"].update(sections)
        return base
    return empty_moussanada_notes()


def save_notes(ym, data):
    with open(notes_path(ym), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@app.route("/api/moussanada/import/<ym>", methods=["POST"])
def moussanada_import(ym):
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "Aucun fichier reçu"}), 400
    kind = request.form.get("kind", "auto")
    ext = os.path.splitext(f.filename)[1].lower()
    ts = load_timeseries()
    src = load_sources(ym)
    try:
        if ext in (".xlsx", ".xlsm"):
            result, ts, src = _parse_moussanada_xlsx(f, ts, src)
        elif ext == ".csv":
            if kind == "auto":
                return jsonify({"ok": False, "error": "Précisez le type de données pour un CSV"}), 400
            result, ts, src = _parse_moussanada_csv(f, kind, ts, src)
        else:
            return jsonify({"ok": False, "error": "Format non supporté (.xlsx, .xlsm, .csv)"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erreur d'analyse du fichier : {e}"}), 400
    save_timeseries(ts)
    save_sources(ym, src)
    return jsonify({"ok": True, "result": result})


@app.route("/api/moussanada/data/<ym>")
def moussanada_data(ym):
    ts = load_timeseries()
    src = load_sources(ym)
    notes = load_notes(ym)
    month_row = ts.get(ym, {})
    return jsonify({
        "month": {
            "tickets": {
                "ouverts": month_row.get("ouverts", 0),
                "resolus": month_row.get("resolus", 0),
                "en_retard": month_row.get("en_retard", 0),
                "clos": month_row.get("clos", 0),
            },
            "durations": {
                "cloture_h": month_row.get("cloture_h", 0),
                "resolution_h": month_row.get("resolution_h", 0),
                "duree_reelle_h": month_row.get("duree_reelle_h", 0),
            },
            "categories": src.get("categories", []),
            "services": src.get("services", []),
            "techniciens": src.get("techniciens", []),
            "demandeurs": src.get("demandeurs", []),
        },
        "timeseries": ts,
        "notes": notes,
    })


@app.route("/api/moussanada/notes/<ym>", methods=["POST"])
def moussanada_save_notes(ym):
    payload = request.json or {}
    current = load_notes(ym)
    current.update(payload)
    save_notes(ym, current)
    return jsonify({"ok": True})


@app.route("/api/moussanada/export-csv-source/<kind>/<ym>")
def moussanada_export_csv_source(kind, ym):
    if kind not in ("categories", "services", "techniciens", "demandeurs"):
        return jsonify({"ok": False, "error": "kind invalide"}), 400
    src = load_sources(ym)
    items = src.get(kind, [])
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["label", "ouverts", "resolus", "en_retard", "fermes"])
    for it in items:
        writer.writerow([it.get("label", ""), it.get("ouverts", 0), it.get("resolus", 0),
                          it.get("en_retard", 0), it.get("fermes", 0)])
    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(mem, as_attachment=True, download_name=f"moussanada_{kind}_{ym}.csv", mimetype="text/csv")


def write_alerts_sheet(wb, module, header_fill, header_font):
    ws = wb.create_sheet("Alertes")
    for i, h in enumerate(["Date", "Mois", "Alerte(s) déclenchée(s)"]):
        c = ws.cell(row=1, column=i + 1, value=h)
        c.fill = header_fill
        c.font = header_font
    row = 2
    for entry in load_alert_history(module):
        for alert in entry.get("alerts", []):
            ws.cell(row=row, column=1, value=entry.get("date", ""))
            ws.cell(row=row, column=2, value=month_label_fr(entry.get("ym", "")) if entry.get("ym") else "")
            ws.cell(row=row, column=3, value=alert)
            row += 1
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 70


def export_xlsx_moussanada(ym):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    data = moussanada_data(ym).get_json()
    m = data["month"]
    ts = data["timeseries"]
    month_label = month_label_fr(ym)

    wb = Workbook()
    ws = wb.active
    ws.title = "KPI"
    header_fill = PatternFill("solid", fgColor="0B3D3A")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=14, bold=True, color="0B3D3A")

    ws["A1"] = f"Rapport Moussanada — {month_label}"
    ws["A1"].font = title_font
    ws.merge_cells("A1:D1")
    labels = [("Tickets ouverts", m["tickets"]["ouverts"]), ("Tickets résolus", m["tickets"]["resolus"]),
              ("Tickets en retard", m["tickets"]["en_retard"]), ("Tickets clos", m["tickets"]["clos"]),
              ("Délai moyen résolution (h)", m["durations"]["resolution_h"]),
              ("Délai moyen clôture (h)", m["durations"]["cloture_h"])]
    for i, (lbl, val) in enumerate(labels):
        ws.cell(row=3 + i, column=1, value=lbl).font = Font(bold=True)
        ws.cell(row=3 + i, column=2, value=val)

    def write_sheet(name, items, cols):
        ws2 = wb.create_sheet(name)
        for i, h in enumerate(cols):
            c = ws2.cell(row=1, column=i + 1, value=h)
            c.fill = header_fill
            c.font = header_font
        for r, it in enumerate(items, start=2):
            ws2.cell(row=r, column=1, value=it.get("label", ""))
            ws2.cell(row=r, column=2, value=it.get("ouverts", 0))
            ws2.cell(row=r, column=3, value=it.get("resolus", 0))
            ws2.cell(row=r, column=4, value=it.get("en_retard", 0))
            ws2.cell(row=r, column=5, value=it.get("fermes", 0))
        ws2.column_dimensions["A"].width = 42
        for col in "BCDE":
            ws2.column_dimensions[col].width = 12

    write_sheet("Catégories", m["categories"], ["Catégorie", "Ouverts", "Résolus", "En retard", "Fermés"])
    write_sheet("Services", m["services"], ["Service", "Ouverts", "Résolus", "En retard", "Fermés"])
    write_sheet("Techniciens", m["techniciens"], ["Technicien", "Ouverts", "Résolus", "En retard", "Fermés"])
    write_sheet("Demandeurs", m["demandeurs"], ["Demandeur", "Ouverts", "Résolus", "En retard", "Fermés"])

    ws3 = wb.create_sheet("Évolution mensuelle")
    for i, h in enumerate(["Mois", "Ouverts", "Résolus", "En retard", "Clos", "Délai résolution (h)", "Délai clôture (h)"]):
        c = ws3.cell(row=1, column=i + 1, value=h)
        c.fill = header_fill
        c.font = header_font
    for r, key in enumerate(sorted(ts.keys()), start=2):
        row = ts[key]
        ws3.cell(row=r, column=1, value=key)
        ws3.cell(row=r, column=2, value=row.get("ouverts", 0))
        ws3.cell(row=r, column=3, value=row.get("resolus", 0))
        ws3.cell(row=r, column=4, value=row.get("en_retard", 0))
        ws3.cell(row=r, column=5, value=row.get("clos", 0))
        ws3.cell(row=r, column=6, value=row.get("resolution_h", 0))
        ws3.cell(row=r, column=7, value=row.get("cloture_h", 0))
    for col in "ABCDEFG":
        ws3.column_dimensions[col].width = 14

    write_alerts_sheet(wb, "moussanada", header_fill, header_font)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"dashboard_moussanada_{ym}.xlsx",
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/modules")
def get_modules():
    return jsonify(MODULES)


# ---------------------------------------------------------------------------
# API - lecture / écriture groupée par mois (par volet)
# ---------------------------------------------------------------------------
@app.route("/api/<module>/month/<ym>")
def get_month(module, ym):
    check_module(module)
    calls = _load(module, "calls", ym, {})
    emails = _load(module, "emails", ym, {})
    analysis = _load(module, "analysis", ym, empty_analysis(ym))
    return jsonify({"calls": calls, "emails": emails, "analysis": analysis})


@app.route("/api/<module>/calls/<ym>", methods=["POST"])
def save_calls(module, ym):
    check_module(module)
    _save(module, "calls", ym, request.json or {})
    return jsonify({"ok": True})


@app.route("/api/<module>/emails/<ym>", methods=["POST"])
def save_emails(module, ym):
    check_module(module)
    _save(module, "emails", ym, request.json or {})
    return jsonify({"ok": True})


@app.route("/api/<module>/analysis/<ym>", methods=["POST"])
def save_analysis(module, ym):
    check_module(module)
    payload = request.json or {}
    payload.setdefault("month_label", month_label_fr(ym))
    _save(module, "analysis", ym, payload)
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# API - Paramètres globaux (agence, logo)
# ---------------------------------------------------------------------------
@app.route("/api/global-settings", methods=["GET"])
def get_global_settings():
    return jsonify(load_global_settings())


@app.route("/api/global-settings", methods=["POST"])
def post_global_settings():
    current = load_global_settings()
    current.update(request.json or {})
    save_global_settings(current)
    return jsonify({"ok": True})


@app.route("/api/config-export")
def config_export():
    user = current_user()
    if user["role"] != "admin":
        return jsonify({"error": "forbidden"}), 403
    bundle = {
        "global_settings": load_global_settings(),
        "modules": {m: load_module_settings(m) for m in MODULES.keys()},
        "users": load_users(),
        "knowledge": load_knowledge(),
    }
    buf = io.BytesIO(json.dumps(bundle, ensure_ascii=False, indent=2).encode("utf-8"))
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(buf, as_attachment=True, download_name=f"config_helpdesk_{stamp}.json", mimetype="application/json")


@app.route("/api/config-import", methods=["POST"])
def config_import():
    user = current_user()
    if user["role"] != "admin":
        return jsonify({"error": "forbidden"}), 403
    f = request.files.get("config")
    if not f:
        return jsonify({"ok": False, "error": "Aucun fichier reçu"}), 400
    try:
        bundle = json.loads(f.read().decode("utf-8"))
    except Exception as e:
        return jsonify({"ok": False, "error": f"Fichier invalide : {e}"}), 400
    if "global_settings" in bundle:
        save_global_settings(bundle["global_settings"])
    for m, s in (bundle.get("modules") or {}).items():
        if m in MODULES:
            save_module_settings(m, s)
    if "users" in bundle:
        save_users(bundle["users"])
    if "knowledge" in bundle:
        save_knowledge(bundle["knowledge"])
    return jsonify({"ok": True})


@app.route("/api/upload-logo", methods=["POST"])
def upload_logo():
    f = request.files.get("logo")
    if not f:
        return jsonify({"ok": False, "error": "Aucun fichier reçu"}), 400
    logo_type = request.form.get("logo_type", "ammps")
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".png", ".jpg", ".jpeg", ".svg"):
        return jsonify({"ok": False, "error": "Format non supporté (png, jpg, svg)"}), 400
    fname = ("sanad_logo" if logo_type == "sanad" else "logo") + ext
    f.save(os.path.join(UPLOAD_DIR, secure_filename(fname)))
    settings = load_global_settings()
    key = "sanad_logo_filename" if logo_type == "sanad" else "logo_filename"
    settings[key] = fname
    save_global_settings(settings)
    return jsonify({"ok": True, "logo_filename": fname, "logo_type": logo_type})


# ---------------------------------------------------------------------------
# API - Paramètres par volet (contacts, salutation, signature)
# ---------------------------------------------------------------------------
@app.route("/api/<module>/settings", methods=["GET"])
def get_module_settings(module):
    check_module(module)
    return jsonify(load_module_settings(module))


@app.route("/api/<module>/settings", methods=["POST"])
def post_module_settings(module):
    check_module(module)
    current = load_module_settings(module)
    current.update(request.json or {})
    save_module_settings(module, current)
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# API - Administration : mois, historique, export CSV brut (par volet)
# ---------------------------------------------------------------------------
@app.route("/api/<module>/months-list")
def months_list(module):
    check_module(module)
    d = module_dir(module)
    months = set()
    if module == "moussanada":
        prefixes = ("sources_", "notes_")
        for fname in os.listdir(d):
            for kind in prefixes:
                if fname.startswith(kind) and fname.endswith(".json"):
                    months.add(fname[len(kind):-5])
        months |= set(load_timeseries().keys())
        out = []
        for ym in sorted(months, reverse=True):
            out.append({
                "ym": ym,
                "label": month_label_fr(ym),
                "has_calls": ym in load_timeseries(),
                "has_emails": os.path.exists(sources_path(ym)),
                "has_analysis": os.path.exists(notes_path(ym)),
            })
        return jsonify(out)

    for fname in os.listdir(d):
        for kind in ("calls_", "emails_", "analysis_"):
            if fname.startswith(kind) and fname.endswith(".json"):
                months.add(fname[len(kind):-5])
    out = []
    for ym in sorted(months, reverse=True):
        out.append({
            "ym": ym,
            "label": month_label_fr(ym),
            "has_calls": os.path.exists(_path(module, "calls", ym)),
            "has_emails": os.path.exists(_path(module, "emails", ym)),
            "has_analysis": os.path.exists(_path(module, "analysis", ym)),
        })
    return jsonify(out)


@app.route("/api/<module>/delete-month/<ym>", methods=["POST"])
def delete_month(module, ym):
    check_module(module)
    removed = []
    if module == "moussanada":
        for p, label in ((sources_path(ym), "sources"), (notes_path(ym), "notes")):
            if os.path.exists(p):
                os.remove(p)
                removed.append(label)
        ts = load_timeseries()
        if ym in ts:
            del ts[ym]
            save_timeseries(ts)
            removed.append("timeseries")
        return jsonify({"ok": True, "removed": removed})

    for kind in ("calls", "emails", "analysis"):
        p = _path(module, kind, ym)
        if os.path.exists(p):
            os.remove(p)
            removed.append(kind)
    return jsonify({"ok": True, "removed": removed})


@app.route("/api/<module>/send-log")
def get_send_log(module):
    check_module(module)
    return jsonify(load_send_log(module))


@app.route("/api/<module>/export-csv/<kind>/<ym>")
def export_csv(module, kind, ym):
    check_module(module)
    if kind not in ("calls", "emails"):
        return jsonify({"ok": False, "error": "kind invalide"}), 400
    data = _load(module, kind, ym, {})
    buf = io.StringIO()
    writer = csv.writer(buf)
    if kind == "calls":
        writer.writerow(["date", "appels", "duree_hms"])
        for date in sorted(data.keys()):
            rec = data[date]
            writer.writerow([date, rec.get("calls", 0), format_hms(duration_seconds(rec))])
    else:
        writer.writerow(["date", "recus", "envoyes"])
        for date in sorted(data.keys()):
            rec = data[date]
            writer.writerow([date, rec.get("received", 0), rec.get("sent", 0)])
    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(mem, as_attachment=True, download_name=f"{module}_{kind}_{ym}.csv", mimetype="text/csv")


# ---------------------------------------------------------------------------
# API - Sauvegarde / restauration complète (par volet)
# ---------------------------------------------------------------------------
@app.route("/api/<module>/export-all")
def export_all(module):
    check_module(module)
    d = module_dir(module)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(d):
            for fn in files:
                full = os.path.join(root, fn)
                arcname = os.path.relpath(full, d)
                zf.write(full, arcname)
    buf.seek(0)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(buf, as_attachment=True, download_name=f"{module}_backup_{stamp}.zip",
                      mimetype="application/zip")


@app.route("/api/<module>/import-all", methods=["POST"])
def import_all(module):
    check_module(module)
    f = request.files.get("backup")
    if not f:
        return jsonify({"ok": False, "error": "Aucun fichier reçu"}), 400
    try:
        with zipfile.ZipFile(f) as zf:
            zf.extractall(module_dir(module))
    except zipfile.BadZipFile:
        return jsonify({"ok": False, "error": "Fichier zip invalide"}), 400
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# API - Vue annuelle / trimestrielle + heatmap jour de semaine (par volet)
# ---------------------------------------------------------------------------
def month_range(start_ym, end_ym):
    sy, sm = map(int, start_ym.split("-"))
    ey, em = map(int, end_ym.split("-"))
    out = []
    y, m = sy, sm
    while (y, m) <= (ey, em) and len(out) < 60:
        out.append(f"{y}-{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


@app.route("/api/<module>/annual/<int:year>")
def annual(module, year):
    check_module(module)
    start_ym = request.args.get("start") or f"{year}-01"
    end_ym = request.args.get("end") or f"{year}-12"
    months = month_range(start_ym, end_ym)

    if module == "moussanada":
        ts = load_timeseries()
        months_out = []
        for ym in months:
            row = ts.get(ym)
            if not row:
                continue
            months_out.append({
                "ym": ym, "label": month_label_fr(ym),
                "ouverts": row.get("ouverts", 0), "resolus": row.get("resolus", 0),
                "en_retard": row.get("en_retard", 0), "clos": row.get("clos", 0),
                "resolution_h": row.get("resolution_h", 0), "cloture_h": row.get("cloture_h", 0),
            })
        return jsonify({"year": year, "start": start_ym, "end": end_ym, "months": months_out})

    months_out = []
    cat_totals = {}
    weekday_load = {i: {"total": 0, "count": 0} for i in range(7)}

    for ym in months:
        calls = _load(module, "calls", ym, None)
        emails = _load(module, "emails", ym, None)
        analysis = _load(module, "analysis", ym, None)
        if calls is None and emails is None and analysis is None:
            continue
        calls = calls or {}
        emails = emails or {}
        analysis = analysis or {}
        problems = analysis.get("problems", [])
        demandes = analysis.get("demandes", [])
        total_recv = sum((d.get("received") or 0) for d in emails.values())
        total_sent = sum((d.get("sent") or 0) for d in emails.values())
        total_calls = sum((d.get("calls") or 0) for d in calls.values())
        total_bugs = sum(p.get("count", 0) for p in problems)
        total_dem = sum(d.get("count", 0) for d in demandes)
        for p in problems:
            cat_totals.setdefault(p.get("label", "?"), 0)
            cat_totals[p.get("label", "?")] += p.get("count", 0)

        all_dates = set(list(calls.keys()) + list(emails.keys()))
        for date_str in all_dates:
            try:
                dow = datetime.strptime(date_str, "%Y-%m-%d").weekday()
            except ValueError:
                continue
            c = (calls.get(date_str, {}).get("calls") or 0)
            e = (emails.get(date_str, {}).get("received") or 0)
            weekday_load[dow]["total"] += (c + e)
            weekday_load[dow]["count"] += 1

        months_out.append({
            "ym": ym, "label": month_label_fr(ym),
            "emails_received": total_recv, "emails_sent": total_sent,
            "calls": total_calls, "bugs": total_bugs, "demandes": total_dem,
        })

    top_categories = sorted(cat_totals.items(), key=lambda x: x[1], reverse=True)[:3]

    heatmap = []
    for i in range(7):
        cnt = weekday_load[i]["count"]
        avg = round(weekday_load[i]["total"] / cnt, 1) if cnt else 0
        heatmap.append({"day": JOURS_FR[i], "avg": avg})

    return jsonify({
        "year": year,
        "months": months_out,
        "top_categories": [{"label": k, "count": v} for k, v in top_categories],
        "weekday_heatmap": heatmap,
    })


# ---------------------------------------------------------------------------
# Génération HTML du rapport
# ---------------------------------------------------------------------------
def logo_data_uri(filename):
    if not filename:
        return None
    path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(path):
        return None
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    mime = "svg+xml" if ext == "svg" else ("jpeg" if ext in ("jpg", "jpeg") else "png")
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/{mime};base64,{b64}"


def report_logos_html(global_settings):
    sanad_uri = logo_data_uri(global_settings.get("sanad_logo_filename"))
    ammps_uri = logo_data_uri(global_settings.get("logo_filename"))
    imgs = ""
    if sanad_uri:
        imgs += f'<img src="{sanad_uri}" style="height:28px;margin-right:12px;vertical-align:middle;">'
    if ammps_uri:
        imgs += f'<img src="{ammps_uri}" style="height:28px;vertical-align:middle;">'
    if not imgs:
        return ""
    return f'<div style="margin-bottom:8px;">{imgs}</div>'


def build_report_html(module, ym, calls, emails, analysis, prev_calls=None, prev_emails=None, prev_analysis=None, charts=None):
    charts = charts or {}
    prev_calls = prev_calls or {}
    prev_emails = prev_emails or {}
    prev_analysis = prev_analysis or {}
    month_label = analysis.get("month_label") or month_label_fr(ym)
    prev_month_label = month_label_fr(prev_ym(ym))

    total_received = sum((d.get("received") or 0) for d in emails.values())
    total_sent = sum((d.get("sent") or 0) for d in emails.values())
    total_calls = sum((d.get("calls") or 0) for d in calls.values())
    total_duration_sec = sum(duration_seconds(d) for d in calls.values())
    avg_call_sec = (total_duration_sec / total_calls) if total_calls else 0

    prev_received = sum((d.get("received") or 0) for d in prev_emails.values())
    prev_sent = sum((d.get("sent") or 0) for d in prev_emails.values())
    prev_calls_total = sum((d.get("calls") or 0) for d in prev_calls.values())

    problems = analysis.get("problems", [])
    demandes = analysis.get("demandes", [])
    weekly = analysis.get("weekly", [])
    constats = analysis.get("constats", [])
    recommandations = analysis.get("recommandations", [])
    keywords = analysis.get("keywords", [])

    total_problems = sum(p.get("count", 0) for p in problems)
    total_demandes = sum(d.get("count", 0) for d in demandes)
    ratio_bd = round(total_problems / total_demandes, 2) if total_demandes else 0

    prev_problems = prev_analysis.get("problems", [])
    prev_demandes = prev_analysis.get("demandes", [])
    prev_total_problems = sum(p.get("count", 0) for p in prev_problems)
    prev_total_demandes = sum(d.get("count", 0) for d in prev_demandes)

    channel_total = total_calls + total_received
    channel_calls_pct = round(total_calls / channel_total * 100) if channel_total else 0
    channel_email_pct = 100 - channel_calls_pct if channel_total else 0

    crit_week = None
    if weekly:
        crit_week = max(weekly, key=lambda w: (w.get("bugs", 0) + w.get("demandes", 0)))

    def pct(n, total):
        return round((n / total) * 100, 1) if total else 0

    def esc(s):
        return html_module.escape(str(s), quote=True) if s is not None else ""

    def delta_span(cur, prev_v, polarity="neutral"):
        if not prev_v and not cur:
            return ""
        if not prev_v:
            return "<div style='font-size:11px;font-weight:700;color:#67737E;margin-top:2px;'>nouveau</div>"
        diff = ((cur - prev_v) / prev_v) * 100
        sign = "+" if diff >= 0 else ""
        if abs(diff) < 1 or polarity == "neutral":
            color = "#67737E"
        else:
            is_up = diff > 0
            is_good = is_up if polarity == "positive" else not is_up
            color = "#25935F" if is_good else "#DC2828"
        return f"<div style='font-size:11px;font-weight:700;color:{color};margin-top:2px;'>{sign}{diff:.0f}% vs {esc(prev_month_label)}</div>"

    def kpi_card(label, value, delta_html="", bg="#F5F7F9", color="#0B4965"):
        return f"""<td style="width:25%;padding:10px;text-align:center;background:{bg};border:1px solid #DAE0E7;">
          <div style="font-size:11px;color:#67737E;">{esc(label)}</div>
          <div style="font-size:18px;font-weight:bold;color:{color};">{esc(value)}</div>
          {delta_html}
        </td>"""

    def img_tag(key, title=""):
        b64 = charts.get(key)
        if not b64:
            return ""
        caption = f'<div style="font-size:11.5px;color:#67737E;text-align:center;margin:4px 0 14px;">{esc(title)}</div>' if title else ""
        return f'<img src="{b64}" style="max-width:100%;border-radius:8px;margin:14px 0 0;border:1px solid #DAE0E7;display:block;">{caption}'

    def rows_cat_problems(items, total):
        out = ""
        for it in items:
            p = pct(it.get("count", 0), total)
            action = it.get("action", "")
            out += f"""
            <tr>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;color:#0D1926;">{esc(it.get('label',''))}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;color:#0D1926;text-align:center;font-weight:bold;">{it.get('count',0)}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;color:#DC2828;text-align:center;">{p}%</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#67737E;">{esc(it.get('desc',''))}{f'<br><em>Action : {esc(action)}</em>' if action else ''}</td>
            </tr>"""
        return out

    def rows_cat_demandes(items, total):
        out = ""
        for it in items:
            p = pct(it.get("count", 0), total)
            out += f"""
            <tr>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;color:#0D1926;">{esc(it.get('label',''))}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;color:#0D1926;text-align:center;font-weight:bold;">{it.get('count',0)}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;color:#25935F;text-align:center;">{p}%</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#67737E;">{esc(it.get('desc',''))}</td>
            </tr>"""
        return out

    def rows_weekly():
        out = ""
        for w in weekly:
            is_crit = crit_week is not None and w is crit_week
            bg = "#FCE9E9" if is_crit else "#FFFFFF"
            out += f"""
            <tr style="background:{bg};">
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;font-weight:{'bold' if is_crit else 'normal'};color:#0D1926;">{esc(w.get('week',''))}{' ⚠️' if is_crit else ''}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;text-align:center;color:#DC2828;font-weight:bold;">{w.get('bugs',0)}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:13px;text-align:center;color:#25935F;font-weight:bold;">{w.get('demandes',0)}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#67737E;">{esc(w.get('obs',''))}</td>
            </tr>"""
        return out

    def list_html(items):
        return "".join(f"<li style='margin-bottom:6px;font-family:Arial,sans-serif;font-size:13px;color:#0D1926;'>{esc(i)}</li>" for i in items)

    keywords_line = ""
    if keywords:
        sorted_kw = sorted(keywords, key=lambda k: k.get("count", 0), reverse=True)
        keywords_line = ", ".join(f"{esc(k.get('word',''))} ({k.get('count',0)})" for k in sorted_kw[:15])

    global_settings = load_global_settings()
    module_label = MODULES[module]["label"]

    html = f"""
    <div style="max-width:760px;margin:0 auto;font-family:Arial,sans-serif;background:#FFFFFF;">
      <div style="background:#0B4965;padding:22px 24px;border-radius:6px 6px 0 0;">
        {report_logos_html(global_settings)}
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">Rapport Support {esc(module_label)}</div>
        <div style="color:#B9D3E0;font-size:13px;margin-top:4px;">{esc(global_settings.get('agency_name',''))} — {esc(MODULES[module]['subtitle'])} — {esc(month_label)}</div>
      </div>

      <div style="padding:20px 24px;border:1px solid #DAE0E7;border-top:none;">
        <table style="width:100%;border-collapse:collapse;margin-bottom:18px;">
          <tr>
            {kpi_card("EMAILS REÇUS", total_received, delta_span(total_received, prev_received, "neutral"))}
            {kpi_card("EMAILS ENVOYÉS", total_sent, delta_span(total_sent, prev_sent, "neutral"))}
            {kpi_card("APPELS", total_calls, delta_span(total_calls, prev_calls_total, "neutral"))}
            {kpi_card("TEMPS COMM. (hh:mm:ss)", format_hms(total_duration_sec), "", color="#0B4965")}
          </tr>
        </table>

        <table style="width:100%;border-collapse:collapse;margin-bottom:12px;">
          <tr>
            {kpi_card("PROBLÈMES TECH.", total_problems, delta_span(total_problems, prev_total_problems, "negative"), bg="#FCE9E9", color="#DC2828")}
            {kpi_card("DEMANDES INFO", total_demandes, delta_span(total_demandes, prev_total_demandes, "neutral"), bg="#E7F5EE", color="#25935F")}
            {kpi_card("RATIO BUGS/DEMANDES", ratio_bd)}
            {kpi_card("SEMAINE CRITIQUE", crit_week.get('week','-') if crit_week else '-', bg="#FEF3E0", color="#B9791F")}
          </tr>
        </table>

        <table style="width:100%;border-collapse:collapse;margin-bottom:18px;">
          <tr>
            {kpi_card("DURÉE MOY./APPEL", format_hms(avg_call_sec))}
            {kpi_card("CANAL DOMINANT", f"☎ {channel_calls_pct}% / ✉ {channel_email_pct}%" if channel_total else "—", bg="#F5F7F9")}
          </tr>
        </table>

        {img_tag("weekly", "Évolution hebdomadaire — Bugs vs Demandes")}
        {img_tag("problems", "Répartition des problèmes techniques")}

        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:20px 0 8px;">Problèmes techniques signalés ({total_problems})</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Type</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Nb</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">%</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Description</td>
          </tr>
          {rows_cat_problems(problems, total_problems)}
        </table>

        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:20px 0 8px;">Demandes d'information reçues ({total_demandes})</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Type</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Nb</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">%</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Description</td>
          </tr>
          {rows_cat_demandes(demandes, total_demandes)}
        </table>

        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:20px 0 8px;">Évolution hebdomadaire</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Semaine</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Bugs</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Demandes</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Observations</td>
          </tr>
          {rows_weekly()}
        </table>

        {"<div style='font-size:14px;font-weight:bold;color:#0B4965;margin:20px 0 8px;'>Mots-clés fréquents</div><div style='font-size:12.5px;color:#67737E;margin-bottom:16px;'>" + keywords_line + "</div>" if keywords_line else ""}

        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:20px 0 8px;">Constats clés</div>
        <ul style="padding-left:18px;margin:0 0 16px;">{list_html(constats)}</ul>

        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:16px 0 8px;">Recommandations</div>
        <ul style="padding-left:18px;margin:0;">{list_html(recommandations)}</ul>
      </div>
    </div>
    """
    return html


def build_email_body(month_label, report_html, greeting, module_settings, module_label="Tarkhiss"):
    signature_parts = [module_settings.get("signature_name", ""), module_settings.get("signature_function", ""), module_settings.get("signature_phone", "")]
    signature_parts = [p for p in signature_parts if p]
    signature_html = "<br>".join(signature_parts)

    return f"""
    <div style="max-width:760px;margin:0 auto;font-family:Arial,sans-serif;">
      <p style="font-size:14px;color:#0D1926;">{greeting}</p>
      <p style="font-size:14px;color:#0D1926;">Veuillez trouver ci-dessous le rapport mensuel du support {module_label} ({month_label}).</p>
      {report_html}
      <p style="font-size:13px;color:#0D1926;margin-top:22px;">Cordialement,{"<br>" + signature_html if signature_html else ""}</p>
    </div>
    """


def build_subject(module, month_label):
    return f"Rapport de Synthèse du Support {MODULES[module]['label']} - {month_label}"


def pct(n, total):
    return round((n / total) * 100, 1) if total else 0


def top_n(items, n=10, key="ouverts"):
    return sorted(items, key=lambda x: x.get(key, 0), reverse=True)[:n]


def build_report_html_moussanada(ym, month_label, charts=None):
    charts = charts or {}
    data = moussanada_data(ym).get_json()
    m = data["month"]
    ts = data["timeseries"]
    notes = data["notes"]
    tickets = m["tickets"]
    durations = m["durations"]

    prev_ym = f"{int(ym[:4]) - (1 if ym[5:] == '01' else 0)}-{12 if ym[5:] == '01' else int(ym[5:]) - 1:02d}"
    prev = ts.get(prev_ym, {})
    prev_month_label = month_label_fr(prev_ym)

    taux_resolution = pct(tickets["resolus"], tickets["ouverts"])
    taux_cloture = pct(tickets["clos"], tickets["ouverts"])
    prev_taux_resolution = pct(prev.get("resolus", 0), prev.get("ouverts", 0))
    prev_taux_cloture = pct(prev.get("clos", 0), prev.get("ouverts", 0))

    top_cat = top_n(m["categories"], 10)
    top_srv = top_n(m["services"], 10)
    techs = sorted(m["techniciens"], key=lambda x: x.get("ouverts", 0), reverse=True)

    def esc(s):
        return html_module.escape(str(s), quote=True) if s is not None else ""

    def delta_span(cur, prev_v, polarity):
        if not prev_v and not cur:
            return ""
        if not prev_v:
            return "<div style='font-size:11px;font-weight:700;color:#67737E;margin-top:2px;'>nouveau</div>"
        diff = ((cur - prev_v) / prev_v) * 100
        sign = "+" if diff >= 0 else ""
        if abs(diff) < 1 or polarity == "neutral":
            color = "#67737E"
        else:
            is_up = diff > 0
            is_good = is_up if polarity == "positive" else not is_up
            color = "#25935F" if is_good else "#DC2828"
        return f"<div style='font-size:11px;font-weight:700;color:{color};margin-top:2px;'>{sign}{diff:.0f}% vs {esc(prev_month_label)}</div>"

    def kpi_card(label, value, delta_html="", bg="#F5F7F9", color="#0B4965"):
        return f"""<td style="width:25%;padding:10px;text-align:center;background:{bg};border:1px solid #DAE0E7;">
          <div style="font-size:11px;color:#67737E;">{esc(label)}</div>
          <div style="font-size:18px;font-weight:bold;color:{color};">{esc(value)}</div>
          {delta_html}
        </td>"""

    def rows_top(items):
        total = sum(i.get("ouverts", 0) for i in items) or 1
        out = ""
        for it in items:
            p = pct(it.get("ouverts", 0), total)
            out += f"""<tr>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;color:#0D1926;">{esc(it.get('label',''))}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;text-align:center;font-weight:bold;">{it.get('ouverts',0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;text-align:center;color:#25935F;">{it.get('resolus',0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#67737E;text-align:center;">{p}%</td>
            </tr>"""
        return out

    def rows_tech():
        out = ""
        n = len(techs) or 1
        avg_o = sum(t.get('ouverts',0) for t in techs)/n
        avg_r = sum(t.get('resolus',0) for t in techs)/n
        for t in techs:
            above = t.get('ouverts',0) >= avg_o
            out += f"""<tr>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;color:#0D1926;">{esc(t.get('label',''))}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;text-align:center;font-weight:bold;">{t.get('ouverts',0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;text-align:center;color:#25935F;">{t.get('resolus',0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12.5px;text-align:center;color:#DC2828;">{t.get('en_retard',0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:11.5px;text-align:center;color:{'#25935F' if above else '#B9791F'};">{'≥ moy.' if above else '< moy.'}</td>
            </tr>"""
        if techs:
            out += f"""<tr style="background:#F5F7F9;">
              <td style="padding:7px 10px;font-family:Arial,sans-serif;font-size:12.5px;font-weight:bold;color:#0B4965;">Moyenne équipe</td>
              <td style="padding:7px 10px;text-align:center;font-weight:bold;">{avg_o:.1f}</td>
              <td style="padding:7px 10px;text-align:center;font-weight:bold;color:#25935F;">{avg_r:.1f}</td>
              <td style="padding:7px 10px;"></td><td style="padding:7px 10px;"></td>
            </tr>"""
        return out

    def list_html(items):
        return "".join(f"<li style='margin-bottom:6px;font-family:Arial,sans-serif;font-size:13px;color:#0D1926;'>{esc(i)}</li>" for i in items)

    def img_tag(key):
        b64 = charts.get(key)
        if not b64:
            return ""
        return f'<img src="{b64}" style="max-width:100%;border-radius:8px;margin:12px 0 4px;border:1px solid #DAE0E7;display:block;">'

    def copilot_box(text):
        if not text or not text.strip():
            return "<div style='font-size:12px;color:#67737E;font-style:italic;margin:6px 0 6px;'>— Analyse Copilot non renseignée pour cette section —</div>"
        paras = "".join(f"<p style='margin:0 0 8px;'>{esc(p)}</p>" for p in text.split("\n") if p.strip())
        return f"""<div style="background:#E7EEF2;border-left:4px solid #0B4965;border-radius:0 8px 8px 0;padding:12px 16px;margin:10px 0 4px;font-size:13px;color:#0D1926;line-height:1.6;">{paras}</div>"""

    def section_title(n, title):
        return f"""<div style="margin:28px 0 10px;padding-bottom:6px;border-bottom:2px solid #0B4965;">
          <span style="font-size:10.5px;font-weight:800;color:#F59F0A;letter-spacing:.6px;">SECTION {n}</span>
          <div style="font-size:17px;font-weight:800;color:#0B4965;">{esc(title)}</div>
        </div>"""

    sections = notes.get("sections", {})
    global_settings = load_global_settings()

    html = f"""
    <div style="max-width:760px;margin:0 auto;font-family:Arial,sans-serif;background:#FFFFFF;">
      <div style="background:#0B4965;padding:22px 24px;border-radius:6px 6px 0 0;">
        {report_logos_html(global_settings)}
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">Rapport Support Moussanada</div>
        <div style="color:#B9D3E0;font-size:13px;margin-top:4px;">{esc(global_settings.get('agency_name',''))} — Helpdesk GLPI — {esc(month_label)}</div>
      </div>
      <div style="padding:20px 24px;border:1px solid #DAE0E7;border-top:none;">

        {section_title(1, "Volume global des tickets")}
        <table style="width:100%;border-collapse:collapse;margin-bottom:6px;"><tr>
          {kpi_card("TICKETS OUVERTS", tickets["ouverts"], delta_span(tickets["ouverts"], prev.get("ouverts", 0), "neutral"))}
          {kpi_card("TICKETS RÉSOLUS", tickets["resolus"], delta_span(tickets["resolus"], prev.get("resolus", 0), "positive"))}
          {kpi_card("TICKETS CLOS", tickets["clos"], delta_span(tickets["clos"], prev.get("clos", 0), "positive"))}
          {kpi_card("EN RETARD", tickets["en_retard"], delta_span(tickets["en_retard"], prev.get("en_retard", 0), "negative"), bg="#FCE9E9", color="#DC2828")}
        </tr></table>
        {img_tag("volume")}
        {img_tag("backlog")}
        {copilot_box(sections.get("volume", ""))}

        {section_title(2, "Évolution des délais de traitement")}
        <table style="width:100%;border-collapse:collapse;margin-bottom:6px;"><tr>
          {kpi_card("TAUX DE RÉSOLUTION", f"{taux_resolution}%", delta_span(taux_resolution, prev_taux_resolution, "positive"), bg="#E7F5EE", color="#25935F")}
          {kpi_card("TAUX DE CLÔTURE", f"{taux_cloture}%", delta_span(taux_cloture, prev_taux_cloture, "positive"), bg="#E7F5EE", color="#25935F")}
          {kpi_card("DÉLAI MOY. RÉSOLUTION", format_dh(durations["resolution_h"]), delta_span(durations["resolution_h"], prev.get("resolution_h", 0), "negative"))}
          {kpi_card("DÉLAI MOY. CLÔTURE", format_dh(durations["cloture_h"]), delta_span(durations["cloture_h"], prev.get("cloture_h", 0), "negative"))}
        </tr></table>
        {img_tag("trend")}
        {copilot_box(sections.get("delays", ""))}

        {section_title(3, f"Principales catégories de demandes en {month_label}")}
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;">Catégorie</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Ouverts</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Résolus</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">%</td>
          </tr>
          {rows_top(top_cat)}
        </table>
        {img_tag("cat")}
        {img_tag("pareto")}
        {copilot_box(sections.get("categories", ""))}

        {section_title(4, "Services les plus demandeurs")}
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;">Service</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Ouverts</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Résolus</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">%</td>
          </tr>
          {rows_top(top_srv)}
        </table>
        {img_tag("srv")}
        {copilot_box(sections.get("services", ""))}

        {section_title(5, "Charge et mobilisation de l'équipe SI")}
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;">Technicien</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Ouverts</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Résolus</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">Retard</td>
            <td style="padding:7px 10px;color:#fff;font-size:11.5px;font-weight:bold;text-align:center;">vs Équipe</td>
          </tr>
          {rows_tech()}
        </table>
        {img_tag("tech")}
        {img_tag("radar")}
        {copilot_box(sections.get("team", ""))}

        {section_title(6, f"Analyse comparative {prev_month_label} / {month_label}")}
        <p style="font-size:12.5px;color:#67737E;margin:0 0 4px;">Comparatif des volumes et des délais moyens entre les deux mois.</p>
        {img_tag("volume")}
        {img_tag("delay")}
        {copilot_box(sections.get("comparative", ""))}

        {section_title(7, "Synthèse finale et conclusion")}
        {("<div style='font-size:13px;font-weight:bold;color:#0B4965;margin:6px 0 6px;'>Constats clés</div><ul style='padding-left:18px;margin:0 0 14px;'>" + list_html(notes.get('constats', [])) + "</ul>") if notes.get('constats') else ""}
        {("<div style='font-size:13px;font-weight:bold;color:#0B4965;margin:6px 0 6px;'>Recommandations</div><ul style='padding-left:18px;margin:0 0 14px;'>" + list_html(notes.get('recommandations', [])) + "</ul>") if notes.get('recommandations') else ""}
        {copilot_box(sections.get("synthesis", ""))}

      </div>
    </div>
    """
    return html


def get_report_html_and_label(module, ym, charts=None):
    if module == "moussanada":
        month_label = month_label_fr(ym)
        return build_report_html_moussanada(ym, month_label, charts), month_label
    calls = _load(module, "calls", ym, {})
    emails = _load(module, "emails", ym, {})
    analysis = _load(module, "analysis", ym, empty_analysis(ym))
    py = prev_ym(ym)
    prev_calls = _load(module, "calls", py, {})
    prev_emails = _load(module, "emails", py, {})
    prev_analysis = _load(module, "analysis", py, empty_analysis(py))
    month_label = analysis.get("month_label") or month_label_fr(ym)
    return build_report_html(module, ym, calls, emails, analysis, prev_calls, prev_emails, prev_analysis, charts), month_label


@app.route("/api/<module>/preview-email/<ym>", methods=["POST"])
def preview_email(module, ym):
    check_module(module)
    payload = request.json or {}
    module_settings = load_module_settings(module)

    report_html, month_label = get_report_html_and_label(module, ym, payload.get("charts"))
    greeting = payload.get("greeting") or module_settings.get("greeting", "Bonjour,")

    body_html = build_email_body(month_label, report_html, greeting, module_settings, MODULES[module]["label"])
    subject = build_subject(module, month_label)
    return jsonify({"subject": subject, "html": body_html})


@app.route("/api/<module>/send-email/<ym>", methods=["POST"])
def send_email(module, ym):
    check_module(module)
    payload = request.json or {}
    module_settings = load_module_settings(module)

    report_html, month_label = get_report_html_and_label(module, ym, payload.get("charts"))
    greeting = payload.get("greeting") or module_settings.get("greeting", "Bonjour,")
    to_list = payload.get("to") or [c["email"] for c in module_settings.get("contacts", []) if c.get("role") == "to"]
    cc_list = payload.get("cc") or [c["email"] for c in module_settings.get("contacts", []) if c.get("role") == "cc"]

    body_html = build_email_body(month_label, report_html, greeting, module_settings, MODULES[module]["label"])
    subject = build_subject(module, month_label)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["To"] = "; ".join(to_list)
    if cc_list:
        msg["Cc"] = "; ".join(cc_list)
    msg.set_content("Ce message nécessite un client compatible HTML.")
    msg.add_alternative(body_html, subtype="html")

    fname = f"rapport_{module}_{ym}.eml"
    fpath = os.path.join(module_dir(module), "exports", fname)
    with open(fpath, "wb") as f:
        f.write(bytes(msg))

    opened = False
    error = None
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(fpath)  # noqa
            opened = True
        elif system == "Darwin":
            subprocess.run(["open", fpath], check=True)
            opened = True
        else:
            subprocess.run(["xdg-open", fpath], check=True)
            opened = True
    except Exception as e:
        error = str(e)

    append_send_log(module, {
        "date": datetime.now().isoformat(timespec="seconds"),
        "ym": ym,
        "subject": subject,
        "to": to_list,
        "cc": cc_list,
        "opened": opened,
    })

    return jsonify({"ok": True, "file": fpath, "opened": opened, "error": error, "subject": subject})


# ---------------------------------------------------------------------------
# API - Export Excel du dashboard (par volet)
# ---------------------------------------------------------------------------
@app.route("/api/<module>/export-xlsx/<ym>")
def export_xlsx(module, ym):
    check_module(module)
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    if module == "moussanada":
        return export_xlsx_moussanada(ym)

    calls = _load(module, "calls", ym, {})
    emails = _load(module, "emails", ym, {})
    analysis = _load(module, "analysis", ym, empty_analysis(ym))
    month_label = analysis.get("month_label") or month_label_fr(ym)

    total_received = sum((d.get("received") or 0) for d in emails.values())
    total_sent = sum((d.get("sent") or 0) for d in emails.values())
    total_calls = sum((d.get("calls") or 0) for d in calls.values())
    total_duration_sec = sum(duration_seconds(d) for d in calls.values())
    problems = analysis.get("problems", [])
    demandes = analysis.get("demandes", [])
    weekly = analysis.get("weekly", [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    header_fill = PatternFill("solid", fgColor="0B3D3A")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=14, bold=True, color="0B3D3A")

    ws["A1"] = f"Rapport Support {MODULES[module]['label']} — {month_label}"
    ws["A1"].font = title_font
    ws.merge_cells("A1:F1")

    ws["A3"] = "Emails reçus"; ws["B3"] = total_received
    ws["A4"] = "Emails envoyés"; ws["B4"] = total_sent
    ws["A5"] = "Total appels"; ws["B5"] = total_calls
    ws["A6"] = "Temps comm. (hh:mm:ss)"; ws["B6"] = format_hms(total_duration_sec)
    for r in range(3, 7):
        ws[f"A{r}"].font = Font(bold=True)

    row = 8
    ws.cell(row=row, column=1, value="Problèmes techniques").font = title_font
    row += 1
    for i, h in enumerate(["Type", "Nombre", "%", "Description", "Action corrective"]):
        c = ws.cell(row=row, column=i + 1, value=h)
        c.fill = header_fill
        c.font = header_font
    total_p = sum(p.get("count", 0) for p in problems)
    row += 1
    for p in problems:
        pct = round(p.get("count", 0) / total_p * 100, 1) if total_p else 0
        ws.cell(row=row, column=1, value=p.get("label", ""))
        ws.cell(row=row, column=2, value=p.get("count", 0))
        ws.cell(row=row, column=3, value=f"{pct}%")
        ws.cell(row=row, column=4, value=p.get("desc", ""))
        ws.cell(row=row, column=5, value=p.get("action", ""))
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Demandes d'information").font = title_font
    row += 1
    for i, h in enumerate(["Type", "Nombre", "%", "Description"]):
        c = ws.cell(row=row, column=i + 1, value=h)
        c.fill = header_fill
        c.font = header_font
    total_d = sum(d.get("count", 0) for d in demandes)
    row += 1
    for d in demandes:
        pct = round(d.get("count", 0) / total_d * 100, 1) if total_d else 0
        ws.cell(row=row, column=1, value=d.get("label", ""))
        ws.cell(row=row, column=2, value=d.get("count", 0))
        ws.cell(row=row, column=3, value=f"{pct}%")
        ws.cell(row=row, column=4, value=d.get("desc", ""))
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Évolution hebdomadaire").font = title_font
    row += 1
    for i, h in enumerate(["Semaine", "Bugs", "Demandes", "Observations"]):
        c = ws.cell(row=row, column=i + 1, value=h)
        c.fill = header_fill
        c.font = header_font
    row += 1
    for w in weekly:
        ws.cell(row=row, column=1, value=w.get("week", ""))
        ws.cell(row=row, column=2, value=w.get("bugs", 0))
        ws.cell(row=row, column=3, value=w.get("demandes", 0))
        ws.cell(row=row, column=4, value=w.get("obs", ""))
        row += 1

    for col, width in zip("ABCDE", [26, 12, 10, 40, 26]):
        ws.column_dimensions[col].width = width

    write_alerts_sheet(wb, module, header_fill, header_font)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"dashboard_{module}_{ym}.xlsx",
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    print("Helpdesk Dashboard -> http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=True)
