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
import shutil
import html as html_module
from glpi_import import (
    glpi_cell_to_seconds, decimal_hours_cell, format_dh, normalize_header,
    match_header, parse_glpi_row_sheet, import_moussanada_xlsx as _parse_moussanada_xlsx,
    import_moussanada_csv as _parse_moussanada_csv,
    parse_glpi_raw_tickets, merge_raw_tickets, compute_heatmap_matrix,
    compute_resolution_hours, bucket_resolution_hours,
    store_ticket_series, store_source_items,
)
from glpi_client import GLPIClient, GLPIConnectionError, test_connection as glpi_test_connection
import glpi_profiles
import tarkhiss_meta as tm
from pchc_import import (
    CATEGORIES as PCHC_CATEGORIES, COLOR_HEX as PCHC_COLOR_HEX,
    status_color as pchc_status_color, import_pchc_xlsx as _parse_pchc_xlsx,
    import_pchc_csv as _parse_pchc_csv, normalize_key as pchc_normalize_key,
)
import base64
import subprocess
import zipfile
from datetime import datetime, timedelta
from email.message import EmailMessage

from functools import wraps

def hex_to_rgbcolor(hex_str):
    """Convertit une couleur hex '#RRGGBB' en RGBColor python-pptx (import différé, léger)."""
    from pptx.dml.color import RGBColor as _RGBColor
    h = hex_str.lstrip("#")
    return _RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
from flask import Flask, jsonify, request, render_template, send_file, abort, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = os.path.join(BASE_DIR, "static", "uploads")
BACKUP_DIR = os.path.join(DATA_DIR, "_backups")
AUDIT_LOG_FILE = os.path.join(DATA_DIR, "audit_log.jsonl")
BACKUP_RETENTION_PER_FILE = 20
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)


def write_json_safely(path, data):
    """Écriture JSON sécurisée : sauvegarde horodatée de l'ancien contenu avant écrasement,
    puis écriture atomique (fichier temporaire + remplacement) pour éviter toute corruption
    en cas de coupure/crash pendant l'écriture. Purge les anciennes sauvegardes au-delà de
    BACKUP_RETENTION_PER_FILE par fichier source."""
    if os.path.exists(path):
        try:
            rel = os.path.relpath(path, DATA_DIR).replace(os.sep, "__")
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            backup_path = os.path.join(BACKUP_DIR, f"{rel}.{stamp}.bak")
            shutil.copy2(path, backup_path)
            prefix = f"{rel}."
            existing = sorted(
                (f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix) and f.endswith(".bak")),
                reverse=True,
            )
            for old in existing[BACKUP_RETENTION_PER_FILE:]:
                try:
                    os.remove(os.path.join(BACKUP_DIR, old))
                except OSError:
                    pass
        except OSError:
            pass  # une sauvegarde ratée ne doit jamais bloquer l'écriture principale

    tmp_path = f"{path}.tmp-{os.getpid()}"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def log_audit(action, details=None, module=None):
    """Journal d'audit append-only (JSONL) : qui a fait quoi et quand. Consulté par
    l'Administrateur via /api/admin/audit-log."""
    try:
        u = current_user()
        user = u["username"] if u else "system"
    except Exception:
        user = "system"
    entry = {
        "ts": datetime.now().isoformat(),
        "user": user,
        "action": action,
        "module": module,
        "details": details or {},
    }
    try:
        with open(AUDIT_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except OSError:
        pass


GLOBAL_SETTINGS_FILE = os.path.join(DATA_DIR, "global_settings.json")

app = Flask(__name__)
app.config["SESSION_PERMANENT"] = True

SECRET_KEY_FILE = os.path.join(DATA_DIR, ".secret_key")
if os.path.exists(SECRET_KEY_FILE):
    with open(SECRET_KEY_FILE, "r") as f:
        app.secret_key = f.read().strip()
else:
    app.secret_key = os.urandom(32).hex()
    with open(SECRET_KEY_FILE, "w") as f:
        f.write(app.secret_key)


@app.before_request
def apply_session_timeout():
    """Durée de session configurable (Administration → Sécurité), en minutes d'inactivité.
    Doit s'exécuter à chaque requête (avant le contrôle d'accès) pour que le compteur
    d'inactivité soit rafraîchi tant que l'utilisateur est actif."""
    session.permanent = True
    try:
        minutes = int(load_global_settings().get("session_timeout_minutes", 240))
    except Exception:
        minutes = 240
    app.permanent_session_lifetime = timedelta(minutes=max(5, minutes))

MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin",
           "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
JOURS_FR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# ---------------------------------------------------------------------------
# Volets (modules) de l'application
# ---------------------------------------------------------------------------
MODULES = {
    "tarkhiss": {"label": "Tarkhiss", "subtitle": "Support Email & Hotline", "ready": True},
    "moussanada": {"label": "Moussanada", "subtitle": "Helpdesk GLPI", "ready": True},
    "pchc": {"label": "Reporting Métier", "subtitle": "PCHC — Produits Cosmétiques et d'Hygiène Corporelle", "ready": True},
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
    write_json_safely(USERS_FILE, users)


def validate_password(password):
    """Politique de mot de passe configurable (Administration → Sécurité). Retourne un message
    d'erreur (str) si le mot de passe ne respecte pas la politique, sinon None."""
    gs = load_global_settings()
    min_len = int(gs.get("password_min_length", 8))
    if len(password or "") < min_len:
        return f"Mot de passe trop court ({min_len} caractères minimum)."
    if gs.get("password_require_digit", True) and not any(c.isdigit() for c in password):
        return "Le mot de passe doit contenir au moins un chiffre."
    if gs.get("password_require_upper", True) and not any(c.isupper() for c in password):
        return "Le mot de passe doit contenir au moins une majuscule."
    return None


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
    if path == "/api/global-settings" and request.method == "GET":
        return  # branding (nom d'app, agence, thème) : lisible avant connexion pour l'écran de login
    if path == "/api/system/reminder-check":
        return  # authentification par jeton dédié (voir la fonction), pas par session — appel headless
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
    if not username:
        return jsonify({"ok": False, "error": "Identifiant requis."}), 400
    pw_error = validate_password(password)
    if pw_error:
        return jsonify({"ok": False, "error": pw_error}), 400
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
        log_audit("login_failed", {"username": payload.get("username")})
        return jsonify({"ok": False, "error": "Identifiants incorrects"}), 401
    session["user_id"] = u["id"]
    log_audit("login", {"username": u["username"]})
    return jsonify({"ok": True, "user": public_user(u)})


@app.route("/api/auth/logout", methods=["POST"])
def auth_logout():
    u = current_user()
    log_audit("logout", {"username": u["username"] if u else None})
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/auth/change-password", methods=["POST"])
def auth_change_password():
    user = current_user()
    payload = request.json or {}
    if not check_password_hash(user["password_hash"], payload.get("old_password") or ""):
        return jsonify({"ok": False, "error": "Ancien mot de passe incorrect"}), 400
    new_pw = payload.get("new_password") or ""
    pw_error = validate_password(new_pw)
    if pw_error:
        return jsonify({"ok": False, "error": pw_error}), 400
    users = load_users()
    for u in users:
        if u["id"] == user["id"]:
            u["password_hash"] = generate_password_hash(new_pw)
    save_users(users)
    log_audit("password_change", {"username": user["username"]})
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
    if not username:
        return jsonify({"ok": False, "error": "Identifiant requis."}), 400
    pw_error = validate_password(password)
    if pw_error:
        return jsonify({"ok": False, "error": pw_error}), 400
    users = load_users()
    if any(u["username"] == username for u in users):
        return jsonify({"ok": False, "error": "Identifiant déjà utilisé"}), 400
    new_id = max([u["id"] for u in users], default=0) + 1
    user = {"id": new_id, "username": username, "password_hash": generate_password_hash(password),
            "name": (payload.get("name") or "").strip() or username, "role": role,
            "modules": payload.get("modules") or list(MODULES.keys()), "active": True}
    users.append(user)
    save_users(users)
    log_audit("user_create", {"username": username, "role": role})
    return jsonify({"ok": True, "user": public_user(user)})


@app.route("/api/users/<int:uid>", methods=["POST"])
def update_user(uid):
    payload = request.json or {}
    if payload.get("new_password"):
        pw_error = validate_password(payload["new_password"])
        if pw_error:
            return jsonify({"ok": False, "error": pw_error}), 400
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
    log_audit("user_update", {"username": found["username"], "fields": list(payload.keys())})
    return jsonify({"ok": True, "user": public_user(found)})


@app.route("/api/users/<int:uid>", methods=["DELETE"])
def delete_user(uid):
    users = load_users()
    deleted = next((u for u in users if u["id"] == uid), None)
    users = [u for u in users if u["id"] != uid]
    save_users(users)
    log_audit("user_delete", {"username": deleted["username"] if deleted else uid})
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
    write_json_safely(notes_thread_path(hotliner_id), thread)


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


@app.route("/api/notes/search")
def notes_search():
    """Recherche dans les Notes (Lot G — recherche globale étendue). Un admin cherche dans
    tous les fils hotliner ; un hotliner ne cherche que dans son propre fil."""
    q = (request.args.get("q") or "").strip().lower()
    if not q:
        return jsonify([])
    user = current_user()
    results = []
    if user["role"] == "admin":
        hotliners = [u for u in load_users() if u.get("role") == "hotliner"]
        for h in hotliners:
            for entry in load_notes_thread(h["id"]):
                if q in (entry.get("text") or "").lower():
                    results.append({**entry, "hotliner_id": h["id"], "hotliner_name": h.get("name", h.get("username"))})
    else:
        for entry in load_notes_thread(user["id"]):
            if q in (entry.get("text") or "").lower():
                results.append({**entry, "hotliner_id": user["id"], "hotliner_name": user.get("name")})
    return jsonify(results[:20])


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
    write_json_safely(KNOWLEDGE_FILE, items)


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


@app.route("/api/<module>/last-sync")
def last_sync(module):
    """Horodatage de la dernière écriture de données pour un volet (calls/emails/analysis,
    ou timeseries/sources pour Moussanada, ou records pour PCHC) — utilisé pour l'indicateur
    'Dernière synchro' du dashboard."""
    check_module(module)
    latest = None
    d = module_dir(module)
    for root, _, files in os.walk(d):
        if os.path.basename(root) == "exports":
            continue
        for fn in files:
            if not fn.endswith(".json"):
                continue
            mtime = os.path.getmtime(os.path.join(root, fn))
            if latest is None or mtime > latest:
                latest = mtime
    if latest is None:
        return jsonify({"last_sync": None})
    return jsonify({"last_sync": datetime.fromtimestamp(latest).isoformat()})


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
    write_json_safely(_path(module, kind, ym), payload)


# ---------------------------------------------------------------------------
# TARKHISS — import direct de l'export Outlook brut (feuille "Mail", macro VBA côté client).
# Recalcule les compteurs reçus/envoyés par jour pour tous les mois présents dans le fichier
# (resynchronisation complète, à re-uploader chaque mois — le fichier grossit avec l'historique).
# ---------------------------------------------------------------------------
def email_meta_stats_file():
    return os.path.join(module_dir("tarkhiss"), "email_meta_stats.json")


def load_email_meta_stats():
    p = email_meta_stats_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def save_email_meta_stats(data):
    write_json_safely(email_meta_stats_file(), data)


def parse_tarkhiss_emails_raw(file_storage):
    """Parse l'export Outlook brut (feuille 'Mail' : Folder Path, Subject, DisplayTo, DisplayCc,
    DateTimeSent, DateTimeReceived, Importance, IsRead, HasAttachments, Preview, Id). Retourne
    (par_jour, stats_globales). \\Inbox\\ = reçus (DateTimeReceived) ; \\Sent Items\\ = envoyés
    (DateTimeSent). Les autres dossiers (Corbeille, Brouillons, Archive) sont ignorés — ce ne
    sont ni des emails reçus ni des emails effectivement envoyés au support."""
    from openpyxl import load_workbook
    wb = load_workbook(file_storage, data_only=True)
    ws = wb["Mail"] if "Mail" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    by_day = {}  # {"YYYY-MM": {"YYYY-MM-DD": {"received": n, "sent": n}}} — clé pleine date,
                 # alignée sur le format utilisé par la saisie manuelle (data/tarkhiss/emails_<ym>.json).
    events = []  # événements détaillés pour tarkhiss_meta (heures, délais de réponse, sujets…)
    total_received = total_sent = 0
    with_attachment = high_importance = counted = 0
    weekday_counts = [0] * 7  # 0=lundi

    for row in rows:
        if not row or not row[0]:
            continue
        folder = row[0]
        if folder not in ("\\Inbox\\", "\\Sent Items\\"):
            continue
        dt = row[5] if folder == "\\Inbox\\" else row[4]
        if not dt:
            continue
        if isinstance(dt, str):
            try:
                dt = datetime.fromisoformat(dt)
            except ValueError:
                continue
        ym = f"{dt.year}-{dt.month:02d}"
        day = f"{dt.year:04d}-{dt.month:02d}-{dt.day:02d}"  # clé pleine date, alignée sur la saisie manuelle
        by_day.setdefault(ym, {})
        by_day[ym].setdefault(day, {"received": 0, "sent": 0})
        if folder == "\\Inbox\\":
            by_day[ym][day]["received"] += 1
            total_received += 1
        else:
            by_day[ym][day]["sent"] += 1
            total_sent += 1
        counted += 1
        events.append({"kind": "received" if folder == "\\Inbox\\" else "sent", "dt": dt, "subject": row[1],
                       "to": row[2], "importance": row[6], "attachment": bool(row[8]), "is_read": row[7]})
        if row[8]:
            with_attachment += 1
        if row[6] == "High":
            high_importance += 1
        weekday_counts[dt.weekday()] += 1

    stats = {
        "total_received": total_received, "total_sent": total_sent,
        "counted": counted,
        "pct_with_attachment": round(with_attachment / counted * 100, 1) if counted else 0,
        "pct_high_importance": round(high_importance / counted * 100, 1) if counted else 0,
        "weekday_counts": weekday_counts,  # [Lun, Mar, Mer, Jeu, Ven, Sam, Dim]
        "months_covered": sorted(by_day.keys()),
        "last_import": datetime.now().isoformat(timespec="seconds"),
    }
    stats["_events"] = events
    return by_day, stats


def raw_imports_dir():
    d = os.path.join(module_dir("tarkhiss"), "raw_imports")
    os.makedirs(d, exist_ok=True)
    return d


def raw_imports_history_file():
    return os.path.join(module_dir("tarkhiss"), "raw_import_history.json")


def load_raw_imports_history():
    p = raw_imports_history_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_raw_import(kind, file_storage, extra=None):
    """Archive une copie du fichier brut uploadé (email .xlsm/.xlsx ou appels .csv), pour
    traçabilité/audit (RGPD : peut être purgé manuellement en supprimant data/tarkhiss/raw_imports/).
    Retourne l'entrée d'historique créée."""
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    ext = os.path.splitext(file_storage.filename or "")[1] or ".bin"
    safe_orig = secure_filename(file_storage.filename or f"import{ext}")
    saved_as = f"{kind}_{ts}_{safe_orig}"
    file_storage.stream.seek(0)
    with open(os.path.join(raw_imports_dir(), saved_as), "wb") as out:
        out.write(file_storage.stream.read())
    file_storage.stream.seek(0)
    entry = {"id": saved_as, "kind": kind, "original_name": file_storage.filename or safe_orig,
              "uploaded_at": datetime.now().isoformat(timespec="seconds"), **(extra or {})}
    history = load_raw_imports_history()
    history.insert(0, entry)
    write_json_safely(raw_imports_history_file(), history[:50])  # 50 derniers imports conservés
    return entry


@app.route("/api/tarkhiss/raw-imports")
def tarkhiss_raw_imports_list():
    return jsonify(load_raw_imports_history())


@app.route("/api/tarkhiss/raw-imports/download/<path:file_id>")
def tarkhiss_raw_imports_download(file_id):
    """Téléchargement d'un fichier brut archivé (import Outlook ou journal d'appels).
    Le nom est validé contre l'historique pour empêcher tout accès hors du dossier dédié."""
    known = {e["id"]: e for e in load_raw_imports_history()}
    safe_id = secure_filename(file_id)
    entry = known.get(safe_id)
    if not entry:
        return jsonify({"error": "Fichier non référencé dans l'historique des imports"}), 404
    fpath = os.path.join(raw_imports_dir(), safe_id)
    if not os.path.exists(fpath):
        return jsonify({"error": "Fichier introuvable sur le disque (a-t-il été nettoyé ?)"}), 404
    return send_file(fpath, as_attachment=True, download_name=entry["original_name"])


@app.route("/api/tarkhiss/import-emails-raw", methods=["POST"])
def tarkhiss_import_emails_raw():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Aucun fichier fourni"}), 400
    f = request.files["file"]
    try:
        by_day, stats = parse_tarkhiss_emails_raw(f)
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erreur d'analyse du fichier : {e}"}), 400
    if not stats["months_covered"]:
        return jsonify({"ok": False, "error": "Aucun email trouvé dans \\Inbox\\ ou \\Sent Items\\ — vérifiez la feuille 'Mail' et la colonne Folder Path."}), 400

    for ym, days in by_day.items():
        emails = _load("tarkhiss", "emails", ym, {})
        for day, counts in days.items():
            emails[day] = {"received": counts["received"], "sent": counts["sent"]}
        _save("tarkhiss", "emails", ym, emails)

    events = stats.pop("_events", [])
    for ym, meta in tm.aggregate_emails(events).items():
        _save("tarkhiss", "meta_emails", ym, meta)
    save_email_meta_stats(stats)
    save_raw_import("emails", f, {"months": stats["months_covered"],
                                   "total_received": stats["total_received"], "total_sent": stats["total_sent"]})
    log_audit("import_emails_raw", {"file": f.filename, "months": stats["months_covered"],
                                     "total_received": stats["total_received"], "total_sent": stats["total_sent"]},
              module="tarkhiss")
    return jsonify({"ok": True, "stats": stats})


@app.route("/api/tarkhiss/email-meta-stats")
def tarkhiss_email_meta_stats():
    return jsonify(load_email_meta_stats() or {})


@app.route("/api/tarkhiss/email-meta/<ym>")
def tarkhiss_email_meta_month(ym):
    return jsonify(_load("tarkhiss", "meta_emails", ym, {}))


@app.route("/api/tarkhiss/call-meta/<ym>")
def tarkhiss_call_meta_month(ym):
    return jsonify(_load("tarkhiss", "meta_calls", ym, {}))


@app.route("/api/tarkhiss/call-meta-stats")
def tarkhiss_call_meta_stats():
    """Synthèse tous mois confondus du journal d'appels importé (compteurs sommables + série mensuelle)."""
    d = module_dir("tarkhiss")
    months = []
    for fname in sorted(os.listdir(d)):
        if fname.startswith("meta_calls_") and fname.endswith(".json"):
            m = _load("tarkhiss", "meta_calls", fname[len("meta_calls_"):-5], {})
            if m:
                c = m.get("counts", {})
                months.append({"ym": m["ym"], "in": c.get("in", 0), "missed": c.get("missed", 0),
                               "out": c.get("out", 0), "answer_rate": m.get("answer_rate"),
                               "median_sec": (m.get("duration") or {}).get("median_sec")})
    if not months:
        return jsonify({})
    tin, tmiss = sum(x["in"] for x in months), sum(x["missed"] for x in months)
    return jsonify({"months": months, "total_in": tin, "total_missed": tmiss,
                    "total_out": sum(x["out"] for x in months),
                    "answer_rate": round(tin / (tin + tmiss) * 100, 1) if (tin + tmiss) else None})


@app.route("/api/tarkhiss/call-heatmap-hourly-png/<ym>")
def tarkhiss_call_heatmap_hourly_png(ym):
    meta = _load("tarkhiss", "meta_calls", ym, {})
    if not meta:
        abort(404)
    buf = chart_heatmap_png(meta["hourly"]["in"], ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
                             [str(h) for h in range(24)], "Appels reçus — jour × heure",
                             figsize=(9, 3), cbar_label="Nb d'appels")
    return send_file(buf, mimetype="image/png")


@app.route("/api/tarkhiss/email-heatmap-hourly-png/<ym>")
def tarkhiss_email_heatmap_hourly_png(ym):
    meta = _load("tarkhiss", "meta_emails", ym, {})
    if not meta:
        abort(404)
    buf = chart_heatmap_png(meta["hourly"]["received"], ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
                             [str(h) for h in range(24)], "Emails reçus — jour × heure",
                             figsize=(9, 3), cbar_label="Nb d'emails")
    return send_file(buf, mimetype="image/png")


@app.route("/api/tarkhiss/import-calls-raw", methods=["POST"])
def tarkhiss_import_calls_raw():
    """Import du journal d'appels brut (CSV du mobile : Name, Phone, Date, Type, Duration…).
    Fichier cumulatif : chaque mois présent est resynchronisé (jours présents écrasés)."""
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Aucun fichier fourni"}), 400
    f = request.files["file"]
    raw_bytes = f.read()
    f.stream.seek(0)
    try:
        calls, skipped = tm.parse_calllog_csv(raw_bytes)
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erreur d'analyse du fichier : {e}"}), 400
    if not calls:
        return jsonify({"ok": False, "error": "Aucun appel exploitable dans le fichier (dates/types non reconnus)."}), 400
    agg = tm.aggregate_calls(calls)
    for ym, block in agg.items():
        existing = _load("tarkhiss", "calls", ym, {})
        for day, rec in block["days"].items():
            merged = dict(existing.get(day, {}))
            merged.update({"calls": rec["calls"], "duration_sec": rec["duration_sec"]})
            merged.pop("duration_min", None)
            existing[day] = merged
        if block["days"]:
            _save("tarkhiss", "calls", ym, existing)
        _save("tarkhiss", "meta_calls", ym, block["meta"])
    tot = {k: sum(b["meta"]["counts"][k] for b in agg.values()) for k in ("in", "missed", "out", "blocked", "rejected")}
    resp = {"months_covered": sorted(agg.keys()), "rows": len(calls), "skipped": skipped,
            "total_in": tot["in"], "total_missed": tot["missed"], "total_out": tot["out"],
            "total_blocked": tot["blocked"], "total_rejected": tot["rejected"],
            "answer_rate": round(tot["in"] / (tot["in"] + tot["missed"]) * 100, 1) if (tot["in"] + tot["missed"]) else None}
    save_raw_import("calls", f, {"months": resp["months_covered"], "rows": len(calls)})
    log_audit("import_calls_raw", {"file": f.filename, "months": resp["months_covered"], "rows": len(calls)}, module="tarkhiss")
    return jsonify({"ok": True, "stats": resp})


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


def compute_tarkhiss_heatmap(calls, ym):
    """Heatmap 'jour de semaine x semaine du mois' du volume d'appels — limite honnête :
    la saisie Tarkhiss est au jour (pas à l'heure), donc contrairement à la heatmap Moussanada
    (jour x heure, à partir des horodatages réels GLPI), celle-ci ne peut pas descendre en
    dessous de la granularité journalière. Elle reste utile pour repérer les jours de semaine
    à forte charge sur le mois."""
    y, m = (int(x) for x in ym.split("-"))
    import calendar
    _, days_in_month = calendar.monthrange(y, m)
    weeks = []
    matrix = []
    current_week = None
    row = None
    for day in range(1, days_in_month + 1):
        wd = datetime(y, m, day).weekday()  # 0=lundi
        week_of_month = (day + datetime(y, m, 1).weekday() - 1) // 7 + 1
        if week_of_month != current_week:
            current_week = week_of_month
            row = [0] * 7
            matrix.append(row)
            weeks.append(f"Sem. {week_of_month}")
        count = (calls.get(f"{y}-{m:02d}-{day:02d}") or {}).get("calls", 0)
        row[wd] = count
    return matrix, weeks


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
        "ui_theme": "flat",             # "flat", "soft", "neu" ou "clay" — écran uniquement, sans effet sur PDF/Excel
        "ui_palette": "ammps",          # "ammps", "ocean", "emerald", "slate", "violet", "crimson"
        # Grands titres des rapports (email/PDF/Excel/PPTX) — édition avancée, Administration
        "report_title_tarkhiss": "Rapport Support Tarkhiss",
        "report_title_moussanada": "Rapport Support Moussanada",
        "report_title_pchc": "Reporting Métier — PCHC",
        "report_show_contact_names": False,  # PDF/Excel Tarkhiss : noms des contacts (RGPD) — masqués par défaut
        # SMTP optionnel — si renseigné, permet l'envoi RÉEL du rappel programmé par email
        # (sans lui, le rappel reste un simple bandeau à l'ouverture de l'app, comme documenté).
        "smtp_host": "", "smtp_port": 587, "smtp_user": "", "smtp_password": "",
        "smtp_use_tls": True, "smtp_from": "",
        "reminder_api_token": "",  # jeton partagé pour l'appel headless (tâche planifiée)
        # Sécurité (Lot 4)
        "session_timeout_minutes": 240,   # déconnexion automatique après N minutes d'inactivité
        "password_min_length": 8,
        "password_require_digit": True,
        "password_require_upper": True,
        # Intégration GLPI (Lot A) — API REST legacy (apirest.php), en complément de l'import CSV
        "glpi_url": "", "glpi_app_token": "", "glpi_user_token": "", "glpi_entity_id": "",
    }
    if os.path.exists(GLOBAL_SETTINGS_FILE):
        with open(GLOBAL_SETTINGS_FILE, "r", encoding="utf-8") as f:
            defaults.update(json.load(f))
    return defaults


def save_global_settings(payload):
    write_json_safely(GLOBAL_SETTINGS_FILE, payload)


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
        "status_colors": {},
        "pchc_thresholds": {"taux_haut": 50, "taux_bas": 20, "backlog_alert": None, "old_dossiers_alert": None},
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
    write_json_safely(module_settings_file(module), payload)


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
    write_json_safely(alert_history_file(module), history[:300])


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
    write_json_safely(send_log_file(module), log[:200])


def send_email_smtp(msg):
    """Envoi réel via SMTP si configuré (Administration). Retourne (ok, error)."""
    import smtplib
    gs = load_global_settings()
    host = gs.get("smtp_host")
    if not host:
        return False, "SMTP non configuré (Administration → Notifications par email)"
    try:
        msg["From"] = gs.get("smtp_from") or gs.get("smtp_user") or "sanad@localhost"
        with smtplib.SMTP(host, int(gs.get("smtp_port") or 587), timeout=15) as server:
            if gs.get("smtp_use_tls", True):
                server.starttls()
            if gs.get("smtp_user"):
                server.login(gs["smtp_user"], gs.get("smtp_password") or "")
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)


@app.route("/api/system/reminder-check", methods=["POST"])
def system_reminder_check():
    """Point d'entrée headless — destiné à être appelé par une tâche planifiée (Windows
    Task Scheduler, cron...) une fois par jour, en dehors de toute session utilisateur.
    Envoie un email RÉEL (SMTP) si le rapport du mois précédent n'a pas encore été envoyé et
    que le jour de rappel configuré est atteint. Protégé par un jeton partagé (pas de session
    navigateur ici). Sans SMTP configuré, ne fait rien (le rappel reste visible dans l'app)."""
    gs = load_global_settings()
    expected_token = gs.get("reminder_api_token") or ""
    provided = request.headers.get("X-Reminder-Token") or (request.json or {}).get("token", "")
    if not expected_token or provided != expected_token:
        return jsonify({"ok": False, "error": "Jeton invalide ou non configuré"}), 403

    today = datetime.now()
    results = []
    for module in ("tarkhiss", "moussanada"):
        settings = load_module_settings(module)
        reminder_day = settings.get("reminder_day", 5)
        py = prev_ym(f"{today.year}-{today.month:02d}")
        already_sent = any(l.get("ym") == py for l in load_send_log(module))
        if already_sent or today.day < reminder_day:
            results.append({"module": module, "sent": False, "reason": "déjà envoyé ou pas encore le jour de rappel"})
            continue
        to_list = [c["email"] for c in settings.get("contacts", []) if c.get("role") == "to"]
        if not to_list:
            results.append({"module": module, "sent": False, "reason": "aucun destinataire configuré"})
            continue
        report_html, month_label = get_report_html_and_label(module, py, None, for_pdf=False)
        greeting = settings.get("greeting", "Bonjour,")
        body_html = build_email_body(month_label, report_html, greeting, settings, MODULES[module]["label"])
        subject = build_subject(module, month_label)
        msg = EmailMessage()
        msg["Subject"] = f"[Rappel automatique] {subject}"
        msg["To"] = "; ".join(to_list)
        cc_list = [c["email"] for c in settings.get("contacts", []) if c.get("role") == "cc"]
        if cc_list:
            msg["Cc"] = "; ".join(cc_list)
        msg.set_content("Ce message nécessite un client compatible HTML.")
        msg.add_alternative(body_html, subtype="html")
        ok, error = send_email_smtp(msg)
        if ok:
            append_send_log(module, {"date": today.isoformat(timespec="seconds"), "ym": py, "subject": subject, "to": to_list, "cc": cc_list, "opened": False, "auto": True})
            log_audit("reminder_email_sent", {"ym": py, "to": to_list}, module=module)
        results.append({"module": module, "sent": ok, "reason": error})
    return jsonify({"ok": True, "results": results})


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
    write_json_safely(timeseries_file(), data)


def sources_path(ym):
    return os.path.join(module_dir("moussanada"), f"sources_{ym}.json")


def load_sources(ym):
    p = sources_path(ym)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"categories": [], "services": [], "techniciens": [], "demandeurs": []}


def save_sources(ym, data):
    write_json_safely(sources_path(ym), data)


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
    write_json_safely(notes_path(ym), data)


# ---------------------------------------------------------------------------
# MOUSSANADA — filtre technicien (Lot D) : exclusion ou whitelist, persistant, appliqué à la
# fois à l'import (direct et CSV) et à la lecture des données pour dashboard/rapports — un
# changement de filtre affecte donc immédiatement les régénérations de rapport, sans ré-import.
# ---------------------------------------------------------------------------
def technicien_filter_file():
    return os.path.join(module_dir("moussanada"), "technicien_filter.json")


def load_technicien_filter():
    defaults = {"mode": "exclude", "technicians": []}  # mode: "exclude" ou "whitelist"
    p = technicien_filter_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            defaults.update(json.load(f))
    return defaults


def save_technicien_filter(data):
    write_json_safely(technicien_filter_file(), data)


def apply_technicien_filter(items):
    """Filtre une liste d'enregistrements techniciens ({"label": <nom>, ...}) selon la
    configuration persistante. Mode 'exclude' : retire les noms listés. Mode 'whitelist' (si la
    liste n'est pas vide) : ne garde que les noms listés. Insensible à la casse/espaces."""
    cfg = load_technicien_filter()
    names = {n.strip().lower() for n in cfg.get("technicians", []) if n.strip()}
    if not names:
        return items
    if cfg.get("mode") == "whitelist":
        return [it for it in items if (it.get("label") or "").strip().lower() in names]
    return [it for it in items if (it.get("label") or "").strip().lower() not in names]


# ---------------------------------------------------------------------------
# MOUSSANADA — tickets bruts (import ticket-par-ticket, distinct des feuilles agrégées)
# Alimente : répartition par Type/Priorité, et la heatmap de charge (jour x heure).
# ---------------------------------------------------------------------------
def raw_tickets_file():
    return os.path.join(module_dir("moussanada"), "tickets_raw.json")


def load_raw_tickets():
    p = raw_tickets_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_raw_tickets(tickets):
    write_json_safely(raw_tickets_file(), tickets)


@app.route("/api/moussanada/import-tickets", methods=["POST"])
def moussanada_import_tickets():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Aucun fichier fourni"}), 400
    f = request.files["file"]
    try:
        new_tickets = parse_glpi_raw_tickets(f)
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erreur d'analyse du fichier : {e}"}), 400
    if not new_tickets:
        return jsonify({"ok": True, "result": {"rows": 0}, "warnings": [
            "0 ticket détecté — vérifiez qu'il s'agit bien d'un export GLPI brut (ticket par ticket) "
            "avec les colonnes ID, Statut, Date d'ouverture, Priorité, Catégorie, Type, etc."
        ]})
    existing = load_raw_tickets()
    merged = merge_raw_tickets(existing, new_tickets)
    save_raw_tickets(merged)
    log_audit("import_tickets", {"file": f.filename, "new": len(new_tickets), "total": len(merged)}, module="moussanada")
    return jsonify({"ok": True, "result": {"rows": len(new_tickets), "total": len(merged)}, "warnings": []})


@app.route("/api/moussanada/heatmap")
def moussanada_heatmap():
    """Heatmap jour x heure des tickets ouverts, calculée à partir des tickets bruts importés.
    Filtrable par période (start/end au format YYYY-MM)."""
    tickets = load_raw_tickets()
    start = request.args.get("start")
    end = request.args.get("end")
    if start or end:
        def in_range(t):
            d = (t.get("date_ouverture") or "")[:7]
            if start and d < start:
                return False
            if end and d > end:
                return False
            return True
        tickets = [t for t in tickets if in_range(t)]
    matrix = compute_heatmap_matrix(tickets)
    type_counts = {}
    priority_counts = {}
    for t in tickets:
        type_counts[t.get("type") or "Non renseigné"] = type_counts.get(t.get("type") or "Non renseigné", 0) + 1
        priority_counts[t.get("priorite") or "Non renseignée"] = priority_counts.get(t.get("priorite") or "Non renseignée", 0) + 1
    return jsonify({
        "matrix": matrix, "total_tickets": len(tickets),
        "type_counts": type_counts, "priority_counts": priority_counts,
    })


@app.route("/api/moussanada/heatmap.png")
def moussanada_heatmap_png():
    """Rendu image (seaborn) de la heatmap, pour affichage direct <img> côté dashboard."""
    tickets = load_raw_tickets()
    start = request.args.get("start")
    end = request.args.get("end")
    if start or end:
        def in_range(t):
            d = (t.get("date_ouverture") or "")[:7]
            if start and d < start:
                return False
            if end and d > end:
                return False
            return True
        tickets = [t for t in tickets if in_range(t)]
    matrix = compute_heatmap_matrix(tickets)
    hours = [f"{h}h" for h in range(24)]
    days = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
    buf = chart_heatmap_png(matrix, days, hours, "Heatmap de charge — tickets ouverts par jour et heure", figsize=(11, 3.8))
    return send_file(buf, mimetype="image/png")


@app.route("/api/tarkhiss/heatmap-png/<ym>")
def tarkhiss_heatmap_png(ym):
    calls = _load("tarkhiss", "calls", ym, {})
    matrix, weeks = compute_tarkhiss_heatmap(calls, ym)
    buf = chart_heatmap_png(
        matrix, weeks, ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
        "Heatmap de charge — appels par jour de semaine et semaine du mois",
        figsize=(8.5, 3.4), cbar_label="Nombre d'appels",
    )
    return send_file(buf, mimetype="image/png")


def _filter_raw_tickets_period(tickets, start=None, end=None):
    def in_range(t):
        d = (t.get("date_ouverture") or "")[:7]
        if start and d < start:
            return False
        if end and d > end:
            return False
        return True
    return [t for t in tickets if in_range(t)] if (start or end) else tickets


@app.route("/api/moussanada/tickets-analytics")
def moussanada_tickets_analytics():
    """Analyses tirées des tickets bruts GLPI (Type, Priorité, distribution des délais, tickets
    en souffrance) — ce que le module reporting natif de GLPI n'agrège pas. Filtrable par
    période (start/end au format YYYY-MM)."""
    tickets = _filter_raw_tickets_period(load_raw_tickets(), request.args.get("start"), request.args.get("end"))

    type_counts = {}
    priority_counts = {}
    for t in tickets:
        type_counts[t.get("type") or "Non renseigné"] = type_counts.get(t.get("type") or "Non renseigné", 0) + 1
        priority_counts[t.get("priorite") or "Non renseignée"] = priority_counts.get(t.get("priorite") or "Non renseignée", 0) + 1

    durations = compute_resolution_hours(tickets)
    delay_buckets = bucket_resolution_hours(durations)

    open_statuses = {"En attente", "En cours (Attribué)", "En cours (Planifié)", "Nouveau"}
    en_souffrance = sorted(
        [t for t in tickets if (t.get("statut") or "") in open_statuses or (t.get("statut") not in ("Résolu", "Clos") and t.get("statut"))],
        key=lambda t: t.get("date_ouverture") or "",
    )

    return jsonify({
        "total_tickets": len(tickets),
        "type_counts": type_counts,
        "priority_counts": priority_counts,
        "delay_buckets": delay_buckets,
        "avg_resolution_h": round(sum(durations) / len(durations), 1) if durations else 0,
        "tickets_en_souffrance": [
            {"id": t["id"], "titre": t.get("titre", ""), "statut": t.get("statut", ""),
             "date_ouverture": t.get("date_ouverture", ""), "technicien": t.get("technicien", ""),
             "priorite": t.get("priorite", "")}
            for t in en_souffrance
        ],
    })


# ---------------------------------------------------------------------------
# MOUSSANADA — détection de doublons d'import (Lot F). Les profils "par étiquette"
# (categories/services/techniciens/demandeurs) REMPLACENT intégralement les données du mois à
# chaque import (voir store_source_items) — silencieusement si CSV et Direct sont utilisés en
# alternance pour le même mois. On avertit donc avant d'écraser un import fait dans l'AUTRE mode.
REPLACE_SEMANTICS_KINDS = {"categories", "services", "techniciens", "demandeurs"}


def import_history_file():
    return os.path.join(module_dir("moussanada"), "import_history.json")


def load_import_history():
    p = import_history_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def record_import_history(kind, ym, mode, source):
    hist = load_import_history()
    hist.insert(0, {"kind": kind, "ym": ym, "mode": mode, "source": source,
                     "date": datetime.now().isoformat(timespec="seconds")})
    write_json_safely(import_history_file(), hist[:300])


def find_conflicting_import(kind, ym, mode):
    """Retourne la dernière entrée d'import connue pour (kind, ym) faite dans l'AUTRE mode
    (CSV vs Direct), s'il y en a une — signe qu'écraser maintenant perdrait une source
    différente sans avertissement."""
    if kind not in REPLACE_SEMANTICS_KINDS:
        return None
    for entry in load_import_history():
        if entry["kind"] == kind and entry["ym"] == ym and entry["mode"] != mode:
            return entry
    return None


@app.route("/api/moussanada/import/<ym>", methods=["POST"])
def moussanada_import(ym):
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "Aucun fichier reçu"}), 400
    kind = request.form.get("kind", "auto")
    force = str(request.form.get("force", "")).lower() in ("1", "true", "yes")

    conflict = find_conflicting_import(kind, ym, "csv")
    if conflict and not force:
        return jsonify({
            "ok": False, "conflict": True,
            "message": f"Ce mois a déjà été importé en Mode Direct GLPI le {conflict['date'][:16].replace('T',' ')} "
                       f"pour '{kind}'. Continuer en Mode CSV remplacera intégralement ces données.",
            "previous": conflict,
        }), 409

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
    record_import_history(kind, ym, "csv", f.filename)
    log_audit("import", {"file": f.filename, "kind": kind, "result": result}, module="moussanada")
    warnings = []
    if isinstance(result, dict) and not any(v for k, v in result.items() if k != "detected_month"):
        warnings.append("Aucune donnée n'a été détectée dans ce fichier — vérifiez qu'il s'agit bien d'un export GLPI avec les feuilles/en-têtes attendus.")
    return jsonify({"ok": True, "result": result, "warnings": warnings})


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
            "techniciens": apply_technicien_filter(src.get("techniciens", [])),
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


# ---------------------------------------------------------------------------
# EXCEL — moteur xlsxwriter + graphiques seaborn (rendus en image, style "moderne")
# ---------------------------------------------------------------------------
SEABORN_PALETTE = ["#0B4965", "#F59F0A", "#25935F", "#DC2828", "#1794CF", "#8B5CF6", "#F2994A", "#94A3AD"]


def _seaborn_setup():
    import matplotlib
    matplotlib.use("Agg")
    import seaborn as sns
    sns.set_theme(style="whitegrid", font="DejaVu Sans")
    sns.set_palette(SEABORN_PALETTE)


def _fig_to_png_buf(fig):
    import matplotlib.pyplot as plt
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def chart_bar_png(categories, series, title, ylabel="", figsize=(7.5, 4.2)):
    """series: dict {libellé_série: [valeurs...]} — barres groupées."""
    import numpy as np
    import matplotlib.pyplot as plt
    _seaborn_setup()
    fig, ax = plt.subplots(figsize=figsize)
    x = np.arange(len(categories))
    n = max(len(series), 1)
    width = 0.8 / n
    for i, (label, values) in enumerate(series.items()):
        ax.bar(x + i * width - 0.4 + width / 2, values, width=width, label=label, color=SEABORN_PALETTE[i % len(SEABORN_PALETTE)])
    ax.set_xticks(x)
    ax.set_xticklabels(categories, rotation=30, ha="right", fontsize=9)
    ax.set_title(title, fontsize=13, fontweight="bold", color="#0B4965", pad=12)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10)
    if n > 1:
        ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.28), ncol=min(n, 4), fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return _fig_to_png_buf(fig)


def chart_pie_png(labels, values, title, figsize=(6.2, 4.6)):
    import matplotlib.pyplot as plt
    _seaborn_setup()
    fig, ax = plt.subplots(figsize=figsize)
    total = sum(values)
    if not total:
        ax.text(0.5, 0.5, "Aucune donnée pour cette période", ha="center", va="center", fontsize=11, color="#94A3AD", transform=ax.transAxes)
        ax.set_title(title, fontsize=13, fontweight="bold", color="#0B4965", pad=12)
        ax.axis("off")
        fig.tight_layout()
        return _fig_to_png_buf(fig)
    colors = [SEABORN_PALETTE[i % len(SEABORN_PALETTE)] for i in range(len(values))]
    wedges, _texts, _autotexts = ax.pie(
        values, autopct=lambda p: f"{p:.0f}%" if p > 0 else "", colors=colors,
        startangle=90, pctdistance=0.78,
        wedgeprops=dict(width=0.55, edgecolor="white", linewidth=1.5),
        textprops=dict(color="#0D1926", fontsize=9, fontweight="bold"),
    )
    ax.set_title(title, fontsize=13, fontweight="bold", color="#0B4965", pad=12)
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.02, 0.5), frameon=False, fontsize=9)
    fig.tight_layout()
    return _fig_to_png_buf(fig)


def chart_line_png(categories, series, title, ylabel="", figsize=(8, 4.2)):
    import matplotlib.pyplot as plt
    _seaborn_setup()
    fig, ax = plt.subplots(figsize=figsize)
    for i, (label, values) in enumerate(series.items()):
        ax.plot(categories, values, marker="o", linewidth=2.4, markersize=5, label=label, color=SEABORN_PALETTE[i % len(SEABORN_PALETTE)])
    ax.set_title(title, fontsize=13, fontweight="bold", color="#0B4965", pad=12)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=10)
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right", fontsize=9)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.28), ncol=min(len(series), 4), fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    return _fig_to_png_buf(fig)


def chart_combo_png(categories, bar_series, line_series, bar_ylabel="", line_ylabel="", title="", figsize=(9, 6.2)):
    """Barres empilées (bar_series) + ligne sur axe secondaire (line_series, une seule série)."""
    import numpy as np
    import matplotlib.pyplot as plt
    _seaborn_setup()
    fig, ax1 = plt.subplots(figsize=figsize)
    bottom = np.zeros(len(categories))
    for i, (label, values) in enumerate(bar_series.items()):
        values = np.array(values, dtype=float)
        ax1.bar(categories, values, bottom=bottom, label=label, color=SEABORN_PALETTE[i % len(SEABORN_PALETTE)])
        bottom += values
    ax1.set_ylabel(bar_ylabel, fontsize=10)
    ax1.tick_params(axis="x", labelrotation=40)
    for lbl in ax1.get_xticklabels():
        lbl.set_ha("right")
        lbl.set_fontsize(9)
    ax1.spines["top"].set_visible(False)

    ax2 = ax1.twinx()
    line_label, line_values = next(iter(line_series.items()))
    ax2.plot(categories, line_values, marker="o", linewidth=2.6, markersize=6, color="#F59F0A", label=line_label)
    ax2.set_ylabel(line_ylabel, fontsize=10)
    ax2.set_ylim(0, max(100, max(line_values, default=0) * 1.15))
    ax2.grid(False)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.5), ncol=2, fontsize=9)
    ax1.set_title(title, fontsize=13, fontweight="bold", color="#0B4965", pad=12)
    fig.subplots_adjust(bottom=0.38)
    return _fig_to_png_buf(fig)


def chart_heatmap_png(matrix, row_labels, col_labels, title, figsize=(10, 3.6), cbar_label="Nombre de tickets"):
    """Heatmap générique (jour x heure, ou jour x semaine). matrix: liste de listes [rows][cols]."""
    import numpy as np
    import matplotlib.pyplot as plt
    _seaborn_setup()
    import seaborn as sns
    fig, ax = plt.subplots(figsize=figsize)
    data = np.array(matrix, dtype=float)
    sns.heatmap(
        data, ax=ax, cmap="YlOrRd", linewidths=0.6, linecolor="white",
        xticklabels=col_labels, yticklabels=row_labels,
        cbar_kws={"label": cbar_label, "shrink": 0.8},
        annot=data.shape[0] * data.shape[1] <= 60, fmt=".0f", annot_kws={"fontsize": 8},
    )
    ax.set_title(title, fontsize=13, fontweight="bold", color="#0B4965", pad=12)
    plt.setp(ax.get_xticklabels(), rotation=0, fontsize=8)
    plt.setp(ax.get_yticklabels(), rotation=0, fontsize=9)
    fig.tight_layout()
    return _fig_to_png_buf(fig)


def xlsx_formats(workbook):
    """Formats xlsxwriter partagés par tous les exports Excel (remplace les styles openpyxl)."""
    return {
        "title": workbook.add_format({"font_size": 14, "bold": True, "font_color": "#0B4965"}),
        "header": workbook.add_format({"bold": True, "bg_color": "#0B4965", "font_color": "#FFFFFF", "border": 1}),
        "bold": workbook.add_format({"bold": True}),
        "default": workbook.add_format({}),
        "pct": workbook.add_format({"num_format": "0.0%"}),
    }


def xlsx_insert_png(worksheet, row, col, png_buf, scale=0.72):
    worksheet.insert_image(row, col, "chart.png", {"image_data": png_buf, "x_scale": scale, "y_scale": scale})


def write_alerts_sheet_xw(workbook, fmts, module):
    ws = workbook.add_worksheet("Alertes")
    for i, h in enumerate(["Date", "Mois", "Alerte(s) déclenchée(s)"]):
        ws.write(0, i, h, fmts["header"])
    row = 1
    for entry in load_alert_history(module):
        for alert in entry.get("alerts", []):
            ws.write(row, 0, entry.get("date", ""))
            ws.write(row, 1, month_label_fr(entry.get("ym", "")) if entry.get("ym") else "")
            ws.write(row, 2, alert)
            row += 1
    ws.set_column("A:A", 14)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 70)


def export_xlsx_moussanada(ym):
    import xlsxwriter

    data = moussanada_data(ym).get_json()
    m = data["month"]
    ts = data["timeseries"]
    month_label = month_label_fr(ym)

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    fmts = xlsx_formats(wb)

    ws = wb.add_worksheet("KPI")
    ws.merge_range("A1:D1", f"{load_global_settings().get('report_title_moussanada') or 'Rapport Moussanada'} — {month_label}", fmts["title"])
    labels = [("Tickets ouverts", m["tickets"]["ouverts"]), ("Tickets résolus", m["tickets"]["resolus"]),
              ("Tickets en retard", m["tickets"]["en_retard"]), ("Tickets clos", m["tickets"]["clos"]),
              ("Délai moyen résolution (h)", m["durations"]["resolution_h"]),
              ("Délai moyen clôture (h)", m["durations"]["cloture_h"])]
    for i, (lbl, val) in enumerate(labels):
        ws.write(2 + i, 0, lbl, fmts["bold"])
        ws.write(2 + i, 1, val)
    ws.set_column("A:A", 30)
    ws.set_column("B:B", 14)

    def write_sheet(name, items, cols):
        ws2 = wb.add_worksheet(name)
        for i, h in enumerate(cols):
            ws2.write(0, i, h, fmts["header"])
        for r, it in enumerate(items, start=1):
            ws2.write(r, 0, it.get("label", ""))
            ws2.write(r, 1, it.get("ouverts", 0))
            ws2.write(r, 2, it.get("resolus", 0))
            ws2.write(r, 3, it.get("en_retard", 0))
            ws2.write(r, 4, it.get("fermes", 0))
        ws2.set_column("A:A", 42)
        ws2.set_column("B:E", 12)
        return ws2

    cat_sheet = write_sheet("Catégories", m["categories"], ["Catégorie", "Ouverts", "Résolus", "En retard", "Fermés"])
    write_sheet("Services", m["services"], ["Service", "Ouverts", "Résolus", "En retard", "Fermés"])
    write_sheet("Techniciens", m["techniciens"], ["Technicien", "Ouverts", "Résolus", "En retard", "Fermés"])
    write_sheet("Demandeurs", m["demandeurs"], ["Demandeur", "Ouverts", "Résolus", "En retard", "Fermés"])

    top_cats = sorted(m["categories"], key=lambda c: -(c.get("ouverts", 0)))[:10]
    if top_cats:
        png = chart_bar_png(
            [c["label"] for c in top_cats],
            {"Ouverts": [c.get("ouverts", 0) for c in top_cats], "Résolus": [c.get("resolus", 0) for c in top_cats]},
            "Top catégories — Ouverts / Résolus",
        )
        xlsx_insert_png(cat_sheet, 1, 7, png)

    ws3 = wb.add_worksheet("Évolution mensuelle")
    for i, h in enumerate(["Mois", "Ouverts", "Résolus", "En retard", "Clos", "Délai résolution (h)", "Délai clôture (h)"]):
        ws3.write(0, i, h, fmts["header"])
    sorted_keys = sorted(ts.keys())
    for r, key in enumerate(sorted_keys, start=1):
        row = ts[key]
        ws3.write(r, 0, key)
        ws3.write(r, 1, row.get("ouverts", 0))
        ws3.write(r, 2, row.get("resolus", 0))
        ws3.write(r, 3, row.get("en_retard", 0))
        ws3.write(r, 4, row.get("clos", 0))
        ws3.write(r, 5, row.get("resolution_h", 0))
        ws3.write(r, 6, row.get("cloture_h", 0))
    ws3.set_column("A:G", 14)

    if sorted_keys:
        png = chart_line_png(
            sorted_keys,
            {
                "Ouverts": [ts[k].get("ouverts", 0) for k in sorted_keys],
                "Résolus": [ts[k].get("resolus", 0) for k in sorted_keys],
                "Clos": [ts[k].get("clos", 0) for k in sorted_keys],
            },
            "Évolution mensuelle — Ouverts / Résolus / Clos", ylabel="Nombre",
        )
        xlsx_insert_png(ws3, 1, 8, png)

    write_alerts_sheet_xw(wb, fmts, "moussanada")

    wb.close()
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"dashboard_moussanada_{ym}.xlsx",
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# DOSSIERS PCHC — stockage, import, KPI, rapports
# ---------------------------------------------------------------------------
def pchc_records_file():
    return os.path.join(module_dir("pchc"), "records.json")


def load_pchc_records():
    p = pchc_records_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"categories": {k: [] for k in PCHC_CATEGORIES}, "last_import": None}


def save_pchc_records(data):
    write_json_safely(pchc_records_file(), data)


def pchc_status_colors_map():
    settings = load_module_settings("pchc")
    return settings.get("status_colors", {})


def pchc_notes_file():
    return os.path.join(module_dir("pchc"), "notes.json")


def load_pchc_notes():
    p = pchc_notes_file()
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"copilot_text": ""}


def save_pchc_notes(data):
    write_json_safely(pchc_notes_file(), data)


@app.route("/api/pchc/notes", methods=["GET"])
def pchc_get_notes():
    return jsonify(load_pchc_notes())


@app.route("/api/pchc/notes", methods=["POST"])
def pchc_save_notes():
    payload = request.json or {}
    save_pchc_notes({"copilot_text": payload.get("copilot_text", "")})
    return jsonify({"ok": True})


SPECIALIZATIONS = {
    "pchc": {"label": "PCHC", "full_label": "Produits Cosmétiques et d'Hygiène Corporelle", "ready": True},
    "ca": {"label": "CA", "full_label": "Compléments Alimentaires", "ready": False},
    "dm": {"label": "DM", "full_label": "Dispositifs Médicaux et Puériculture", "ready": False},
}


@app.route("/api/pchc/specializations")
def pchc_specializations():
    return jsonify(SPECIALIZATIONS)


def pchc_default_period():
    today = datetime.now()
    return f"{today.year}-01-01", today.strftime("%Y-%m-%d")


def filter_by_period(records, start, end):
    return [r for r in records if r.get("date_depot") and start <= r["date_depot"] <= end]


def backlog_bucket(date_depot, today):
    try:
        d = datetime.strptime(date_depot, "%Y-%m-%d")
    except (ValueError, TypeError):
        return None
    age = (today - d).days
    if age <= 30:
        return "0-30"
    if age <= 60:
        return "31-60"
    if age <= 90:
        return "61-90"
    return ">90"


def compute_pchc_dashboard(start, end):
    data = load_pchc_records()
    cmap = pchc_status_colors_map()
    settings = load_module_settings("pchc")
    th = settings.get("pchc_thresholds", {"taux_haut": 50, "taux_bas": 20, "backlog_alert": None, "old_dossiers_alert": None})
    today = datetime.now()

    categories_out = {}
    global_total = global_delivered = global_encours = global_bloque = 0
    entity_totals = {}
    entity_delivered = {}
    entity_encours = {}
    entity_refuse = {}
    backlog_global = {"0-30": 0, "31-60": 0, "61-90": 0, ">90": 0}
    alerts = []

    for cat_key, cfg in PCHC_CATEGORIES.items():
        all_recs = data["categories"].get(cat_key, [])
        recs = filter_by_period(all_recs, start, end)
        total = len(recs)
        delivered = encours = bloque = 0
        status_counts = {}
        backlog_cat = {"0-30": 0, "31-60": 0, "61-90": 0, ">90": 0}

        for r in recs:
            color = pchc_status_color(r.get("statut", ""), cmap)
            status_counts.setdefault(r.get("statut", "?"), {"count": 0, "color": color})
            status_counts[r.get("statut", "?")]["count"] += 1

            entity = r.get("entity") or "—"
            entity_totals[entity] = entity_totals.get(entity, 0) + 1

            if color == "green":
                delivered += 1
                entity_delivered[entity] = entity_delivered.get(entity, 0) + 1
            elif color == "red":
                bloque += 1
                entity_refuse[entity] = entity_refuse.get(entity, 0) + 1
            else:
                encours += 1
                entity_encours[entity] = entity_encours.get(entity, 0) + 1
                b = backlog_bucket(r.get("date_depot"), today)
                if b:
                    backlog_cat[b] += 1
                    backlog_global[b] += 1

        taux = round(delivered / total * 100, 1) if total else 0
        if taux >= th.get("taux_haut", 50):
            voyant = "green"
        elif taux >= th.get("taux_bas", 20):
            voyant = "yellow"
        else:
            voyant = "red"

        categories_out[cat_key] = {
            "label": cfg["label"], "total": total, "delivered": delivered,
            "encours": encours, "bloque": bloque, "taux": taux, "voyant": voyant,
            "status_counts": status_counts, "backlog": backlog_cat,
        }
        global_total += total
        global_delivered += delivered
        global_encours += encours
        global_bloque += bloque

        if total and th.get("taux_bas") is not None and taux < th.get("taux_bas"):
            alerts.append(f"{cfg['label']} : taux de délivrance critique ({taux}%)")

    global_taux = round(global_delivered / global_total * 100, 1) if global_total else 0
    old_dossiers = backlog_global.get(">90", 0)
    if th.get("old_dossiers_alert") is not None and old_dossiers > th.get("old_dossiers_alert"):
        alerts.append(f"{old_dossiers} dossiers en cours depuis plus de 90 jours")
    if th.get("backlog_alert") is not None and global_encours > th.get("backlog_alert"):
        alerts.append(f"Stock de dossiers en cours élevé ({global_encours})")

    def top10(d):
        return sorted(d.items(), key=lambda x: x[1], reverse=True)[:10]

    return {
        "start": start, "end": end,
        "categories": categories_out,
        "summary": {
            "total": global_total, "delivered": global_delivered, "encours": global_encours,
            "bloque": global_bloque, "taux": global_taux,
        },
        "backlog": backlog_global,
        "top10": {
            "total": [{"entity": k, "count": v} for k, v in top10(entity_totals)],
            "delivered": [{"entity": k, "count": v} for k, v in top10(entity_delivered)],
            "encours": [{"entity": k, "count": v} for k, v in top10(entity_encours)],
            "refuse": [{"entity": k, "count": v} for k, v in top10(entity_refuse)],
        },
        "alerts": alerts,
        "last_import": data.get("last_import"),
    }


@app.route("/api/pchc/import", methods=["POST"])
def pchc_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "error": "Aucun fichier reçu"}), 400
    kind = request.form.get("kind", "auto")
    ext = os.path.splitext(f.filename)[1].lower()
    data = load_pchc_records()
    result = {}
    try:
        if ext == ".xlsx":
            parsed = _parse_pchc_xlsx(f)
            if not parsed:
                return jsonify({"ok": False, "error": "Aucune feuille reconnue dans ce fichier"}), 400
            for cat_key, recs in parsed.items():
                data["categories"][cat_key] = recs
                result[cat_key] = len(recs)
        elif ext == ".csv":
            if kind == "auto" or kind not in PCHC_CATEGORIES:
                return jsonify({"ok": False, "error": "Précisez la catégorie pour un CSV"}), 400
            recs = _parse_pchc_csv(f, kind)
            data["categories"][kind] = recs
            result[kind] = len(recs)
        else:
            return jsonify({"ok": False, "error": "Format non supporté (.xlsx, .csv)"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"Erreur d'analyse du fichier : {e}"}), 400
    data["last_import"] = datetime.now().isoformat(timespec="seconds")
    save_pchc_records(data)
    log_audit("import", {"file": f.filename, "kind": kind, "result": result}, module="pchc")

    warnings = []
    if not any(result.values()):
        warnings.append("0 ligne importée — vérifiez le nom des feuilles/colonnes attendues pour cette catégorie.")

    from pchc_import import DEFAULT_STATUS_COLORS
    custom_map = pchc_status_colors_map()
    known = set(pchc_normalize_key(k) for k in custom_map.keys()) | set(DEFAULT_STATUS_COLORS.keys())
    all_statuses = set()
    new_statuses = set()
    for cat_key in result.keys():
        for r in data["categories"].get(cat_key, []):
            if r.get("statut"):
                all_statuses.add(r["statut"])
                if pchc_normalize_key(r["statut"]) not in known:
                    new_statuses.add(r["statut"])

    return jsonify({
        "ok": True, "result": result, "last_import": data["last_import"],
        "statuses_found": sorted(all_statuses), "new_statuses": sorted(new_statuses),
        "warnings": warnings,
    })


@app.route("/api/pchc/data")
def pchc_data():
    start = request.args.get("start") or pchc_default_period()[0]
    end = request.args.get("end") or pchc_default_period()[1]
    return jsonify(compute_pchc_dashboard(start, end))


@app.route("/api/pchc/records/<category>")
def pchc_records_endpoint(category):
    if category not in PCHC_CATEGORIES:
        return jsonify({"error": "catégorie invalide"}), 400
    start = request.args.get("start") or pchc_default_period()[0]
    end = request.args.get("end") or pchc_default_period()[1]
    data = load_pchc_records()
    recs = filter_by_period(data["categories"].get(category, []), start, end)
    cmap = pchc_status_colors_map()
    for r in recs:
        r["color"] = pchc_status_color(r.get("statut", ""), cmap)
    return jsonify(recs)


@app.route("/api/pchc/status-colors", methods=["GET"])
def pchc_get_status_colors():
    from pchc_import import DEFAULT_STATUS_COLORS
    settings = load_module_settings("pchc")
    custom = settings.get("status_colors", {})
    data = load_pchc_records()
    observed = set()
    for recs in data["categories"].values():
        for r in recs:
            if r.get("statut"):
                observed.add(r["statut"])
    merged = {}
    for s in observed:
        merged[s] = custom.get(s) or DEFAULT_STATUS_COLORS.get(normalize_header(s).replace(" ", ""), None) or pchc_status_color(s, custom)
    return jsonify({"mapping": merged, "palette": PCHC_COLOR_HEX})


@app.route("/api/pchc/status-colors", methods=["POST"])
def pchc_save_status_colors():
    payload = request.json or {}
    settings = load_module_settings("pchc")
    settings["status_colors"] = payload.get("mapping", {})
    save_module_settings("pchc", settings)
    return jsonify({"ok": True})


def pchc_report_filename(end_date_str=None):
    ref = datetime.now()
    first_of_month = ref.replace(day=1)
    last_day_prev_month = first_of_month - timedelta(days=1)
    return f"Tarkhiss_Rapport_{last_day_prev_month.strftime('%Y%m%d')}.xlsx"


@app.route("/api/pchc/export-xlsx")
def pchc_export_xlsx():
    start = request.args.get("start") or pchc_default_period()[0]
    end = request.args.get("end") or pchc_default_period()[1]
    buf = build_pchc_xlsx(start, end)
    return send_file(buf, as_attachment=True, download_name=pchc_report_filename(),
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def build_pchc_xlsx(start, end):
    import xlsxwriter

    d = compute_pchc_dashboard(start, end)
    data = load_pchc_records()

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    fmts = xlsx_formats(wb)

    ws = wb.add_worksheet("Dashboard Exécutif")
    ws.merge_range("A1:E1", f"{load_global_settings().get('report_title_pchc') or 'Reporting Métier PCHC'} — {start} au {end}", fmts["title"])
    rows = [
        ("Total dossiers", d["summary"]["total"]),
        ("Délivrés / Validés", d["summary"]["delivered"]),
        ("En cours", d["summary"]["encours"]),
        ("Bloqués / Refusés", d["summary"]["bloque"]),
        ("Taux de délivrance global", f"{d['summary']['taux']}%"),
    ]
    for i, (label, val) in enumerate(rows, start=2):
        ws.write(i, 0, label, fmts["bold"])
        ws.write(i, 1, val)
    row = len(rows) + 4

    ws.write(row, 0, "Backlog par ancienneté", fmts["title"])
    row += 1
    for i, h in enumerate(["Ancienneté", "Nombre"]):
        ws.write(row, i, h, fmts["header"])
    row += 1
    backlog_labels = ["0-30j", "31-60j", "61-90j", ">90j"]
    backlog_values = [d["backlog"]["0-30"], d["backlog"]["31-60"], d["backlog"]["61-90"], d["backlog"][">90"]]
    for label, val in zip(backlog_labels, backlog_values):
        ws.write(row, 0, label)
        ws.write(row, 1, val)
        row += 1

    summary_row = row + 1
    ws.write(summary_row, 0, "Synthèse globale", fmts["title"])
    summary_header_row = summary_row + 1
    for i, h in enumerate(["Statut", "Nombre"]):
        ws.write(summary_header_row, i, h, fmts["header"])
    summary_labels = ["Délivrés", "En cours", "Bloqués/Refusés"]
    summary_values = [d["summary"]["delivered"], d["summary"]["encours"], d["summary"]["bloque"]]
    for i, (label, val) in enumerate(zip(summary_labels, summary_values)):
        ws.write(summary_header_row + 1 + i, 0, label)
        ws.write(summary_header_row + 1 + i, 1, val)

    ws.set_column("A:A", 32)
    ws.set_column("B:E", 14)

    xlsx_insert_png(ws, 2, 3, chart_bar_png(backlog_labels, {"Dossiers": backlog_values}, "Backlog par ancienneté", ylabel="Nombre", figsize=(6, 3.8)))
    xlsx_insert_png(ws, 20, 3, chart_pie_png(summary_labels, summary_values, "Synthèse globale des dossiers"))

    # Synthèse par module — table + graphique combo (barres empilées + ligne taux de délivrance)
    synth_row = summary_header_row + 6
    ws.write(synth_row, 0, "Synthèse par module", fmts["title"])
    synth_header_row = synth_row + 1
    for i, h in enumerate(["Module", "Délivrés", "En cours", "Bloqués", "Taux (%)"]):
        ws.write(synth_header_row, i, h, fmts["header"])
    synth_labels, synth_delivered, synth_encours, synth_bloque, synth_taux = [], [], [], [], []
    for i, (cat_key, cfg) in enumerate(PCHC_CATEGORIES.items()):
        cat = d["categories"][cat_key]
        r = synth_header_row + 1 + i
        ws.write(r, 0, cfg["label"])
        ws.write(r, 1, cat["delivered"])
        ws.write(r, 2, cat["encours"])
        ws.write(r, 3, cat["bloque"])
        ws.write(r, 4, cat["taux"])
        synth_labels.append(cfg["label"]); synth_delivered.append(cat["delivered"])
        synth_encours.append(cat["encours"]); synth_bloque.append(cat["bloque"]); synth_taux.append(cat["taux"])

    if synth_labels:
        png = chart_combo_png(
            synth_labels,
            {"Délivrés": synth_delivered, "En cours": synth_encours, "Bloqués": synth_bloque},
            {"Taux de délivrance": synth_taux},
            bar_ylabel="Nombre de dossiers", line_ylabel="Taux (%)",
            title="Synthèse par module — volume & taux de délivrance",
        )
        xlsx_insert_png(ws, synth_header_row + len(synth_labels) + 3, 0, png, scale=0.85)

    for cat_key, cfg in PCHC_CATEGORIES.items():
        cat = d["categories"][cat_key]
        ws2 = wb.add_worksheet(f"KPI {cfg['label'][:22]}")
        ws2.write(0, 0, cfg["label"], fmts["title"])
        stats = [("Total", cat["total"]), ("Délivrés", cat["delivered"]), ("En cours", cat["encours"]),
                 ("Bloqués/Refusés", cat["bloque"]), ("Taux de délivrance", f"{cat['taux']}%")]
        for i, (label, val) in enumerate(stats, start=2):
            ws2.write(i, 0, label, fmts["bold"])
            ws2.write(i, 1, val)
        row2 = len(stats) + 4
        for i, h in enumerate(["Statut", "Nombre"]):
            ws2.write(row2, i, h, fmts["header"])
        row2 += 1
        status_labels, status_values = [], []
        for statut, info in sorted(cat["status_counts"].items(), key=lambda x: -x[1]["count"]):
            ws2.write(row2, 0, statut)
            ws2.write(row2, 1, info["count"])
            status_labels.append(statut); status_values.append(info["count"])
            row2 += 1
        ws2.set_column("A:A", 40)
        ws2.set_column("B:B", 14)

        if status_labels:
            png = chart_pie_png(status_labels, status_values, f"Répartition des statuts — {cfg['label']}")
            xlsx_insert_png(ws2, 2, 3, png)

    for cat_key, cfg in PCHC_CATEGORIES.items():
        recs = filter_by_period(data["categories"].get(cat_key, []), start, end)
        ws3 = wb.add_worksheet(f"Brut {cat_key}"[:31])
        headers = ["Référence", "Entité", "Détails", "Date dépôt", "Statut"]
        for i, h in enumerate(headers):
            ws3.write(0, i, h, fmts["header"])
        for r, rec in enumerate(recs, start=1):
            ws3.write(r, 0, rec.get("ref", ""))
            ws3.write(r, 1, rec.get("entity", ""))
            ws3.write(r, 2, rec.get("details", ""))
            ws3.write(r, 3, rec.get("date_depot", ""))
            ws3.write(r, 4, rec.get("statut", ""))
        for col, width in zip("ABCDE", [22, 30, 40, 14, 22]):
            ws3.set_column(f"{col}:{col}", width)

    wb.close()
    buf.seek(0)
    return buf


def build_report_html_pchc(start, end, charts=None, for_pdf=False):
    charts = charts or {}
    d = compute_pchc_dashboard(start, end)
    s = d["summary"]
    notes = load_pchc_notes()
    global_settings = load_global_settings()

    def esc(x):
        return html_module.escape(str(x), quote=True) if x is not None else ""

    def kpi_card(label, value, bg="#F5F7F9", color="#0B4965"):
        return f"""<td style="width:20%;padding:10px;text-align:center;background:{bg};border:1px solid #DAE0E7;">
          <div style="font-size:10.5px;color:#55616B;">{esc(label)}</div>
          <div style="font-size:17px;font-weight:bold;color:{color};">{esc(value)}</div>
        </td>"""

    def img_tag(key, title=""):
        b64 = charts.get(key)
        if not b64:
            return ""
        caption = f'<div style="font-size:11.5px;color:#55616B;text-align:center;margin:4px 0 14px;">{esc(title)}</div>' if title else ""
        return f'<img src="{b64}" style="max-width:100%;border-radius:8px;margin:14px 0 0;border:1px solid #DAE0E7;display:block;">{caption}'

    voyant_emoji = {"green": "🟢", "yellow": "🟡", "red": "🔴"}

    cat_rows = ""
    for cat_key, cfg in PCHC_CATEGORIES.items():
        cat = d["categories"][cat_key]
        cat_rows += f"""<tr>
          <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-size:12.5px;">{voyant_emoji.get(cat['voyant'],'')} {esc(cfg['label'])}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;text-align:center;font-weight:bold;">{cat['total']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;text-align:center;color:#25935F;">{cat['delivered']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;text-align:center;color:#1794CF;">{cat['encours']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;text-align:center;color:#DC2828;">{cat['bloque']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;text-align:center;font-weight:bold;">{cat['taux']}%</td>
        </tr>"""

    alerts_html = ""
    if d["alerts"]:
        items = "".join(f"<li style='margin-bottom:5px;font-size:13px;color:#8A2A32;'>{esc(a)}</li>" for a in d["alerts"])
        alerts_html = f"""<div style="background:#FCE9E9;border-left:4px solid #DC2828;border-radius:0 8px 8px 0;padding:12px 16px;margin:16px 0;">
          <div style="font-weight:bold;color:#8A2A32;font-size:13px;margin-bottom:6px;">⚠️ Alertes</div>
          <ul style="margin:0;padding-left:18px;">{items}</ul>
        </div>"""

    html = f"""
    <div style="max-width:760px;margin:0 auto;font-family:Arial,sans-serif;background:#FFFFFF;">
      <div style="background:#0B4965;padding:22px 24px;border-radius:6px 6px 0 0;">
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">{esc(global_settings.get('report_title_pchc') or 'Reporting Métier — PCHC')}</div>
        <div style="color:#B9D3E0;font-size:13px;margin-top:4px;">{esc(global_settings.get('agency_name',''))} — Gestion des dossiers réglementaires — du {esc(start)} au {esc(end)}</div>
      </div>
      <div style="padding:20px 24px;border:1px solid #DAE0E7;border-top:none;">

        <table style="width:100%;border-collapse:collapse;margin-bottom:16px;"><tr>
          {kpi_card("TOTAL DOSSIERS", s["total"])}
          {kpi_card("DÉLIVRÉS", s["delivered"], bg="#E7F5EE", color="#25935F")}
          {kpi_card("EN COURS", s["encours"], bg="#E6F5FB", color="#1794CF")}
          {kpi_card("BLOQUÉS/REFUSÉS", s["bloque"], bg="#FCE9E9", color="#DC2828")}
          {kpi_card("TAUX DE DÉLIVRANCE", f"{s['taux']}%")}
        </tr></table>

        {alerts_html}

        {f'''<div class="pdf-break"></div>
        <div style="margin:6px 0 4px;padding:10px 14px;background:#F5F7F9;border:1px solid #DAE0E7;border-radius:6px;">
          <div style="font-size:10.5px;font-weight:800;color:#0B4965;letter-spacing:.4px;text-transform:uppercase;margin-bottom:6px;">Sommaire</div>
          <a href="#sec-graphs" style="display:block;font-size:12px;color:#135A7D;text-decoration:none;padding:2px 0;">1. Analyse graphique</a>
          <a href="#sec-synth" style="display:block;font-size:12px;color:#135A7D;text-decoration:none;padding:2px 0;">2. Synthèse par module</a>
          {'<a href="#sec-analysis" style="display:block;font-size:12px;color:#135A7D;text-decoration:none;padding:2px 0;">3. Analyse</a>' if notes.get("copilot_text") else ""}
        </div>''' if for_pdf else ""}

        <a name="sec-graphs"></a>
        <div class="pdf-break"></div>
        <div style="font-size:15px;font-weight:800;color:#0B4965;margin:4px 0 4px;">Analyse graphique</div>
        {img_tag("summary", "Répartition globale des statuts")}
        {img_tag("backlog", "Ancienneté du backlog")}

        <a name="sec-synth"></a>
        <div class="pdf-break"></div>
        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Synthèse par module</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Module</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Total</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Délivrés</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">En cours</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Bloqués</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Taux</td>
          </tr>
          {cat_rows}
        </table>

        {img_tag("top10", "Top 10 opérateurs par volume de dossiers")}

        {'<a name="sec-analysis"></a><div class="pdf-break"></div><div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Analyse</div>' + "".join(f"<p style='font-size:13px;color:#0D1926;line-height:1.6;'>{esc(p)}</p>" for p in notes.get("copilot_text","").split(chr(10)) if p.strip()) if notes.get("copilot_text") else ""}

      </div>
    </div>
    """
    return html


@app.route("/api/pchc/preview-email", methods=["POST"])
def pchc_preview_email():
    payload = request.json or {}
    start = payload.get("start") or pchc_default_period()[0]
    end = payload.get("end") or pchc_default_period()[1]
    module_settings = load_module_settings("pchc")
    greeting = payload.get("greeting") or module_settings.get("greeting", "Bonjour,")
    report_html = build_report_html_pchc(start, end, payload.get("charts"))
    body_html = build_email_body(f"{start} au {end}", report_html, greeting, module_settings, "Reporting Métier (PCHC)")
    subject = f"Rapport de Synthèse Reporting Métier PCHC - du {start} au {end}"
    return jsonify({"subject": subject, "html": body_html})


@app.route("/api/pchc/send-email", methods=["POST"])
def pchc_send_email():
    payload = request.json or {}
    start = payload.get("start") or pchc_default_period()[0]
    end = payload.get("end") or pchc_default_period()[1]
    module_settings = load_module_settings("pchc")
    greeting = payload.get("greeting") or module_settings.get("greeting", "Bonjour,")
    to_list = payload.get("to") or [c["email"] for c in module_settings.get("contacts", []) if c.get("role") == "to"]
    cc_list = payload.get("cc") or [c["email"] for c in module_settings.get("contacts", []) if c.get("role") == "cc"]

    report_html = build_report_html_pchc(start, end, payload.get("charts"))
    body_html = build_email_body(f"{start} au {end}", report_html, greeting, module_settings, "Reporting Métier (PCHC)")
    subject = f"Rapport de Synthèse Reporting Métier PCHC - du {start} au {end}"

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["To"] = "; ".join(to_list)
    if cc_list:
        msg["Cc"] = "; ".join(cc_list)
    msg.set_content("Ce message nécessite un client compatible HTML.")
    msg.add_alternative(body_html, subtype="html")

    xlsx_buf = build_pchc_xlsx(start, end)
    fname_xlsx = pchc_report_filename()
    msg.add_attachment(xlsx_buf.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=fname_xlsx)

    fname = f"rapport_pchc_{start}_{end}.eml"
    fpath = os.path.join(module_dir("pchc"), "exports", fname)
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

    append_send_log("pchc", {
        "date": datetime.now().isoformat(timespec="seconds"),
        "ym": f"{start}_{end}", "subject": subject, "to": to_list, "cc": cc_list, "opened": opened,
        "file": fname,
    })
    return jsonify({"ok": True, "file": fpath, "opened": opened, "error": error, "subject": subject})


@app.route("/api/pchc/export-pptx", methods=["POST"])
def pchc_export_pptx():
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    import base64

    payload = request.json or {}
    start = payload.get("start") or pchc_default_period()[0]
    end = payload.get("end") or pchc_default_period()[1]
    charts = payload.get("charts") or {}
    d = compute_pchc_dashboard(start, end)
    s = d["summary"]
    global_settings = load_global_settings()
    NAVY = RGBColor(0x0B, 0x49, 0x65)
    AMBER = RGBColor(0xF5, 0x9F, 0x0A)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    GREY = RGBColor(0x55, 0x61, 0x6B)

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def add_bg(slide, color=WHITE):
        rect = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        rect.fill.solid()
        rect.fill.fore_color.rgb = color
        rect.line.fill.background()
        slide.shapes._spTree.remove(rect._element)
        slide.shapes._spTree.insert(2, rect._element)
        return rect

    def add_text(slide, text, left, top, width, height, size=18, bold=False, color=NAVY, align=PP_ALIGN.LEFT):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "Arial"
        return box

    # Slide 1 — page de garde
    slide = prs.slides.add_slide(blank)
    add_bg(slide, NAVY)
    add_text(slide, global_settings.get("app_name") or "SANAD", Inches(0.8), Inches(2.6), Inches(6), Inches(0.8), size=40, bold=True, color=WHITE)
    add_text(slide, global_settings.get("report_title_pchc") or "Reporting Métier — PCHC", Inches(0.8), Inches(3.4), Inches(8), Inches(0.7), size=26, bold=True, color=AMBER)
    add_text(slide, f"Période du {start} au {end}", Inches(0.8), Inches(4.1), Inches(8), Inches(0.5), size=15, color=WHITE)
    add_text(slide, global_settings.get("agency_name", ""), Inches(0.8), Inches(6.7), Inches(8), Inches(0.4), size=12, color=WHITE)

    # Slide 2 — synthèse exécutive
    slide = prs.slides.add_slide(blank)
    add_text(slide, "Synthèse exécutive", Inches(0.6), Inches(0.35), Inches(8), Inches(0.6), size=26, bold=True, color=NAVY)
    kpis = [("Total dossiers", s["total"], NAVY), ("Délivrés", s["delivered"], hex_to_rgbcolor(PCHC_COLOR_HEX["green"])),
            ("En cours", s["encours"], hex_to_rgbcolor(PCHC_COLOR_HEX["blue"])), ("Bloqués/Refusés", s["bloque"], hex_to_rgbcolor(PCHC_COLOR_HEX["red"])),
            ("Taux délivrance", f"{s['taux']}%", NAVY)]
    x = Inches(0.6)
    for label, val, color in kpis:
        box = slide.shapes.add_shape(1, x, Inches(1.2), Inches(2.3), Inches(1.3))
        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xF5, 0xF7, 0xF9)
        box.line.color.rgb = RGBColor(0xDA, 0xE0, 0xE7)
        add_text(slide, label, x + Inches(0.15), Inches(1.3), Inches(2.0), Inches(0.4), size=11, color=GREY)
        add_text(slide, str(val), x + Inches(0.15), Inches(1.7), Inches(2.0), Inches(0.7), size=26, bold=True, color=color)
        x += Inches(2.5)
    if d["alerts"]:
        add_text(slide, "⚠️ Alertes", Inches(0.6), Inches(2.9), Inches(4), Inches(0.4), size=15, bold=True, color=hex_to_rgbcolor(PCHC_COLOR_HEX["red"]))
        alert_text = "\n".join(f"• {a}" for a in d["alerts"][:6])
        add_text(slide, alert_text, Inches(0.6), Inches(3.35), Inches(11), Inches(1.8), size=13, color=NAVY)

    # Slide 3 — KPI par module
    slide = prs.slides.add_slide(blank)
    add_text(slide, "KPI par module", Inches(0.6), Inches(0.35), Inches(8), Inches(0.6), size=26, bold=True, color=NAVY)
    y = Inches(1.2)
    voyant_color = {
        "green": hex_to_rgbcolor(PCHC_COLOR_HEX["green"]),
        "yellow": hex_to_rgbcolor(PCHC_COLOR_HEX["yellow"]),
        "red": hex_to_rgbcolor(PCHC_COLOR_HEX["red"]),
    }
    for cat_key, cfg in PCHC_CATEGORIES.items():
        cat = d["categories"][cat_key]
        dot = slide.shapes.add_shape(9, Inches(0.7), y + Inches(0.12), Inches(0.18), Inches(0.18))
        dot.fill.solid(); dot.fill.fore_color.rgb = voyant_color.get(cat["voyant"], GREY)
        dot.line.fill.background()
        add_text(slide, cfg["label"], Inches(1.0), y, Inches(4.2), Inches(0.4), size=13, bold=True, color=NAVY)
        add_text(slide, f"Total: {cat['total']}  ·  Délivrés: {cat['delivered']}  ·  En cours: {cat['encours']}  ·  Bloqués: {cat['bloque']}  ·  Taux: {cat['taux']}%",
                  Inches(5.3), y, Inches(7.2), Inches(0.4), size=12, color=GREY)
        y += Inches(0.55)

    # Slides graphiques (si fournis)
    for key, title in [("summary","Répartition globale des statuts"), ("backlog","Ancienneté du backlog"), ("top10","Top 10 opérateurs")]:
        b64 = charts.get(key)
        if not b64:
            continue
        slide = prs.slides.add_slide(blank)
        add_text(slide, title, Inches(0.6), Inches(0.35), Inches(10), Inches(0.6), size=24, bold=True, color=NAVY)
        try:
            img_data = base64.b64decode(b64.split(",")[1])
            img_stream = io.BytesIO(img_data)
            slide.shapes.add_picture(img_stream, Inches(1.2), Inches(1.2), height=Inches(5.6))
        except Exception:
            pass

    # Slide finale — conclusions
    slide = prs.slides.add_slide(blank)
    add_bg(slide, NAVY)
    add_text(slide, "Conclusions", Inches(0.8), Inches(0.6), Inches(6), Inches(0.7), size=28, bold=True, color=WHITE)
    concl = f"Taux de délivrance global : {s['taux']}%\nDossiers en cours : {s['encours']}\nDossiers bloqués/refusés : {s['bloque']}"
    add_text(slide, concl, Inches(0.8), Inches(1.6), Inches(10), Inches(2), size=16, color=AMBER)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    stamp = datetime.now().strftime("%Y%m%d")
    return send_file(buf, as_attachment=True, download_name=f"SANAD_PCHC_RevueDirection_{stamp}.pptx",
                      mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


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


@app.route("/api/<module>/heatmap/<ym>")
def tarkhiss_heatmap(module, ym):
    check_module(module)
    if module != "tarkhiss":
        return jsonify({"error": "heatmap non disponible pour ce volet"}), 400
    calls = _load(module, "calls", ym, {})
    matrix, weeks = compute_tarkhiss_heatmap(calls, ym)
    return jsonify({"matrix": matrix, "weeks": weeks, "days": ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]})


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
# API - Journal d'audit (Administrateur uniquement, cf. auth_gate)
# ---------------------------------------------------------------------------
@app.route("/api/admin/audit-log")
def get_audit_log():
    limit = min(int(request.args.get("limit", 200)), 1000)
    entries = []
    if os.path.exists(AUDIT_LOG_FILE):
        with open(AUDIT_LOG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    continue
    entries.reverse()
    return jsonify(entries[:limit])


# ---------------------------------------------------------------------------
# API - RGPD (recherche / export / anonymisation d'une personne nommée dans les données)
# Périmètre couvert : tickets bruts Moussanada (champ "demandeur") et dossiers PCHC (champ
# "entity" — raison sociale/opérateur). Les feuilles agrégées mensuelles Moussanada
# (Catégories/Services/Techniciens/Demandeurs) ne sont PAS couvertes par l'anonymisation :
# ce sont des totaux par étiquette, pas des enregistrements individuels — limite assumée.
# ---------------------------------------------------------------------------
def _rgpd_search(query):
    q = (query or "").strip().lower()
    if not q:
        return {"tickets": [], "pchc_records": []}
    tickets = [t for t in load_raw_tickets() if q in (t.get("demandeur") or "").lower()]
    pchc_data = load_pchc_records()
    matches_pchc = []
    for cat_key, recs in pchc_data.get("categories", {}).items():
        for r in recs:
            if q in (r.get("entity") or "").lower():
                matches_pchc.append({**r, "categorie": cat_key})
    return {"tickets": tickets, "pchc_records": matches_pchc}


@app.route("/api/admin/rgpd/search")
def rgpd_search():
    result = _rgpd_search(request.args.get("q"))
    return jsonify({
        "query": request.args.get("q", ""),
        "tickets_count": len(result["tickets"]), "tickets": result["tickets"][:50],
        "pchc_count": len(result["pchc_records"]), "pchc_records": result["pchc_records"][:50],
    })


@app.route("/api/admin/rgpd/export")
def rgpd_export():
    q = request.args.get("q", "")
    result = _rgpd_search(q)
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["source", "champ_nom", "id_ou_ref", "statut", "date", "detail"])
    for t in result["tickets"]:
        writer.writerow(["Ticket Moussanada", t.get("demandeur", ""), t.get("id", ""), t.get("statut", ""), t.get("date_ouverture", ""), t.get("titre", "")])
    for r in result["pchc_records"]:
        writer.writerow(["Dossier PCHC", r.get("entity", ""), r.get("ref", ""), r.get("statut", ""), r.get("date_depot", ""), r.get("details", "")])
    log_audit("rgpd_export", {"query": q, "tickets": len(result["tickets"]), "pchc": len(result["pchc_records"])})
    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(mem, as_attachment=True, download_name=f"rgpd_export_{secure_filename(q) or 'recherche'}.csv", mimetype="text/csv")


@app.route("/api/admin/rgpd/anonymize", methods=["POST"])
def rgpd_anonymize():
    """Remplace le nom recherché par '[Anonymisé RGPD]' dans les tickets bruts et dossiers PCHC
    correspondants. Action irréversible (hors restauration depuis une sauvegarde horodatée) —
    tracée dans le journal d'audit avec le nombre d'enregistrements touchés."""
    q = ((request.json or {}).get("q") or "").strip()
    if not q:
        return jsonify({"ok": False, "error": "Requête vide"}), 400
    anon_label = "[Anonymisé RGPD]"

    tickets = load_raw_tickets()
    touched_tickets = 0
    for t in tickets:
        if q.lower() in (t.get("demandeur") or "").lower():
            t["demandeur"] = anon_label
            touched_tickets += 1
    if touched_tickets:
        save_raw_tickets(tickets)

    pchc_data = load_pchc_records()
    touched_pchc = 0
    for cat_key, recs in pchc_data.get("categories", {}).items():
        for r in recs:
            if q.lower() in (r.get("entity") or "").lower():
                r["entity"] = anon_label
                touched_pchc += 1
    if touched_pchc:
        save_pchc_records(pchc_data)

    log_audit("rgpd_anonymize", {"query": q, "tickets_touched": touched_tickets, "pchc_touched": touched_pchc})
    return jsonify({"ok": True, "tickets_touched": touched_tickets, "pchc_touched": touched_pchc})


# ---------------------------------------------------------------------------
# API - Intégration GLPI (Lot A — fondation : connexion, sans profils d'import pour l'instant)
# ---------------------------------------------------------------------------
@app.route("/api/admin/glpi/test-connection", methods=["POST"])
def glpi_test_connection_route():
    """Bouton "Tester la connexion" du panneau Admin > Intégrations. Utilise soit les
    paramètres déjà enregistrés, soit ceux fournis dans le corps de la requête (permet de
    tester avant d'enregistrer)."""
    payload = request.json or {}
    gs = load_global_settings()
    base_url = payload.get("glpi_url") or gs.get("glpi_url")
    app_token = payload.get("glpi_app_token") or gs.get("glpi_app_token")
    user_token = payload.get("glpi_user_token") or gs.get("glpi_user_token")
    entity_id = payload.get("glpi_entity_id") or gs.get("glpi_entity_id")
    ok, message = glpi_test_connection(base_url, app_token, user_token, entity_id or None)
    log_audit("glpi_test_connection", {"url": base_url, "ok": ok, "message": message})
    return jsonify({"ok": ok, "message": message})


@app.route("/api/admin/glpi/profiles")
def glpi_list_profiles():
    """Liste des profils d'import disponibles, pour peupler le sélecteur du panneau Admin."""
    return jsonify(glpi_profiles.list_profiles())


@app.route("/api/moussanada/technicien-filter", methods=["GET"])
def get_technicien_filter():
    return jsonify(load_technicien_filter())


@app.route("/api/moussanada/technicien-filter", methods=["POST"])
def set_technicien_filter():
    payload = request.json or {}
    mode = payload.get("mode")
    if mode not in ("exclude", "whitelist"):
        return jsonify({"ok": False, "error": "mode doit être 'exclude' ou 'whitelist'"}), 400
    technicians = [str(t).strip() for t in (payload.get("technicians") or []) if str(t).strip()]
    data = {"mode": mode, "technicians": technicians}
    save_technicien_filter(data)
    log_audit("technicien_filter_update", data, module="moussanada")
    return jsonify({"ok": True, **data})


@app.route("/api/moussanada/technicien-names")
def moussanada_technicien_names():
    """Liste (non filtrée) des noms de techniciens actuellement connus dans les données déjà
    importées, pour faciliter la saisie du filtre côté Admin (éviter de retaper les noms).
    Parcourt directement les fichiers sources_<mois>.json existants (les techniciens sont
    stockés indépendamment de la série mensuelle des tickets)."""
    names = set()
    d = module_dir("moussanada")
    for fn in os.listdir(d):
        if fn.startswith("sources_") and fn.endswith(".json"):
            ym = fn[len("sources_"):-len(".json")]
            src = load_sources(ym)
            for t in src.get("techniciens", []):
                if t.get("label"):
                    names.add(t["label"])
    return jsonify(sorted(names))


@app.route("/api/moussanada/glpi-import-direct", methods=["POST"])
def moussanada_glpi_import_direct():
    """Mode DIRECT (Lot C/E) : appel API GLPI -> agrégation -> stockage JSON SANAD, sans fichier
    intermédiaire. Converge vers les mêmes fonctions de stockage (store_ticket_series /
    store_source_items) que le mode CSV historique — voir glpi_import.py et glpi_profiles.py."""
    payload = request.json or {}
    profile_key = payload.get("profile")
    if profile_key not in glpi_profiles.PROFILES:
        return jsonify({"ok": False, "error": f"Profil inconnu : {profile_key}"}), 400

    gs = load_global_settings()
    base_url = gs.get("glpi_url")
    app_token = gs.get("glpi_app_token")
    user_token = gs.get("glpi_user_token")
    entity_id = gs.get("glpi_entity_id") or None
    if not base_url or not app_token or not user_token:
        return jsonify({"ok": False, "error": "Configuration GLPI incomplète (Administration → Intégrations GLPI)"}), 400

    date_start = payload.get("date_start")
    date_end = payload.get("date_end")
    force = bool(payload.get("force"))

    profile = glpi_profiles.PROFILES[profile_key]
    target = profile["target"]
    ym = payload.get("ym")
    if target.startswith("source_items:"):
        kind = target.split(":", 1)[1]
        if not ym:
            return jsonify({"ok": False, "error": "Paramètre 'ym' requis pour ce profil (mois cible, format YYYY-MM)"}), 400
        conflict = find_conflicting_import(kind, ym, "direct")
        if conflict and not force:
            return jsonify({
                "ok": False, "conflict": True,
                "message": f"Ce mois a déjà été importé en Mode CSV le {conflict['date'][:16].replace('T',' ')} "
                           f"pour '{kind}'. Continuer en Mode Direct remplacera intégralement ces données.",
                "previous": conflict,
            }), 409

    try:
        client = GLPIClient(base_url, app_token, user_token, entity_id)
        client.init_session()
        try:
            records, total_raw = glpi_profiles.run_profile(client, profile_key, date_start, date_end)
        finally:
            client.kill_session()
    except GLPIConnectionError as e:
        log_audit("glpi_import_direct_failed", {"profile": profile_key, "error": str(e)}, module="moussanada")
        return jsonify({"ok": False, "error": str(e)}), 502

    if target == "ticket_series":
        ts = load_timeseries()
        store_ticket_series(records, ts)
        save_timeseries(ts)
    elif target.startswith("source_items:"):
        kind = target.split(":", 1)[1]
        src = load_sources(ym)
        store_source_items(kind, records, src)
        save_sources(ym, src)
        record_import_history(kind, ym, "direct", f"GLPI:{profile_key}")
    else:
        return jsonify({"ok": False, "error": f"Cible de stockage non gérée : {target}"}), 500

    log_audit("glpi_import_direct", {"profile": profile_key, "raw_tickets": total_raw, "records": len(records)}, module="moussanada")
    return jsonify({"ok": True, "profile": profile_key, "raw_tickets": total_raw, "records_stored": len(records)})


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

    for kind in ("calls", "emails", "analysis", "meta_calls", "meta_emails"):
        p = _path(module, kind, ym)
        if os.path.exists(p):
            os.remove(p)
            removed.append(kind)
    return jsonify({"ok": True, "removed": removed})


@app.route("/api/<module>/send-log")
def get_send_log(module):
    check_module(module)
    return jsonify(load_send_log(module))


@app.route("/api/<module>/send-log/download/<path:filename>")
def download_sent_export(module, filename):
    """Téléchargement d'un email déjà généré (.eml, avec sa pièce jointe Excel/PDF incluse),
    référencé dans le journal d'envoi (Lot F). Le nom de fichier est validé contre le contenu
    du journal pour empêcher tout accès en dehors du dossier exports/ du volet."""
    check_module(module)
    known_files = {entry.get("file") for entry in load_send_log(module) if entry.get("file")}
    safe_name = secure_filename(filename)
    if safe_name not in known_files:
        return jsonify({"error": "Fichier non référencé dans le journal d'envoi de ce volet"}), 404
    fpath = os.path.join(module_dir(module), "exports", safe_name)
    if not os.path.exists(fpath):
        return jsonify({"error": "Fichier introuvable sur le disque (a-t-il été nettoyé ?)"}), 404
    return send_file(fpath, as_attachment=True, download_name=safe_name)


@app.route("/api/<module>/export-csv/<kind>/<ym>")
def export_csv(module, kind, ym):
    check_module(module)
    if kind not in ("calls", "emails", "problems", "demandes", "weekly"):
        return jsonify({"ok": False, "error": "kind invalide"}), 400
    buf = io.StringIO()
    writer = csv.writer(buf)
    if kind in ("problems", "demandes", "weekly"):
        analysis = _load(module, "analysis", ym, empty_analysis(ym))
        rows = analysis.get(kind, [])
        if kind == "weekly":
            writer.writerow(["semaine", "bugs", "demandes", "observations"])
            for w in rows:
                writer.writerow([w.get("week", ""), w.get("bugs", 0), w.get("demandes", 0), w.get("obs", "")])
        else:
            writer.writerow(["type", "nombre", "description"] + (["action"] if kind == "problems" else []))
            for r in rows:
                row = [r.get("label", ""), r.get("count", 0), r.get("desc", "")]
                if kind == "problems":
                    row.append(r.get("action", ""))
                writer.writerow(row)
        mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
        return send_file(mem, as_attachment=True, download_name=f"{module}_{kind}_{ym}.csv", mimetype="text/csv")

    data = _load(module, kind, ym, {})
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


@app.route("/api/pchc/export-csv-raw")
def pchc_export_csv_raw():
    """Export CSV brut de tous les dossiers PCHC (toutes catégories), pour analyse ad hoc
    hors SANAD (Power BI, Excel avancé...)."""
    start = request.args.get("start") or pchc_default_period()[0]
    end = request.args.get("end") or pchc_default_period()[1]
    data = load_pchc_records()
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["categorie", "reference", "entite", "details", "date_depot", "statut"])
    for cat_key, cfg in PCHC_CATEGORIES.items():
        recs = filter_by_period(data["categories"].get(cat_key, []), start, end)
        for r in recs:
            writer.writerow([cfg["label"], r.get("ref", ""), r.get("entity", ""), r.get("details", ""), r.get("date_depot", ""), r.get("statut", "")])
    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(mem, as_attachment=True, download_name=f"pchc_dossiers_{start}_{end}.csv", mimetype="text/csv")


@app.route("/api/moussanada/export-csv-tickets")
def moussanada_export_csv_tickets():
    """Export CSV brut des tickets GLPI importés (ticket par ticket), pour analyse ad hoc."""
    tickets = _filter_raw_tickets_period(load_raw_tickets(), request.args.get("start"), request.args.get("end"))
    buf = io.StringIO()
    writer = csv.writer(buf)
    cols = ["id", "titre", "statut", "date_ouverture", "priorite", "demandeur", "groupe_demandeur",
            "technicien", "categorie", "ttr", "type", "derniere_modification"]
    writer.writerow(cols)
    for t in tickets:
        writer.writerow([t.get(c, "") for c in cols])
    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(mem, as_attachment=True, download_name="moussanada_tickets_bruts.csv", mimetype="text/csv")


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


@app.route("/api/moussanada/category-trend")
def moussanada_category_trend():
    """Évolution du volume des catégories de tickets sur plusieurs mois (contrairement au Pareto
    du dashboard, figé sur le mois courant). S'appuie sur les fichiers sources_<mois>.json déjà
    conservés à chaque import — aucune donnée supplémentaire à collecter côté GLPI."""
    start_ym = request.args.get("start")
    end_ym = request.args.get("end")
    if not start_ym or not end_ym:
        ts = load_timeseries()
        known_months = sorted(ts.keys())
        if not known_months:
            return jsonify({"months": [], "categories": [], "series": {}})
        end_ym = known_months[-1]
        start_idx = max(0, len(known_months) - 6)
        start_ym = known_months[start_idx]
    months = month_range(start_ym, end_ym)

    per_month = {}
    all_labels = set()
    for ym in months:
        src = load_sources(ym)
        cats = {c["label"]: c.get("ouverts", 0) for c in src.get("categories", [])}
        per_month[ym] = cats
        all_labels.update(cats.keys())

    # Ne garder que le Top 8 des catégories (en volume cumulé) pour un graphique lisible
    totals = {label: sum(per_month[ym].get(label, 0) for ym in months) for label in all_labels}
    top_labels = [l for l, _ in sorted(totals.items(), key=lambda kv: kv[1], reverse=True)[:8]]

    series = {label: [per_month[ym].get(label, 0) for ym in months] for label in top_labels}
    return jsonify({
        "months": months, "month_labels": [month_label_fr(ym) for ym in months],
        "categories": top_labels, "series": series,
    })


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


def build_tarkhiss_advanced_stats_html(meta_calls, meta_emails, show_contact_names):
    """Section 'Statistiques avancées' (PDF/Excel/écran uniquement, jamais l'email) — issue du
    journal d'appels et de l'export Outlook bruts (heures réelles, délai de réponse, rappels,
    sujets/mots-clés fréquents). Vide si aucun des deux imports n'a été fait pour ce mois."""
    if not meta_calls and not meta_emails:
        return ""

    def esc(s):
        return html_module.escape(str(s), quote=True) if s is not None else ""

    def fmt_min(v):
        if v is None:
            return "—"
        h, m = divmod(int(round(v)), 60)
        return f"{h} h {m:02d}" if h else f"{m} min"

    def mini_table(headers, rows):
        head = "".join(f'<td style="padding:6px 8px;color:#fff;font-size:11px;font-weight:bold;">{esc(h)}</td>' for h in headers)
        body = "".join(
            "<tr>" + "".join(f'<td style="padding:5px 8px;border-bottom:1px solid #DAE0E7;font-size:11.5px;color:#0D1926;">{esc(c)}</td>' for c in row) + "</tr>"
            for row in rows
        ) or f'<tr><td colspan="{len(headers)}" style="padding:6px 8px;font-size:11.5px;color:#55616B;">Aucune donnée</td></tr>'
        return f'<table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;margin-bottom:12px;"><tr style="background:#0B4965;">{head}</tr>{body}</table>'

    blocks = []

    if meta_calls:
        c = meta_calls.get("counts", {})
        dur = meta_calls.get("duration", {})
        cb = meta_calls.get("callback", {})
        hourly_png = base64.b64encode(chart_heatmap_png(
            meta_calls["hourly"]["in"], ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"], [str(h) for h in range(24)],
            "Appels reçus — répartition jour × heure", figsize=(9.5, 3), cbar_label="Nb d'appels",
        ).read()).decode()
        callers_rows = [[(e["label"] if show_contact_names else mask_call_label(e)), e["calls"], e["missed"], format_hms(e["duration_sec"])]
                        for e in meta_calls.get("top_callers", [])[:8]]
        blocks.append(f"""
        <div style="font-size:13px;font-weight:800;color:#0B4965;margin:14px 0 8px;">☎ Journal d'appels — détail horaire</div>
        <p style="font-size:11.5px;color:#55616B;margin:0 0 8px;">Import du journal d'appels brut. Heures locales (Maroc).</p>
        <img src="data:image/png;base64,{hourly_png}" style="width:100%;max-width:640px;display:block;margin:0 auto 12px;" />
        <table style="width:100%;border-collapse:collapse;margin-bottom:12px;">
          <tr>
            <td style="padding:4px 8px;font-size:11.5px;">Taux de décroché : <strong>{meta_calls.get('answer_rate') if meta_calls.get('answer_rate') is not None else '—'}%</strong></td>
            <td style="padding:4px 8px;font-size:11.5px;">Manqués : <strong>{c.get('missed', 0)}</strong></td>
            <td style="padding:4px 8px;font-size:11.5px;">Durée médiane : <strong>{format_hms(dur.get('median_sec') or 0)}</strong></td>
            <td style="padding:4px 8px;font-size:11.5px;">Durée max : <strong>{format_hms(dur.get('max_sec') or 0)}</strong></td>
          </tr>
        </table>
        <div style="font-size:11.5px;color:#55616B;margin:0 0 6px;">Rappel des appels manqués ({cb.get('missed_total', 0)}) : {cb.get('called_back_by_us', 0)} rappelés par nous, {cb.get('client_called_again', 0)} le client a rerappelé, {cb.get('no_callback_24h', 0)} sans suite sous 24 h.</div>
        <div style="font-size:11.5px;font-weight:bold;color:#0B4965;margin:10px 0 4px;">Top appelants{'(anonymisé)' if not show_contact_names else ''}</div>
        {mini_table(["Contact", "Appels", "Manqués", "Durée cumulée"], callers_rows)}
        """)

    if meta_emails:
        resp = meta_emails.get("response", {})
        hourly_png = base64.b64encode(chart_heatmap_png(
            meta_emails["hourly"]["received"], ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"], [str(h) for h in range(24)],
            "Emails reçus — répartition jour × heure", figsize=(9.5, 3), cbar_label="Nb d'emails",
        ).read()).decode()
        subj_rows = [[s["subject"][:70], s["count"]] for s in meta_emails.get("top_subjects", [])[:6]]
        words_line = " · ".join(f"{w['word']} ({w['count']})" for w in meta_emails.get("top_words", [])[:10])
        rcpt_rows = [[(r["name"] if show_contact_names else "•••• (masqué)"), r["count"]] for r in meta_emails.get("top_recipients", [])[:6]]
        blocks.append(f"""
        <div style="font-size:13px;font-weight:800;color:#0B4965;margin:18px 0 8px;">✉ Emails — détail horaire &amp; délai de réponse</div>
        <img src="data:image/png;base64,{hourly_png}" style="width:100%;max-width:640px;display:block;margin:0 auto 12px;" />
        <table style="width:100%;border-collapse:collapse;margin-bottom:8px;">
          <tr>
            <td style="padding:4px 8px;font-size:11.5px;">Délai médian (heures ouvrées) : <strong>{fmt_min(resp.get('median_biz_min'))}</strong></td>
            <td style="padding:4px 8px;font-size:11.5px;">P90 : <strong>{fmt_min(resp.get('p90_biz_min'))}</strong></td>
            <td style="padding:4px 8px;font-size:11.5px;">Sans réponse sous 7 j : <strong>{meta_emails.get('unanswered', 0)}</strong></td>
            <td style="padding:4px 8px;font-size:11.5px;">Non lus : <strong>{meta_emails.get('pct_unread', 0)}%</strong></td>
          </tr>
        </table>
        <p style="font-size:10.5px;color:#8A94A3;margin:0 0 10px;">{esc(resp.get('note', ''))} Base : {resp.get('business_hours', '')}. Correspondance sujet↔sujet : {resp.get('match_rate') if resp.get('match_rate') is not None else '—'}% des échanges appariés.</p>
        <div style="font-size:11.5px;font-weight:bold;color:#0B4965;margin:8px 0 4px;">Sujets les plus fréquents</div>
        {mini_table(["Sujet", "Occurrences"], subj_rows)}
        <div style="font-size:11.5px;font-weight:bold;color:#0B4965;margin:8px 0 4px;">Mots-clés fréquents</div>
        <p style="font-size:11.5px;color:#0D1926;margin:0 0 10px;">{esc(words_line) or 'Aucune donnée'}</p>
        <div style="font-size:11.5px;font-weight:bold;color:#0B4965;margin:8px 0 4px;">Top destinataires{'(anonymisé)' if not show_contact_names else ''}</div>
        {mini_table(["Destinataire", "Emails envoyés"], rcpt_rows)}
        """)

    return f"""
        <a name="sec-advanced"></a>
        <div class="pdf-break"></div>
        <div style="font-size:15px;font-weight:800;color:#0B4965;margin:4px 0 4px;">Statistiques avancées — appels &amp; emails</div>
        <p style="font-size:11px;color:#8A94A3;margin:0 0 8px;">Basé sur les imports bruts (journal d'appels / export Outlook). Non inclus dans le corps de l'email.</p>
        {''.join(blocks)}
    """


def mask_call_label(entry):
    return entry["label"] if not entry.get("named") else "•••• (masqué)"


def build_report_html(module, ym, calls, emails, analysis, prev_calls=None, prev_emails=None, prev_analysis=None, charts=None, for_pdf=False):
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

    # Heatmap de charge — limite honnête : la saisie Tarkhiss est au jour (pas à l'heure), donc
    # cette heatmap croise jour de semaine x semaine du mois (pas jour x heure comme Moussanada,
    # qui dispose d'horodatages réels via l'import GLPI brut).
    heatmap_matrix, heatmap_weeks = compute_tarkhiss_heatmap(calls, ym)
    heatmap_has_data = any(any(row) for row in heatmap_matrix)
    heatmap_b64 = ""
    if heatmap_has_data:
        heatmap_b64 = base64.b64encode(chart_heatmap_png(
            heatmap_matrix, heatmap_weeks, ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
            "Heatmap de charge — appels par jour de semaine et semaine du mois",
            figsize=(8, 3.2), cbar_label="Nombre d'appels",
        ).read()).decode()

    # Statistiques avancées (journal d'appels + export Outlook brut) — issues de tarkhiss_meta.
    # N'apparaissent PAS dans l'email (for_pdf=False y est toujours utilisé) : uniquement PDF/Excel/écran.
    show_contact_names = load_global_settings().get("report_show_contact_names", False)
    adv_html = ""
    if for_pdf:
        meta_calls = _load("tarkhiss", "meta_calls", ym, {})
        meta_emails = _load("tarkhiss", "meta_emails", ym, {})
        adv_html = build_tarkhiss_advanced_stats_html(meta_calls, meta_emails, show_contact_names)

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
            return "<div style='font-size:11px;font-weight:700;color:#55616B;margin-top:2px;'>nouveau</div>"
        diff = ((cur - prev_v) / prev_v) * 100
        sign = "+" if diff >= 0 else ""
        if abs(diff) < 1 or polarity == "neutral":
            color = "#55616B"
        else:
            is_up = diff > 0
            is_good = is_up if polarity == "positive" else not is_up
            color = "#25935F" if is_good else "#DC2828"
        return f"<div style='font-size:11px;font-weight:700;color:{color};margin-top:2px;'>{sign}{diff:.0f}% vs {esc(prev_month_label)}</div>"

    def kpi_card(label, value, delta_html="", bg="#F5F7F9", color="#0B4965"):
        return f"""<td style="width:25%;padding:10px;text-align:center;background:{bg};border:1px solid #DAE0E7;">
          <div style="font-size:11px;color:#55616B;">{esc(label)}</div>
          <div style="font-size:18px;font-weight:bold;color:{color};">{esc(value)}</div>
          {delta_html}
        </td>"""

    def img_tag(key, title=""):
        b64 = charts.get(key)
        if not b64:
            return ""
        caption = f'<div style="font-size:11.5px;color:#55616B;text-align:center;margin:4px 0 14px;">{esc(title)}</div>' if title else ""
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
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#55616B;">{esc(it.get('desc',''))}{f'<br><em>Action : {esc(action)}</em>' if action else ''}</td>
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
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#55616B;">{esc(it.get('desc',''))}</td>
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
              <td style="padding:8px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#55616B;">{esc(w.get('obs',''))}</td>
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
    report_title = global_settings.get(f"report_title_{module}") or f"Rapport Support {module_label}"

    html = f"""
    <div style="max-width:760px;margin:0 auto;font-family:Arial,sans-serif;background:#FFFFFF;">
      <div style="background:#0B4965;padding:22px 24px;border-radius:6px 6px 0 0;">
        {report_logos_html(global_settings)}
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">{esc(report_title)}</div>
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

        {(lambda _entries: f'''<div class="pdf-break"></div>
        <div style="margin:6px 0 4px;padding:10px 14px;background:#F5F7F9;border:1px solid #DAE0E7;border-radius:6px;">
          <div style="font-size:10.5px;font-weight:800;color:#0B4965;letter-spacing:.4px;text-transform:uppercase;margin-bottom:6px;">Sommaire</div>
          {"".join(f'<a href="#{aid}" style="display:block;font-size:12px;color:#135A7D;text-decoration:none;padding:2px 0;">{i}. {label}</a>' for i, (aid, label) in enumerate(_entries, 1))}
        </div>''')([
            ("sec-graphs", "Analyse graphique"), ("sec-problems", "Problèmes techniques"),
            ("sec-demandes", "Demandes d'information"), ("sec-weekly", "Évolution hebdomadaire"),
            *([("sec-heatmap", "Heatmap de charge")] if heatmap_has_data else []),
            *([("sec-advanced", "Statistiques avancées — appels & emails")] if adv_html else []),
            ("sec-synth", "Constats & recommandations"),
        ]) if for_pdf else ""}

        <a name="sec-graphs"></a>
        <div class="pdf-break"></div>
        <div style="font-size:15px;font-weight:800;color:#0B4965;margin:4px 0 4px;">Analyse graphique</div>
        {img_tag("weekly", "Évolution hebdomadaire — Bugs vs Demandes")}
        {img_tag("problems", "Répartition des problèmes techniques")}

        <a name="sec-problems"></a>
        <div class="pdf-break"></div>
        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Problèmes techniques signalés ({total_problems})</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Type</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Nb</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">%</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Description</td>
          </tr>
          {rows_cat_problems(problems, total_problems)}
        </table>

        <a name="sec-demandes"></a>
        <div class="pdf-break"></div>
        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Demandes d'information reçues ({total_demandes})</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Type</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Nb</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">%</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Description</td>
          </tr>
          {rows_cat_demandes(demandes, total_demandes)}
        </table>

        <a name="sec-weekly"></a>
        <div class="pdf-break"></div>
        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Évolution hebdomadaire</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #DAE0E7;">
          <tr style="background:#0B4965;">
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Semaine</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Bugs</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;text-align:center;">Demandes</td>
            <td style="padding:8px 10px;color:#fff;font-size:12px;font-weight:bold;">Observations</td>
          </tr>
          {rows_weekly()}
        </table>

        {"<div style='font-size:14px;font-weight:bold;color:#0B4965;margin:20px 0 8px;'>Mots-clés fréquents</div><div style='font-size:12.5px;color:#55616B;margin-bottom:16px;'>" + keywords_line + "</div>" if keywords_line else ""}

        {'<a name="sec-heatmap"></a><div class="pdf-break"></div><div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Heatmap de charge</div><p style="font-size:12px;color:#55616B;margin:0 0 10px;">Volume d\'appels par jour de semaine et semaine du mois (granularité journalière — Tarkhiss ne journalise pas l\'heure des appels).</p><img src="data:image/png;base64,' + heatmap_b64 + '" style="width:100%;max-width:620px;display:block;margin:0 auto 14px;" />' if heatmap_has_data else ""}

        {adv_html}

        <a name="sec-synth"></a>
        <div class="pdf-break"></div>
        <div style="font-size:14px;font-weight:bold;color:#0B4965;margin:4px 0 8px;">Constats clés</div>
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


def build_report_html_moussanada(ym, month_label, charts=None, for_pdf=False):
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

    # Heatmap de charge (jour x heure) — seulement si des tickets bruts ont été importés pour
    # ce mois (fonctionnalité optionnelle, distincte de l'import agrégé habituel).
    raw_tickets_for_month = [t for t in load_raw_tickets() if (t.get("date_ouverture") or "")[:7] == ym]
    heatmap_matrix = compute_heatmap_matrix(raw_tickets_for_month) if raw_tickets_for_month else None
    heatmap_b64 = ""
    if heatmap_matrix:
        _hours = [f"{h}h" for h in range(24)]
        _days = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
        heatmap_b64 = base64.b64encode(chart_heatmap_png(heatmap_matrix, _days, _hours, "Heatmap de charge — tickets ouverts par jour et heure", figsize=(9.5, 3.4)).read()).decode()

    def esc(s):
        return html_module.escape(str(s), quote=True) if s is not None else ""

    def delta_span(cur, prev_v, polarity):
        if not prev_v and not cur:
            return ""
        if not prev_v:
            return "<div style='font-size:11px;font-weight:700;color:#55616B;margin-top:2px;'>nouveau</div>"
        diff = ((cur - prev_v) / prev_v) * 100
        sign = "+" if diff >= 0 else ""
        if abs(diff) < 1 or polarity == "neutral":
            color = "#55616B"
        else:
            is_up = diff > 0
            is_good = is_up if polarity == "positive" else not is_up
            color = "#25935F" if is_good else "#DC2828"
        return f"<div style='font-size:11px;font-weight:700;color:{color};margin-top:2px;'>{sign}{diff:.0f}% vs {esc(prev_month_label)}</div>"

    def kpi_card(label, value, delta_html="", bg="#F5F7F9", color="#0B4965"):
        return f"""<td style="width:25%;padding:10px;text-align:center;background:{bg};border:1px solid #DAE0E7;">
          <div style="font-size:11px;color:#55616B;">{esc(label)}</div>
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
              <td style="padding:7px 10px;border-bottom:1px solid #DAE0E7;font-family:Arial,sans-serif;font-size:12px;color:#55616B;text-align:center;">{p}%</td>
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
            return "<div style='font-size:12px;color:#55616B;font-style:italic;margin:6px 0 6px;'>— Analyse Copilot non renseignée pour cette section —</div>"
        paras = "".join(f"<p style='margin:0 0 8px;'>{esc(p)}</p>" for p in text.split("\n") if p.strip())
        return f"""<div style="background:#E7EEF2;border-left:4px solid #0B4965;border-radius:0 8px 8px 0;padding:12px 16px;margin:10px 0 4px;font-size:13px;color:#0D1926;line-height:1.6;">{paras}</div>"""

    def section_title(n, title):
        # En PDF, la section 1 doit démarrer sur sa propre page (après le sommaire, sur sa page
        # dédiée) ; sans sommaire (email), la section 1 reste à la suite de l'en-tête.
        break_div = '<div class="pdf-break"></div>' if (n > 1 or for_pdf) else ""
        return f"""<a name="sec-{n}"></a>{break_div}<div style="margin:4px 0 10px;padding-bottom:6px;border-bottom:2px solid #0B4965;">
          <span style="font-size:10.5px;font-weight:800;color:#F59F0A;letter-spacing:.6px;">SECTION {n}</span>
          <div style="font-size:17px;font-weight:800;color:#0B4965;">{esc(title)}</div>
        </div>"""

    moussanada_toc_titles = [
        "Volume global des tickets", "Évolution des délais de traitement",
        f"Principales catégories de demandes en {month_label}", "Services les plus demandeurs",
        "Charge et mobilisation de l'équipe SI", f"Analyse comparative {prev_month_label} / {month_label}",
    ]
    if heatmap_matrix:
        moussanada_toc_titles.append("Heatmap de charge (jour x heure)")
    moussanada_toc_titles.append("Synthèse finale et conclusion")
    moussanada_toc_html = f"""<div class="pdf-break"></div><div style="margin:10px 0 4px;padding:10px 14px;background:#F5F7F9;border:1px solid #DAE0E7;border-radius:6px;">
      <div style="font-size:10.5px;font-weight:800;color:#0B4965;letter-spacing:.4px;text-transform:uppercase;margin-bottom:6px;">Sommaire</div>
      {"".join(f'<a href="#sec-{i+1}" style="display:block;font-size:12px;color:#135A7D;text-decoration:none;padding:2px 0;">{i+1}. {esc(t)}</a>' for i, t in enumerate(moussanada_toc_titles))}
    </div>""" if for_pdf else ""

    sections = notes.get("sections", {})
    global_settings = load_global_settings()
    report_title = global_settings.get("report_title_moussanada") or "Rapport Support Moussanada"

    html = f"""
    <div style="max-width:760px;margin:0 auto;font-family:Arial,sans-serif;background:#FFFFFF;">
      <div style="background:#0B4965;padding:22px 24px;border-radius:6px 6px 0 0;">
        {report_logos_html(global_settings)}
        <div style="color:#FFFFFF;font-size:20px;font-weight:bold;">{esc(report_title)}</div>
        <div style="color:#B9D3E0;font-size:13px;margin-top:4px;">{esc(global_settings.get('agency_name',''))} — Helpdesk GLPI — {esc(month_label)}</div>
      </div>
      <div style="padding:20px 24px;border:1px solid #DAE0E7;border-top:none;">

        {moussanada_toc_html}

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
        <p style="font-size:12.5px;color:#55616B;margin:0 0 4px;">Comparatif des volumes et des délais moyens entre les deux mois.</p>
        {img_tag("volume")}
        {img_tag("delay")}
        {copilot_box(sections.get("comparative", ""))}

        {section_title(7, "Heatmap de charge (jour x heure)") if heatmap_matrix else ""}
        {"<p style='font-size:12.5px;color:#55616B;margin:0 0 10px;'>Répartition des tickets ouverts par jour de semaine et heure de la journée — basée sur les tickets bruts importés (colonne Date d'ouverture). Utile pour dimensionner les plages de présence hotline.</p><img src='data:image/png;base64," + heatmap_b64 + "' style='width:100%;max-width:680px;display:block;margin:0 auto 14px;' />" if heatmap_matrix else ""}

        {section_title(8 if heatmap_matrix else 7, "Synthèse finale et conclusion")}
        {("<div style='font-size:13px;font-weight:bold;color:#0B4965;margin:6px 0 6px;'>Constats clés</div><ul style='padding-left:18px;margin:0 0 14px;'>" + list_html(notes.get('constats', [])) + "</ul>") if notes.get('constats') else ""}
        {("<div style='font-size:13px;font-weight:bold;color:#0B4965;margin:6px 0 6px;'>Recommandations</div><ul style='padding-left:18px;margin:0 0 14px;'>" + list_html(notes.get('recommandations', [])) + "</ul>") if notes.get('recommandations') else ""}
        {copilot_box(sections.get("synthesis", ""))}

      </div>
    </div>
    """
    return html


def get_report_html_and_label(module, ym, charts=None, for_pdf=False):
    if module == "moussanada":
        month_label = month_label_fr(ym)
        return build_report_html_moussanada(ym, month_label, charts, for_pdf=for_pdf), month_label
    calls = _load(module, "calls", ym, {})
    emails = _load(module, "emails", ym, {})
    analysis = _load(module, "analysis", ym, empty_analysis(ym))
    py = prev_ym(ym)
    prev_calls = _load(module, "calls", py, {})
    prev_emails = _load(module, "emails", py, {})
    prev_analysis = _load(module, "analysis", py, empty_analysis(py))
    month_label = analysis.get("month_label") or month_label_fr(ym)
    return build_report_html(module, ym, calls, emails, analysis, prev_calls, prev_emails, prev_analysis, charts, for_pdf=for_pdf), month_label


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
        "file": fname,
    })

    return jsonify({"ok": True, "file": fpath, "opened": opened, "error": error, "subject": subject})


# ---------------------------------------------------------------------------
# API - Export Excel du dashboard (par volet)
# ---------------------------------------------------------------------------
def render_pdf_bytes(inner_html, landscape=False, header_label=None):
    """Génère le PDF final. Utilise WeasyPrint si disponible (meilleur rendu CSS moderne,
    flexbox/grid, en-tête récurrent via CSS Paged Media standard) ; sinon repli automatique
    sur xhtml2pdf (moteur historique, zéro dépendance système). WeasyPrint nécessite GTK3 sous
    Windows — voir le README pour l'installation ; sans ça, l'app continue de fonctionner
    normalement avec xhtml2pdf."""
    size = "A4 landscape" if landscape else "A4"
    header_block = ""
    if header_label:
        header_block = f"""<div id="pdfRunningHeader" style="font-size:9px;color:#55616B;
          border-bottom:1px solid #DAE0E7;padding-bottom:4px;">
          <strong style="color:#0B4965;">SANAD — AMMPS/DSID</strong> &nbsp;·&nbsp; {header_label}
        </div>"""

    try:
        import weasyprint
        top_margin = "2.1cm" if header_label else "1.4cm"
        header_css = ""
        if header_label:
            header_css = """
              @top-left { content: element(pdfrunhead); }
            """
            header_block = f'<div style="position:running(pdfrunhead);">{header_block}</div>'
        doc = f"""<html><head><meta charset="utf-8">
        <style>
          @page {{ size: {size}; margin: {top_margin} 1.4cm 1.4cm 1.4cm;{header_css} }}
          body {{ font-family: Arial, sans-serif; }}
          img {{ max-width: 100%; }}
          .pdf-break {{ page-break-before: always; }}
        </style>
        </head><body>{header_block}{inner_html}</body></html>"""
        pdf_bytes = weasyprint.HTML(string=doc).write_pdf()
        buf = io.BytesIO(pdf_bytes)
        buf.seek(0)
        return buf
    except Exception:
        pass  # WeasyPrint absent ou en échec (ex. GTK3 manquant sous Windows) → repli xhtml2pdf

    from xhtml2pdf import pisa
    frame_rule = ""
    if header_label:
        frame_rule = """
          @frame header_frame {
            -pdf-frame-content: pdfRunningHeader;
            top: 0.6cm; left: 1.4cm; width: 18cm; height: 0.9cm;
          }"""
    doc = f"""<html><head><meta charset="utf-8">
    <style>
      @page {{ size: {size}; margin: {"2.1cm" if header_label else "1.4cm"} 1.4cm 1.4cm 1.4cm;{frame_rule} }}
      body {{ font-family: Arial, sans-serif; }}
      img {{ max-width: 100%; }}
      .pdf-break {{ page-break-before: always; }}
    </style>
    </head><body>{header_block}{inner_html}</body></html>"""
    buf = io.BytesIO()
    pisa.CreatePDF(src=doc, dest=buf, encoding="utf-8")
    buf.seek(0)
    return buf


@app.route("/api/<module>/export-pdf/<ym>", methods=["POST"])
def export_pdf(module, ym):
    check_module(module)
    payload = request.json or {}
    report_html, month_label = get_report_html_and_label(module, ym, payload.get("charts"), for_pdf=True)
    header_label = f"{MODULES.get(module,{}).get('label', module)} — {month_label}"
    buf = render_pdf_bytes(report_html, header_label=header_label)
    return send_file(buf, as_attachment=True, download_name=f"rapport_{module}_{ym}.pdf", mimetype="application/pdf")


@app.route("/api/pchc/export-pdf", methods=["POST"])
def pchc_export_pdf():
    payload = request.json or {}
    start = payload.get("start") or pchc_default_period()[0]
    end = payload.get("end") or pchc_default_period()[1]
    report_html = build_report_html_pchc(start, end, payload.get("charts"), for_pdf=True)
    header_label = f"Reporting Métier — du {start} au {end}"
    buf = render_pdf_bytes(report_html, header_label=header_label)
    return send_file(buf, as_attachment=True, download_name=f"rapport_pchc_{start}_{end}.pdf", mimetype="application/pdf")


@app.route("/api/<module>/export-xlsx/<ym>")
def export_xlsx(module, ym):
    check_module(module)
    import xlsxwriter

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

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    fmts = xlsx_formats(wb)
    ws = wb.add_worksheet("Dashboard")

    _gs = load_global_settings()
    _report_title = _gs.get(f"report_title_{module}") or f"Rapport Support {MODULES[module]['label']}"
    ws.merge_range("A1:F1", f"{_report_title} — {month_label}", fmts["title"])

    ws.write(2, 0, "Emails reçus", fmts["bold"]); ws.write(2, 1, total_received)
    ws.write(3, 0, "Emails envoyés", fmts["bold"]); ws.write(3, 1, total_sent)
    ws.write(4, 0, "Total appels", fmts["bold"]); ws.write(4, 1, total_calls)
    ws.write(5, 0, "Temps comm. (hh:mm:ss)", fmts["bold"]); ws.write(5, 1, format_hms(total_duration_sec))

    row = 7
    ws.write(row, 0, "Problèmes techniques", fmts["title"])
    row += 1
    for i, h in enumerate(["Type", "Nombre", "%", "Description", "Action corrective"]):
        ws.write(row, i, h, fmts["header"])
    total_p = sum(p.get("count", 0) for p in problems)
    row += 1
    problem_labels, problem_values = [], []
    for p in problems:
        pct = round(p.get("count", 0) / total_p * 100, 1) if total_p else 0
        ws.write(row, 0, p.get("label", "")); ws.write(row, 1, p.get("count", 0))
        ws.write(row, 2, f"{pct}%"); ws.write(row, 3, p.get("desc", "")); ws.write(row, 4, p.get("action", ""))
        problem_labels.append(p.get("label", "")); problem_values.append(p.get("count", 0))
        row += 1

    row += 1
    ws.write(row, 0, "Demandes d'information", fmts["title"])
    row += 1
    for i, h in enumerate(["Type", "Nombre", "%", "Description"]):
        ws.write(row, i, h, fmts["header"])
    total_d = sum(d.get("count", 0) for d in demandes)
    row += 1
    for d in demandes:
        pct = round(d.get("count", 0) / total_d * 100, 1) if total_d else 0
        ws.write(row, 0, d.get("label", "")); ws.write(row, 1, d.get("count", 0))
        ws.write(row, 2, f"{pct}%"); ws.write(row, 3, d.get("desc", ""))
        row += 1

    row += 1
    ws.write(row, 0, "Évolution hebdomadaire", fmts["title"])
    row += 1
    for i, h in enumerate(["Semaine", "Bugs", "Demandes", "Observations"]):
        ws.write(row, i, h, fmts["header"])
    row += 1
    weekly_labels, weekly_bugs, weekly_demandes = [], [], []
    for w in weekly:
        ws.write(row, 0, w.get("week", "")); ws.write(row, 1, w.get("bugs", 0))
        ws.write(row, 2, w.get("demandes", 0)); ws.write(row, 3, w.get("obs", ""))
        weekly_labels.append(w.get("week", "")); weekly_bugs.append(w.get("bugs", 0)); weekly_demandes.append(w.get("demandes", 0))
        row += 1

    ws.set_column("A:A", 26)
    ws.set_column("B:B", 12)
    ws.set_column("C:C", 10)
    ws.set_column("D:D", 40)
    ws.set_column("E:E", 26)

    if weekly_labels:
        png = chart_bar_png(weekly_labels, {"Bugs": weekly_bugs, "Demandes": weekly_demandes}, "Évolution hebdomadaire — Bugs vs Demandes", ylabel="Nombre")
        xlsx_insert_png(ws, 2, 7, png)

    if problem_labels:
        png = chart_pie_png(problem_labels, problem_values, "Répartition des problèmes techniques")
        xlsx_insert_png(ws, 20, 7, png)

    if module == "tarkhiss":
        write_tarkhiss_advanced_sheet_xw(wb, fmts, ym)

    write_alerts_sheet_xw(wb, fmts, module)

    wb.close()
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f"dashboard_{module}_{ym}.xlsx",
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def write_tarkhiss_advanced_sheet_xw(workbook, fmts, ym):
    """Feuille 'Stats avancées' (appels + emails bruts) — absente si aucun des deux imports
    n'a été fait pour ce mois. Jamais reprise dans le corps de l'email (Excel uniquement)."""
    meta_calls = _load("tarkhiss", "meta_calls", ym, {})
    meta_emails = _load("tarkhiss", "meta_emails", ym, {})
    if not meta_calls and not meta_emails:
        return
    show_names = load_global_settings().get("report_show_contact_names", False)
    ws = workbook.add_worksheet("Stats avancées")
    ws.set_column("A:A", 30)
    ws.set_column("B:E", 16)
    row = 0
    ws.write(row, 0, "Statistiques avancées — journal d'appels & export Outlook", fmts["title"]); row += 2

    if meta_calls:
        c, dur, cb = meta_calls.get("counts", {}), meta_calls.get("duration", {}), meta_calls.get("callback", {})
        ws.write(row, 0, "Appels — synthèse", fmts["title"]); row += 1
        for label, val in [
            ("Reçus (décrochés)", c.get("in", 0)), ("Manqués", c.get("missed", 0)),
            ("Émis", c.get("out", 0)), ("Bloqués", c.get("blocked", 0)),
            ("Taux de décroché", f"{meta_calls.get('answer_rate')}%" if meta_calls.get("answer_rate") is not None else "—"),
            ("Durée moyenne", format_hms(dur.get("avg_sec") or 0)), ("Durée médiane", format_hms(dur.get("median_sec") or 0)),
            ("Durée P90", format_hms(dur.get("p90_sec") or 0)), ("Durée max (un appel)", format_hms(dur.get("max_sec") or 0)),
            ("Appelants uniques", meta_calls.get("unique_callers", 0)),
            ("Appelants récurrents (≥5 appels)", meta_calls.get("recurrent_callers", 0)),
            ("Rappelés par nous (sous 24h)", cb.get("called_back_by_us", 0)),
            ("Client a rerappelé (sous 24h)", cb.get("client_called_again", 0)),
            ("Manqués sans suite (24h)", cb.get("no_callback_24h", 0)),
        ]:
            ws.write(row, 0, label); ws.write(row, 1, val); row += 1
        row += 1
        ws.write(row, 0, "Distribution des durées", fmts["header"]); ws.write(row, 1, "Nb appels", fmts["header"]); row += 1
        for d in dur.get("distribution", []):
            ws.write(row, 0, d["label"]); ws.write(row, 1, d["count"]); row += 1
        row += 1
        ws.write(row, 0, "Top appelants", fmts["header"]); ws.write(row, 1, "Appels", fmts["header"])
        ws.write(row, 2, "Manqués", fmts["header"]); ws.write(row, 3, "Durée cumulée", fmts["header"]); row += 1
        for e in meta_calls.get("top_callers", []):
            label = e["label"] if show_names else mask_call_label(e)
            ws.write(row, 0, label); ws.write(row, 1, e["calls"]); ws.write(row, 2, e["missed"]); ws.write(row, 3, format_hms(e["duration_sec"])); row += 1
        png = chart_heatmap_png(meta_calls["hourly"]["in"], ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
                                 [str(h) for h in range(24)], "Appels reçus — jour × heure", figsize=(9, 3.2), cbar_label="Nb d'appels")
        xlsx_insert_png(ws, row + 1, 0, png, scale=0.85)
        row += 20

    if meta_emails:
        resp = meta_emails.get("response", {})
        ws.write(row, 0, "Emails — délai de réponse (indicatif)", fmts["title"]); row += 1
        for label, val in [
            ("Reçus", meta_emails.get("received", 0)), ("Envoyés", meta_emails.get("sent", 0)),
            ("Fils de discussion détectés", meta_emails.get("threads", 0)),
            ("Répondus (appariés)", resp.get("replied", 0)),
            ("Taux d'appariement", f"{resp.get('match_rate')}%" if resp.get("match_rate") is not None else "—"),
            ("Sans réponse sous 7 j", meta_emails.get("unanswered", 0)),
            ("Délai médian (heures ouvrées)", format_hms((resp.get("median_biz_min") or 0) * 60)),
            ("Délai P90 (heures ouvrées)", format_hms((resp.get("p90_biz_min") or 0) * 60)),
            ("Non lus", f"{meta_emails.get('pct_unread', 0)}%"),
            ("Avec pièce jointe", f"{meta_emails.get('attachments', 0)}"),
        ]:
            ws.write(row, 0, label); ws.write(row, 1, val); row += 1
        row += 1
        ws.write(row, 0, "Sujets fréquents", fmts["header"]); ws.write(row, 1, "Occurrences", fmts["header"]); row += 1
        for s in meta_emails.get("top_subjects", []):
            ws.write(row, 0, s["subject"]); ws.write(row, 1, s["count"]); row += 1
        row += 1
        ws.write(row, 0, "Mots-clés fréquents", fmts["header"]); ws.write(row, 1, "Occurrences", fmts["header"]); row += 1
        for w in meta_emails.get("top_words", []):
            ws.write(row, 0, w["word"]); ws.write(row, 1, w["count"]); row += 1
        row += 1
        ws.write(row, 0, "Top destinataires", fmts["header"]); ws.write(row, 1, "Emails", fmts["header"]); row += 1
        for r in meta_emails.get("top_recipients", []):
            ws.write(row, 0, r["name"] if show_names else "•••• (masqué)"); ws.write(row, 1, r["count"]); row += 1
        png = chart_heatmap_png(meta_emails["hourly"]["received"], ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
                                 [str(h) for h in range(24)], "Emails reçus — jour × heure", figsize=(9, 3.2), cbar_label="Nb d'emails")
        xlsx_insert_png(ws, row + 1, 0, png, scale=0.85)


if __name__ == "__main__":
    print("Helpdesk Dashboard -> http://127.0.0.1:5050")
    app.run(host="127.0.0.1", port=5050, debug=True)
