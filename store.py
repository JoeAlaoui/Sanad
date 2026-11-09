# -*- coding: utf-8 -*-
"""
store.py — couche d'accès unique aux données de SANAD (migration JSON → SQLite, étape 2).

TOUT accès aux données persistantes passe ici. Deux backends interchangeables (même interface,
mêmes résultats) ; la clé d'un document est son chemin de fichier historique :
  - « json »   (défaut) : fichiers JSON dans data/, sauvegarde horodatée + écriture atomique ;
  - « sqlite » : un seul fichier data/sanad.db (mode WAL) — voir SCHEMA ci-dessous.
Choix : variable d'environnement SANAD_STORE=json|sqlite (SANAD_DB_PATH = emplacement du .db ; sous WSL le placer
sur le système de fichiers Linux, pas sur /mnt/d : 9p rend chaque écriture ~100× plus lente) (jamais activer sur les données réelles
avant l'étape 3 : outil de migration vérifié + sauvegardes adaptées).

Interface : read_json, write_json, exists, remove, names, mtime, latest_mtime, append_jsonl,
read_jsonl, export_documents, snapshot_to, reopen, backend_name.
Les chemins hors du dossier de données restent de simples fichiers JSON (même en mode sqlite).

Ce module ne dépend ni de Flask ni de app.py (importable par les outils de migration et les tests).
"""
import json
import os
import re
import shutil
import sqlite3
import threading
from datetime import datetime

_cfg = {"data_dir": None, "backup_dir": None, "retention": 20, "backend": "json", "db_path": None}
_lock = threading.RLock()
_conn = None

# Documents « liste d'objets » qui grossissent avec le temps : lignes dédiées (table tickets) plutôt
# qu'un seul gros document. Clé = chemin relatif au dossier de données.
COLLECTIONS = {"moussanada/tickets_raw.json": "tickets"}

SCHEMA = """
CREATE TABLE IF NOT EXISTS docs (            -- un document JSON par clé (réglages, notes, séries, sources, appels, emails…)
    key        TEXT PRIMARY KEY,             -- chemin relatif à data/, séparateur « / » (ex. tarkhiss/calls_2026-10.json)
    data       TEXT NOT NULL,                -- JSON (ou NULL-marker « @tickets » pour une collection à lignes dédiées)
    updated_at REAL NOT NULL                 -- epoch secondes
);
CREATE TABLE IF NOT EXISTS docs_history (    -- anciennes versions (équivalent des .bak), rétention par clé
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    key      TEXT NOT NULL,
    saved_at REAL NOT NULL,
    data     TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_docs_history_key ON docs_history(key, id);
CREATE TABLE IF NOT EXISTS tickets (         -- tickets GLPI bruts, une ligne par ticket (ordre d'import conservé par seq)
    seq            INTEGER PRIMARY KEY AUTOINCREMENT,
    id             TEXT,
    month          TEXT,                     -- AAAA-MM de date_ouverture
    date_ouverture TEXT,
    statut         TEXT,
    data           TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_tickets_month  ON tickets(month);
CREATE INDEX IF NOT EXISTS ix_tickets_statut ON tickets(statut);
CREATE TABLE IF NOT EXISTS journal (         -- journaux append-only (audit_log.jsonl…)
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    ts   TEXT,
    data TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_journal_name_ts ON journal(name, ts);
"""


def configure(data_dir, backup_dir, retention=20, backend=None, db_path=None):
    """Appelé une fois par app.py au démarrage. backend : « json » | « sqlite » (défaut : SANAD_STORE ou json)."""
    global _conn
    backend = (backend or os.environ.get("SANAD_STORE") or "json").strip().lower()
    if backend not in ("json", "sqlite"):
        raise ValueError(f"SANAD_STORE inconnu : {backend!r} (json | sqlite)")
    with _lock:
        if _conn is not None:
            _conn.close()
            _conn = None
        _cfg.update(data_dir=data_dir, backup_dir=backup_dir, retention=retention, backend=backend,
                    db_path=db_path or os.environ.get("SANAD_DB_PATH") or os.path.join(data_dir, "sanad.db"))
        if backend == "sqlite":
            _connect()


def backend_name():
    return _cfg["backend"]


def db_path():
    """Chemin du fichier SQLite (SANAD_DB_PATH, sinon data/sanad.db) ; None en mode json."""
    return _cfg["db_path"] if _cfg["backend"] == "sqlite" else None


def _connect():
    global _conn
    if _conn is None:
        os.makedirs(os.path.dirname(_cfg["db_path"]), exist_ok=True)
        _conn = sqlite3.connect(_cfg["db_path"], check_same_thread=False, isolation_level=None, timeout=15)
        _conn.execute("PRAGMA journal_mode=WAL")
        _conn.execute("PRAGMA synchronous=NORMAL")
        _conn.execute("PRAGMA busy_timeout=15000")
        _conn.executescript(SCHEMA)
    return _conn


def close():
    """Referme la base (avant remplacement du fichier .db) ; la prochaine opération la rouvre."""
    global _conn
    with _lock:
        if _conn is not None:
            _conn.close()
            _conn = None


def reopen():
    """Referme et rouvre la base (après remplacement du fichier .db par une restauration)."""
    global _conn
    with _lock:
        if _conn is not None:
            _conn.close()
            _conn = None
        if _cfg["backend"] == "sqlite":
            _connect()


def _key(path):
    """Clé SQLite d'un chemin, ou None s'il est hors du dossier de données (→ simple fichier)."""
    if _cfg["backend"] != "sqlite" or not _cfg["data_dir"]:
        return None
    rel = os.path.relpath(path, _cfg["data_dir"])
    if rel.startswith("..") or os.path.isabs(rel):
        return None
    return rel.replace(os.sep, "/")


def _dumps(data):
    return json.dumps(data, ensure_ascii=False, indent=2)


def _tx(fn):
    """Exécute fn(conn) dans une transaction d'écriture immédiate (verrou global + verrou SQLite)."""
    with _lock:
        c = _connect()
        c.execute("BEGIN IMMEDIATE")
        try:
            out = fn(c)
            c.execute("COMMIT")
            return out
        except BaseException:
            c.execute("ROLLBACK")
            raise


# ----------------------------------------------------------------------------- documents
def exists(path):
    k = _key(path)
    if k is None:
        return os.path.exists(path)
    with _lock:
        return _connect().execute("SELECT 1 FROM docs WHERE key=?", (k,)).fetchone() is not None


def read_json(path, default=None):
    """Contenu JSON du document, ou `default` s'il n'existe pas (JSON corrompu : l'erreur remonte, comme avant)."""
    k = _key(path)
    if k is None:
        if not os.path.exists(path):
            return default
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    with _lock:
        c = _connect()
        row = c.execute("SELECT data FROM docs WHERE key=?", (k,)).fetchone()
        if row is None:
            return default
        if k in COLLECTIONS:
            return [json.loads(r[0]) for r in c.execute("SELECT data FROM tickets ORDER BY seq")]
    return json.loads(row[0])


def _backup(path):
    data_dir, backup_dir = _cfg["data_dir"], _cfg["backup_dir"]
    if not (backup_dir and os.path.exists(path)):
        return
    try:
        rel = os.path.relpath(path, data_dir).replace(os.sep, "__")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        shutil.copy2(path, os.path.join(backup_dir, f"{rel}.{stamp}.bak"))
        prefix = f"{rel}."
        existing = sorted((f for f in os.listdir(backup_dir) if f.startswith(prefix) and f.endswith(".bak")), reverse=True)
        for old in existing[_cfg["retention"]:]:
            try:
                os.remove(os.path.join(backup_dir, old))
            except OSError:
                pass
    except OSError:
        pass  # une sauvegarde ratée ne doit jamais bloquer l'écriture principale


def _put_tickets(c, rows):
    c.execute("DELETE FROM tickets")
    c.executemany("INSERT INTO tickets(id, month, date_ouverture, statut, data) VALUES (?,?,?,?,?)",
                  [(str(t.get("id", "")), (t.get("date_ouverture") or "")[:7], t.get("date_ouverture"), t.get("statut"), json.dumps(t, ensure_ascii=False))
                   for t in rows])


def write_json(path, data):
    """Écriture sécurisée : ancien contenu sauvegardé (rétention par document), puis écriture atomique."""
    k = _key(path)
    if k is None:
        with _lock:
            _backup(path)
            tmp_path = f"{path}.tmp-{os.getpid()}-{threading.get_ident()}"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, path)
        return
    now = datetime.now().timestamp()

    def op(c):
        old = c.execute("SELECT data FROM docs WHERE key=?", (k,)).fetchone()
        if k in COLLECTIONS:
            _put_tickets(c, data)
            c.execute("INSERT OR REPLACE INTO docs(key, data, updated_at) VALUES (?,?,?)", (k, "@" + COLLECTIONS[k], now))
            return
        if old is not None:
            c.execute("INSERT INTO docs_history(key, saved_at, data) VALUES (?,?,?)", (k, now, old[0]))
            c.execute("DELETE FROM docs_history WHERE key=? AND id NOT IN (SELECT id FROM docs_history WHERE key=? ORDER BY id DESC LIMIT ?)",
                      (k, k, _cfg["retention"]))
        c.execute("INSERT OR REPLACE INTO docs(key, data, updated_at) VALUES (?,?,?)", (k, _dumps(data), now))
    _tx(op)


def remove(path):
    """Supprime un document ; True s'il existait."""
    k = _key(path)
    if k is None:
        with _lock:
            if os.path.exists(path):
                os.remove(path)
                return True
            return False

    def op(c):
        if c.execute("SELECT 1 FROM docs WHERE key=?", (k,)).fetchone() is None:
            return False
        if k in COLLECTIONS:
            c.execute("DELETE FROM tickets")
        c.execute("DELETE FROM docs WHERE key=?", (k,))
        return True
    return _tx(op)


def names(directory, pattern=None):
    """Noms des documents d'un dossier (non récursif), triés ; `pattern` = expression régulière (re.match)."""
    rx = re.compile(pattern) if pattern else None
    k = _key(os.path.join(directory, "x"))
    if k is None:
        if not os.path.isdir(directory):
            return []
        return sorted(fn for fn in os.listdir(directory) if not rx or rx.match(fn))
    prefix = k[:-1]
    with _lock:
        rows = _connect().execute("SELECT key FROM docs WHERE substr(key,1,?)=?", (len(prefix), prefix)).fetchall()
    found = sorted(r[0][len(prefix):] for r in rows if "/" not in r[0][len(prefix):])
    # les fichiers non JSON du dossier (exports, imports bruts…) restent sur le disque
    disk = [fn for fn in (os.listdir(directory) if os.path.isdir(directory) else [])
            if os.path.isfile(os.path.join(directory, fn)) and not fn.endswith(".json")]   # d'éventuels JSON d'origine ne comptent plus
    return sorted(n for n in set(found) | set(disk) if not rx or rx.match(n))


def mtime(path):
    """Date de dernière écriture (epoch) du document, ou None."""
    k = _key(path)
    if k is None:
        return os.path.getmtime(path) if os.path.exists(path) else None
    with _lock:
        row = _connect().execute("SELECT updated_at FROM docs WHERE key=?", (k,)).fetchone()
    return row[0] if row else None


def latest_mtime(directory, skip_dirs=("exports",)):
    """Dernière écriture de tout document JSON du dossier (récursif, hors `skip_dirs`), ou None."""
    k = _key(os.path.join(directory, "x"))
    if k is None:
        latest = None
        for root, _, files in os.walk(directory):
            if os.path.basename(root) in skip_dirs:
                continue
            for fn in files:
                if fn.endswith(".json"):
                    t = os.path.getmtime(os.path.join(root, fn))
                    latest = t if latest is None or t > latest else latest
        return latest
    prefix = k[:-1]
    with _lock:
        rows = _connect().execute("SELECT key, updated_at FROM docs WHERE substr(key,1,?)=?", (len(prefix), prefix)).fetchall()
    times = [t for key, t in rows if not any(f"/{s}/" in "/" + key for s in skip_dirs)]
    return max(times) if times else None


def export_documents(directory):
    """[(nom relatif, octets)] des documents d'un dossier (récursif) pour les archives par volet.
    Mode json : liste vide (les documents SONT des fichiers déjà parcourus par l'appelant)."""
    k = _key(os.path.join(directory, "x"))
    if k is None:
        return []
    prefix = k[:-1]
    with _lock:
        rows = _connect().execute("SELECT key FROM docs WHERE substr(key,1,?)=?", (len(prefix), prefix)).fetchall()
    out = []
    for (key,) in rows:
        rel = key[len(prefix):]
        out.append((rel, _dumps(read_json(os.path.join(_cfg["data_dir"], *key.split("/")), None)).encode("utf-8")))
    return out


# ----------------------------------------------------------------------------- tickets : accès filtré (indexé en sqlite)
_CLOSED = {"resolu", "clos", "closed", "solved"}


def _norm_status(v):
    import unicodedata
    s = unicodedata.normalize("NFD", str(v or "")).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


def _wanted(t, start, end, open_only):
    d = (t.get("date_ouverture") or "")[:7]
    if start and d < start:
        return False
    if end and d > end:
        return False
    if open_only:
        st = _norm_status(t.get("statut"))
        return bool(st) and st not in _CLOSED
    return True


def tickets(path, start=None, end=None, open_only=False):
    """Tickets bruts, filtrés par mois d'ouverture (start/end = AAAA-MM, bornes incluses) et/ou non résolus/clos.
    sqlite : le filtre est fait par la base (index mois/statut), seuls les tickets utiles sont lus ; json : lecture
    complète puis même filtre. Mêmes résultats, même ordre, dans les deux modes."""
    k = _key(path)
    if k not in COLLECTIONS:
        return [t for t in read_json(path, []) if _wanted(t, start, end, open_only)]
    where, args = [], []
    if start:
        where.append("month >= ?")
        args.append(start)
    if end:
        where.append("month <= ?")
        args.append(end)
    if open_only:
        where.append("statut IS NOT NULL AND TRIM(statut) <> '' AND lower(statut) NOT IN ('resolu','résolu','clos','closed','solved')")
    sql = "SELECT data FROM tickets" + (" WHERE " + " AND ".join(where) if where else "") + " ORDER BY seq"
    with _lock:
        rows = _connect().execute(sql, args).fetchall()
    return [t for t in (json.loads(r[0]) for r in rows) if _wanted(t, start, end, open_only)]


def tickets_count(path):
    """Nombre total de tickets bruts importés (sans les lire)."""
    k = _key(path)
    if k not in COLLECTIONS:
        return len(read_json(path, []))
    with _lock:
        return _connect().execute("SELECT COUNT(*) FROM tickets").fetchone()[0]


# ----------------------------------------------------------------------------- journaux append-only
def append_jsonl(path, entry):
    """Ajoute une entrée à un journal (une ligne JSON dans le fichier, ou une ligne de la table journal)."""
    k = _key(path)
    if k is None:
        with _lock:
            with open(path, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        return
    _tx(lambda c: c.execute("INSERT INTO journal(name, ts, data) VALUES (?,?,?)",
                            (k, str(entry.get("ts") or ""), json.dumps(entry, ensure_ascii=False))))


def read_jsonl(path):
    """Entrées du journal, de la plus ancienne à la plus récente (lignes invalides ignorées)."""
    k = _key(path)
    if k is None:
        out = []
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            out.append(json.loads(line))
                        except json.JSONDecodeError:
                            continue
        return out
    with _lock:
        rows = _connect().execute("SELECT data FROM journal WHERE name=? ORDER BY id", (k,)).fetchall()
    return [json.loads(r[0]) for r in rows]


# ----------------------------------------------------------------------------- sauvegarde de la base
def snapshot_to(dest):
    """Copie COHÉRENTE de la base (API de sauvegarde SQLite) vers `dest` ; False en mode json."""
    if _cfg["backend"] != "sqlite":
        return False
    with _lock:
        c = _connect()
        out = sqlite3.connect(dest)
        try:
            c.backup(out)
        finally:
            out.close()
    return True
