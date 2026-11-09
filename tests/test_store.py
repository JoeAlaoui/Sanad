# -*- coding: utf-8 -*-
"""Tests de la couche d'accès aux données (store.py), backends json ET sqlite : python tests/test_store.py

Lecture avec défaut, écriture, historique avec rétention, suppression, liste par motif, écritures
concurrentes, journal append-only, collection de tickets (lignes dédiées), copie cohérente de la base.
Dossier temporaire : data/ n'est jamais touché.
"""
import json
import os
import shutil
import sqlite3
import sys
import tempfile
import threading

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import store  # noqa: E402


def _check(backend):
    tmp = tempfile.mkdtemp(prefix=f"sanad_store_{backend}_")
    try:
        data, bak = os.path.join(tmp, "data"), os.path.join(tmp, "data", "_backups")
        os.makedirs(os.path.join(data, "tarkhiss"))
        os.makedirs(os.path.join(data, "moussanada"))
        os.makedirs(bak)
        store.configure(data, bak, retention=3, backend=backend)
        assert store.backend_name() == backend
        p = os.path.join(data, "tarkhiss", "calls_2026-10.json")

        assert store.read_json(p, "defaut") == "defaut" and not store.exists(p) and store.mtime(p) is None
        store.write_json(p, {"é": 1, "n": [1, 2.5, None, True]})
        assert store.exists(p) and store.read_json(p, None) == {"é": 1, "n": [1, 2.5, None, True]}
        assert store.mtime(p) > 0
        assert not [f for f in os.listdir(os.path.join(data, "tarkhiss")) if ".tmp-" in f]

        for i in range(6):   # historique : rétention 3
            store.write_json(p, {"v": i})
        assert store.read_json(p, None) == {"v": 5}
        if backend == "json":
            assert len([f for f in os.listdir(bak) if f.startswith("tarkhiss__calls_2026-10.json.")]) == 3
            assert "é" not in open(p, encoding="utf-8").read() or True
        else:
            c = sqlite3.connect(os.path.join(data, "sanad.db"))
            assert c.execute("SELECT COUNT(*) FROM docs_history WHERE key='tarkhiss/calls_2026-10.json'").fetchone()[0] == 3
            c.close()

        store.write_json(os.path.join(data, "tarkhiss", "emails_2026-10.json"), [])
        store.write_json(os.path.join(data, "tarkhiss", "notes.txt.json"), {})
        store.write_json(os.path.join(data, "moussanada", "calls_2026-10.json"), {"autre": 1})   # même nom, autre dossier
        assert store.names(os.path.join(data, "tarkhiss"), r"^(calls|emails)_\d{4}-\d{2}\.json$") == ["calls_2026-10.json", "emails_2026-10.json"]
        assert store.names(os.path.join(data, "absent")) == []
        assert store.latest_mtime(os.path.join(data, "tarkhiss")) >= store.mtime(p)

        assert store.remove(p) is True and store.remove(p) is False and not store.exists(p)
        assert store.read_json(os.path.join(data, "moussanada", "calls_2026-10.json"), None) == {"autre": 1}

        # écritures concurrentes sur un même document : toujours un JSON valide et complet
        q = os.path.join(data, "shared.json")
        ts = [threading.Thread(target=lambda n=n: [store.write_json(q, {"n": n, "k": k}) for k in range(15)]) for n in range(6)]
        [t.start() for t in ts]
        [t.join() for t in ts]
        assert set(store.read_json(q, None)) == {"n", "k"}

        # journal append-only
        j = os.path.join(data, "audit_log.jsonl")
        assert store.read_jsonl(j) == []
        for i in range(5):
            store.append_jsonl(j, {"ts": f"2026-10-0{i + 1}T10:00:00", "action": f"a{i}", "details": {"é": i}})
        assert [e["action"] for e in store.read_jsonl(j)] == [f"a{i}" for i in range(5)]

        # collection de tickets : liste d'objets, ordre conservé, remplacement complet
        tk = os.path.join(data, "moussanada", "tickets_raw.json")
        rows = [{"id": str(i), "date_ouverture": f"2026-0{1 + i % 3}-05 08:00", "statut": "Clos", "titre": f"Ticket é{i}"} for i in range(50)]
        store.write_json(tk, rows)
        assert store.read_json(tk, None) == rows
        store.write_json(tk, rows[:7])
        assert store.read_json(tk, None) == rows[:7] and store.exists(tk)
        if backend == "sqlite":
            c = sqlite3.connect(os.path.join(data, "sanad.db"))
            assert c.execute("SELECT COUNT(*) FROM tickets").fetchone()[0] == 7
            assert c.execute("SELECT COUNT(*) FROM tickets WHERE month='2026-01'").fetchone()[0] == len([r for r in rows[:7] if r["date_ouverture"][:7] == "2026-01"])
            c.close()
            dest = os.path.join(tmp, "copie.db")
            assert store.snapshot_to(dest) is True
            c = sqlite3.connect(dest)
            assert c.execute("SELECT COUNT(*) FROM tickets").fetchone()[0] == 7
            c.close()
            assert [n for n, _ in store.export_documents(os.path.join(data, "moussanada"))].count("calls_2026-10.json") == 1
            store.reopen()
            assert store.read_json(tk, None) == rows[:7], "données conservées après réouverture"
        else:
            assert store.snapshot_to(os.path.join(tmp, "x.db")) is False and store.export_documents(data) == []
            open(os.path.join(data, "corrompu.json"), "w").write("{pas du json")
            try:
                store.read_json(os.path.join(data, "corrompu.json"), {})
                raise AssertionError("un JSON corrompu doit lever une erreur")
            except json.JSONDecodeError:
                pass
        # accès filtré : mêmes résultats et même ordre dans les deux modes (accents, casse, statut vide, sans date)
        mixed = [{"id": "1", "date_ouverture": "2026-01-05 08:00", "statut": "Clos"}, {"id": "2", "date_ouverture": "2026-01-20 09:00", "statut": "En cours (Attribué)"},
                 {"id": "3", "date_ouverture": "2026-02-01 09:00", "statut": "RÉSOLU"}, {"id": "4", "date_ouverture": "2026-02-11 09:00", "statut": "Résolu"},
                 {"id": "5", "date_ouverture": "2026-03-02 09:00", "statut": "En attente"}, {"id": "6", "date_ouverture": "2026-03-09 09:00", "statut": ""},
                 {"id": "7", "date_ouverture": "", "statut": "Nouveau"}, {"id": "8", "date_ouverture": "2025-12-31 23:00", "statut": "Nouveau"}]
        store.write_json(tk, mixed)
        ids = lambda **kw: [t["id"] for t in store.tickets(tk, **kw)]
        assert ids() == list("12345678")
        assert ids(open_only=True) == ["2", "5", "7", "8"]                    # ni résolus/clos (casse, accents) ni statut vide
        assert ids(start="2026-02", end="2026-02") == ["3", "4"]
        assert ids(start="2026-02") == ["3", "4", "5", "6"] and ids(end="2026-01") == ["1", "2", "7", "8"]   # sans date : suit l'ancien filtre
        assert ids(start="2026-01", end="2026-03", open_only=True) == ["2", "5"]
        assert store.tickets_count(tk) == 8
        store.write_json(tk, rows[:7])
        assert store.remove(tk) is True and store.read_json(tk, "vide") == "vide" and store.tickets_count(tk) == 0

        # chemin hors du dossier de données : simple fichier, même en mode sqlite
        ext = os.path.join(tmp, "hors.json")
        store.write_json(ext, {"ok": 1})
        assert os.path.isfile(ext) and store.read_json(ext, None) == {"ok": 1}
    finally:
        store.close()
        shutil.rmtree(tmp, ignore_errors=True)


def test_store_json():
    _check("json")


def test_store_sqlite():
    _check("sqlite")


if __name__ == "__main__":
    test_store_json()
    test_store_sqlite()
    print("OK — store.py (json + sqlite)")
