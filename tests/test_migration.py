# -*- coding: utf-8 -*-
"""Migration JSON → SQLite (tools/migrate_json_to_sqlite.py) : python tests/test_migration.py

Sur un dossier de données temporaire : simulation, migration, rapport, refus si la base existe, refus
sans rien écrire si un JSON est corrompu, JSON d'origine intacts, et refus des restaurations
d'archives d'un autre mode de stockage (snapshot JSON ↔ SQLite).
"""
import hashlib
import json
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import migrate_json_to_sqlite as M  # noqa: E402
import store  # noqa: E402


def _tree(base):
    os.makedirs(os.path.join(base, "moussanada", "exports"))
    os.makedirs(os.path.join(base, "tarkhiss"))
    W = lambda rel, obj: open(os.path.join(base, *rel.split("/")), "w", encoding="utf-8").write(json.dumps(obj, ensure_ascii=False, indent=2))
    W("global_settings.json", {"agency_name": "AMMPS — é"})
    W("moussanada/timeseries.json", {"2026-01": {"ouverts": 10, "resolus": 8, "en_retard": 2, "clos": 7}, "2026-02": {"ouverts": 5, "resolus": 5, "en_retard": 0, "clos": 5}})
    W("moussanada/tickets_raw.json", [{"id": str(i), "date_ouverture": "2026-01-05 08:00", "statut": "Clos"} for i in range(30)])
    W("tarkhiss/calls_2026-01.json", {"2026-01-05": {"calls": 3}})
    open(os.path.join(base, "moussanada", "exports", "x.eml"), "w").write("ne doit pas être migré")
    open(os.path.join(base, "audit_log.jsonl"), "w", encoding="utf-8").write(
        "".join(json.dumps({"ts": f"2026-01-0{i}T10:00:00", "action": f"a{i}"}) + "\n" for i in range(1, 6)) + "ligne cassée\n")


def _hashes(base):
    return {os.path.relpath(os.path.join(r, f), base): hashlib.sha256(open(os.path.join(r, f), "rb").read()).hexdigest()
            for r, _, fs in os.walk(base) for f in fs if not f.startswith("sanad.db")}


def test_migration():
    tmp = tempfile.mkdtemp(prefix="sanad_mig_")
    try:
        data = os.path.join(tmp, "data")
        _tree(data)
        before = _hashes(data)

        rep = M.migrate(data, dry_run=True, quiet=True)
        assert rep["dry_run"] and not os.path.exists(os.path.join(data, "sanad.db")) and not os.path.exists(os.path.join(data, "sanad.db.migrating"))

        rep = M.migrate(data, quiet=True)
        assert rep["documents"] == 4 and rep["tickets"] == 30 and rep["journal_entries"] == 5 and rep["journal_lines_skipped"] == 1
        assert rep["by_folder"] == {".": 1, "moussanada": 2, "tarkhiss": 1}, rep["by_folder"]
        assert os.path.isfile(os.path.join(data, "sanad.db")) and not os.path.exists(os.path.join(data, "sanad.db.migrating"))
        assert _hashes(data) == before, "les fichiers d'origine ne doivent pas changer"

        store.configure(data, None, 20, backend="sqlite")   # relecture par un store neuf
        assert len(store.read_json(os.path.join(data, "moussanada", "tickets_raw.json"), [])) == 30
        assert [e["action"] for e in store.read_jsonl(os.path.join(data, "audit_log.jsonl"))] == ["a1", "a2", "a3", "a4", "a5"]
        assert "x.eml" not in store.read_json(os.path.join(data, "moussanada", "timeseries.json"), {})
        store.close()

        # aller-retour : export SQLite → JSON identique aux documents et au journal d'origine
        import export_sqlite_to_json as X
        out = os.path.join(tmp, "retour")
        assert X.export(data, out) == (4, 5)
        for rel in ("global_settings.json", "moussanada/timeseries.json", "moussanada/tickets_raw.json", "tarkhiss/calls_2026-01.json"):
            assert json.load(open(os.path.join(out, *rel.split("/")), encoding="utf-8")) == json.load(open(os.path.join(data, *rel.split("/")), encoding="utf-8")), rel
        assert [json.loads(l)["action"] for l in open(os.path.join(out, "audit_log.jsonl"), encoding="utf-8")] == ["a1", "a2", "a3", "a4", "a5"]
        # base hors de data/ (SANAD_DB_PATH, sous WSL : système de fichiers Linux)
        alt = os.path.join(tmp, "ailleurs", "x.db")
        os.environ["SANAD_DB_PATH"] = alt
        try:
            store.configure(data, None, 20, backend="sqlite")
            assert store.db_path() == alt and os.path.isfile(alt)
            store.write_json(os.path.join(data, "tarkhiss", "t.json"), {"ok": 1})
            store.close()
        finally:
            del os.environ["SANAD_DB_PATH"]

        try:
            M.migrate(data, quiet=True)
            raise AssertionError("une base existante doit être refusée")
        except M.MigrationError as e:
            assert "existe déjà" in str(e)

        os.remove(os.path.join(data, "sanad.db"))   # JSON corrompu : échec, aucune base créée
        open(os.path.join(data, "tarkhiss", "calls_2026-02.json"), "w").write("{cassé")
        try:
            M.migrate(data, quiet=True)
            raise AssertionError("un JSON corrompu doit bloquer la migration")
        except M.MigrationError as e:
            assert "calls_2026-02.json" in str(e)
        assert not os.path.exists(os.path.join(data, "sanad.db")) and not os.path.exists(os.path.join(data, "sanad.db.migrating"))
    finally:
        store.close()
        shutil.rmtree(tmp, ignore_errors=True)


def test_snapshot_cross_mode_refused():
    """Archive d'un mode restaurée dans l'autre : refus clair (via de vrais serveurs)."""
    import io
    import zipfile
    from test_pchc_export import ADMIN, Server
    old = os.environ.get("SANAD_STORE")
    try:
        for mode, other in (("json", "sqlite"), ("sqlite", "json")):
            os.environ["SANAD_STORE"] = mode
            with Server() as srv:
                a = srv.session()
                assert a.post(srv.url + "/api/auth/setup", json=ADMIN).status_code == 200
                a.post(srv.url + "/api/moussanada/settings", json={"greeting": "Bonjour"})
                snap = a.post(srv.url + "/api/admin/snapshots").json()["name"]
                raw = a.get(srv.url + f"/api/admin/snapshots/{snap}/download").content
                assert ("data/sanad.db" in zipfile.ZipFile(io.BytesIO(raw)).namelist()) == (mode == "sqlite")
                os.environ["SANAD_STORE"] = other
                with Server() as srv2:
                    b = srv2.session()
                    assert b.post(srv2.url + "/api/auth/setup", json=ADMIN).status_code == 200
                    r = b.post(srv2.url + "/api/admin/snapshots/upload", files={"file": ("autre.zip", raw)})
                    name = r.json().get("name")
                    assert name, r.text
                    r = b.post(srv2.url + f"/api/admin/snapshots/{name}/restore")
                    assert r.status_code == 400 and "Archive du mode" in r.json()["error"], r.text
                os.environ["SANAD_STORE"] = mode
    finally:
        if old is None:
            os.environ.pop("SANAD_STORE", None)
        else:
            os.environ["SANAD_STORE"] = old


if __name__ == "__main__":
    test_migration()
    test_snapshot_cross_mode_refused()
    print("OK — migration et compatibilité des snapshots")
