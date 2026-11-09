# -*- coding: utf-8 -*-
"""Sauvegardes, restauration et archives par volet avec le backend choisi (SANAD_STORE=json|sqlite) :
python tests/test_store_ops.py   puis   SANAD_STORE=sqlite python tests/test_store_ops.py

Serveur réel isolé : snapshot → modification → restauration (réglages ET tickets), export-all /
import-all d'un volet, journal d'audit, dernière synchro.
"""
import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_pchc_export import ADMIN, Server  # noqa: E402

CSV = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "reference", "glpi (1).csv")


def test_ops():
    with Server() as srv:
        U = srv.url
        a = srv.session()
        assert a.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        assert a.get(U + "/api/moussanada/last-sync").json()["last_sync"] is None
        assert a.post(U + "/api/moussanada/settings", json={"greeting": "Bonjour A,"}).status_code == 200
        with open(CSV, "rb") as f:
            assert a.post(U + "/api/moussanada/import-tickets", files={"file": ("glpi.csv", f)}).json()["result"]["rows"] == 55
        assert a.get(U + "/api/moussanada/last-sync").json()["last_sync"]
        n_tickets = a.get(U + "/api/moussanada/late-tickets").json()["open_total"]

        snap = a.post(U + "/api/admin/snapshots").json()
        assert snap["ok"], snap
        a.post(U + "/api/moussanada/settings", json={"greeting": "Bonjour B,"})
        a.post(U + "/api/moussanada/settings", json={"late_tickets": {"top_n": 5, "excluded_tickets": ["1369"]}})
        assert a.get(U + "/api/moussanada/settings").json()["greeting"] == "Bonjour B,"

        # export-all / import-all d'un volet
        z = a.get(U + "/api/moussanada/export-all")
        names = zipfile.ZipFile(io.BytesIO(z.content)).namelist()
        assert "settings.json" in names and "tickets_raw.json" in names, names
        a.post(U + "/api/moussanada/settings", json={"greeting": "Bonjour C,"})
        assert a.post(U + "/api/moussanada/import-all", files={"backup": ("b.zip", z.content)}).json()["ok"]
        assert a.get(U + "/api/moussanada/settings").json()["greeting"] == "Bonjour B,"

        # restauration du snapshot : retour à l'état d'avant (réglages + tickets), base rouverte
        r = a.post(U + f"/api/admin/snapshots/{snap['name']}/restore").json()
        assert r["ok"], r
        assert a.get(U + "/api/moussanada/settings").json()["greeting"] == "Bonjour A,"
        assert a.get(U + "/api/moussanada/late-tickets").json()["open_total"] == n_tickets
        # l'application continue d'écrire après la restauration
        assert a.post(U + "/api/moussanada/settings", json={"greeting": "Bonjour D,"}).status_code == 200
        assert a.get(U + "/api/moussanada/settings").json()["greeting"] == "Bonjour D,"

        audit = a.get(U + "/api/admin/audit-log?limit=100").json()
        # le journal fait partie de la sauvegarde : restauré à l'état du snapshot, puis la restauration y est tracée
        assert "snapshot_restore" in {e["action"] for e in audit}
        srv.log.flush()
        log = open(os.path.join(srv.dir, "server.log"), encoding="utf-8", errors="replace").read()
        assert "Traceback" not in log, log[-1500:]


if __name__ == "__main__":
    test_ops()
    print(f"OK — sauvegardes/restauration ({os.environ.get('SANAD_STORE', 'json')})")
