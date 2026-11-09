# -*- coding: utf-8 -*-
"""Durcissement post-bascule SQLite : python tests/test_hardening.py   (aussi avec SANAD_STORE=sqlite)

- GET /api/global-settings : public (nom, logos, palette) sans connexion ; secrets (jetons GLPI, mot de passe
  SMTP, jeton de rappel) réservés à l'administrateur ;
- miroir data/network.json lu par Lancer-SANAD.bat (port) ;
- tools/backup_sqlite.py : copie cohérente, intégrité vérifiée, rotation.
"""
import json
import os
import shutil
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_pchc_export import ADMIN, Server  # noqa: E402

PWD = "Passw0rd#2026"
SECRETS = ("glpi_app_token", "glpi_user_token", "smtp_password", "reminder_api_token")


def test_global_settings_and_launcher_port():
    with Server() as srv:
        U = srv.url
        admin = srv.session()
        assert admin.post(U + "/api/auth/setup", json=ADMIN).status_code == 200
        assert admin.post(U + "/api/global-settings", json={"glpi_app_token": "APPSECRET", "glpi_user_token": "USERSECRET", "smtp_password": "PWSECRET",
                                                            "reminder_api_token": "REMINDSECRET", "agency_name": "AMMPS test", "session_timeout_minutes": 99}).status_code == 200
        for name, role in (("sup", "superviseur"), ("hot", "hotliner")):
            assert admin.post(U + "/api/users", json={"username": name, "password": PWD, "role": role, "name": name, "modules": ["tarkhiss", "moussanada"]}).status_code == 200
        sup, hot = srv.session({"username": "sup", "password": PWD}), srv.session({"username": "hot", "password": PWD})

        anon = srv.session().get(U + "/api/global-settings")
        assert anon.status_code == 200
        pub = anon.json()
        assert pub["agency_name"] == "AMMPS test" and not set(SECRETS) & set(pub) and "session_timeout_minutes" not in pub and "smtp_host" not in pub
        assert not any(x in anon.text for x in ("APPSECRET", "USERSECRET", "PWSECRET", "REMINDSECRET"))
        for who, s in (("superviseur", sup), ("hotliner", hot)):
            r = s.get(U + "/api/global-settings").json()
            assert not set(SECRETS) & set(r) and r["session_timeout_minutes"] == 99, who
        full = admin.get(U + "/api/global-settings").json()
        assert [full[k] for k in SECRETS] == ["APPSECRET", "USERSECRET", "PWSECRET", "REMINDSECRET"]
        # l'administrateur peut toujours enregistrer sans écraser les jetons
        assert admin.post(U + "/api/global-settings", json={"agency_name": "AMMPS 2"}).status_code == 200
        assert admin.get(U + "/api/global-settings").json()["glpi_app_token"] == "APPSECRET"

        # miroir du port pour le lanceur (.bat) : écrit à chaque enregistrement de la configuration réseau (et au démarrage de `python app.py`)
        mirror = os.path.join(srv.dir, "data", "network.json")
        assert not os.path.exists(mirror)
        port = admin.get(U + "/api/system/network-info").json()["configured_port"]
        assert admin.post(U + "/api/system/network-config", json={"host": "127.0.0.1", "port": port}).status_code == 200
        assert json.load(open(mirror)) == {"host": "127.0.0.1", "port": port}


def test_backup_sqlite():
    import backup_sqlite
    import store
    tmp = tempfile.mkdtemp(prefix="sanad_bk_")
    try:
        data = os.path.join(tmp, "data")
        os.makedirs(data)
        db = os.path.join(tmp, "sanad.db")
        store.configure(data, None, 20, backend="sqlite", db_path=db)
        store.write_json(os.path.join(data, "a.json"), {"é": 1})
        store.append_jsonl(os.path.join(data, "audit_log.jsonl"), {"ts": "2026-01-01", "action": "x"})
        dest = os.path.join(tmp, "copies")
        made = []
        for _ in range(4):
            made.append(backup_sqlite.backup(dest, db, keep=2))
            time.sleep(1.1)   # noms horodatés à la seconde
        kept = sorted(os.listdir(dest))
        assert len(kept) == 2 and kept == [os.path.basename(m) for m in made[-2:]], kept
        store.close()
        store.configure(data, None, 20, backend="sqlite", db_path=made[-1])   # la copie est une base complète et lisible
        assert store.read_json(os.path.join(data, "a.json"), None) == {"é": 1}
        assert [e["action"] for e in store.read_jsonl(os.path.join(data, "audit_log.jsonl"))] == ["x"]
        store.close()
    finally:
        store.close()
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    test_global_settings_and_launcher_port()
    test_backup_sqlite()
    print("OK — durcissement (réglages publics/secrets, port du lanceur, sauvegarde SQLite)")
