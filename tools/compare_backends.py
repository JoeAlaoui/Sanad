# -*- coding: utf-8 -*-
"""
Répétition de migration : compare l'application en mode JSON et en mode SQLite sur la MÊME copie de données.

    python tools/compare_backends.py [dossier data]        (défaut : data/ du projet, LU seulement)

Copie data/ (sans utilisateurs, clé secrète, sauvegardes, snapshots) dans deux serveurs isolés, migre l'un
vers SQLite (tools/migrate_json_to_sqlite.py), puis appelle les mêmes API GET sur les deux et exige des
réponses identiques (tickets, séries, sources, annuel, qualité, journal d'audit, aperçus d'email…).
Les données du dossier source ne sont jamais modifiées.
"""
import json
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tests"))
from test_pchc_export import ADMIN, Server  # noqa: E402

VOLATILE = re.compile(r'"(generated|as_of|ts|last_sync|imported_at|date|updated_at|created|size)"')


def months(data, prefix, sub):
    d = os.path.join(data, sub)
    return sorted({m.group(1) for fn in os.listdir(d) for m in [re.match(rf"^{prefix}_(\d{{4}}-\d{{2}})\.json$", fn)] if m}) if os.path.isdir(d) else []


def endpoints(data):
    tk = months(data, "calls", "tarkhiss") or months(data, "analysis", "tarkhiss")
    ms = months(data, "sources", "moussanada")
    out = ["/api/modules", "/api/tarkhiss/months-list", "/api/moussanada/months-list", "/api/tarkhiss/settings", "/api/moussanada/settings",
           "/api/global-settings", "/api/moussanada/late-tickets", "/api/moussanada/tickets-analytics", "/api/moussanada/category-trend?start=2026-01&end=2026-12",
           "/api/moussanada/annual/2026?start=2026-01&end=2026-12", "/api/tarkhiss/annual/2026?start=2026-01&end=2026-12",
           "/api/moussanada/annual-email-preview?start=2026-01&end=2026-12", "/api/tarkhiss/annual-email-preview?start=2026-01&end=2026-12",
           "/api/moussanada/send-log", "/api/tarkhiss/send-log", "/api/moussanada/alert-history", "/api/notes", "/api/knowledge"]
    out += [f"/api/tarkhiss/month/{ym}" for ym in tk] + [f"/api/tarkhiss/analysis/{ym}" for ym in tk]
    out += [f"/api/moussanada/data/{ym}" for ym in ms] + [f"/api/moussanada/preview-email/{ym}" for ym in ms[-1:]]
    return out


def strip(o):
    """Retire les valeurs qui varient d'un appel à l'autre (horodatages)."""
    if isinstance(o, dict):
        return {k: strip(v) for k, v in o.items() if not VOLATILE.fullmatch(f'"{k}"')}
    if isinstance(o, list):
        return [strip(v) for v in o]
    return o


def fetch(srv, path):
    s = srv.session() if not hasattr(srv, "_admin") else srv._admin
    srv._admin = s
    if path.endswith("annual-email-preview") or "preview-email" in path:
        pass
    r = s.post(srv.url + path, json={}) if "preview-email/" in path else s.get(srv.url + path)
    try:
        body = r.json()
    except ValueError:
        body = r.text
    return r.status_code, strip(body)


def main(data):
    data = os.path.abspath(data)
    paths = endpoints(data)
    results = {}
    for mode in ("json", "sqlite"):
        os.environ["SANAD_STORE"] = mode
        with Server(seed_all=data) as srv:
            admin = srv.session()
            assert admin.post(srv.url + "/api/auth/setup", json=ADMIN).status_code == 200
            srv._admin = admin
            results[mode] = {p: fetch(srv, p) for p in paths}
    bad = []
    for p in paths:
        a, b = results["json"][p], results["sqlite"][p]
        if a != b:
            bad.append(p)
        print(("  = " if a == b else "  ≠ ") + f"{p}  [{a[0]}/{b[0]}]")
    if bad:
        for p in bad[:5]:
            sa, sb = json.dumps(results["json"][p], ensure_ascii=False, sort_keys=True), json.dumps(results["sqlite"][p], ensure_ascii=False, sort_keys=True)
            i = next((k for k in range(min(len(sa), len(sb))) if sa[k] != sb[k]), min(len(sa), len(sb)))
            print(f"\nÉCART {p}\n  json   …{sa[max(0, i - 80):i + 120]}\n  sqlite …{sb[max(0, i - 80):i + 120]}")
        raise SystemExit(f"\n{len(bad)} réponse(s) différente(s) sur {len(paths)}")
    print(f"\nOK : {len(paths)} réponses identiques entre JSON et SQLite")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "data"))
