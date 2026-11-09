# -*- coding: utf-8 -*-
"""
Mesure de volumétrie JSON vs SQLite (store.py) sur des données synthétiques construites à partir des
vrais tickets : python tools/bench_store.py [multiplicateurs…]   (défaut : 1 10 50)

Pour chaque volume : taille sur disque, lecture complète des tickets, écriture complète, écriture d'un petit
document (réglage), ajout au journal d'audit, lecture d'un mois de calls_*.json parmi 120 mois.
Travaille dans un dossier temporaire ; data/ n'est jamais touché.
"""
import json
import os
import shutil
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import store  # noqa: E402

SRC = os.path.join(ROOT, "data", "moussanada", "tickets_raw.json")


def timed(fn, n=1):
    t = time.perf_counter()
    for _ in range(n):
        fn()
    return (time.perf_counter() - t) / n * 1000


def run(backend, tickets):
    tmp = tempfile.mkdtemp(prefix=f"bench_{backend}_")
    try:
        data = os.path.join(tmp, "data")
        os.makedirs(os.path.join(data, "moussanada"))
        os.makedirs(os.path.join(data, "tarkhiss"))
        store.configure(data, os.path.join(tmp, "bak"), 20, backend=backend)
        os.makedirs(os.path.join(tmp, "bak"), exist_ok=True)
        tk = os.path.join(data, "moussanada", "tickets_raw.json")
        for m in range(120):   # 10 ans de mois de saisie Tarkhiss
            store.write_json(os.path.join(data, "tarkhiss", f"calls_{2016 + m // 12}-{m % 12 + 1:02d}.json"),
                             {f"2026-01-{d:02d}": {"calls": d, "duration": "00:10:00"} for d in range(1, 29)})
        w_all = timed(lambda: store.write_json(tk, tickets))
        r_all = timed(lambda: store.read_json(tk, []), 3)
        r_open = timed(lambda: store.tickets(tk, open_only=True), 3)
        r_month = timed(lambda: store.tickets(tk, "2026-04", "2026-04"), 3)
        settings = os.path.join(data, "moussanada", "settings.json")
        store.write_json(settings, {"greeting": "Bonjour", "x": list(range(200))})
        w_small = timed(lambda: store.write_json(settings, {"greeting": "Bonjour", "x": list(range(200))}), 20)
        audit = os.path.join(data, "audit_log.jsonl")
        for i in range(5000):
            store.append_jsonl(audit, {"ts": "2026-10-01T10:00:00", "action": "x", "details": {"i": i}})
        a_one = timed(lambda: store.append_jsonl(audit, {"ts": "2026-10-01T10:00:00", "action": "y", "details": {}}), 50)
        a_read = timed(lambda: store.read_jsonl(audit), 3)
        month = timed(lambda: store.read_json(os.path.join(data, "tarkhiss", "calls_2020-05.json"), {}), 50)
        store.close()
        size = sum(os.path.getsize(os.path.join(r, f)) for r, _, fs in os.walk(data) for f in fs)
        return dict(taille_mo=size / 1e6, ecriture_tickets_ms=w_all, lecture_tickets_ms=r_all, lecture_ouverts_ms=r_open, lecture_un_mois_ms=r_month, ecriture_reglage_ms=w_small,
                    ajout_audit_ms=a_one, lecture_audit_5000_ms=a_read, lecture_mois_ms=month)
    finally:
        store.close()
        shutil.rmtree(tmp, ignore_errors=True)


def main(mults):
    base = json.load(open(SRC, encoding="utf-8"))
    print(f"Base : {len(base)} tickets réels\n")
    labels = [("taille_mo", "taille disque (Mo)"), ("lecture_tickets_ms", "lecture de tous les tickets (ms)"), ("lecture_ouverts_ms", "lecture des tickets ouverts seulement (ms)"), ("lecture_un_mois_ms", "lecture d'un seul mois (ms)"), ("ecriture_tickets_ms", "écriture de tous les tickets (ms)"),
              ("ecriture_reglage_ms", "écriture d'un petit réglage (ms)"), ("ajout_audit_ms", "ajout d'une entrée d'audit (ms)"),
              ("lecture_audit_5000_ms", "lecture de 5 000 entrées d'audit (ms)"), ("lecture_mois_ms", "lecture d'un mois (calls_*.json, ms)")]
    for k in mults:
        tickets = [dict(t, id=f"{t.get('id', '')}-{i}") for i in range(k) for t in base]
        res = {b: run(b, tickets) for b in ("json", "sqlite")}
        print(f"=== ×{k} : {len(tickets)} tickets")
        print(f"  {'':<42}{'JSON':>10}{'SQLite':>10}")
        for key, lab in labels:
            print(f"  {lab:<42}{res['json'][key]:>10.1f}{res['sqlite'][key]:>10.1f}")
        print()


if __name__ == "__main__":
    main([int(a) for a in sys.argv[1:]] or [1, 10, 50])
