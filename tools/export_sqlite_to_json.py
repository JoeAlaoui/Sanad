# -*- coding: utf-8 -*-
"""
Export SQLite → JSON (retour arrière sûr après une bascule en SQLite).

    python tools/export_sqlite_to_json.py <dossier data> <dossier de sortie> [--db chemin.db]

Réécrit dans <dossier de sortie> tous les documents de la base (mêmes chemins relatifs que data/) et le
journal d'audit (audit_log.jsonl). Pour revenir au mode JSON SANS perdre les écritures faites depuis la
bascule : exporter vers un dossier neuf, vérifier, puis le substituer à data/ (en gardant l'ancien) et
relancer sans SANAD_STORE. Ne modifie ni la base ni data/.
"""
import json
import os
import sqlite3
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import store  # noqa: E402


def export(data_dir, out_dir, db_path=None):
    data_dir, out_dir = os.path.abspath(data_dir), os.path.abspath(out_dir)
    if os.path.exists(out_dir) and os.listdir(out_dir):
        raise SystemExit(f"{out_dir} n'est pas vide : choisissez un dossier neuf.")
    store.configure(data_dir, None, 20, backend="sqlite", db_path=db_path)
    path = store.db_path()
    if not os.path.exists(path):
        raise SystemExit(f"Base introuvable : {path}")
    c = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    keys = [r[0] for r in c.execute("SELECT key FROM docs ORDER BY key")]
    names = [r[0] for r in c.execute("SELECT DISTINCT name FROM journal")]
    c.close()
    for key in keys:
        doc = store.read_json(os.path.join(data_dir, *key.split("/")), None)
        target = os.path.join(out_dir, *key.split("/"))
        os.makedirs(os.path.dirname(target), exist_ok=True)
        with open(target, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
    n = 0
    for name in names:
        target = os.path.join(out_dir, *name.split("/"))
        os.makedirs(os.path.dirname(target), exist_ok=True)
        with open(target, "w", encoding="utf-8") as f:
            for e in store.read_jsonl(os.path.join(data_dir, *name.split("/"))):
                f.write(json.dumps(e, ensure_ascii=False) + "\n")
                n += 1
    store.close()
    print(f"OK : {len(keys)} documents et {n} entrées de journal exportés vers {out_dir}")
    return len(keys), n


if __name__ == "__main__":
    argv = sys.argv[1:]
    db = argv[argv.index("--db") + 1] if "--db" in argv else None
    args = [a for a in argv if not a.startswith("--") and a != db]
    if len(args) != 2:
        raise SystemExit(__doc__)
    export(args[0], args[1], db)
