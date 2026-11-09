# -*- coding: utf-8 -*-
"""
Migration JSON → SQLite.

    python tools/migrate_json_to_sqlite.py <dossier data> [--db chemin.db] [--dry-run] [--report rapport.json]

1. Lit et VALIDE tous les documents JSON du dossier (hors _backups, _snapshots, exports, raw_imports,
   uploads) et le journal d'audit ; au moindre fichier illisible, rien n'est écrit et la liste est donnée.
2. Écrit dans une base TEMPORAIRE (sanad.db.migrating) puis, seulement si tout est vérifié, la renomme en
   sanad.db (jamais de base à moitié remplie). Refuse si sanad.db existe déjà (la supprimer à la main).
3. Vérifie : chaque document relu == fichier d'origine ; journal : même nombre d'entrées, même ordre ;
   tickets : même nombre ; séries mensuelles : mêmes totaux ; comptage par dossier.
--dry-run : étapes 1 et 3 sur une base en mémoire/temporaire, sans créer sanad.db.
Les fichiers JSON d'origine ne sont JAMAIS modifiés ni supprimés (retour arrière = SANAD_STORE=json).
⚠ Sur les données réelles : arrêter le serveur, sauvegarder data/, puis migrer ; ou répéter sur une COPIE.
"""
import json
import os
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import store  # noqa: E402

SKIP_DIRS = {"_backups", "_snapshots", "exports", "raw_imports", "uploads"}
TS_FIELDS = ("ouverts", "resolus", "en_retard", "clos")


class MigrationError(Exception):
    pass


def scan(data_dir):
    """(documents, journaux) : chemins complets des .json et .jsonl du dossier de données."""
    docs, journals = [], []
    for root, dirs, files in os.walk(data_dir):
        dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS)
        for fn in sorted(files):
            full = os.path.join(root, fn)
            if ".tmp-" in fn or fn.startswith("sanad.db"):
                continue
            if fn.endswith(".json"):
                docs.append(full)
            elif fn.endswith(".jsonl"):
                journals.append(full)
    return docs, journals


def _read_journal(path):
    entries, bad = [], 0
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    bad += 1
    return entries, bad


def migrate(data_dir, db_path=None, dry_run=False, quiet=False):
    """Migre `data_dir` ; retourne le rapport (dict). Lève MigrationError sans rien laisser derrière soi."""
    t0 = time.time()
    data_dir = os.path.abspath(data_dir)
    final = os.path.abspath(db_path or os.path.join(data_dir, "sanad.db"))
    if os.path.exists(final) and not dry_run:
        raise MigrationError(f"{final} existe déjà : migration déjà faite ? Supprimez-le (après vérification) pour recommencer.")
    docs, journals = scan(data_dir)
    if not docs:
        raise MigrationError(f"Aucun document JSON dans {data_dir}")

    # 1. lecture + validation
    loaded, errors = {}, []
    for full in docs:
        try:
            with open(full, "r", encoding="utf-8") as f:
                loaded[full] = json.load(f)
        except (OSError, ValueError) as e:
            errors.append(f"{os.path.relpath(full, data_dir)} : {e}")
    if errors:
        raise MigrationError("Documents illisibles, rien n'a été écrit :\n  " + "\n  ".join(errors))
    journal_data, skipped_lines = {}, 0
    for full in journals:
        journal_data[full], bad = _read_journal(full)
        skipped_lines += bad

    # 2. écriture dans une base temporaire
    tmp_db = final + ".migrating"
    for ext in ("", "-wal", "-shm"):
        if os.path.exists(tmp_db + ext):
            os.remove(tmp_db + ext)
    store.configure(data_dir, None, 20, backend="sqlite", db_path=tmp_db)
    try:
        for full, content in loaded.items():
            store.write_json(full, content)
        for full, entries in journal_data.items():
            for e in entries:
                store.append_jsonl(full, e)

        # 3. vérifications
        problems = []
        by_folder = {}
        for full, content in loaded.items():
            folder = os.path.relpath(os.path.dirname(full), data_dir).replace(os.sep, "/")
            by_folder[folder] = by_folder.get(folder, 0) + 1
            if store.read_json(full, "ABSENT") != content:
                problems.append(f"document différent : {os.path.relpath(full, data_dir)}")
        for full, entries in journal_data.items():
            if store.read_jsonl(full) != entries:
                problems.append(f"journal différent : {os.path.relpath(full, data_dir)}")
        n_tickets = 0
        for full, content in loaded.items():
            if os.path.relpath(full, data_dir).replace(os.sep, "/") in store.COLLECTIONS:
                n_tickets = len(content)
                if len(store.read_json(full, [])) != n_tickets:
                    problems.append("nombre de tickets différent")
            if os.path.basename(full) == "timeseries.json":
                back = store.read_json(full, {})
                for field in TS_FIELDS:
                    a = sum((r.get(field) or 0) for r in content.values() if isinstance(r, dict))
                    b = sum((r.get(field) or 0) for r in back.values() if isinstance(r, dict))
                    if a != b:
                        problems.append(f"total « {field} » de la série mensuelle : {a} ≠ {b}")
        if problems:
            raise MigrationError("Vérification échouée, base non créée :\n  " + "\n  ".join(problems))
    finally:
        store.close()
    size = os.path.getsize(tmp_db)
    if dry_run:
        for ext in ("", "-wal", "-shm"):
            if os.path.exists(tmp_db + ext):
                os.remove(tmp_db + ext)
    else:
        os.replace(tmp_db, final)
        for ext in ("-wal", "-shm"):
            if os.path.exists(tmp_db + ext):
                os.remove(tmp_db + ext)
    report = {"data_dir": data_dir, "db": None if dry_run else final, "dry_run": dry_run, "documents": len(loaded),
              "by_folder": dict(sorted(by_folder.items())), "journal_entries": sum(len(v) for v in journal_data.values()),
              "journal_lines_skipped": skipped_lines, "tickets": n_tickets, "db_bytes": size,
              "json_bytes": sum(os.path.getsize(p) for p in list(loaded) + list(journal_data)), "seconds": round(time.time() - t0, 2)}
    if not quiet:
        print(("SIMULATION OK" if dry_run else "MIGRATION OK") + f" : {report['documents']} documents, {report['journal_entries']} entrées de journal, "
              f"{n_tickets} tickets — base {size / 1e6:.2f} Mo (JSON {report['json_bytes'] / 1e6:.2f} Mo) en {report['seconds']} s")
        for folder, n in report["by_folder"].items():
            print(f"  {folder or '.':<28}{n:>5} document(s)")
        if skipped_lines:
            print(f"  ⚠ {skipped_lines} ligne(s) de journal illisible(s) ignorée(s)")
    return report


if __name__ == "__main__":
    argv = sys.argv[1:]
    flags = {a for a in argv if a.startswith("--")}
    opt = lambda name: argv[argv.index(name) + 1] if name in argv else None
    skip = {opt("--db"), opt("--report")}
    args = [a for a in argv if not a.startswith("--") and a not in skip]
    if not args:
        raise SystemExit(__doc__)
    try:
        rep = migrate(args[0], opt("--db"), dry_run="--dry-run" in flags)
    except MigrationError as e:
        raise SystemExit(f"ÉCHEC : {e}")
    if opt("--report"):
        with open(opt("--report"), "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=2)
