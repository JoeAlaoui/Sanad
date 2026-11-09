# -*- coding: utf-8 -*-
"""
Sauvegarde cohérente de la base SANAD (SQLite) vers un dossier de votre choix — à faire tourner même
serveur démarré (API de sauvegarde de SQLite, pas une simple copie de fichier).

    python tools/backup_sqlite.py <dossier de destination> [--db chemin.db] [--keep 30]

Crée sanad_AAAAMMJJ_HHMMSS.db, vérifie son intégrité (PRAGMA integrity_check) et ne garde que les `--keep`
plus récentes (défaut 30). La base (défaut : SANAD_DB_PATH ou ~/sanad-data/sanad.db) n'est jamais modifiée.
Destination conseillée : un dossier synchronisé ou externe au PC (OneDrive, clé USB, partage réseau).
⚠ La base contient les comptes utilisateurs et les jetons : protéger la destination.
"""
import glob
import os
import sqlite3
import sys
from datetime import datetime


def backup(dest_dir, db_path=None, keep=30):
    db_path = db_path or os.environ.get("SANAD_DB_PATH") or os.path.expanduser("~/sanad-data/sanad.db")
    if not os.path.exists(db_path):
        raise SystemExit(f"Base introuvable : {db_path}")
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, f"sanad_{datetime.now():%Y%m%d_%H%M%S}.db")
    src = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True, timeout=30)
    out = sqlite3.connect(dest)
    try:
        src.backup(out)
        ok = out.execute("PRAGMA integrity_check").fetchone()[0]
        n = out.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
    finally:
        out.close()
        src.close()
    if ok != "ok":
        os.remove(dest)
        raise SystemExit(f"Sauvegarde invalide ({ok}) : supprimée")
    for old in sorted(glob.glob(os.path.join(dest_dir, "sanad_*.db")))[:-keep]:
        os.remove(old)
    print(f"OK : {dest} ({os.path.getsize(dest) / 1e6:.1f} Mo, {n} documents, intégrité ok)")
    return dest


if __name__ == "__main__":
    argv = sys.argv[1:]
    opt = lambda name: argv[argv.index(name) + 1] if name in argv else None
    skip = {opt("--db"), opt("--keep")}
    args = [a for a in argv if not a.startswith("--") and a not in skip]
    if len(args) != 1:
        raise SystemExit(__doc__)
    backup(args[0], opt("--db"), int(opt("--keep") or 30))
