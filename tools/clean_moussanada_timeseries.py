# -*- coding: utf-8 -*-
"""Nettoie la série mensuelle Moussanada (data/moussanada/timeseries.json) des clés qui ne sont
pas des mois AAAA-MM (ex. noms de techniciens enregistrés par un ancien import GLPI).

    python tools/clean_moussanada_timeseries.py            # simulation : n'écrit rien
    python tools/clean_moussanada_timeseries.py --apply    # applique (sauvegarde + audit)

--apply passe par save_timeseries() → write_json_safely (copie horodatée de l'ancien fichier
dans data/_backups, écriture atomique) et trace l'opération dans le journal d'audit.
Les mois réels ne sont jamais modifiés.
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ.setdefault("SANAD_NO_AUTO_SNAPSHOT", "1")
import app as A  # noqa: E402

MONTH = re.compile(r"^\d{4}-\d{2}$")


def main(apply):
    ts = A.load_timeseries()
    bad = {k: v for k, v in ts.items() if not MONTH.match(k)}
    kept = sorted(k for k in ts if MONTH.match(k))
    print(f"Fichier : {A.timeseries_file()}")
    print(f"Mois conservés ({len(kept)}) : {', '.join(kept) or '—'}")
    print(f"Clés à supprimer ({len(bad)}) :")
    for k in sorted(bad):
        print(f"  - {k!r} : {json.dumps(bad[k], ensure_ascii=False)}")
    if not bad:
        print("Rien à nettoyer.")
        return
    if not apply:
        print("\nSIMULATION — aucune écriture. Relancer avec --apply pour appliquer.")
        return
    A.save_timeseries({k: v for k, v in ts.items() if MONTH.match(k)})
    A.log_audit("moussanada_timeseries_cleanup", {"removed_keys": sorted(bad), "kept_months": len(kept)}, module="moussanada")
    print(f"\nAPPLIQUÉ — {len(bad)} clé(s) supprimée(s) ; ancienne version sauvegardée dans data/_backups.")


if __name__ == "__main__":
    main("--apply" in sys.argv[1:])
