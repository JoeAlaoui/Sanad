"""Importe l'application depuis une copie isolée : data/ n'est jamais ouvert en écriture.
users.json, .secret_key et les sauvegardes ne sont pas copiés."""
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def load_app():
    work = tempfile.mkdtemp(prefix="charte_", dir=os.environ.get("CHARTE_SCRATCH") or None)
    for fn in os.listdir(ROOT):
        if fn.endswith(".py"):
            shutil.copy2(os.path.join(ROOT, fn), work)
    for d in ("templates", "pchc_template"):
        shutil.copytree(os.path.join(ROOT, d), os.path.join(work, d))
    shutil.copytree(os.path.join(ROOT, "static"), os.path.join(work, "static"), ignore=shutil.ignore_patterns("uploads"))
    os.makedirs(os.path.join(work, "data"))
    for d in ("moussanada", "pchc", "tarkhiss"):
        shutil.copytree(os.path.join(ROOT, "data", d), os.path.join(work, "data", d))
    shutil.copy2(os.path.join(ROOT, "data", "global_settings.json"), os.path.join(work, "data"))
    os.environ.update(SANAD_NO_AUTO_SNAPSHOT="1", SANAD_NO_OUTLOOK="1")
    os.chdir(work)
    sys.path.insert(0, work)
    import app
    return app, work
