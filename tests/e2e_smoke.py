"""Smoke test navigateur (interface unique) : python tests/e2e_smoke.py
Serveur lancé sur :5050, compte admin / Admin#12345678. Prérequis : pip install playwright && playwright install chromium"""
import json, os
from playwright.sync_api import sync_playwright
B = "http://127.0.0.1:5050"
TABS = ["dashboard", "calls", "emails", "analysis", "annual", "prompts", "p-dashboard", "p-import", "p-livrables", "p-analysis",
        "admin", "users", "notes", "knowledge", "quality", "compliance", "backups"]
MTABS = ["m-dashboard", "m-sources", "m-analysis", "m-annual", "m-guide"]
with sync_playwright() as p:
    b = p.chromium.launch()
    ctx = b.new_context(viewport={"width": 1400, "height": 900})
    H = {"Content-Type": "application/json"}
    cred = json.dumps({"username": "admin", "password": "Admin#12345678", "name": "Admin"})
    if ctx.request.post(B + "/api/auth/setup", data=cred, headers=H).status != 200:
        ctx.request.post(B + "/api/auth/login", data=cred, headers=H)
    # Hors-ligne : SANAD_CDN_DIR=dossier contenant node_modules/chart.js et chartjs-plugin-datalabels
    cdn = os.environ.get("SANAD_CDN_DIR")
    if cdn:
        ctx.route("**/chart.umd.min.js", lambda r: r.fulfill(path=f"{cdn}/node_modules/chart.js/dist/chart.umd.js", content_type="application/javascript"))
        ctx.route("**/chartjs-plugin-datalabels.min.js", lambda r: r.fulfill(path=f"{cdn}/node_modules/chartjs-plugin-datalabels/dist/chartjs-plugin-datalabels.min.js", content_type="application/javascript"))
    pg = ctx.new_page(); errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)[:200]))
    pg.on("console", lambda m: errs.append(m.text[:200]) if m.type == "error" else None)
    pg.goto(B + "/"); pg.wait_for_timeout(1500)
    pg.keyboard.press("Escape")
    bad = 0
    for mod, tabs in (("tarkhiss", TABS), ("moussanada", MTABS)):
        if mod == "moussanada":
            pg.click('#moduleSwitch button[data-parent="moussanada"]'); pg.wait_for_timeout(1200)
        for t in tabs:
            errs.clear()
            pg.evaluate(f"activateTab('{t}')"); pg.wait_for_timeout(700)
            ok = pg.evaluate(f"document.getElementById('tab-{t}').classList.contains('active')")
            print(t, "ok" if ok and not errs else "ERR", errs[:2]); bad += 0 if ok and not errs else 1
    # Revue de direction PCHC : l'onglet affiche la synthèse si un import existe, sinon le message vide
    pg.click('#moduleSwitch button[data-parent="tarkhiss"]'); pg.wait_for_timeout(1000)
    pg.evaluate("activateTab('p-livrables')"); pg.wait_for_timeout(1500)
    txt = pg.inner_text("#plSynthesis")
    ok = "TOTAL" in txt or "Aucun import" in txt
    print("p-livrables contenu", "ok" if ok else "ERR", txt[:80]); bad += 0 if ok else 1
    b.close()
    raise SystemExit(1 if bad else 0)
