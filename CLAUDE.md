# SANAD — contexte pour Claude Code

Application web locale **Flask + fichiers JSON** (aucune base) de pilotage du Helpdesk SI de l'AMMPS (DSID).
Volets : **Tarkhiss** (Support Email & Hotline), **Moussanada** (Helpdesk GLPI), **Reporting Métier PCHC**.
UI : **legacy unique** = `templates/index.html` + `static/app.js` (JS vanilla, Chart.js CDN) + `static/style.css`. Pas de framework.
Thème **Nolvira** (défaut sur `/`) = couche `static/nolvira/nolvira.css|js` injectée par la route `/` ; `/classic` = ancien thème brut (rollback).
Lire `HANDOFF.md` (contexte complet) avant toute modification importante.

## Utilisateur : Youssef (PMO IT / Qualité IT, AMMPS)
- Français. Réponses concises, structurées, directement exploitables. Pas de théorie longue, pas de répétitions.
- Code complet et fonctionnel, simple, maintenable, avec instructions d'exécution. Hypothèses permises.
- Qualité : SOP auditables, KPI, contrôles, traçabilité.
- Après chaque livraison : proposer suggestions / améliorations / ajouts (par rôle admin/superviseur/hotliner si pertinent) et dire ce qui n'est pas testé.

## Règles fermes
- Stack simple : JSON > base de données ; JS vanilla > frameworks. NE PAS réintroduire React/shadcn (supprimés volontairement).
- Écritures JSON atomiques via `write_json_safely` ; actions sensibles tracées par `log_audit`.
- PDF toujours générés côté serveur (WeasyPrint prioritaire, xhtml2pdf en repli), jamais `window.print()`.
- Aucune mention « SANAD » dans les fichiers exportés et les emails (nom de l'agence à la place) ; l'interface garde le nom.
- Pas de score SANAD, pas de colonne « Résolus » Tarkhiss, pas de délai de réponse/Prompt E Tarkhiss (ni dans le PPTX), pas d'éditeur libre de templates.
- Rapports : charte AMMPS dans `ammps_charte.py` (composants) et `rapports_mensuels.py` (Tarkhiss/Moussanada : PDF, PPTX, email) ; PCHC = hybride dans `pchc_deliverables.py`. Salutation et « Cordialement » toujours hors gabarit, à gauche. Images d'email en `data:` à l'aperçu, en CID dans le `.eml`.
- Tickets en retard Moussanada (`retards.py`, `retards_xlsx.py`) : retard = non résolu/clos ET TTR (échéance) dépassée, depuis « Import tickets bruts ». TOP N de l'email (réglage `late_tickets`, techniciens exclus jamais nommés) ; l'Excel joint (gabarit `static/charte/modele_tickets_en_retard.xlsx`, rempli, pas généré) et le bouton du dashboard écartent aussi ces techniciens, tout comme les KPI de l'email (le dashboard, lui, montre tout, écartés grisés).
- Polices : PDF = Montserrat (titres) + **Inter** (texte, remplace Roboto) ; graphiques matplotlib en Inter ; texte des emails/PPTX/Excel en Arial (seule police garantie sur les postes). Fichiers dans `static/charte/`.
- Emails : titre en texte, aucun logo dans le corps (logo uniquement dans les PDF/PPTX/Excel joints).
- Rôles : admin (tout) · superviseur (lecture seule sauf API admin interdite) · hotliner (Notes, Base de connaissances, GET Tarkhiss limités). Verrouillage côté API dans `auth_gate`.
- Données et archives contiennent `users.json` et `.secret_key` : ne jamais les committer.

## Lancer / tester
```bash
python -m pip install -r requirements.txt     # Windows : toujours `python -m pip`
python app.py                                 # http://127.0.0.1:5050 (premier lancement : création de l'admin)
SANAD_NO_AUTO_SNAPSHOT=1 python app.py        # désactive le thread de sauvegarde auto (tests)
python tests/e2e_smoke.py                     # Playwright : parcourt tous les onglets (admin / Admin#12345678)
python tests/test_pchc_export.py              # Reporting Métier PCHC : calcul, Excel = référence, serveur isolé, .eml
python tests/test_tarkhiss_annual.py          # Vue annuelle Tarkhiss : PDF/PPTX/Excel, email, droits, aucune mention SANAD
python tests/test_rapports_mensuels.py        # Rapports mensuels charte AMMPS : routes PDF/PPTX/email, droits, aucune mention SANAD
python tests/test_retards.py                  # Tickets en retard Moussanada : calcul (reference/glpi (1).csv), TOP email + exclusions, Excel joint, droits
python tools/build_retards_template.py        # reconstruit static/charte/modele_tickets_en_retard.xlsx (gabarit Excel des retards)
# hors-ligne : SANAD_CDN_DIR=<dossier avec node_modules/chart.js@4.4.4 + chartjs-plugin-datalabels@2.2.0>
```
Tester tout changement avec un vrai serveur + `curl`/`requests` (et Playwright pour l'UI) avant de livrer. Lancer le serveur et les requêtes dans la **même** commande shell.
**Livraison** : modifier directement les fichiers du projet (pas d'archive zip), puis redémarrer le serveur SANAD de Youssef (il tourne dans WSL : `.venv/bin/python app.py`, port 5050) pour que le code soit pris en compte. Ne jamais toucher `data/` (données réelles) : tests dans une copie isolée.
Tout bouton existant qui produit un livrable doit appeler le même générateur que le nouveau (ex. Excel PCHC du Dashboard → gabarit).

## Pièges connus
- Apostrophes françaises non échappées dans les chaînes JS ; accolade orpheline cassant `app.js` → `node --check static/app.js` après édition.
- matplotlib plante sur camembert à valeurs nulles (géré dans `chart_pie_png`).
- `moussanada_data()` utilise `jsonify` → besoin d'un contexte d'app hors requête.
- Imports GLPI : séparer parse et stockage (CSV et mode direct convergent). Collision TTR/Technicien dans le parseur CSV GLPI (corrigée).
- `pkill` dans la même commande shell tue le shell (exit 144).
