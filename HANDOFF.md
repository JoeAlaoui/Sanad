# HANDOFF SANAD — état au 28/10/2026 (lot 10)

## 1. Produit
SANAD = application web locale de pilotage du Helpdesk SI de l'AMMPS (Agence Marocaine du Médicament et des Produits de Santé), DSID. Multi-utilisateurs, multi-volets, JSON uniquement.
- **Tarkhiss** : Support Email & Hotline (saisie appels/emails par jour, analyse mensuelle, dashboard, vue annuelle, prompts Copilot, rapport email/PDF/Excel).
- **Moussanada** : Helpdesk GLPI (import tickets agrégés CSV/direct API GLPI, dashboard, analyse & Copilot, vue annuelle, guide GLPI).
- **Reporting Métier PCHC** (sous-volet de Tarkhiss dans la sidebar) : import Excel 5 feuilles, dashboard, analyse, exports Excel/CSV/PDF/PPTX « Revue Direction », email avec Excel joint.
- Transverse : Administration, Utilisateurs, Notes, Base de connaissances, Qualité des données, SOP & preuves d'audit, Sauvegardes.

## 2. Historique résumé
Application d'origine livrée par lots (legacy). Puis migration vers React 19 + TS + Tailwind v4 + shadcn/ui (lots 2 à 9, avec pages Qualité, Sauvegardes, SOP, dashboards « copie carbone »). **Décision de Youssef (20/10/2026) : l'ancienne interface est meilleure → retour à la legacy, React/shadcn SUPPRIMÉ, et les nouvelles fonctions portées dans la legacy.** (lot 10). Les anciennes pages React n'existent plus dans le code.

## 3. Architecture
```
app.py            Flask (~4500 l., ~98 routes) : auth, modules, API JSON, rapports HTML/PDF/Excel/PPTX, emails .eml
ops.py            ops.register(app, module) : qualité, snapshots, SOP, dossier de preuves
glpi_client.py / glpi_import.py / glpi_profiles.py   API REST GLPI + parseurs
pchc_import.py    import Excel PCHC
tarkhiss_meta.py  import brut journal d'appels (CSV) et export Outlook (.xlsm) → stats horaires
templates/index.html   toute l'UI (onglets .tab-panel#tab-<nom>, sidebar .nav-item[data-tab][data-module][data-roles])
static/app.js     toute la logique (~4850 l.) ; static/style.css (tokens CSS « shadcn-like » zinc, palettes configurables)
static/design-kits/    kits de thème fournis par Youssef (Pastel, AMMPS inspired, AMMPS Modern 2026)
mockups/          6 maquettes statiques de comparaison de style (build.py les génère ; index.html = grille)
tests/e2e_smoke.py
data/<module>/    JSON (calls_YYYY-MM, emails_..., analysis_..., settings), data/_snapshots, data/_backups, audit_log.jsonl, users.json, .secret_key
```
Fonctions clés (`app.py`) : `write_json_safely`, `log_audit`, `_load/_save(module, kind, ym, default)`, `check_module`, `load_global_settings` (POST /api/global-settings fusionne), `load_module_settings`, `auth_gate` (before_request : réseau local/loopback, session, rôles), `NETWORK_STATE` (partage réseau), `get_report_html_and_label(module, ym, charts, for_pdf)` → `build_report_html` (Tarkhiss) / `build_report_html_moussanada` / PCHC ; `server_report_charts(module, ym)` (fallback matplotlib si le client n'envoie pas d'images) ; `chart_*_png` ; `render_pdf_bytes` ; `report_subtitle`, `report_wrap_style`, `build_email_body`, `build_subject`.
Routes utiles : `/api/<module>/month/<ym>`, `/calls/<ym>`, `/emails/<ym>`, `/analysis/<ym>`, `/settings`, `/months-list`, `/last-sync`, `/alert-history`, `/preview-email/<ym>`, `/send-email/<ym>` (payload to, cc, greeting, charts ; crée le .eml + log d'envoi), `/export-pdf/<ym>`, `/export-xlsx/<ym>`, `/api/system/network-info|network-config`, `/api/admin/*` (audit-log, rgpd search/anonymize, glpi test/import, snapshots, sop, audit-pack), `/api/quality/overview`, `/api/notes`, `/api/knowledge`.
UI : `activateTab(tab)` → `renderTab(tab)` ; modules via `#moduleSwitch` et `switchModule` ; mois global `#monthSelect` (`loadMonth` Tarkhiss / `loadMoussanadaMonth`) ; `state`, `mState`, `moduleSettings`, `globalSettings`, `currentUser`.

## 4. Fonctionnalités par rôle
- **admin** : tout (saisie, import, envoi email, administration, utilisateurs, sauvegardes, SOP, RGPD, GLPI direct, partage réseau, journal d'audit).
- **superviseur** : lecture seule des dashboards/vues annuelles/exports (Excel, PDF), comparatifs N-1, Qualité des données. Pas de saisie, envoi, admin.
- **hotliner** : Notes + Base de connaissances (volets assignés) ; Tarkhiss en lecture limitée.
Topbar legacy : fil d'Ariane, ★ favori, mois ←/→, recherche globale, indicateur de sauvegarde et de synchro, rôle, « Partager SANAD » (admin), « Compact », « ? Aide » (onboarding par rôle), déconnexion.

## 5. Ajouts du lot 10 (portés dans la legacy)
- Onglets **Qualité des données** (`tab-quality`, `renderQuality`), **SOP & preuves d'audit** (`tab-compliance`, `loadCompliance`), **Sauvegardes** (`tab-backups`, `loadBackups`) — groupe « Qualité & conformité » de la sidebar ; backend dans `ops.py`.
- Sauvegardes : ZIP de `data/` + logos avec manifeste SHA-256 ; vérification d'intégrité AVANT restauration + snapshot « pre-restore » ; auto quotidienne (thread, rétention `snapshot_retention`, `snapshot_auto`) ; ⚠ contiennent `users.json` et `.secret_key`.
- SOP (PDF/Word, métadonnées `sop_ref/version/owner/approver/date`) et dossier de preuves mensuel (PDF ou ZIP : PDF + SOP + CSV audit/envois + manifeste) ; jours fériés `quality_holidays`.
- **Couverture email/PDF d'origine** : wrapper `max-width:760px;margin:0 auto`, sous-titre par défaut `{agence} — {sous_titre} — {mois}` (variables `{agence} {sous_titre} {mois} {module} {periode}`), titre/sous-titre éditables en Administration > Édition avancée ; option `email_full_width` (défaut false) pour un email pleine largeur (jamais le PDF).
- Emails : si le client n'envoie pas de graphiques, `server_report_charts` les génère (Tarkhiss : weekly, problems ; Moussanada : volume, delay, trend, backlog, cat, pareto, srv, tech, radar). Clés `email_sections` incluant `chart_<key>`.
- Titres/sous-titres de rapport éditables, mois vides masqués dans les graphiques annuels/séries.

## 5 bis. Lot 11 — Reporting Métier PCHC : import brut + livrables (28/10/2026)
- `pchc_raw.py` : `parse_workbook()` (pur, openpyxl read_only, en-têtes ligne 5, colonnes détectées par en-tête avec repli B/H/J/F/D, rapport : lignes, feuilles manquantes/vides/ignorées, statuts inconnus, coquilles, contrôle de signature des colonnes → **bloquant**, forçable) et **`pchc_compute()`, calcul unique** (sémantique COUNTA/NB.SI insensible à la casse/SOMME.SI des formules Excel) utilisé par Excel, PPTX, PDF, e-mail et UI. Valeurs brutes stockées telles quelles ; coquilles regroupées au calcul (comme les formules).
- `pchc_excel.py` : **copie conforme par gabarit**. Référence = `reference/Rapport Tarkhiss_PCHC_Metier_YTD.xlsx` (enregistrée par Excel, feuilles brutes masquées). `tools/build_pchc_template.py` en tire `pchc_template/dashboard_gabarit.xlsx` (+ `.json`) : la référence à l'identique, feuilles brutes vidées, chaînes partagées réduites. À chaque export, `_TemplateWriter` réécrit seulement les feuilles brutes, les `<v>` des formules (format Excel : 17 chiffres, notation E sous 0,1, SOMME séquentielle), les caches des graphiques et la date ; tout le reste (16 graphiques, styles, MFC Excel 2010, mise en page) est identique **à l'octet**. Seul écart voulu (décision Youssef) : le classeur s'ouvre sur le Dashboard, et non sur le dernier onglet consulté dans la référence. Les valeurs de chaque cellule viennent de `_Builder` (port openpyxl de `reference/build_dashboard_pchc_reference.py`, aussi utilisé en repli si le gabarit manque). Version allégée : feuilles brutes et calcChain retirées, formules qui les lisaient → valeurs. Si la référence change : relancer l'outil puis les tests. Pourquoi pas XlsxWriter/openpyxl : ils ne savent pas reproduire le XML des graphiques et MFC enregistrés par Excel.
- Import : le classeur de référence complet s'importe tel quel (feuilles « KPI … »/« Dashboard » ignorées, noms exacts prioritaires).
- `pchc_deliverables.py` : PPTX 16:9 (9 diapos : garde, synthèse, chiffres clés & alertes, 5 KPI, méthodologie), PDF A4 paysage (WeasyPrint, repli xhtml2pdf, « Page n / N »), corps e-mail, `.eml` multipart/mixed (X-Unsent: 1, RFC 2231).
- `app.py` : `POST /api/pchc/import` (.xlsx/.xlsm → nouveau pipeline + synchro `records.json` du Dashboard « période » ; CSV inchangé), `GET /api/pchc/summary`, `/api/pchc/export/xlsx|pptx|pdf[?date=&light=1]`, `/api/pchc/eml-preview`, `/api/pchc/eml?xlsx=&pptx=&pdf=&light=`. Stockage `data/pchc/raw/{volet}.json`, `meta.json`, `history.json`. Snapshot `pchc-import` (via `app.extensions["sanad_create_snapshot"]`, ops.py) + `log_audit` à chaque import/export/eml. Réglage `pchc_enabled` (Administration › Édition avancée) : API `/api/pchc/*` fermées pour tous si décoché.
- UI : onglet **Reporting Métier › Revue de direction** (`tab-p-livrables`, `renderPLivrables`).
- Tests : `tests/test_pchc_export.py` (6 tests, dont égalité octet par octet des parties non régénérées et serveur réel en copie isolée), `tests/pchc_compare.py` (comparaison cellule par cellule). Classeur brut d'exemple : `reference/3282e5c0-dashboardTarkhisss.xlsx.xlsx`. Vérifié aussi hors tests via Excel (COM) : recalcul complet 0 écart / 0 erreur, export PDF des 6 feuilles identique au pixel à la référence (hors bruit de rendu d'Excel, présent aussi entre deux exports de la référence elle-même).

## 5 ter. Lot 12 — Tarkhiss : rapport de la Vue annuelle (28/10/2026)
- Onglet **Vue annuelle** (Tarkhiss › Reporting Hotline & Emails) : boutons Excel / PDF / PPTX + « Envoyer par email » (admin) avec choix des pièces jointes (PDF coché par défaut, PPTX, Excel). Hotliner : boutons masqués et API refusées ; superviseur : exports seulement.
- `app.py` : `tarkhiss_annual_data()` (source unique écran + rapports), `build_tarkhiss_annual(fmt, start, end)` (même fonction pour téléchargement et pièces jointes), routes `GET /api/tarkhiss/annual-report/<pdf|pptx|xlsx>?start&end`, `GET /api/tarkhiss/annual-email-preview`, `POST /api/tarkhiss/annual-send-email` (.eml + journal d'envoi + audit `annual_export` / `annual_send_email`).
- `tarkhiss_annual.py` : corps de l'email = éléments de la Vue annuelle (KPI, tendance, charge par jour, top 3, détail par mois), page de garde PDF, PPTX (garde, 2 diapos de synthèse, 2 diapos par mois), feuille Excel « Synthèse » (graphique natif) + une feuille par mois (`write_tarkhiss_month_sheet`, factorisée avec l'export mensuel).
- PDF annuel : garde + synthèse + chaque mois en 1-2 pages = rapport mensuel en mode `compact` (même design et mêmes sections que l'email mensuel, graphiques côte à côte), numérotation « Page n / N ».
- **Plus aucune mention « SANAD » dans les exports et envois** : logo SANAD retiré des rapports, en-tête PDF = nom de l'agence, pieds de page PCHC, Message-ID, propriétés Excel, expéditeur SMTP par défaut, SOP/dossier de preuves (réf. par défaut `SOP-DSID-HELPDESK-001`, `DSID-QUAL-…`). L'interface garde le nom SANAD.
- Test : `tests/test_tarkhiss_annual.py` (serveur isolé alimenté par une copie de `data/tarkhiss`).

## 5 quater. Lot 13 — Outlook, Analyse mensuelle allégée, descriptifs annuels (28/10/2026)
- **Ouverture Outlook** : tous les .eml (Tarkhiss/Moussanada mensuel, PCHC « Envoyer », PCHC Revue de direction, Vue annuelle) passent par `new_draft_email()` (en-tête `X-Unsent: 1` = brouillon) et `open_in_outlook()` : application associée aux .eml par Windows (sur le poste de Youssef : **nouvel Outlook**, l'Outlook classique n'étant pas configuré), y compris serveur sous WSL (`cmd.exe /c start`). Vérifié par Youssef le 28/10/2026 : le nouvel Outlook ouvre le .eml en brouillon modifiable (bouton Envoyer). `SANAD_NO_OUTLOOK=1` neutralise l'ouverture (tests).
- **Analyse mensuelle Tarkhiss** : « Action corrective », « Nuage de mots-clés », « Notes internes » et « Prompt F » retirés de la saisie, du Dashboard, du rapport email/PDF, de l'Excel et du CSV. Les valeurs déjà saisies restent dans les JSON (conservées à l'enregistrement, jamais affichées).
- **Vue annuelle — descriptifs** : un texte + un prompt Copilot par bloc (`tarkhiss_annual.BLOCKS` : indicateurs, charge, top3, detail, synthese) ; stockage `data/tarkhiss/annual_notes.json` par période ; `GET/POST /api/tarkhiss/annual-notes` (admin écrit, superviseur lit, hotliner exclu) ; textes repris dans l'écran, l'email, le PDF, le PPTX (diapos « Analyse de la période », 2 blocs/diapo) et l'Excel (feuille Synthèse).

## 6. Décisions verrouillées (ne pas remettre en cause sans accord)
Pas de score SANAD ; PPTX seulement Reporting Métier ; pas d'éditeur libre de templates ; PDF toujours serveur (WeasyPrint obligatoire, xhtml2pdf repli) ; colonne « Résolus » Tarkhiss supprimée ; champ délai de réponse et Prompt E supprimés ; JSON plutôt que base ; JS vanilla plutôt que frameworks ; écritures atomiques ; actions SANAD-spécifiques PCHC exclues de certains tours de suggestions.

## 7. Façon de travailler
Lot par lot, confirmation entre lots ; options A/B proposées pour les décisions d'architecture ; tests réels (curl + serveur lancé) avant livraison ; **livraison directe dans `D:\Apps\sanad_app` (fichiers modifiés sur place, PAS d'archive zip) puis redémarrage du serveur SANAD (Flask ne recharge pas le code)** ; suggestions systématiques après chaque livraison ; dire explicitement les limites non testées.

## 8. En cours / à décider
1. **Choix du style visuel** : 6 maquettes dans `mockups/` (HeroUI, Untitled UI, Linear/Geist sombre, Fluent 2/M365, Pastel, AMMPS Modern 2026). Youssef doit trancher ; ensuite appliquer le style à la legacy en ne changeant que les tokens CSS (`static/style.css` : couleurs, rayons, ombres, polices) — pas de réécriture JS. Le sélecteur de palette existe déjà (`ui_palette`).
2. Vérifier avec les vraies données (Calllog-export, sanad_RawSource_TarkhissEmails.xlsm, dashboardTarkhiss.xlsx fournis dans le projet claude.ai) : dashboards, imports, emails.
3. Non testé : restauration d'une archive venant d'une autre installation, envoi réel Outlook (`xdg-open`/`os.startfile`), PDF sous Windows sans GTK3 (repli xhtml2pdf), import GLPI direct sur un vrai serveur, rendu de Moussanada avec gros volumes.
4. Test Playwright : Chart.js est chargé depuis jsdelivr ; hors-ligne utiliser `SANAD_CDN_DIR`.

## 9. Idées de suite (à proposer, non faites)
- Sauvegarde auto avant chaque import ; indicateur « dernière sauvegarde » sur le dashboard.
- Envoi email avec PDF en pièce jointe (case à cocher) ; envoi SMTP réel du rappel programmé (réglages SMTP déjà présents, optionnels).
- Journal des alertes de seuil consultable + export ; tableau de bord superviseur « santé des données ».
- Tests unitaires pytest sur les calculs (métriques mensuelles, scores qualité, parseurs) et CI simple.
- Accessibilité (contrastes, focus, aria) et mode sombre du thème choisi.
- Refactor progressif de `app.js` (4850 l.) en modules ES sans build.

## 10. Prompt de reprise (à coller dans Claude Code, à la racine du projet)
> Lis CLAUDE.md puis HANDOFF.md. Lance l'app (`SANAD_NO_AUTO_SNAPSHOT=1 python app.py`) et `python tests/e2e_smoke.py` pour vérifier l'état. Ne réintroduis ni React ni shadcn. Prochaine tâche : [décrire]. Teste avec un vrai serveur avant de conclure, puis propose suggestions et limites non testées.
