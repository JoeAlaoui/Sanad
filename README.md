# SANAD — Plateforme locale de pilotage du helpdesk

Application web multi-utilisateurs construite avec Flask et JavaScript natif. Elle regroupe le
support email et hotline, la gestion de tickets GLPI et le reporting métier PCHC. Les données sont
accessibles par une couche unique avec stockage JSON ou SQLite sélectionnable.

## Installation et lancement

```powershell
cd D:\Apps\sanad_app
python -m pip install -r requirements.txt
python app.py
```

Ouvrir <http://127.0.0.1:5050> et laisser le terminal ouvert. Au premier lancement, créer le compte
administrateur. Sous Windows, utiliser `python -m pip` plutôt que `pip` seul.

Les PDF sont générés côté serveur : WeasyPrint est utilisé en priorité, avec xhtml2pdf en solution
de repli. Sous Windows, WeasyPrint peut nécessiter l’installation des bibliothèques GTK3 (Pango,
Cairo et GDK-Pixbuf) ; sans elles, le repli PDF reste disponible.

## Rôles

| Rôle | Principaux droits |
|---|---|
| Administrateur | Saisie, imports, rapports, gestion des comptes, paramètres, sauvegardes et fonctions d’administration |
| Superviseur | Consultation, rapports et exports autorisés ; pas de saisie, d’envoi ni d’administration |
| Hotliner | Notes, base de connaissances et accès en lecture limités aux volets qui lui sont attribués |

Les restrictions sont appliquées côté API, et pas uniquement par masquage de l’interface.

## Fonctionnalités

### Support email et hotline

- Saisie et analyse mensuelles, tableaux de bord, vue annuelle, comparatifs de périodes, alertes et
  libellés de KPI personnalisables.
- Import cumulatif des exports Outlook `.xlsm` et des journaux d’appels `.csv`, avec recalcul des
  compteurs journaliers et statistiques détaillées.
- Analyses avancées : activité par jour et heure, taux de décroché, durées et rappels d’appels
  manqués, tendances et sujets fréquents. Les noms de contacts sont masqués par défaut dans les
  rapports détaillés.
- Vue annuelle avec descriptions et prompts par rubrique, ainsi que des exports PDF, Excel et
  PowerPoint et une préparation d’email.
- Les sections et graphiques des emails mensuels sont configurables. Les messages sont créés en
  brouillon `.eml` pour être ouverts et modifiés dans l’application de messagerie associée.

### Gestion des tickets GLPI

- Imports agrégés Excel/CSV et imports de tickets individuels ; indicateurs, tendances, catégories,
  services, techniciens, backlog et délais.
- Vue annuelle et rapports sur une période libre, avec détail mensuel, catégories, demandeurs et
  services. Un classeur annuel agrégé peut également être importé, consulté et utilisé pour générer
  les rapports.
- Connexion directe à l’API REST GLPI, profils d’import, pagination, test de connexion et contrôle
  des doublons entre sources.
- Filtrage des techniciens appliqué aux tableaux de bord et aux rapports, sans réimporter les
  tickets.
- Suivi des retards à partir des tickets bruts importés : un ticket est en retard si sa résolution
  ou sa clôture n’est pas enregistrée et si son échéance TTR est dépassée. Les tickets sans
  échéance sont signalés séparément.
- Le TOP et les KPI de retard des emails, ainsi que le classeur Excel joint ou téléchargé, peuvent
  exclure les techniciens configurés. Le classeur est rempli à partir du gabarit fourni.

### Reporting métier PCHC

- Import de classeurs `.xlsx`/`.xlsm` et de fichiers CSV, contrôles de structure et de qualité,
  synthèse et tableaux de bord.
- Les anomalies de structure sont signalées avant import ; les avertissements et les statuts
  inconnus sont visibles pour faciliter le contrôle des données.
- Calcul partagé entre l’interface et les livrables. Export Excel basé sur un gabarit, version
  allégée disponible, rapports PDF et PowerPoint, et préparation d’emails avec pièces jointes au
  format `.eml`.
- Les imports et exports sont journalisés ; une sauvegarde peut être créée avant un import.

### Stockage et migration

- `store.py` fournit une interface commune aux deux backends. Le mode JSON conserve les fichiers
  historiques ; le mode SQLite regroupe les documents et journaux dans une base avec journalisation
  WAL et index pour filtrer les tickets par période et statut.
- Le backend se choisit avec `SANAD_STORE=json` ou `SANAD_STORE=sqlite`. En mode SQLite, le chemin
  est défini par `SANAD_DB_PATH` (par défaut `data/sanad.db`). Pour une instance sous WSL, placer la
  base dans le système de fichiers Linux plutôt que sous `/mnt/d`.
- Migration contrôlée des fichiers JSON vers SQLite :

  ```bash
  python tools/migrate_json_to_sqlite.py <dossier-data> --dry-run
  python tools/migrate_json_to_sqlite.py <dossier-data>
  ```

  La migration valide les documents avant la création de la base, vérifie les données migrées,
  refuse d’écraser une base existante et conserve les fichiers JSON d’origine. Sauvegarder les
  données avant la migration et tester d’abord sur une copie.
- `tools/export_sqlite_to_json.py` permet de préparer un retour au backend JSON. Les snapshots
  SQLite incluent une copie cohérente de la base ; la restauration d’une archive créée avec un
  backend différent est refusée.
- Sous Windows avec WSL, `Lancer-SANAD.bat` démarre l’application en SQLite. `Sauvegarde-SANAD.bat`
  crée une copie cohérente de la base vers le dossier de sauvegarde choisi ; planifier son exécution
  régulière et protéger les copies, qui contiennent des données de comptes et de configuration.

### Qualité, conformité et administration

- Tableau de qualité des données et contrôles mensuels.
- Gestion de procédures documentées et génération de dossiers mensuels de preuves, avec exports
  PDF ou ZIP et manifeste d’intégrité.
- Sauvegardes manuelles et automatiques, vérification d’intégrité avant restauration et création
  d’une sauvegarde de sécurité avant celle-ci.
- Journal d’audit des connexions, imports, exports et actions d’administration ; outils de recherche
  et d’anonymisation des données personnelles.
- Administration des utilisateurs, paramètres, notes et base de connaissances.

## Accès sur le réseau local

L’administrateur peut activer le partage de l’application sur le réseau local et transmettre le lien
affiché par l’interface. Le partage n’est pas destiné à exposer l’application à Internet. Ne
l’activer que sur un réseau de confiance ; les utilisateurs se connectent avec leur propre compte.

## Données et sécurité

Les données restent locales. En mode JSON, les documents sont stockés sous `data/`. En mode SQLite,
la base est placée au chemin configuré par `SANAD_DB_PATH`. Les fichiers locaux de comptes et la clé
de signature des sessions restent exclus du dépôt Git. Les sauvegardes peuvent contenir des données
de comptes et de configuration et doivent être protégées. Aucun mot de passe n’est stocké en clair :
les mots de passe sont hachés.

Après remplacement de `data/.secret_key`, les cookies de session précédents ne sont plus valides ;
les utilisateurs doivent se reconnecter. Cette rotation ne modifie pas les mots de passe.

## Tests

Depuis la racine du dépôt, avec les dépendances installées :

```powershell
python tests/test_retards.py
python tests/test_rapports_mensuels.py
python tests/test_pchc_export.py
python tests/e2e_smoke.py
python tests/test_store.py
python tests/test_store_ops.py
python tests/test_migration.py
python tests/test_hardening.py
```

Les scripts de test des rapports annuels sont également disponibles dans le dossier `tests/`. Les
tests de bout en bout utilisent un serveur et des données isolés. Pour désactiver le thread de
sauvegarde automatique dans les tests qui lancent l’application :

```powershell
$env:SANAD_NO_AUTO_SNAPSHOT = "1"
python app.py
```

Le navigateur charge Chart.js depuis un CDN. Pour les tests sans réseau, la variable
`SANAD_CDN_DIR` peut pointer vers un dossier contenant les dépendances Chart.js hors ligne.

## Limites connues

- L’intégration directe GLPI a été testée contre un serveur simulé ; vérifier les champs de recherche
  sur l’instance GLPI réelle avant de l’utiliser en production.
- La génération PDF sous Windows sans GTK3 repose sur le moteur de repli ; son rendu peut différer.
- Les analyses de première réponse email et de rappel des appels manqués sont indicatives, car les
  exports sources ne fournissent pas toujours d’identifiant fiable de conversation.
