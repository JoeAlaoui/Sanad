# SANAD — Plateforme de pilotage du Helpdesk SI (AMMPS / DSID)

Application web locale (Flask + JSON, aucune base de données), multi-volets, multi-utilisateurs.
"SANAD" désigne l'application complète : **Tarkhiss** (Support Email & Hotline), **Moussanada**
(Helpdesk GLPI) et **Reporting Métier PCHC**. Nom, sous-titre, agence et titres de rapport sont
personnalisables dans Administration → Édition avancée.

## 1. Installation & lancement

```bash
cd sanad_app
python -m pip install -r requirements.txt
python app.py
```
Ouvrir : **http://127.0.0.1:5050**. Laisser le terminal ouvert.

**⚠️ Windows — installation** : utiliser systématiquement `python -m pip install -r requirements.txt`
(avec `python -m` devant) plutôt que `pip` seul, qui a échoué sur certains postes (pip introuvable
dans le PATH alors que Python l'est).

**⚠️ Windows — génération PDF (WeasyPrint)** : WeasyPrint est le moteur PDF principal (meilleur
rendu que l'ancien moteur, conservé en repli automatique : xhtml2pdf). Il dépend de bibliothèques
système absentes de Windows par défaut (GTK3 : Pango, Cairo, GDK-Pixbuf). Si l'export PDF échoue
avec une erreur `libgobject`/`cairo`/`pango`, installer le runtime GTK3 pour Windows (voir la
documentation officielle WeasyPrint, section "Windows"). **Tant que GTK3 n'est pas installé,
l'application bascule automatiquement sur xhtml2pdf** — aucune action requise pour que l'export
PDF continue de fonctionner en attendant.

## 2. Premier lancement — création du compte administrateur

Au premier accès, l'application demande de créer le compte **Administrateur / Responsable
Helpdesk**. Ce compte gère ensuite les autres utilisateurs dans l'onglet **Utilisateurs**.

## 3. Rôles

| Rôle | Accès |
|---|---|
| **Administrateur** | Accès complet : saisie, imports, rapports, administration, comptes, tous les volets |
| **Superviseur** | Lecture seule : Dashboard, Vue annuelle, export PDF/Excel des volets. Pas de saisie, pas d'envoi d'email, pas d'administration |
| **Hotliner** | **Base de connaissances** et **Notes**, limité aux volets assignés par l'admin, + **vue lecture seule** du Dashboard et de la Vue annuelle **Tarkhiss uniquement** (jamais Moussanada/PCHC, jamais de saisie/export/envoi) si "Tarkhiss" fait partie de ses volets assignés |

## 4. Volets

### 4.1 Tarkhiss — Support Email & Hotline

**Heatmap jour × heure (appels et emails) recadrée aux horaires d'ouverture** : affichage
limité aux jours ouvrés (Lun-Ven) et à la plage 08h-17h (≈08h30-16h30, horaires Tarkhiss),
sur l'écran, le PDF et l'Excel. Seule la heatmap est recadrée — le taux de décroché, le délai
de réponse et les autres statistiques restent calculés sur l'intégralité des données importées,
week-ends et heures creuses compris.
- Saisie mensuelle appels/emails, analyse (problèmes, demandes, constats, recommandations),
  dashboard, vue annuelle, heatmap jour de semaine × semaine du mois, comparatif N vs N-1,
  alertes de seuil configurables, prompts Copilot, rapport email/PDF/Excel.
- **Import Outlook brut (.xlsm)** : à partir d'un export de la boîte mail (feuille "Mail" —
  Folder Path, Subject, DisplayTo/Cc, DateTimeSent/Received, Importance, IsRead, HasAttachments,
  Preview, Id), recalcule automatiquement les compteurs jour par jour (`\Inbox\` = reçus,
  `\Sent Items\` = envoyés ; Deleted/Drafts/Archive ignorés). Fichier cumulatif : chaque
  réimport resynchronise les mois couverts.
- **Import journal d'appels brut (.csv)** : à partir d'un export du journal d'appels mobile
  (Name, Phone, Date, Type, Duration…), recalcule les compteurs jour par jour (appels
  reçus/durée). Même principe cumulatif que l'import Outlook.
- **Statistiques avancées** (issues des deux imports bruts ci-dessus, PDF/Excel/écran
  uniquement — **jamais dans le corps de l'email**) : heatmap jour × heure (appels et emails),
  taux de décroché, distribution des durées d'appel, rappel des appels manqués sous 24h, top
  appelants ; délai de première réponse email (indicatif — apparié par sujet, pas d'ID de
  conversation dans l'export), sujets et mots-clés fréquents, top destinataires. Les noms de
  contacts sont **masqués par défaut** dans ces rapports (case à cocher RGPD en Administration
  pour les afficher en clair).
- **Traçabilité des imports** : chaque fichier brut importé (Outlook, journal d'appels) est
  archivé et re-téléchargeable depuis Administration → Imports bruts Tarkhiss (50 derniers
  conservés).

### 4.2 Moussanada — Helpdesk GLPI
- Imports GLPI agrégés (xlsx/csv), KPI, Pareto, radar, backlog, comparatifs, tendance des
  catégories dans le temps.
- Import CSV ticket-par-ticket (tickets bruts) → heatmap jour × heure, répartition par Type,
  distribution des délais de résolution, liste nominative des tickets en souffrance.
- **Intégration GLPI (Mode Direct)** : client REST (initSession/search/killSession, pagination),
  profils d'import déclaratifs (tickets, catégories, demandeurs, techniciens), interface
  unifiée "Mode CSV / Mode Direct GLPI", panneau Admin → Intégrations GLPI (URL, App-Token,
  User-Token, entity_id, test de connexion). Détection de doublons CSV↔Direct (confirmation
  requise pour écraser un import déjà présent).
- Filtre techniciens (exclusion/whitelist), appliqué à la lecture : un changement de filtre se
  répercute immédiatement sur dashboard, Excel, PDF et email, sans ré-import.
- **Limite connue** : profils GLPI testés contre un serveur simulé — les identifiants de champs
  de recherche (`glpi_profiles.py`) correspondent à un GLPI 11.x standard, à vérifier sur
  l'instance réelle (Configuration → Listes). Règle "ticket en retard" pour l'import direct =
  ouvert depuis plus de 30 jours (hypothèse, ajustable).

### 4.3 Reporting Métier PCHC
- Import Excel/CSV, voyants, table statut → couleur paramétrable, backlog par ancienneté,
  top 10 opérateurs, export Excel/PDF/PPTX ("Revue de Direction")/email.
- Seul volet disposant d'un export PowerPoint (décision de conception : pas de PPTX pour
  Tarkhiss/Moussanada).

## 5. Partager SANAD sur le réseau local

Bouton **🔗 Partager SANAD** (menu utilisateur, réservé à l'administrateur) : bascule entre accès
**local uniquement** (défaut) et accès **réseau local**, avec un port configurable. Une fois
enregistré, la fenêtre affiche le lien à transmettre (ex. `http://192.168.1.42:5050`) basé sur
l'IP détectée du poste. Chaque collègue qui ouvre ce lien depuis le **même réseau** (même
Wi-Fi/switch) arrive sur l'écran de connexion SANAD et utilise son propre compte.

**Bascule instantanée, sans redémarrage** : le processus écoute en permanence sur toutes les
interfaces ; c'est un contrôle applicatif (pas le système d'exploitation) qui refuse les
connexions non locales tant que le partage n'est pas activé. Activer/désactiver le partage
réseau prend donc effet immédiatement. **Seul un changement de port** force un vrai redémarrage
du processus — SANAD le fait alors lui-même automatiquement (quelques secondes d'interruption,
reconnexion automatique du navigateur), sans jamais vous demander de relancer `python app.py`
manuellement. Le réglage est conservé dans `data/network_config.json`.

**⚠️ Sécurité** : le mode réseau expose SANAD à quiconque est connecté au même réseau local (pas
à Internet). L'authentification et les rôles habituels restent en vigueur ; à réserver à un
réseau de confiance (bureau, VPN d'entreprise). Techniquement, le port reste toujours ouvert sur
toutes les interfaces réseau de la machine (nécessaire pour la bascule instantanée) — seul
l'accès applicatif est refusé en mode local, ce qui diffère d'un port réellement fermé.

## 6. Personnalisation de l'écran (Administration)

- **Un seul design, retravaillé en profondeur** (Neumorphism doux — plus de sélection de thème) :
  sidebar à indicateur d'onglet actif par liseré coloré (plus lisible qu'un bloc plein), cartes
  KPI avec barre d'accent supérieure selon le statut (normal/alerte/succès), tableaux avec
  lignes zébrées et survol, boutons et champs avec retour visuel au clic/focus, ombres douces
  cohérentes sur l'ensemble des cartes. Les kits de thèmes explorés en cours de route (Ant
  Design Pro, Pastel, AMMPS Vert Institutionnel, AMMPS Moderne 2026) sont conservés dans
  `static/design-kits/` à titre de référence, mais ne sont plus proposés comme choix actifs.
- **Aperçu live** : cliquer sur un thème ou une palette l'applique immédiatement à l'écran entier
  (avant même d'enregistrer), pour comparer sans aller-retour. Si vous quittez l'onglet
  Administration sans cliquer sur "Enregistrer", l'écran revient automatiquement au thème/palette
  réellement sauvegardés.
- **6 palettes de couleur** restent configurables : AMMPS (défaut), Océan, Émeraude, Ardoise,
  Violet, Bordeaux — avec aperçu live au clic, avant même d'enregistrer.
- Ces réglages sont **purement visuels côté écran** : aucun effet sur les rapports
  PDF/Excel/PowerPoint, qui conservent toujours la charte AMMPS officielle.
- **Navigation par mois** : flèches ←/→ à côté du sélecteur de mois pour avancer/reculer d'un
  mois et recharger automatiquement les données, avec une transition douce (fondu + léger
  glissement) au changement plutôt qu'un rafraîchissement brut.
- Bouton **densité compacte** (topbar) pour resserrer les tableaux volumineux, mémorisé par
  navigateur.
- Raccourcis clavier : `/` (recherche globale), `Échap` (ferme les fenêtres), `Ctrl+S`/`Cmd+S`
  sur un dashboard (exporte le PDF du volet actif).
- Favoris ★ (barre latérale, jusqu'à 8, mémorisés par navigateur), fil d'ariane, recherche
  globale (mois, catégories, contacts, base de connaissances, Notes), glisser-déposer sur les
  imports, notifications navigateur, aperçu PDF avant téléchargement, CSS d'impression,
  indicateur "dernière synchro", onboarding par rôle.

## 7. Gouvernance / fiabilité

- Sauvegardes automatiques horodatées avant chaque écriture (`data/_backups/`, 20
  versions/fichier), écriture atomique.
- Journal d'audit consultable en Administration (`data/audit_log.jsonl`) : connexions, imports,
  gestion des comptes, actions RGPD.
- RGPD : recherche/export CSV, anonymisation des noms demandeurs GLPI et entités PCHC, masquage
  par défaut des contacts Tarkhiss dans les rapports détaillés (case à cocher pour lever le
  masquage). **Limite connue** : les feuilles agrégées mensuelles historiques ne sont pas
  couvertes par l'anonymisation rétroactive.
- SMTP optionnel + endpoint headless `POST /api/system/reminder-check` (jeton `X-Reminder-Token`)
  pour tâche planifiée de rappel d'envoi mensuel.
- Export/import de la configuration complète (agence, logo, contacts, salutations, signatures,
  seuils, comptes, base de connaissances) en un fichier `.json`, distinct des données mensuelles.
- Historique des envois d'email avec téléchargement des `.eml` (protégé contre le path
  traversal) ; historique des imports bruts Tarkhiss avec re-téléchargement (même protection).

## 8. Notes techniques

- Authentification par session Flask (cookie), mots de passe hashés (`werkzeug.security`), clé
  de session persistée dans `data/.secret_key`. Politique de mot de passe et timeout de session
  configurables (Administration).
- Stockage JSON à plat dans `data/<volet>/` — aucune base de données. Comptes : `data/users.json`.
  Notes : `data/notes/thread_<id>.json`. Base de connaissances : `data/knowledge.json`.
- Dépendances (`requirements.txt`) : Flask, openpyxl (lecture des imports), xlsxwriter (génération
  Excel), matplotlib + seaborn (graphiques en images PNG), python-pptx (PPTX PCHC), weasyprint
  (PDF principal), xhtml2pdf (repli), requests (client GLPI).
- Aucune donnée envoyée à l'extérieur : tout reste local à la machine qui exécute l'application.

## 9. Historique des lots livrés (repères)

- **Lots A–G** : navigation/UI (accordéon, favoris, recherche globale, thèmes/palettes,
  densité, raccourcis, drag & drop, notifications), alertes de seuil, export/import de
  configuration, rappel programmé, export PNG des graphiques.
- **Lots GLPI (Mode Direct)** : client REST, profils d'import déclaratifs, pipeline en deux
  couches (parsing pur → stockage) partagé CSV/Direct, panneau Admin dédié, filtre techniciens
  persistant, détection de doublons.
- **Import Outlook brut** puis **exploitation avancée** : heatmap horaire, délai de réponse
  indicatif, sujets/mots-clés fréquents, top destinataires.
- **Import journal d'appels brut** : heatmap horaire, taux de décroché, rappels sous 24h, top
  appelants — restitué en PDF/Excel/écran, exclu de l'email.
- **Thème Ant Design Pro**, aperçu live des thèmes/palettes, navigation par flèches ←/→ avec
  transition douce au changement de mois.
- **Thèmes Pastel / AMMPS Vert Institutionnel / AMMPS Moderne 2026** (d'après kits fournis),
  heatmap horaire recadrée aux horaires d'ouverture, vue lecture seule Tarkhiss pour le
  hotliner, partage réseau local.
- **Design unifié** : suppression de la sélection de thème (un seul design retravaillé —
  sidebar, cartes, tableaux, boutons), partage réseau à bascule instantanée avec redémarrage
  automatique géré par l'application (changement de port uniquement).

## 10. Limites connues (honnêteté sur les points non vérifiés)

- Rendu visuel jamais vérifié dans un navigateur réel en conditions de développement (sandbox
  sans affichage) : structure, CSS et données sont validés, pas le rendu pixel par pixel.
- Module GLPI testé uniquement contre un serveur simulé, jamais contre l'instance GLPI réelle
  de l'AMMPS.
- WeasyPrint non testé sur poste Windows sans GTK3 (le repli xhtml2pdf est fonctionnel mais son
  rendu visuel sous Windows n'a pas été contrôlé).
- Délai de première réponse email et rappels d'appels manqués : indicateurs **indicatifs**
  (appariement par sujet / par numéro et fenêtre de temps, pas d'identifiant de conversation
  réel dans les exports sources).
