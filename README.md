# SANAD — Plateforme de pilotage du Helpdesk SI (AMMPS)

Application web locale (Flask + JSON), multi-volets (Tarkhiss, Moussanada, Reporting Métier
PCHC), multi-utilisateurs, sans base de données. Le nom "SANAD" désigne désormais l'application
complète (et non plus seulement le volet historique Tarkhiss) — personnalisable dans
Administration → Édition avancée.

## 1. Installation & lancement

```bash
cd sanad_app
python -m pip install -r requirements.txt
python app.py
```
Ouvrir : **http://127.0.0.1:5050**. Laisser le terminal ouvert.

**⚠️ Remarque Windows** : la commande `pip install -r requirements.txt` seule a échoué sur certains
postes (pip introuvable dans le PATH alors que Python l'est). Utiliser systématiquement
`python -m pip install -r requirements.txt` (avec `python -m` devant) plutôt que `pip` seul —
cela invoque pip via l'interpréteur Python actif et évite ce problème.

**⚠️ WeasyPrint (génération PDF) sous Windows** : WeasyPrint est le moteur PDF principal depuis
cette version (meilleur rendu que l'ancien moteur xhtml2pdf, conservé en repli automatique).
Il dépend de bibliothèques système (GTK3 : Pango, Cairo, GDK-Pixbuf) absentes de Windows par
défaut. Si l'export PDF échoue avec une erreur liée à `libgobject`/`cairo`/`pango`, installer le
runtime GTK3 pour Windows (voir la documentation officielle WeasyPrint, section "Windows") puis
relancer l'application. **Tant que GTK3 n'est pas installé, l'application bascule automatiquement
sur xhtml2pdf** (rendu légèrement moins soigné mais fonctionnel) — aucune action requise pour que
l'export PDF continue de fonctionner en attendant.

## 2. Premier lancement — création du compte administrateur

Au premier accès, l'application demande de créer le compte **Administrateur / Responsable
Helpdesk** (identifiant + mot de passe). C'est ce compte qui gère ensuite les autres
utilisateurs dans l'onglet **Utilisateurs**.

## 3. Rôles

| Rôle | Accès |
|---|---|
| **Administrateur** (Responsable Helpdesk) | Accès complet : alimentation des données, génération de rapports, administration, gestion des comptes, tous les volets |
| **Superviseur** | Lecture seule : Dashboard, Vue annuelle, export PDF/Excel des deux volets. Pas de saisie, pas d'envoi d'email, pas d'administration |
| **Hotliner** | Accès restreint : **Base de connaissances** et **Notes** uniquement, limité aux volets qui lui sont assignés (par l'admin, à la création du compte) |

Les comptes se gèrent dans **Utilisateurs** (visible par l'Administrateur uniquement) :
créer un compte, assigner un rôle, assigner les volets accessibles (Hotliner), désactiver un
compte, réinitialiser un mot de passe.

### Notes (Hotliner ↔ Administrateur)
Messagerie simple à deux, par hotliner : le hotliner écrit une note (ex. pendant un appel),
l'administrateur y répond. Chaque hotliner a son propre fil ; l'admin choisit le hotliner dans
un menu déroulant.

### Base de connaissances
Articles (titre + contenu) classés par volet, consultables par tous, modifiables par
l'Administrateur uniquement. Structure volontairement simple — dites-moi si vous voulez des
catégories, pièces jointes, historique de versions, etc.

## 4. Les 7 fonctionnalités ajoutées

1. **Détection auto du mois du snapshot** — à l'import d'un `.xlsx` Moussanada, l'app détecte
   le mois le plus récent présent dans la feuille Tickets et propose de basculer dessus si
   différent du mois sélectionné.
2. **Comparaison technicien vs équipe** — tableau dédié sur le dashboard Moussanada (écart en %
   par rapport à la moyenne de l'équipe) + repris dans l'email/rapport.
3. **Alertes de seuil** — configurables par volet (Administration) : taux de résolution
   minimum, stock en retard maximum, délai moyen maximum. Bannière rouge sur le dashboard en
   cas de dépassement.
4. **Export/import de la configuration complète** — agence, logo, contacts, salutations,
   signatures, seuils, comptes utilisateurs et base de connaissances, en un fichier `.json`
   (Administration). Distinct de la sauvegarde des données mensuelles.
5. **Rappel programmé** — jour du mois configurable (Administration) à partir duquel la
   bannière de rappel s'affiche si le rapport du mois précédent n'est pas envoyé, avec
   notification navigateur si autorisée. **Limite honnête** : comme l'application ne tourne
   pas en tâche de fond, ce rappel se déclenche à l'ouverture de l'app, pas par email
   automatique pendant que l'app est fermée.
6. **Recherche globale** — barre de recherche en haut de l'application : mois, catégories,
   services, techniciens, problèmes/demandes, contacts, articles de la base de connaissances.
   Clique sur un résultat → ouvre directement l'onglet concerné.
7. **Export PNG des graphiques** — bouton "⬇ PNG" sur chaque graphique (dashboard et vue
   annuelle, les deux volets).

## 5. Reste inchangé

Toutes les fonctionnalités précédentes (Tarkhiss complet, Moussanada complet — import GLPI,
KPI, Pareto, radar, Top 10, Guide GLPI, email avec À/Cc/salutation/signature, Excel/PDF) sont
conservées à l'identique, désormais protégées par le contrôle d'accès par rôle.

## 6. Notes techniques

- Authentification par session Flask (cookie), mots de passe hashés (`werkzeug.security`),
  clé de session générée et persistée automatiquement dans `data/.secret_key`.
- Comptes : `data/users.json`. Notes : `data/notes/thread_<id>.json`. Base de connaissances :
  `data/knowledge.json`.
- Aucune donnée envoyée à l'extérieur, tout reste local.
- Sauvegardes horodatées automatiques avant chaque écriture (`data/_backups/`, 20 versions/fichier).
- Journal d'audit : `data/audit_log.jsonl` (connexions, imports, gestion des comptes, RGPD).

## 7. Personnalisation de l'écran (Administration)

- **4 thèmes visuels** : Flat (défaut), Neumorphism doux (recommandé si vous sortez du flat),
  Neumorphism complet et Claymorphism (ces deux derniers sont des options assumées, pas des
  recommandations — voir l'avertissement affiché dans l'écran Administration).
- **6 palettes de couleur**, indépendantes du thème choisi : AMMPS (défaut), Océan, Émeraude,
  Ardoise, Violet, Bordeaux. Neumorphism complet et Claymorphism gardent leur propre teinte
  fixe (la palette ne s'applique pas par-dessus ces deux thèmes).
- Ces réglages sont **purement visuels côté écran** : ils n'ont aucun effet sur les rapports
  PDF/Excel/PowerPoint, qui conservent toujours la charte AMMPS officielle.
- Bouton **densité compacte** (topbar) pour resserrer les tableaux volumineux, mémorisé par
  navigateur.
- Raccourcis clavier : `/` (focus recherche globale), `Échap` (ferme les fenêtres ouvertes),
  `Ctrl+S` / `Cmd+S` sur un dashboard (exporte le PDF du volet actif).
- Glisser-déposer disponible sur toutes les zones d'import de fichiers.
