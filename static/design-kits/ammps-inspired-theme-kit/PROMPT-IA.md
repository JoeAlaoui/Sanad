# Prompt maître à donner à une IA de développement

Tu es un lead designer UI et développeur frontend senior. Applique au projet existant un thème institutionnel moderne inspiré d’un portail public de santé marocain, sans copier de logo, d’images, de textes, de code propriétaire ou d’identité protégée.

## Objectif visuel
Créer une interface sobre, rassurante, accessible et premium, dominée par un vert institutionnel profond, du blanc, des gris chauds très clairs et un rouge réservé aux alertes importantes.

## Source de vérité
1. Utilise `design-tokens.json` pour toutes les valeurs de design.
2. Utilise `theme.css` pour les fondations.
3. Utilise `components.css` pour les composants.
4. Reproduis les comportements de `theme.js` dans le framework du projet.
5. Ne crée aucune couleur, ombre, rayon ou espacement arbitraire si un token existe.

## Règles de design
- Largeur de contenu maximale : 1200 px.
- Grilles aérées, fonds blancs ou gris très clair.
- Vert principal `#0d4823`; vert action `#0d6b3c`; rouge alerte `#df2328`.
- Titres bleu-gris `#24324a`; texte courant `#667085`.
- Boutons principaux en forme de pilule, transitions 200 à 300 ms.
- Cartes blanches avec bordure subtile, rayon 16 à 18 px et ombre douce.
- Badges légers, formulaires généreux, focus visible.
- Header sticky avec topbar facultative, navigation, CTA d’accès et menu mobile accessible.
- Sections types : hero, cartes de services, recherche, actualités, processus, FAQ, CTA et footer.
- Ne pas utiliser plus d’une famille d’icônes.
- Ne pas utiliser `!important` sauf justification documentée.
- Pas de CSS inline.

## Responsive
- Mobile first.
- 0–575 px : une colonne, boutons pleine largeur si utile.
- 576–767 px : une ou deux colonnes selon le contenu.
- 768–991 px : navigation mobile/tablette.
- 992–1199 px : grille desktop compacte.
- 1200 px et plus : grille desktop complète.
- Tester à 360, 768, 1024 et 1440 px.

## Accessibilité
- WCAG 2.2 AA.
- HTML sémantique, landmarks, labels et messages d’erreur.
- Navigation clavier complète, focus visible, fermeture Escape des menus.
- `aria-expanded`, `aria-controls` et rôles appropriés.
- Respect de `prefers-reduced-motion`.
- Contraste minimum 4.5:1 pour le texte normal.
- Le français est la langue par défaut; prévoir arabe RTL et anglais.

## Architecture attendue
- Transformer chaque motif en composant indépendant.
- Centraliser les données de navigation, cartes, actualités et FAQ.
- Séparer structure, données, style et interactions.
- Aucun contenu AMMPS réel dans les données de démonstration.
- Conserver la stack actuelle du projet sauf incompatibilité majeure.
- Si le projet utilise React, produire des composants fonctionnels TypeScript.
- Si le projet est en HTML classique, produire du JavaScript ES2022 sans jQuery.

## Composants obligatoires
`TopBar`, `Header`, `MegaMenu`, `MobileNav`, `Hero`, `ServiceCard`, `SearchPanel`, `NewsCard`, `ProcessSteps`, `Accordion`, `CTA`, `Footer`, `DataTableToolbar`, `Button`, `Badge`, `Input`, `Select`, `Modal`.

## Critères d’acceptation
- Aucun débordement horizontal à 360 px.
- Aucun style inline.
- Tous les états hover, focus, active, disabled et error sont couverts.
- Lighthouse : accessibilité ≥ 95, bonnes pratiques ≥ 95.
- Les animations sont désactivables.
- Les images ont dimensions, `alt`, lazy-loading hors hero, et formats modernes.
- Le résultat doit être visuellement cohérent avec `demo.html`.

Commence par auditer le projet, puis liste les fichiers modifiés. Applique ensuite le thème, exécute les tests disponibles et termine par un récapitulatif des choix, limites et éventuels écarts.
