# Guide d’intégration

## HTML / PHP / Laravel Blade
Copier les trois fichiers de thème dans le dossier d’assets, les charger dans le layout principal, puis extraire chaque section en partial/composant. Dans Blade, placer les données de navigation dans un tableau de configuration et utiliser des boucles plutôt que du HTML dupliqué.

## React / Next.js
- Convertir chaque bloc de `demo.html` en composant TSX.
- Importer `theme.css` globalement et `components.css` soit globalement, soit le transformer en CSS Modules.
- Remplacer `theme.js` par de l’état React.
- Pour les accordéons et menus, conserver les attributs ARIA.
- Utiliser `next/image` pour les images et les dictionnaires i18n pour FR/AR/EN.

## Vue / Nuxt
Créer `AppHeader.vue`, `HeroSection.vue`, `ServiceCard.vue`, `FaqAccordion.vue` et `AppFooter.vue`. Piloter les états avec `ref()` et centraliser les tokens dans CSS.

## Angular
Créer des composants standalone; utiliser le CDK pour menu/dialog/accessibilité si disponible. Charger les tokens dans `styles.css`.

## Bootstrap existant
Le kit fonctionne sans Bootstrap. Si Bootstrap est conservé, évitez les collisions en préfixant les composants ou en migrant progressivement. Les variables du thème peuvent alimenter les variables Sass Bootstrap.

## Tailwind
Reporter les tokens dans `tailwind.config` et construire des composants plutôt que de recopier de longues chaînes de classes. Conserver la sémantique et les états ARIA.

## Checklist
- Remplacer la marque de démonstration.
- Ajouter les vraies polices avec licences appropriées.
- Vérifier les contrastes.
- Ajouter RTL pour l’arabe.
- Optimiser les images en AVIF/WebP.
- Tester clavier, lecteur d’écran et mobile.
- Charger DataTables/TanStack uniquement sur les pages concernées.
