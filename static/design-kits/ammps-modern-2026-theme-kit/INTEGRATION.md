# Intégration du thème

## HTML, PHP ou Laravel
Chargez `theme.css` dans le layout principal et `theme.js` avec `defer`. Transformez les sections de `mockpage.html` en partials ou composants Blade. Centralisez les menus et cartes dans des structures de données.

## React ou Next.js
Transformez chaque section en composant TypeScript : `Header`, `Hero`, `QuickActions`, `ServiceGrid`, `DataShowcase`, `Updates`, `FAQ`, `CTA` et `Footer`. Remplacez les événements de `theme.js` par l’état React. Les tokens peuvent rester en CSS ou être exportés vers votre configuration Tailwind.

## Vue ou Nuxt
Créez des composants Vue et utilisez `ref()` pour la navigation, la recherche et la FAQ. Conservez les attributs ARIA présents dans la démonstration.

## Angular
Créez des composants standalone. Utilisez Angular CDK pour les menus, dialogues et la gestion du focus si disponible.

## Vérifications
- Remplacer la marque générique.
- Ajouter des polices juridiquement utilisables.
- Vérifier les contrastes et la navigation clavier.
- Ajouter les contenus FR, EN et AR avec `dir="rtl"`.
- Convertir les médias en AVIF/WebP.
- Tester à 360, 768, 1024 et 1440 px.
