# Kit de thème institutionnel vert et rouge

Ce kit traduit les principes visuels observables dans le code source fourni en un système de design réutilisable, sans dépendre des fichiers propriétaires du site d’origine.

## Contenu
- `STACK-TECHNIQUE.md` : audit technique détaillé, avec niveaux de certitude.
- `PROMPT-IA.md` : prompt prêt à transmettre à une IA de développement.
- `design-tokens.json` : couleurs, typographie, rayons, ombres et espacements.
- `theme.css` : variables et fondations globales.
- `components.css` : composants réutilisables.
- `theme.js` : interactions accessibles sans jQuery.
- `demo.html` : démonstration autonome.
- `MIGRATION.md` : intégration dans différents frameworks.

## Utilisation rapide
```html
<link rel="stylesheet" href="theme.css">
<link rel="stylesheet" href="components.css">
<script src="theme.js" defer></script>
```

Ouvrir `demo.html` dans un navigateur. Le kit ne charge aucune ressource distante.

## Principe juridique et technique
Le kit reproduit une direction artistique générale, pas le code minifié, les images, le logo, les textes ni les polices propriétaires du site. Remplacez les contenus de démonstration par vos propres éléments.
