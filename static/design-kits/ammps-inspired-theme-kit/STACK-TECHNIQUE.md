# Audit exhaustif de la stack visible

## Résumé
Le code fourni décrit un site institutionnel rendu côté serveur ou généré par un CMS/framework, puis enrichi côté navigateur par une stack classique HTML5, CSS, Bootstrap 5, jQuery et de nombreux plugins. Le front n’est pas une SPA React/Vue/Angular dans l’extrait.

## 1. Couche document et SEO
- HTML5 (`<!doctype html>`, balises sémantiques `header`, `nav`, `section`, `footer`).
- Métadonnées SEO classiques : titre, description, auteur, viewport.
- Open Graph et Twitter Cards.
- Favicon et Apple Touch Icons.
- Vérification Google Search Console.
- Google Analytics 4 via `gtag.js`, identifiant `G-4DH5VCECLH`.

## 2. CSS et système de mise en page
### Confirmé
- Bootstrap 5 : classes `container`, `row`, `col-*`, `navbar-expand-lg`, `collapse`, `data-bs-toggle`, accordéon et utilitaires responsive.
- Feuilles propres au thème : `style.min-1.css`, `responsive.min-1.css`, `accounting.css`.
- Bundle fournisseur : `vendors.min-1.css`.
- CSS d’icônes : `icon.min-1.css`.
- CSS inline important, organisé par blocs fonctionnels et media queries.
- Design responsive aux seuils 575, 767/768, 991, 1199, 1399, 1500 et 1700 px.
- Variables CSS déjà présentes dans le thème externe, par exemple `--green-dark`, `--dark-gray`.

### Typographie
- Famille custom `cerebri-sans-medium` et `cerebri-sans-semibold` visible dans les règles.
- Préconnexion à Google Fonts, mais la police Google exacte ne peut pas être déterminée depuis l’extrait.
- Plusieurs familles d’icônes coexistent : Font Awesome, Feather Icons et Bootstrap Icons.

## 3. JavaScript
### Confirmé
- jQuery 3.7.1.
- Bundle `vendors.min-1.js` et logique du thème dans `main-1.js`.
- JavaScript natif pour préchargeur, menu « Accès en ligne » et comportement mobile.
- Bootstrap JavaScript très probable et pratiquement confirmé par les attributs `data-bs-*`; il peut être inclus dans `vendors.min-1.js`.
- Configuration déclarative des animations via attributs `data-anime`.
- Effets de parallaxe via `data-bottom-top`, `data-top-bottom` et `data-parallax-background-ratio`.

### Carrousels et animations
- Swiper : classes `swiper`, `swiper-wrapper`, `swiper-slide`, options JSON, autoplay, pagination, navigation et breakpoints.
- Moteur d’animation du thème : attribut `data-anime` avec translate, scale, rotateX, opacity, easing et stagger.
- Un moteur de scroll/parallaxe est présent dans le bundle fournisseur, mais sa bibliothèque exacte ne peut pas être affirmée à partir du seul HTML.

## 4. Tableaux et exports
- DataTables 1.13.8.
- Adaptateur Bootstrap 5 pour DataTables.
- DataTables Responsive 2.5.0.
- DataTables Buttons 2.4.2.
- JSZip 3.10.1 pour les exports Excel/ZIP.
- pdfmake 0.2.7 et `vfs_fonts.js` pour les exports PDF.
- Exports CSV, Excel, PDF paysage A4 et impression.
- Recherche globale, filtres exacts par colonne, pagination et comptage dynamique.

## 5. Composants fonctionnels visibles
- Préchargeur avec pulsation du logo.
- Barre supérieure, liens prioritaires, réseaux sociaux et langues.
- Barre de navigation sticky/responsive.
- Méga-menus multi-colonnes et menu simple.
- Menu mobile Bootstrap Collapse.
- Dropdown d’accès aux services.
- Recherche avancée plein écran et recherche rapide.
- Hero/slider plein écran avec images desktop/mobile via `<picture>`.
- Cartes de services, actualités et notes.
- Timeline/processus en trois étapes.
- Marquee horizontal fondé sur Swiper.
- FAQ avec Bootstrap Collapse/Accordion.
- CTA avec image parallax.
- Footer en cartes, responsive.
- Curseur personnalisé et indicateur de progression de scroll.

## 6. Images et performance
- Images raster PNG/JPEG et SVG.
- Images responsives avec `<picture>` sur le slider.
- `object-fit: cover` et positions adaptées au mobile.
- Préchargeur de page.
- CSS et JS minifiés et regroupés.
- Les dimensions d’images, `loading="lazy"`, `fetchpriority`, WebP/AVIF et CSP ne sont pas visibles ou systématiques dans l’extrait.

## 7. Accessibilité observable
### Bonnes bases
- `lang="en"` présent, mais incohérent avec le contenu majoritairement français.
- Quelques `aria-label`, `aria-controls`, `aria-expanded` et rôles.
- Navigation clavier activée dans Swiper.
- Textes alternatifs sur plusieurs images.

### Risques
- Liens `javascript:void(0)` et liens `#`.
- Dropdown personnalisé non décrit avec ARIA complet.
- Gestion du focus non démontrée pour les menus.
- Curseur personnalisé potentiellement gênant.
- Contrastes à vérifier, notamment textes gris et éléments transparents.
- Animations sans règle visible `prefers-reduced-motion`.
- Images décoratives parfois avec `alt` vide, parfois sans attribut explicite.
- Le menu mobile neutralise les liens avec `preventDefault` sans montrer une logique clavier équivalente.

## 8. Backend, CMS, base de données et hébergement
Non déterminables avec certitude depuis cet extrait : langage serveur, framework backend, CMS, moteur de templates, base de données, serveur web, cache/CDN, pipeline CI/CD et hébergeur. Les URLs propres et le contenu dynamique suggèrent un rendu serveur/CMS, mais identifier Laravel, WordPress, Drupal ou une autre solution serait spéculatif sans en-têtes HTTP, cookies, fichiers manifestes ou code serveur.

## 9. Dette technique identifiable
- Beaucoup de CSS inline et plusieurs blocs `<style>` dispersés.
- Règles répétées ou contradictoires pour le slider et les breakpoints.
- Multiplication des systèmes d’icônes.
- Dépendance simultanée à jQuery et au JavaScript natif.
- Sélecteurs très spécifiques et nombreux `!important`.
- Gestion de menu mobile ajoutant des écouteurs à chaque redimensionnement.
- Éléments DataTables chargés globalement même si le tableau n’est pas présent sur toutes les pages.
- Contenu, structure et présentation fortement couplés.

## 10. Stack moderne recommandée pour reproduire le design
- Front : HTML/CSS/TypeScript ou React/Next.js selon le projet.
- Design system : tokens CSS + composants documentés.
- Layout : CSS Grid/Flexbox; Bootstrap facultatif.
- Icônes : une seule bibliothèque, par exemple Lucide ou Font Awesome.
- Slider : Swiper uniquement si nécessaire.
- Tables : TanStack Table côté React, ou DataTables si l’existant jQuery doit être conservé.
- Animations : CSS et IntersectionObserver; respecter `prefers-reduced-motion`.
- Internationalisation : FR/AR/EN avec `lang`, `dir="rtl"` et contenus séparés.
- Qualité : ESLint, Stylelint, Prettier, tests Playwright, Lighthouse et axe-core.
