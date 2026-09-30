# Thème "Moderne" — Design System portable

Extrait de l'application MIRSAAD. Inspiré d'une maquette EdTech (sidebar, cartes pastel
rotatives, coins très arrondis, accent violet). Livré ici sous forme de fichier CSS autonome
(`theme-moderne.css`), sans dépendance, réutilisable dans n'importe quel projet HTML/CSS.

## Installation

1. Copier `theme-moderne.css` dans votre projet.
2. Le charger dans le `<head>` :
   ```html
   <link rel="stylesheet" href="theme-moderne.css">
   ```
3. (Optionnel mais recommandé) Charger la police **Inter** pour un rendu fidèle :
   ```html
   <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
   ```
   Sans elle, le CSS retombe automatiquement sur la police système (`-apple-system`, `Segoe UI`...).

## Palette de couleurs

| Rôle | Variable CSS | Valeur |
|---|---|---|
| Accent principal | `--tm-accent` | `#8b6cff` |
| Accent (survol/actif) | `--tm-accent-dark` | `#7752f5` |
| Accent doux (fond) | `--tm-accent-soft` | `#e8e0fe` |
| Texte principal | `--tm-ink` | `#201c33` |
| Texte secondaire | `--tm-ink-soft` | `#6b6580` |
| Fond de page | `--tm-bg` | `#f6f5fb` |
| Fond des cartes | `--tm-surface` | `#ffffff` |
| Bordure fine | `--tm-border` | `#eeecf7` |

**Palette pastel rotative** (cartes statistiques) :

| Nom | Variable | Valeur |
|---|---|---|
| Pêche | `--tm-pastel-1` | `#ffead0` |
| Lavande | `--tm-pastel-2` | `#e8e0fe` |
| Menthe | `--tm-pastel-3` | `#dcf6e2` |
| Ciel | `--tm-pastel-4` | `#dcedff` |

**Couleurs sémantiques** (statuts — à garder identiques quel que soit le thème visuel de
l'application, pour ne jamais faire dépendre la lecture d'une criticité/d'un statut du thème
choisi) :

| Rôle | Variable | Valeur |
|---|---|---|
| Danger / critique | `--tm-danger` | `#b3261e` |
| Avertissement | `--tm-warning` | `#b06000` |
| Succès | `--tm-success` | `#1e7a4c` |
| Neutre | `--tm-neutral` | `#4b5568` |

## Forme

- Rayon des cartes/panneaux : `18px` (`--tm-radius-lg`)
- Rayon des champs/boutons non-pilule : `12px` (`--tm-radius-md`)
- Boutons et badges : forme pilule (`--tm-radius-pill` = `999px`)
- Ombre douce standard : `--tm-shadow-soft` = `0 2px 14px rgba(90, 70, 160, 0.07)`

## Structure de page type

```html
<body class="tm-body">
    <aside class="tm-sidebar">
        <div class="tm-brand">
            <span class="tm-brand-mark">M</span>
            <span class="tm-brand-name">MonApp</span>
        </div>
        <nav class="tm-nav">
            <a href="#" class="active">Tableau de bord</a>
            <a href="#">Section 2</a>
            <a href="#">Section 3</a>
        </nav>
    </aside>

    <main class="tm-content">
        <div class="tm-stat-grid">
            <div class="tm-stat-card tm-stat-card--1">
                <div class="tm-stat-value">128</div>
                <div class="tm-stat-label">Exemple de métrique</div>
            </div>
            <div class="tm-stat-card tm-stat-card--2">
                <div class="tm-stat-value">42%</div>
                <div class="tm-stat-label">Autre métrique</div>
            </div>
        </div>

        <div class="tm-panel">
            <div class="tm-panel-header">Titre de section</div>
            <div class="tm-panel-body">
                <table class="tm-table">
                    <thead><tr><th>Colonne A</th><th>Colonne B</th></tr></thead>
                    <tbody><tr><td>Valeur 1</td><td>Valeur 2</td></tr></tbody>
                </table>
                <a href="#" class="tm-btn tm-btn-primary">Action principale</a>
                <a href="#" class="tm-btn">Action secondaire</a>
            </div>
        </div>
    </main>
</body>
```

La sidebar (`.tm-sidebar`) est en position fixe et repasse automatiquement en barre horizontale
sous 900px de large (media query incluse dans le fichier).

## Composants disponibles

| Classe | Usage |
|---|---|
| `.tm-body` | Fond de page, typographie de base |
| `.tm-sidebar` / `.tm-brand` / `.tm-nav` | Navigation latérale |
| `.tm-content` | Zone de contenu principale (décalée pour la sidebar) |
| `.tm-stat-grid` / `.tm-stat-card` (+ `--1` à `--4`) | Cartes statistiques pastel |
| `.tm-panel` / `.tm-panel-header` / `.tm-panel-body` | Blocs de contenu |
| `.tm-btn` / `.tm-btn-primary` | Boutons |
| `.tm-input` / `.tm-select` / `.tm-textarea` | Champs de formulaire |
| `.tm-table` | Tableaux |
| `.tm-badge` (+ `-danger`, `-warning`, `-success`, `-neutral`) | Étiquettes de statut |
| `.tm-progress` / `.tm-progress-fill` | Barre de progression |

## Adapter la palette

Toutes les couleurs passent par des variables CSS custom properties définies sur `:root`. Pour
personnaliser (autre couleur d'accent, par exemple), il suffit de surcharger les variables après
l'import du fichier :

```html
<link rel="stylesheet" href="theme-moderne.css">
<style>
    :root {
        --tm-accent: #ff6b9d;      /* rose au lieu de violet */
        --tm-accent-dark: #e94f85;
        --tm-nav-active-bg: #ffe0ec;
    }
</style>
```

## Origine

Ce thème a été développé pour **MIRSAAD** (plateforme de pilotage de portefeuille de services
SI), en s'inspirant d'une maquette de type EdTech (sidebar, cartes de cours pastel). Il est
publié ici sous forme de fichier autonome pour être repris dans d'autres applications internes.
