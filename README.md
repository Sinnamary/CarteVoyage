# CarteVoyage

Carte interactive et statistiques de voyage, générées à partir d'un fichier Excel de planning.

**Site en ligne :** [https://sinnamary.github.io/CarteVoyage/web/](https://sinnamary.github.io/CarteVoyage/web/)

- [Carte](https://sinnamary.github.io/CarteVoyage/web/index.html) — points géolocalisés, filtres par jour, trajets à pied ou en voiture
- [Statistiques](https://sinnamary.github.io/CarteVoyage/web/stats.html) — distances, durées, budget, répartition par jour et par ville
- [Contrôle](https://sinnamary.github.io/CarteVoyage/web/inspect.html) — cohérence des JSON (carte, stats, overview)

## Principe

```
Google Drive (.xlsx)  →  preparer_excel  →  corrections sur Drive  →  generer_site  →  web/  →  publier.ps1  →  GitHub Pages
```

**Fichier de base** : le classeur `.xlsx` sur Google Drive (`data/drive_config.json`). Les enrichissements automatiques (vue d'ensemble, géolocalisation) y sont réécrits. C'est ce fichier qui sert à générer le site.

**Copie locale** : `excel/Voyage Aout 2026.xlsx` est la copie de travail utilisée pendant la phase 1 (téléchargement Drive → enrichissements → renvoi sur Drive). Elle n'est pas versionnée dans Git (voir [Fichiers Excel et sauvegardes](#fichiers-excel-et-sauvegardes-locales)).

**Git** : le programme complet vit dans le dépôt Git ; la **publication** (`publier.ps1`) n'envoie que le dossier `web/` sur GitHub Pages.

## Workflow : Drive → site → publication

Le processus est volontairement découpé en **trois phases** séparées, avec des pauses pour corriger sur Google Drive et vérifier le site en local avant publication.

**Convention :** lancer un script **sans argument** exécute l'étape complète du workflow. Les paramètres (`--skip-*`, `-WebOnly`, etc.) ne servent qu'à **contourner ponctuellement** une étape.

### Schéma

```
┌─────────────────────────────────────────────────────────────────┐
│  PHASE 1 — preparer_excel.ps1                                   │
│  Lire Drive + backup → enrichissements → vérifications          │
│  → écriture sur Google Drive                                    │
└───────────────────────────┬─────────────────────────────────────┘
                            │
              ┌─────────────▼─────────────┐
              │  Erreurs de vérification ? │
              └─────────────┬─────────────┘
                    oui     │     non
              ┌─────────────▼─────────────┐
              │ Corriger sur Google Drive  │◄────┐
              │ Relancer phase 1             │     │
              └─────────────┬───────────────┘     │
                            │ non                 │
┌───────────────────────────▼─────────────────────────────────────┐
│  PHASE 2 — generer_site.ps1                                     │
│  Générer carte, stats, contrôle dans web/                       │
└───────────────────────────┬─────────────────────────────────────┘
                            │
              ┌─────────────▼─────────────┐
              │  Contrôle local OK ?       │
              │  (web/index.html…)         │
              └─────────────┬─────────────┘
                    non     │     oui
              ┌─────────────┘             │
              │ Corriger sur Google Drive │
              │ Phases 1 + 2              ├──────┘
              └───────────────────────────┘
                            │ oui
┌───────────────────────────▼─────────────────────────────────────┐
│  PHASE 3 — publier.ps1                                          │
│  Commit + push GitHub Pages (site uniquement, pas l'Excel)      │
└─────────────────────────────────────────────────────────────────┘
```

### Commandes

| Phase | Action | Commande |
|-------|--------|----------|
| **1** | Lire Drive, backup, vérifier, enrichir, réécrire sur Drive | `.\preparer_excel.ps1` |
| | *Si erreurs : corriger le `.xlsx` sur Google Drive, puis relancer la phase 1* | |
| **2** | Générer le site web local | `.\generer_site.ps1` |
| | *Contrôler dans le navigateur* | `start web/index.html` |
| | *Si le site ne convient pas : corriger sur Drive, relancer phases 1 et 2* | |
| **3** | Publier en ligne | `.\publier.ps1` |

Équivalent Python :

```powershell
python preparer_excel.py
python generer_site.py
.\publier.ps1
```

### `sync_excel.ps1` — orchestrateur

Raccourci qui enchaîne les phases sans logique métier propre : il appelle `preparer_excel.py`, puis `generer_site.py`, et éventuellement `publier.ps1`.

```powershell
.\sync_excel.ps1                  # phases 1 + 2 (défaut)
.\sync_excel.ps1 -Publish          # phases 1 + 2 + 3
.\sync_excel.ps1 -PrepareOnly      # phase 1 seulement
.\sync_excel.ps1 -WebOnly          # phase 2 seulement
.\sync_excel.ps1 -Publish -Message "Mise à jour carte jour 5"
```

Équivalent manuel du comportement par défaut :

```powershell
.\sync_excel.ps1
start web/index.html
.\publier.ps1
```

| Option | Effet | Sauvegardes locales (`excel/backups/`) |
|--------|-------|----------------------------------------|
| *(aucune)* | Phases 1 + 2 | **Oui** — via la phase 1 |
| `-PrepareOnly` | Phase 1 seulement | **Oui** |
| `-WebOnly` | Phase 2 seulement | **Non** — la phase 1 est ignorée |
| `-Publish` | Phases 1 + 2 + `publier.ps1` | **Oui** — sauvegardes Excel avant publication |

**Sans argument, chaque script exécute son étape du workflow.** Les paramètres ne servent qu'à contourner ponctuellement une étape (voir `python preparer_excel.py --help` ou `python generer_site.py --help`).

### Prérequis Google Drive

1. Google Drive pour ordinateur installé et le classeur synchronisé localement (fichier `.xlsx`, pas un raccourci `.gsheet`).
2. Copier `data/drive_config.example.json` vers `data/drive_config.json`.
3. Renseigner `source_path` avec le chemin complet sur le disque (ex. `G:/Mon Drive/Voyages/Voyage Aout 2026.xlsx`).

La phase 1 copie le fichier Drive vers `excel/` (copie de travail + backup horodaté dans `excel/backups/`), enrichit ce classeur, vérifie la structure, puis **réécrit le fichier de base sur Google Drive**. La phase 2 lit ce fichier Drive pour produire `web/`.

### Phase 1 — ce que fait `preparer_excel.py`

```
preparer_excel.py
  ├── outils/sync_excel_from_drive.py   → lecture Drive + backup local
  ├── outils/sync_listes_validations.py → listes déroulantes alignées sur Listes
  ├── site_web/build_overview.py        → feuille Vue d'ensemble régénérée
  ├── outils/verify_planning_workbook.py → contrôle structure (arrêt si erreur)
  ├── outils/geocode_excel.py           → Latitude / Longitude (Nominatim)
  └── outils/sync_excel_to_drive.py     → écriture du classeur enrichi sur Drive
```

| Modification Excel | Détail |
|--------------------|--------|
| **Vue d'ensemble** | Recalculée depuis les feuilles `Jour N` ; ne pas l'éditer à la main ; config dans `data/overview_config.json` |
| **Géolocalisation** | Colonnes `Latitude`, `Longitude`, `Lien` pour les nouveaux lieux (sauf `--geocode-force`) |
| **Listes déroulantes** | Validations remises en phase avec la feuille masquée `Listes` |

Si `verify_planning_workbook.py` signale des erreurs, le script s'arrête **avant** l'écriture sur Drive et **avant** toute génération web. Corrigez les feuilles `Jour N` sur Google Drive, puis relancez la phase 1.

### Phase 2 — ce que fait `generer_site.py`

```
generer_site.py
  ├── site_web/build_map.py      → data/voyages.json + web/index.html (lit le fichier Drive)
  ├── site_web/build_stats.py    → data/stats.json + web/stats.html
  └── site_web/build_inspect.py  → data/inspect.json + web/inspect.html
```

Par défaut, lit directement le fichier de base sur Google Drive (sans re-téléchargement). Aucune modification du classeur Excel, donc **aucune sauvegarde** dans `excel/backups/` à cette étape. Avec `--drive-pull`, un téléchargement Drive → `excel/` est effectué avant la génération (backup horodaté inclus).

### Fichiers Excel et sauvegardes locales

#### Deux emplacements du classeur

| Emplacement | Rôle |
|-------------|------|
| **Google Drive** (`source_path` dans `data/drive_config.json`) | Fichier de référence : celui que vous éditez et partagez. Source pour la phase 2. |
| **`excel/Voyage Aout 2026.xlsx`** | Copie de travail locale (nom par défaut, constante `DEFAULT_EXCEL_NAME`). Utilisée pendant la phase 1 pour les enrichissements avant réécriture sur Drive. |

Sans `data/drive_config.json`, les phases 1 et 2 travaillent uniquement sur `excel/Voyage Aout 2026.xlsx` (pas de sync Drive).

La copie dans `excel/` peut être **légèrement en retard** par rapport à Drive si vous avez modifié le fichier sur Drive sans relancer `preparer_excel.ps1`. C'est normal : relancez la phase 1 pour resynchroniser.

#### Sauvegardes dans `excel/backups/`

Les sauvegardes sont des **copies de sécurité locales** (pas sur un autre disque). Elles sont créées automatiquement par la phase 1 ; `sync_excel.ps1` ne fait rien lui-même, il délègue à `preparer_excel.py`.

| Fichier produit | Quand | Script |
|---------------|-------|--------|
| `*.backup.YYYYMMDD-HHMMSS.xlsx` | Avant de remplacer `excel/Voyage Aout 2026.xlsx` par la version Drive | `sync_excel_from_drive.py` |
| `*.backup.xlsx` | Avant chaque écriture du classeur local (listes, vue d'ensemble, géocodage) — **écrasé** à chaque nouvelle écriture | `sync_listes_validations.py`, `build_overview.py`, `geocode_excel.py` |
| `*.drive.backup.YYYYMMDD-HHMMSS.xlsx` | Avant d'écraser le fichier sur Google Drive | `sync_excel_to_drive.py` |

Exemples concrets :

- `Voyage Aout 2026.backup.20260612-073756.xlsx` — ancienne copie locale avant téléchargement Drive
- `Voyage Aout 2026.backup.xlsx` — état juste avant la dernière modification locale
- `Voyage Aout 2026.drive.backup.20260612-075928.xlsx` — ancienne version Drive avant push

**Nettoyage :** vous pouvez supprimer les anciens `.xlsx` dans `excel/backups/` pour libérer de l'espace ; le pipeline en recréera au prochain run. Conservez `.gitkeep` (dossier versionné vide). Les backups ne sont pas dans Git (`.gitignore` : `*.xlsx`).

## Démarrage rapide

```powershell
# Environnement
python -m venv .venv
.venv\Scripts\activate          # Windows
pip install -r scripts/requirements.txt

# Workflow complet (Google Drive configuré dans data/drive_config.json)
.\preparer_excel.ps1
.\generer_site.ps1
start web/index.html            # contrôle local avant publication
.\publier.ps1

# Raccourci : phases 1 + 2 en une commande
.\sync_excel.ps1
start web/index.html
.\publier.ps1
```

Sans `data/drive_config.json`, les phases 1 et 2 travaillent automatiquement sur `excel/` (aucun paramètre requis).

### Contournements ponctuels

À n'utiliser qu'en exception — le comportement par défaut suit le workflow :

```bash
python preparer_excel.py --help    # ex. --skip-geocode, --geocode-force
python generer_site.py --help      # ex. --drive-pull, --no-osrm
```

### Géocodage

Lors d'un `preparer_excel.py` normal, `geocode_excel.py` ne contacte **Nominatim que pour les lieux sans Latitude/Longitude**. Les points déjà géolocalisés sont ignorés (sauf `--geocode-force`). Les lignes « Trajet … » / « Retour … » ne sont jamais géocodées.

## Synchronisation Google Drive (détail)

```powershell
python scripts/outils/sync_excel_from_drive.py   # Drive → excel/ (backup horodaté)
python scripts/outils/sync_excel_to_drive.py     # excel/ → Drive (intégré à preparer_excel)
python scripts/outils/sync_excel_from_drive.py --dry-run
```

**Note :** le fichier source doit être un vrai `.xlsx` sur le disque. Un Google Sheet natif (fichier `.gsheet`) n'est pas copiable directement — enregistrez-le au format Excel sur Drive.

## Hébergement (GitHub Pages)

Le site est **entièrement statique** : GitHub Pages sert le dossier `web/` sans exécuter de code serveur. Les données du voyage sont embarquées dans les pages HTML ; le fichier Excel n'est jamais exposé en ligne.

```
generer_site.py  →  web/  →  publier.ps1  →  git push  →  GitHub Actions  →  GitHub Pages
```

| Page | URL |
|------|-----|
| Carte | [sinnamary.github.io/CarteVoyage/web/index.html](https://sinnamary.github.io/CarteVoyage/web/index.html) |
| Statistiques | [sinnamary.github.io/CarteVoyage/web/stats.html](https://sinnamary.github.io/CarteVoyage/web/stats.html) |
| Contrôle | [sinnamary.github.io/CarteVoyage/web/inspect.html](https://sinnamary.github.io/CarteVoyage/web/inspect.html) |

Le préfixe `/web/` vient de la structure du dépôt : le site généré vit dans le sous-dossier `web/`.

### `publier.ps1`

Ne régénère **pas** le site : lancer `generer_site.ps1` (ou `sync_excel.ps1`) au préalable, puis vérifier dans le navigateur.

```powershell
.\publier.ps1
.\publier.ps1 "Mise à jour carte jour 5"
```

Comportement :

1. `git add web/` — seuls les fichiers du site sont indexés
2. Si aucune modification dans `web/` : message d'avertissement, pas de commit
3. `git commit` puis `git push` vers le dépôt distant
4. GitHub Actions (`.github/workflows/pages.yml`) déploie `web/` sur Pages (délai habituel : 1 à 2 minutes)

**Prérequis :** dépôt Git avec remote configuré, droits d'écriture sur `master`, site déjà généré et validé localement.

### GitHub Actions

Workflow **Publier la page web sur GitHub Pages** :

- Déclenché à chaque push sur `master`, ou manuellement (`workflow_dispatch`)
- Déploie le dossier `web/` tel que commité — **ne régénère pas** le site
- Un seul déploiement à la fois (`cancel-in-progress`)

## Versionnement Git

Deux usages distincts :

| Quoi | Où | Comment |
|------|-----|---------|
| **Programme complet** | dépôt Git (local ou distant) | scripts, docs, `data/`, configuration d'exemple… |
| **Site en ligne** | GitHub Pages | dossier `web/` uniquement |

### Sauvegarder le programme

Le dépôt Git contient tout le code et la doc. Fichiers **exclus** par `.gitignore` :

- `excel/*.xlsx` — classeur et backups (reste sur Google Drive)
- `data/drive_config.json` — chemin personnel vers votre Drive
- `.venv/` — environnement Python

Sauvegarde manuelle du programme (quand vous modifiez scripts ou doc) :

```powershell
git add -A
git commit -m "Evolution du pipeline"
git push          # optionnel : vers votre dépôt distant
```

Un dépôt **uniquement local** (`git init` sans `remote`) convient aussi : le programme y est versionné ; seul `publier.ps1` nécessite un remote pour l'hébergement en ligne.

> Voir aussi la section [Hébergement (GitHub Pages)](#hébergement-github-pages) pour le détail de `publier.ps1` et GitHub Actions.

## Structure

```
CarteVoyage/
├── preparer_excel.py / .ps1    # Phase 1 : Excel (Drive ou local)
├── generer_site.py / .ps1      # Phase 2 : génération web/
├── sync_excel.ps1              # Orchestrateur (phases 1 + 2 + option -Publish)
├── publier.ps1                 # Phase 3 : publication GitHub Pages
├── excel/                      # Copie de travail (non versionnée)
│   ├── Voyage Aout 2026.xlsx
│   └── backups/
├── data/
│   ├── drive_config.json       # Chemin Google Drive (local, non versionné)
│   ├── drive_config.example.json
│   ├── overview_config.json
│   ├── voyages.json, stats.json, inspect.json
│   └── geocode_cache.json, route_stats_cache.json, …
├── scripts/
│   ├── pipeline_common.py
│   ├── requirements.txt
│   ├── site_web/               # build_map, build_stats, build_inspect, build_overview
│   └── outils/                 # géocodage, sync Drive, vérifications Excel
├── web/                        # Site statique (publié sur GitHub Pages)
│   ├── index.html, stats.html, inspect.html
│   └── assets/
├── .github/workflows/pages.yml # Déploiement GitHub Pages
├── logos/
└── docs/
```

## Scripts

### Racine

| Script | Phase | Rôle |
|--------|-------|------|
| `preparer_excel.ps1` | 1 | Lecture Drive, vérifications, enrichissements, écriture Drive |
| `generer_site.ps1` | 2 | Génération `web/` depuis le fichier de base |
| `publier.ps1` | 3 | Commit + push du dossier `web/` (GitHub Pages) |
| `sync_excel.ps1` | 1–3 | Orchestrateur : enchaîne `preparer_excel` + `generer_site` ; `-Publish` pour la phase 3 |

### `scripts/site_web/` — site web

| Script | Rôle |
|--------|------|
| `build_overview.py` | Régénère la feuille Excel `Vue d'ensemble` (phase 1) |
| `build_map.py` | Génère `voyages.json` et `web/index.html` |
| `build_stats.py` | Génère `stats.json` et `web/stats.html` |
| `build_inspect.py` | Génère `inspect.json` et `web/inspect.html` |

### `scripts/outils/` — utilitaires

Voir [scripts/outils/README.md](scripts/outils/README.md).

## Configuration vue d'ensemble

La feuille Excel **Vue d'ensemble** est **générée automatiquement** par `preparer_excel.py` (via `build_overview.py`). Ne pas la modifier à la main.

Fichier `data/overview_config.json` : titre, date de départ, domicile et marqueurs de vérification.

## Fichier Excel

- **Où le modifier :** sur Google Drive (fichier pointé par `data/drive_config.json`), pas dans `excel/` sauf en mode sans Drive
- **Copie locale :** `excel/Voyage Aout 2026.xlsx` — voir [Fichiers Excel et sauvegardes](#fichiers-excel-et-sauvegardes-locales)
- Feuilles : **`Vue d'ensemble`** (générée), `Listes` (masquée), `Jour 1` … `Jour N`
- Colonnes principales : `N° étape`, `Lieu`, `Nature`, `Catégorie`, `Quartier`, `Ville`, `Prix (€)`, horaires…
- Colonnes carte (auto) : `Latitude`, `Longitude`, `Lien`
- Saisir les étapes `.10` en **format texte** (`@`) pour éviter les collisions d'ordre

## Documentation

Spécification complète : [docs/CAHIER_DES_CHARGES.md](docs/CAHIER_DES_CHARGES.md)

## Dépendances externes

- [Nominatim](https://nominatim.openstreetmap.org/) — géocodage (scripts)
- [OSRM](https://project-osrm.org/) — itinéraires pied/voiture (navigateur et stats)
- [OpenStreetMap](https://www.openstreetmap.org/) — tuiles carte
- [Leaflet](https://leafletjs.com/) 1.9.4 — cartographie
