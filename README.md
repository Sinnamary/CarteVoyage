# CarteVoyage

Carte interactive et statistiques de voyage, générées à partir d'un fichier Excel de planning.

**Site en ligne :** [https://sinnamary.github.io/CarteVoyage/web/](https://sinnamary.github.io/CarteVoyage/web/)

- [Carte](https://sinnamary.github.io/CarteVoyage/web/index.html) — points géolocalisés, filtres par jour, trajets à pied ou en voiture
- [Statistiques](https://sinnamary.github.io/CarteVoyage/web/stats.html) — distances, durées, budget, répartition par jour et par ville
- [Contrôle](https://sinnamary.github.io/CarteVoyage/web/inspect.html) — cohérence des JSON (carte, stats, overview)

## Principe

```
Excel → generer_site.py → JSON + HTML statique → navigateur (Leaflet)
```

Une seule source de vérité : le classeur Excel. Pas de serveur applicatif — le site est entièrement statique.

## Démarrage rapide

```bash
# Environnement
python -m venv .venv
.venv\Scripts\activate          # Windows
pip install -r scripts/requirements.txt

# Placer le fichier Excel dans excel/ (par défaut : Voyage Aout 2026.xlsx)
# — ou le synchroniser depuis Google Drive (voir ci-dessous)

python generer_site.py

# Vérification locale
start web/index.html
start web/stats.html
start web/inspect.html
```

### Options utiles

```bash
python generer_site.py --skip-geocode   # saute toute l'étape geocodage
python generer_site.py --skip-verify    # sans contrôle du classeur
python generer_site.py --no-osrm        # stats sans OSRM
python generer_site.py --geocode-force  # re-géocoder même les lieux déjà renseignés
```

### Géocodage

Lors d'un `generer_site.py` normal, `geocode_excel.py` ne contacte **Nominatim que pour les lieux sans Latitude/Longitude** dans Excel. Les points déjà géolocalisés sont ignorés (sauf avec `--geocode-force`). Les lignes « Trajet … » / « Retour … » ne sont jamais géocodées.

## Synchronisation Google Drive

Si le classeur est partagé sur Google Drive et monté comme disque local (Google Drive pour ordinateur) :

1. Copier `data/drive_config.example.json` vers `data/drive_config.json`
2. Renseigner `source_path` avec le chemin complet sur le disque Google Drive (ex. `G:/Mon Drive/Voyages/Voyage Aout 2026.xlsx`)
3. Lancer la chaîne complète :

```powershell
.\sync_excel.ps1
```

Étapes séparées :

```powershell
python scripts/outils/sync_excel_from_drive.py   # copie + backup horodaté
python generer_site.py
.\publier.ps1
```

L'ancienne version locale est sauvegardée dans `excel/backups/` avant chaque remplacement (`*.backup.YYYYMMDD-HHMMSS.xlsx`).

**Note :** le fichier source doit être un vrai `.xlsx` sur le disque. Un Google Sheet natif (fichier `.gsheet`) n'est pas copiable directement — enregistrez-le au format Excel sur Drive ou travaillez avec une copie `.xlsx` synchronisée.

## Publication (étape séparée)

`generer_site.py` ne pousse rien sur Git. Une fois le site vérifié en local :

```powershell
.\publier.ps1
.\publier.ps1 "Message de commit"
```

`publier.ps1` enregistre et envoie les fichiers déjà générés (`web/`, `data/`, etc.) sur GitHub Pages.

## Structure

```
CarteVoyage/
├── generer_site.py     # Génération locale (point d'entrée)
├── sync_excel.ps1      # Sync Drive → generer_site → publier
├── publier.ps1         # Publication Git uniquement
├── excel/              # Classeur source (non versionné)
├── data/               # JSON et rapports générés
├── scripts/
│   ├── site_web/       # Génération du site web
│   └── outils/         # Excel, géocodage, vérification
├── web/                # Site statique
└── docs/               # Documentation détaillée
```

## Scripts

### Racine

| Script | Rôle |
|--------|------|
| `generer_site.py` | Pipeline local Excel → site web |
| `publier.ps1` | Commit + push GitHub (après vérif. locale) |

### `scripts/site_web/` — site web

| Script | Rôle |
|--------|------|
| `build_overview.py` | Régénère la feuille Excel `Vue d'ensemble` |
| `build_map.py` | Génère `voyages.json` et `web/index.html` |
| `build_stats.py` | Génère `stats.json` et `web/stats.html` |
| `build_inspect.py` | Génère `inspect.json` et `web/inspect.html` |

### `scripts/outils/` — utilitaires

Voir [scripts/outils/README.md](scripts/outils/README.md) pour le détail.

| Script / module | Rôle |
|-----------------|------|
| `excel_utils.py` | Bibliothèque partagée (colonnes, feuilles Jour, JSON carte) |
| `overview_config.py` | Chargement de `data/overview_config.json` |
| `geocode_excel.py` | Géocode les lieux sans coordonnées (Nominatim) |
| `sync_listes_validations.py` | Resynchronise les listes déroulantes Excel |
| `verify_planning_workbook.py` | Vérifie la structure du classeur |

## Configuration vue d'ensemble

La feuille Excel **Vue d'ensemble** est **générée automatiquement** à chaque exécution de `generer_site.py` (via `build_overview.py`). Ne pas la modifier à la main : les changements seront écrasés à la prochaine génération.

Fichier `data/overview_config.json` : titre, date de départ, domicile et marqueurs de vérification. Produit aussi `data/overview.json` (snapshot calculé).

## Fichier Excel

- Feuilles : **`Vue d'ensemble`** (générée automatiquement), `Listes` (masquée), `Jour 1` … `Jour N`
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
