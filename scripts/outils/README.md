# scripts/outils — utilitaires Excel

Modules et scripts Python partagés par `preparer_excel.py`, `generer_site.py` et les générateurs du dossier `scripts/site_web/`.

## Rôle dans le pipeline

**Phase 1 — `preparer_excel.py`**

```
preparer_excel.py
  ├── sync_excel_from_drive.py
  ├── sync_listes_validations.py   (optionnel : --skip-sync)
  ├── site_web/build_overview.py   (optionnel : --skip-overview)
  ├── verify_planning_workbook.py  (optionnel : --skip-verify)
  ├── geocode_excel.py             (optionnel : --skip-geocode)
  └── sync_excel_to_drive.py
```

**Phase 2 — `generer_site.py`**

```
generer_site.py
  ├── site_web/build_map.py        (lit le fichier Drive de base)
  ├── site_web/build_stats.py
  └── site_web/build_inspect.py
  └── sync_excel_from_drive.py     (contournement : --drive-pull)
```

Tous les scripts acceptent un chemin Excel en argument ; le défaut est `excel/Voyage Aout 2026.xlsx` (constante `DEFAULT_EXCEL_NAME` dans `excel_utils.py`).

## Fichiers

### Modules (bibliothèques)

| Fichier | Rôle |
|---------|------|
| `excel_utils.py` | Lecture/écriture du classeur : colonnes, feuilles Jour, listes déroulantes, export JSON carte (`build_voyage_data`), helpers hébergement et trajets (`is_trajet_line`, `lodging_*`), sauvegarde |
| `overview_config.py` | Chargement de `data/overview_config.json` (date de départ, feuille vue d'ensemble, marqueurs de vérif.) |

### Scripts exécutables

| Script | Rôle | Appelé par |
|--------|------|------------|
| `sync_listes_validations.py` | Resynchronise les plages de validation des feuilles Jour avec la feuille masquée `Listes` | `preparer_excel.py` |
| `verify_planning_workbook.py` | Contrôle la structure du classeur (feuilles, colonnes, ordres, combos) | `preparer_excel.py` |
| `geocode_excel.py` | Géocode via Nominatim les lieux sans Latitude/Longitude ; ignore les trajets déjà géolocalisés | `preparer_excel.py` |
| `sync_excel_from_drive.py` | Copie le `.xlsx` depuis Google Drive vers `excel/` (backup horodaté) | `preparer_excel.py`, `generer_site.py` |
| `sync_excel_to_drive.py` | Réécrit le classeur enrichi sur Google Drive (fichier de base) | `preparer_excel.py` (après géocodage) |
| `excel_workbook_sync.py` | Copie sécurisée partagée (validation, backup, écriture atomique) | modules ci-dessus |

## Utilisation directe

Depuis la racine du projet (avec l'environnement virtuel activé) :

```bash
# Synchroniser les listes déroulantes
python scripts/outils/sync_listes_validations.py
python scripts/outils/sync_listes_validations.py "excel/Mon voyage.xlsx" --dry-run

# Vérifier le classeur (code de sortie 1 si erreur bloquante)
python scripts/outils/verify_planning_workbook.py
python scripts/outils/verify_planning_workbook.py "excel/Mon voyage.xlsx"

# Géocoder les lieux manquants
python scripts/outils/geocode_excel.py
python scripts/outils/geocode_excel.py --dry-run
python scripts/outils/geocode_excel.py --force

# Synchroniser depuis / vers Google Drive (disque local)
python scripts/outils/sync_excel_from_drive.py
python scripts/outils/sync_excel_to_drive.py
python scripts/outils/sync_excel_from_drive.py --dry-run
python scripts/outils/sync_excel_from_drive.py --source "G:/Mon Drive/Voyage Aout 2026.xlsx"
```

Configuration Drive : `data/drive_config.json` (voir `data/drive_config.example.json`).

Les scripts ajoutent `scripts/` au `sys.path` pour importer `outils.*` ; lancez-les depuis la racine du dépôt ou via `preparer_excel.py` / `generer_site.py`.

## Structure Excel attendue

Feuilles obligatoires (dans cet ordre relatif pour `Liens` / `Listes`) :

1. `Vue d'ensemble` — synthèse (régénérée par `build_overview.py`)
2. `Liens` — URLs et libellés (manuel, jamais modifiée par le pipeline)
3. `Listes` — listes déroulantes (masquée)
4. `Jour 1` … `Jour 12`

Colonnes principales des feuilles Jour (ligne 2) : voir `PLANNING_COLUMNS` dans `excel_utils.py`. Colonnes carte ajoutées automatiquement : `Latitude`, `Longitude`, `Lien`.

## Fichiers produits

| Fichier | Producteur |
|---------|------------|
| `data/geocode_cache.json` | `geocode_excel.py` |
| `data/geocode_errors.csv` | `geocode_excel.py` |
| `excel/backups/*.backup.xlsx` | `backup_excel()` avant toute écriture Excel |
| `excel/backups/*.backup.YYYYMMDD-HHMMSS.xlsx` | `sync_excel_from_drive.py` avant remplacement local |
| `excel/backups/*.drive.backup.YYYYMMDD-HHMMSS.xlsx` | `sync_excel_to_drive.py` avant écrasement du fichier Drive |
| `data/overview_config.json` | configuration utilisateur (lu par `overview_config.py`) |

## Dépendances

Listées dans `scripts/requirements.txt` : `openpyxl`, `requests` (géocodage Nominatim).

## Documentation complémentaire

- [README principal](../../README.md) — démarrage rapide et publication
- [docs/CAHIER_DES_CHARGES.md](../../docs/CAHIER_DES_CHARGES.md) — spécification détaillée
- [data/hardcode_generer_site.txt](../../data/hardcode_generer_site.txt) — éléments encore codés en dur dans le pipeline
