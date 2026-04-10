# Extraction de la hierarchie CIM-10

Ce dossier contient un script Python qui lit le fichier Excel `cim10_atih_FR_2025.xlsx` et genere le fichier CSV `cim10_hierarchie.csv`.

## Prerequis

- Python 
- Et la dépendance `openpyxl`

## Installation

Depuis ce dossier :

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install openpyxl
```

Si le dossier `.venv` existe deja, il suffit de le reactiver :

```bash
source .venv/bin/activate
```

## Lancement

```bash
python script.py \
  --input ./cim10_atih_FR_2025.xlsx \
  --output-dir .
```

Le script doit afficher une sortie proche de :

```text
Source : /chemin/vers/cim10_atih_FR_2025.xlsx
Output : /chemin/vers/cim10_hierarchie.csv
Lines : 19075
```

Le fichier généré contient l'entête suivant :

```csv
code;label;parent
```

## Options utiles

- `--output-name NOM.csv` : change le nom du fichier CSV genere.
- `--sheet "Nom de feuille"` : force la lecture d'une feuille Excel precise.
