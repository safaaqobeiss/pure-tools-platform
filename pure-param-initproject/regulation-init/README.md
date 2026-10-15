# Regulation Project Initializer

Python tool with a Tkinter GUI that generates **Regulation** and/or **Output** Excel project files for new regulatory frameworks, and optionally publishes them to GitLab.

## Prerequisites

- Python 3.10+
- `pip install -r requirements.txt`
- GitLab token with `api` scope (for publishing)

## Quick Start

```bash
pip install -r requirements.txt
python src/main.py
```

A launcher window appears with two options:

1. ** Jira PDF Import** — Automatic field extraction from a Jira ticket (PDF)
2. ** Manual Entry** — Step-by-step interactive wizard

## Mode 1: Jira PDF Import (automatic)

1. Select a PDF file exported from Jira (XRAY Test Set ticket `PURE_RQM-*`)
2. Fields are extracted automatically (code, version, dates, author, etc.)
3. A summary is displayed in the window
4. Click ** Generate files** to create the project locally
5. Optional: ** Push to GitLab** to publish the files


## Mode 2: Manual Entry (wizard)

The wizard guides you through 7 steps:

| Step | Description |
|------|-------------|
| 1 | Regulatory code (e.g. COREP, FINREP) |
| 2 | Project type — Regulation, Output, or Both |
| 3 | Git project name |
| 4 | Regulation file selection (with editable target names) |
| 5 | Output file selection (with editable target names) |
| 6 | GitLab publishing (branch, token) |
| 7 | Summary & Generation |



## Generated Structure

```
generated_projects/<project-name>/
├── regulation/
│   ├── pure-dictionary-<CODE>.xlsx
│   ├── pure-regulation-<CODE>.xlsx
│   ├── pure-transformationRules-<CODE>.xlsx
│   ├── pure-transformedDataStructure-<CODE>.xlsm
│   └── pure-uiDashboard-<CODE>.xlsx
└── output/
    ├── pure-regulatoryOutput-<CODE>-<OUTPUT>.xlsx
    └── pure-outputStructure-<OUTPUT>-<STRUCT>.xlsx
```

## GitLab Publishing

The GitLab dialog prompts for:

- **Target branch** (default: `main`)
- **GitLab token** (prefilled from `GITLAB_TOKEN`)
- **Regulation project name** (default: `pure-regulation-<name>`)
- **Output project name** (default: `pure-output-<name>`)

Token resolution order:
1. `GITLAB_TOKEN` environment variable
2. Manual entry in the GUI


## Scripts utilitaires

| Script | Description |
|--------|-------------|
| `setup_env.bat` | Vérifie Python, installe pip et les dépendances |
| `run_local.bat` | Lance l'application (`py src\main.py`) |
| `build_exe.bat` | Génère un exécutable standalone via PyInstaller (`dist\regulation-init.exe`) |

## Build exécutable

Pour créer un `.exe` autonome (ne nécessite pas Python sur la machine cible) :

```bash
build_exe.bat
```

L'exécutable est généré dans `dist\regulation-init.exe`.

## Tests e2e

Des données de test de bout en bout sont disponibles dans :

- `src/regulation/e2e-test-data/` — fichiers d'entrée et résultats attendus pour le module Regulation
- `src/Output/e2e-test-data/` — fichiers d'entrée et résultats attendus pour le module Output

## Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl` | Read/write Excel files |
| `pymupdf` | Extract text from Jira PDFs |
| `requests` | GitLab API calls |
| `pyyaml` | YAML configuration |
