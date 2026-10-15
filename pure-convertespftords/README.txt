================================================================================
ESPF to RDS Conversion Tool — README
================================================================================

1. OVERVIEW
-----------
This tool converts ESPF fixed-width files (.sXX) into RDS-compliant CSV files
for two dataset types: CATR (frbdfruba_catr) and ENGAG (frbdfruba_engag).

It reads ESPF lines, extracts header and variable zone fields based on
position/length definitions, applies mapping rules, and outputs CSV files
following the official RDS column structure.


2. REQUIREMENTS
---------------
- Python 3.10+
- openpyxl (installed in .venv)

Setup:
  .venv\Scripts\activate
  pip install openpyxl


3. FILES
--------
convert_espf_to_rds.py
    Main conversion script.

2025-11-21_Correspondance_entrees_specifiques CATR & creditContracts_NW.xlsx
    Mapping file: defines which ESPF zone maps to which RDS column.
    Sheets: "catr" (CATR mappings), "Engag" (ENGAG mappings).

pure-transformedDataStructure-frbdfruba.xlsm
    RDS structure file: defines all RDS columns and their order.
    Sheet: "structuresTransformations".

RUBA_BDF_FR.xlsx
    ESPF structure file: defines zone positions and lengths per ESPF type.
    Sheet: "Spins".

expected_reference.txt
    Reference file documenting the expected output format (column order,
    rules, sample data) used as the basis for output ordering.

_compare_results.py
    Utility to compare generated CSV output against reference files.


4. USAGE
--------

Interactive mode (with dialogs):
  python convert_espf_to_rds.py

  Flow:
    1. File dialog: select ESPF input file (.sXX)
    2. Parameter dialog: enter Institution (default: RUBA)
       and Periodicity (default: monthly)
    3. Folder dialog: select output directory
    4. Conversion runs automatically

Command-line mode (no dialogs):
  python convert_espf_to_rds.py ^
      --input-sxx <your_file.sXX> ^
      --output-dir output ^
      --institution-code RUBA ^
      --periodicity monthly ^
      --no-dialog

Analysis mode:
  python convert_espf_to_rds.py --input-sxx <your_file.sXX> --analyze --no-dialog --output-dir output
  python convert_espf_to_rds.py --input-sxx <your_file.sXX> --analyze-only --no-dialog

All CLI arguments:
  --input-sxx           Path to ESPF input file
  --output-dir          Directory for generated CSV files
  --mapping-file        Path to mapping workbook (default: 2025-11-21_...xlsx)
  --espf-structure-file Path to ESPF structure workbook (default: RUBA_BDF_FR.xlsx)
  --rds-structure-file  Path to RDS structure workbook (default: pure-...xlsm)
  --periodicity         Reporting periodicity (default: monthly)
  --institution-code    Institution code (default: RUBA)
  --analyze             Analyze ESPF format before conversion
  --analyze-only        Only analyze, no CSV generation
  --no-dialog           Disable file picker dialogs


5. OUTPUT
---------
Two CSV files are generated:
  - frbdfruba_catr.csv   (25 columns)
  - frbdfruba_engag.csv  (33 columns)

Column order follows the Expected order (defined in EXPECTED_COLUMNS in the
script, documented in expected_reference.txt). Only mapped columns and known
reference columns are included.

Columns with no defined mapping are filled with <NoMapping>.
Columns with a mapping but no data in the source are left empty.

The first 7 columns are always:
  datasetCode, reportingSessionPeriodicity, reportingSessionDate,
  institutionCode, datasourceCode, identifier, commonDatasetRecordId


6. HOW IT WORKS
---------------

Input parsing:
  - Each line of the ESPF file is a fixed-width record.
  - The common header (first 228 bytes) contains shared fields:
    DTSITU (date), CDFILI (datasource), CDDEVI (currency),
    NOCPCI (internal account), VLRFCE (reference value).
  - The FMESTD field (position 163, length 5) identifies the record type
    (CATR or ENGAG).
  - Variable zones follow the header, with positions/lengths defined
    in RUBA_BDF_FR.xlsx (Spins sheet).

Mapping:
  - The mapping file links ESPF zone names to RDS column names.
  - CATR: 16 active mappings (MONTCONTR, DUREEINI, AGENTECO, NATCAT,
    MATURITE, INDEXREF, TESE1-9, SPREAD, SENSMONT).
  - ENGAG: 23 active mappings (country codes, amounts, instrument types, etc.).

Value formatting:
  - Signed numeric zones (starting with + or -) are divided by a scale
    factor (e.g. MONTCONTR / 100, TESE1-9 / 1000000, SPREAD / 10000).
  - Amount zones (MTENCOURS, MTDUREE1-4) default to "0" when blank.
  - Leading zeros are stripped from numeric values.

Column ordering:
  - Output includes only BASE columns + expected columns per type.
  - CATR: 18 columns after the 7 base columns (25 total).
  - ENGAG: 26 columns after the 7 base columns (33 total).
  - Order follows the EXPECTED_COLUMNS constant in the script
    (documented in expected_reference.txt).


7. CONFIGURATION CHANGES
-------------------------

To change mapping (add/modify a zone-to-column correspondence):
  Edit the mapping Excel file, sheet "catr" or "Engag", column C (SBPRR).
  No code change needed.

To change column order:
  Edit the EXPECTED_COLUMNS dict in convert_espf_to_rds.py.
  This defines which columns appear and in what order for each type.

To add a new ESPF type:
  1. Add zone definitions in RUBA_BDF_FR.xlsx (Spins sheet).
  2. Add a mapping sheet in the mapping Excel file.
  3. Add the type to the TARGETS dict in convert_espf_to_rds.py.

To change decimal scaling:
  Edit the DECIMAL_SCALES dict in convert_espf_to_rds.py.
