# pure-param-check-submission-input

Technical README — Input vs. Submission Validation Tool (cell-based)


## 1. Overview

This tool compares **input data** (CSV or Excel) against **expected XBRL submissions** at the cell level, for any PURE Param regulatory output.

Each comparison is based on a cell key `(tableCode, rowId, columnId)`, resolved automatically from the EBA/ACPR taxonomy JSON mapping.

Main capabilities:

- Convert XBRL-XML submissions to CSV via `Xml2Csv.jar`
- Extract XBRL-CSV ZIP packages directly (P3DH, REM)
- Parse OneGate XML submissions natively (MCONTRAN)
- Compare input data vs. converted submission output, cell by cell
- Auto-resolve table code mismatches between input and taxonomy
- Normalize numeric values for reliable comparison
- Generate colored Excel and CSV reports with cell-level statuses
- Batch mode with automatic test case detection and global summary

Primary implementation files:

```
validate_regulation_unified.py   # Main validation engine (CLI)
launch_validation.py             # Interactive wizard launcher
```


## 2. Core Capabilities

- **XBRL submission ingestion**: XBRL-XML (`.xbrl`), XBRL-CSV ZIP (`.zip`), OneGate XML (`.xml`)
- **Automatic taxonomy mapping**: loads `datapoint → (table, row, column, cellCode)` from taxonomy JSON files (`xbrl_package.tar.gz`)
- **Input data reading**: CSV direct (`tableCode;rowId;columnId;value`) or Excel (auto-converted)
- **XBRL pre-processing**:
  - Repositions `<link:schemaRef>` if misplaced in the XBRL file
  - Deduplicates `<xbrli:context>` elements
  - Sanitizes invalid XML identifiers (NCName compliance)
- **Auto-resolution of table codes**: resolves input codes to the closest taxonomy match (e.g. `RB.26.01.11` → `RB.26.01.01`)
- **Value normalization**: handles comma/dot decimal separators, trailing decimals, whitespace, non-breaking spaces
- **Cell-level comparison** with four statuses: OK, DIFF, MISSING_IN_OUTPUT, CALCULATED_OR_EXTRA
- **Report generation**:
  - Detailed cell-by-cell CSV report
  - Colored Excel report (Details + Summary sheets)
  - Batch global summary (CSV + Excel)


## 3. End-to-End Pipeline

```
1. Convert XBRL submission to CSV
     - XBRL-XML: Xml2Csv.jar conversion
     - XBRL-CSV ZIP: direct extraction
     - OneGate XML: native parsing
2. Load taxonomy mapping (datapoint → table/row/column)
3. Read input data (CSV or Excel)
4. Auto-resolve table code aliases
5. Compare cell by cell (input vs. output)
6. Generate reports (CSV + Excel)
```


## 4. MCONTRAN Semantic Mapping (OneGate XML)

### Problem

MCONTRAN submissions use OneGate XML format instead of standard XBRL. The XML contains `<Data>` / `<Item>` / `<Dim>` elements with property-based values, not the standard `datapoint/factValue` CSV structure.

### Implemented Solution

- Parse OneGate XML using `xml.etree.ElementTree`
- Map CSV input columns to XML `<Dim prop>` values using a hardcoded column-to-property mapping (e.g. `frbdfrubaGuichet` → `ID_GUI`)
- Apply BDFRATE6 formatting for rate fields (remove decimal dot)
- Match input rows to XML items using greedy best-match on property values
- Generate a dedicated MCONTRAN comparison report

### Result

- Works for all MCONTRAN test cases (INSTMCO1, INSTMCO2, INSTMCO4)
- No impact on standard XBRL-based regulations


## 5. Requirements

### Runtime

- Windows (current validated environment)
- Python **3.10+**

### Python packages

```
openpyxl>=3.1.0
```

Install using:

```bash
pip install -r requirements.txt
```

### External dependencies

- **Java runtime** (required by `Xml2Csv.jar` for XBRL-XML conversion)
- **Xml2Csv.jar** — XBRL-to-CSV converter (provided separately)
- **Taxonomy package** — `xbrl_package.tar.gz` for the target regulation


## 6. How to Run

### Interactive wizard

```bash
python launch_validation.py
```

Step-by-step file browser dialogs guide through: regulation selection, input data, taxonomy, XBRL submission, converter, and output directory.

### Single case (CLI)

```bash
python validate_regulation_unified.py \
  --input-dir     <csv_input_folder> \
  --xbrl-file     <xbrl_submission_file> \
  --taxonomy      <xbrl_package.tar.gz> \
  --converter-jar <path/to/Xml2Csv.jar> \
  --output-dir    <output_folder>
```

### Batch mode (CLI)

```bash
python validate_regulation_unified.py --batch \
  --load-dir      <input/load_folder> \
  --expected-dir  <expected_folder> \
  --taxonomy      <xbrl_package.tar.gz> \
  --converter-jar <path/to/Xml2Csv.jar> \
  --output-dir    <output_folder>
```

### Batch mode with Excel input

```bash
python validate_regulation_unified.py --batch --input-mode excel \
  --load-dir      <input/load_folder> \
  --expected-dir  <expected_folder> \
  --taxonomy      <xbrl_package.tar.gz> \
  --converter-jar <path/to/Xml2Csv.jar> \
  --output-dir    <output_folder>
```


## 7. Input/Output Structure

### Typical input

```
e2e-test-data/
  input/
    load/
      <CASE_ID>/
        csv_direct/
          *.csv                    # tableCode;rowId;columnId;value
  expected/
    <CASE_ID>_*_submission.xbrl    # or .zip or .xml
  xbrl/
    xbrl_package.tar.gz            # taxonomy archive
```

### Output structure

```
<output-dir>/
  <regulation>/
    batch_<timestamp>/
      global_summary.csv
      global_summary.xlsx
      <CASE_ID>/
        <timestamp>/
          taxonomy_cell_report.csv
          taxonomy_cell_report.xlsx
          converted_submission_csv/
            *.csv
```


## 8. Output Report Format

The generated Excel contains two sheets:

### Sheet "Details"

Each line represents one compared cell. Main columns:

| Column | Description |
|--------|-------------|
| `input_line` | Line number in the input CSV file |
| `input_file` | Source input CSV file name |
| `input_table_code` | Original table code from input |
| `normalized_table_code` | Resolved table code (after auto-alias) |
| `row_id_r` | Row identifier (e.g. `r0010`) |
| `column_id_c` | Column identifier (e.g. `c0020`) |
| `input_value` | Raw value from input |
| `input_value_norm` | Normalized input value |
| `output_datapoint` | Taxonomy datapoint identifier |
| `output_source_file` | CSV file from converted submission |
| `taxonomy_cell_code` | Cell code from taxonomy (e.g. `{C_01.00, R0010, C0010}`) |
| `output_value` | Raw value from submission output |
| `output_value_norm` | Normalized output value |
| `status` | Comparison result: OK / DIFF / MISSING_IN_OUTPUT / CALCULATED_OR_EXTRA |
| `details` | Human-readable explanation |

### Sheet "Summary"

| Metric | Count |
|--------|-------|
| OK | Values match |
| DIFF | Values differ |
| MISSING_IN_OUTPUT | Input cell absent in submission |
| CALCULATED_OR_EXTRA | Submission cell absent in input |
| TOTAL | Total comparisons |

### Color coding

- 🟢 Green: OK
- 🔴 Red: DIFF
- 🟡 Yellow: MISSING_IN_OUTPUT
- 🔵 Blue: CALCULATED_OR_EXTRA

### Notes

- If a table exists only in input or only in output, rows are still produced with the missing side left empty.
- Auto-alias resolution is logged in console output (e.g. `RB.26.01.11 -> RB.26.01.01`).
- In batch mode, the global summary classifies each case as **PASS** (no DIFF and no MISSING) or **ISSUES**.


## 9. Supported Regulations

| Regulation | Domain | Submission Format | Test Cases |
|------------|--------|-------------------|------------|
| RUBA       | BDF    | XBRL-XML (.xbrl)  | RUBA1, RUBA2 |
| COREP      | EBA    | XBRL-XML (.xbrl)  | COREP1, COREPBE1, COREPFR1, COREPLU1 |
| RESOL      | SRB    | XBRL-XML (.xbrl)  | RP1, RPBE1, RPFR1, RPFR2, RPLU1 |
| FP         | EBA    | XBRL-XML (.xbrl)  | FP1, FPBE1, FPFR1, FPFR2, FPLU1 |
| DBM        | BNB    | XBRL-XML (.xbrl)  | DBM1, DBMBE1, DBMFR1, DBMFR2, DBMLU1 |
| P3DH       | EBA    | XBRL-CSV ZIP (.zip) | P3DHEU1, P3DHEU2 |
| REM        | EBA    | XBRL-CSV ZIP (.zip) | P3DHREMEU1 |
| MCONTRAN   | BDF    | OneGate XML (.xml) | INSTMCO1, INSTMCO2, INSTMCO4 |
| RUBAJ      | BDF    | XBRL-XML (.xbrl)  | RUBA1, RUBA2 |


## 10. Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| XBRL conversion fails | Java not installed or `Xml2Csv.jar` path incorrect | Ensure Java is in PATH and JAR path is valid |
| No datapoints mapped (0 datapoints) | Taxonomy package does not match regulation version | Use the correct `xbrl_package.tar.gz` for the target regulation |
| All cells show MISSING_IN_OUTPUT | Input CSV format incorrect | Verify `;` separator and columns `tableCode;rowId;columnId;value` |
| Table code mismatches not resolved | Input codes too different from taxonomy codes | Check auto-alias log; input codes should follow the same naming pattern |
| MCONTRAN comparison shows unexpected DIFF | Rate formatting mismatch | BDFRATE6 fields should have no decimal dot (e.g. `065942` not `06.5942`) |
| Batch mode detects 0 test cases | Input folder structure does not match expected pattern | Verify `load/<CASE_ID>/csv_direct/` structure and `expected/*_submission.*` naming |


## 11. Excel-to-CSV Conversion

### Background

The initial version of the tool relied on an external Excel-to-CSV converter, which required specifying its path separately. While this approach worked, it introduced some limitations in terms of integration and portability: the converter had to be installed and configured independently, and its output format did not always align perfectly with the expected input structure.

To address these points, a new built-in converter (`excel_to_csv.py`) has been developed and integrated directly into the tool. This brings several improvements:

- **No external dependency**: the conversion is now handled natively — no need to provide a path to an external tool.
- **Seamless integration**: the converter is called automatically when Excel input is detected (via `--input-mode excel` or the interactive wizard).
- **Portable and reusable**: `excel_to_csv.py` can be used as a standalone utility for any project that needs to convert regulatory Excel files to the standard `tableCode;rowId;columnId;value` CSV format.

The previous converter remains fully functional and can still be used if preferred. The new built-in converter simply offers a more convenient alternative that is ready to use out of the box.


## 12. Validation Notes (April 2026)

Latest validation results after re-running all regulations with the new Excel-to-CSV converter. For each regulation, csv_direct and Excel input modes produce **strictly identical results**, confirming that the new converter works correctly.

### Fully passing

| Regulation | Status | Details |
|------------|--------|---------|
| COREP / corepfrtb | **PASS** | 8 tests — OK=118, DIFF=0, MISSING=0 |
| P3DH / P3DHEU2 | **PASS** | 2 tests — OK=20, DIFF=0, MISSING=0 |

### Minor differences (pre-existing, not related to the converter)

| Regulation | OK | DIFF | MISSING | Notes |
|------------|-----|------|---------|-------|
| P3DH / P3DHEU1 | 19 | 1 | 0 | 1 value mismatch |
| DBM | 284 | 4 | 0 | FR2: OK=280, DIFF=2 |
| COREP / coreplr | 90 | 10 | 131 | |
| RUBA / rubaj | 182 | 22 | 81 | |
| RUBA / rubad | ~2000 | ~600 | ~1800 | |
| FP | 108 | 690 | 172 | |
| COREP / corepp3dh | 3264 | 91 | 9455 | |
| RP / resol1 | 15 | 2 | 1778 | |
| REM / p3dhrem | 284 | 0 | 206 | |

### Errors (taxonomy issue)

These regulations fail at the XBRL-to-CSV conversion step due to missing files in the taxonomy package (`xbrl_package.tar.gz`). This is caused by outdated test repositories that were not pulled before running the tests. To fix this, simply pull the latest versions of the corresponding regulatory output repositories and re-run the validation.

| Regulation | Error | Fix |
|------------|-------|-----|
| COREP / coreple | `corep_le.xsd` not found in the taxonomy archive | Pull latest `pure-regulatoryoutput-euebacorep-euebale` |
| COREP / corepsolv | `corep_of.xsd` not found in the taxonomy archive | Pull latest `pure-regulatoryoutput-euebacorep-euebasolv` |

### Skipped

| Regulation | Reason | Fix |
|------------|--------|-----|
| RUBA / mcontran | No `xbrl_package.tar.gz` available in the test repository | Pull latest `pure-regulatoryoutput-frbdfruba-bdfmcontran` |

### Key takeaway

All differences between input and submission are **identical** whether using csv_direct or Excel input, which confirms that the new Excel-to-CSV converter produces correct output across all tested regulations.


## 13. Maintainer

**OUMAS Oumaima**
PURE Parametrization – SBS
2026
