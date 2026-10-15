# =============================================================================
# ESPF to RDS Conversion Tool
# =============================================================================
# Converts ESPF fixed-width files (.sXX) into RDS-compliant CSV files.
# Supports two dataset types: CATR (frbdfruba_catr) and ENGAG (frbdfruba_engag).
#
# Each line in an ESPF file is a fixed-width record where fields are extracted
# by byte position and length. The tool reads mapping rules from Excel files
# to translate ESPF zone names into RDS column names, applies value formatting
# (decimal scaling, sign handling, leading-zero stripping), and writes the
# result as CSV with a predefined column order.
# =============================================================================

from __future__ import annotations

import argparse
import csv
import math
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook


@dataclass(frozen=True)
class FieldSpec:
    """Describes one field (zone) in the ESPF fixed-width format.

    Attributes:
        zone: The ESPF zone identifier (e.g. "MONTCONTR", "AGENTECO").
        position: 1-based byte offset where the field starts in the line.
        length: Number of characters the field occupies.
    """
    zone: str
    position: int  # 1-based (ZES_POS)
    length: int    # ZES_LON


# -------------------------------------------------------------------------
# Constants
# -------------------------------------------------------------------------

# Maps internal type names to RDS dataset codes used in the output CSV.
TARGETS = {
    "CATR": "frbdfruba_catr",
    "ENGAG": "frbdfruba_engag",
}

# Reverse lookup: RDS dataset code -> internal type name.
DATASET_TO_TYPE = {v: k for k, v in TARGETS.items()}

# The first 7 columns of every output CSV, always present in this order.
BASE_COLUMNS = [
    "datasetCode",
    "reportingSessionPeriodicity",
    "reportingSessionDate",
    "institutionCode",
    "datasourceCode",
    "identifier",
    "commonDatasetRecordId",
]

# Common header zones (ESPF fixed format, shared across all ESPF types).
# Positions and lengths defined by the ESTD system standard.
COMMON_HEADER_ZONES: Dict[str, FieldSpec] = {
    "CDSESS": FieldSpec("CDSESS", 1, 2),
    "CDSCTE": FieldSpec("CDSCTE", 3, 5),
    "CDSSSO": FieldSpec("CDSSSO", 8, 5),
    "CDFILI": FieldSpec("CDFILI", 13, 5),
    "VLRFCE": FieldSpec("VLRFCE", 18, 60),
    "NOCHRO": FieldSpec("NOCHRO", 78, 7),
    "DTSITU": FieldSpec("DTSITU", 94, 8),
    "TYESTD": FieldSpec("TYESTD", 102, 1),
    "CDETAT": FieldSpec("CDETAT", 103, 1),
    "CDHABI": FieldSpec("CDHABI", 104, 10),
    "CDDEVI": FieldSpec("CDDEVI", 121, 5),
    "NOCPCI": FieldSpec("NOCPCI", 126, 15),
    "ARGUCM": FieldSpec("ARGUCM", 141, 15),
    "FMESTD": FieldSpec("FMESTD", 163, 5),
    "CODEIEM": FieldSpec("CODEIEM", 168, 1),
    "ZONETRACE": FieldSpec("ZONETRACE", 169, 60),
}

# Common header zones → RDS output field names.
# These zones are not in the mapping file (target=None) but their values
# must appear in the output under the corresponding RDS field names.
COMMON_HEADER_TARGETS: Dict[str, str] = {
    "DTSITU": "reportingSessionDate",
    "CDFILI": "datasourceCode",
    "CDDEVI": "currency",
    "NOCPCI": "internalAccount",
    "VLRFCE": "referenceValue",
}

# Decimal scaling for signed numeric ESPF zones.
DECIMAL_SCALES: Dict[str, int] = {
    "MONTCONTR": 100,
    "TESE1": 1_000_000, "TESE2": 1_000_000, "TESE3": 1_000_000,
    "TESE4": 1_000_000, "TESE5": 1_000_000, "TESE6": 1_000_000,
    "TESE7": 1_000_000, "TESE8": 1_000_000, "TESE9": 1_000_000,
    "SPREAD": 10_000,
    "MTENCOURS": 1, "MTDUREE1": 1, "MTDUREE2": 1, "MTDUREE3": 1, "MTDUREE4": 1,
}

# Zones where blank/spaces → "0".
AMOUNT_ZONES = {"MTENCOURS", "MTDUREE1", "MTDUREE2", "MTDUREE3", "MTDUREE4"}

# Expected column order (after BASE_COLUMNS) for each type.
# This defines which columns appear in the output CSV and in what order.
# CATR has 18 columns after the 7 base columns (25 total).
# ENGAG has 26 columns after the 7 base columns (33 total).
EXPECTED_COLUMNS: Dict[str, List[str]] = {
    "CATR": [
        "frbdfrubaContractAmount",
        "frbdfrubaNatureOfthetermAccount",
        "frbdfrubaMaturityRange",
        "frbdfrubaReferenceMarketIndex",
        "frbdfrubaInterestRateSlice1",
        "frbdfrubaInterestRateSlice2",
        "frbdfrubaInterestRateSlice3",
        "frbdfrubaInterestRateSlice4",
        "frbdfrubaInterestRateSlice5",
        "frbdfrubaInterestRateSlice6",
        "frbdfrubaInterestRateSlice7",
        "frbdfrubaInterestRateSlice8",
        "frbdfrubaInterestRateSlice9",
        "frbdfrubaSpread",
        "frbdfrubaCounterpartInstitutionalSector",
        "frbdfrubaDurin",
        "currency",
        "originalMaturityInMonths",
    ],
    "ENGAG": [
        "frbdfrubaCountryOfimmediateRiskOrCounterparty",
        "frbdfrubaCountryOfResidenceOfUltimateRisk",
        "frbdfrubaResidenceOfaBranchOrSubsidiaryAbroad",
        "frbdfrubaCounterpartInstitutionalSector",
        "frbdfrubaEconomicCodeOfUltimateRisk",
        "frbdfrubaPortfolioType",
        "frbdfrubaInstrumentType",
        "frbdfrubaTypeOfParticipation",
        "frbdfrubaTypeOfSecurity",
        "frbdfrubaTypeOfGuarantee",
        "frbdfrubaTypeOfEngagement",
        "frbdfrubaInitialDurationOfCommitmentLessthanOrGreaterthanOneyear",
        "frbdfrubaCurrenciesfamilyimmediaterisk",
        "frbdfrubaCurrenciesfamilyultimaterisk",
        "frbdfrubaLocalCurrencyIndicatorCountryOfImmediateRisk",
        "frbdfrubaLocalCurrencyIndicatorCountryOfUltimateRisk",
        "frbdfrubaAmountOfOutstandingTransactionAndCoverage",
        "frbdfrubaAmountRemainingTermEqualOrLessThanOneYear",
        "frbdfrubaAmountRemainingTermBetweenOneAndTwoYear",
        "frbdfrubaAmountRemainingTermMoreThanTwoYear",
        "frbdfrubaAmountOfResidualDurationNotAllocated",
        "internalAccount",
        "amountDirection",
        "currency",
        "tymnt",
        "referenceValue",
    ],
}


# -------------------------------------------------------------------------
# Utility functions
# -------------------------------------------------------------------------

def normalize_type(value: str) -> str:
    """Normalize an ESPF type identifier to uppercase (e.g. 'catr' -> 'CATR')."""
    return value.strip().upper()


def slice_value(line: str, position: int, length: int) -> str:
    """Extract a field from a fixed-width line by position (1-based) and length.

    Returns the trimmed value, or an empty string if the position exceeds
    the line length.
    """
    start = max(0, position - 1)  # Convert 1-based to 0-based index
    end = start + max(0, length)
    if start >= len(line):
        return ""
    return line[start:end].strip()


def format_value(raw: str, zone_name: str) -> str:
    """Format a raw ESPF value based on the zone type."""
    raw = raw.strip()
    if not raw:
        return "0" if zone_name in AMOUNT_ZONES else ""

    # Signed numeric value: starts with '+' or '-' and has a defined scale factor.
    # The integer part is divided by the scale to produce a decimal.
    # Example: MONTCONTR "+0201514717" / 100 = 2015147.17
    if raw[0] in ("+", "-") and zone_name in DECIMAL_SCALES:
        sign = -1 if raw[0] == "-" else 1
        digits = raw[1:]
        try:
            int_val = int(digits) * sign
        except ValueError:
            return raw  # Not a valid number, return as-is
        scale = DECIMAL_SCALES[zone_name]
        if scale == 1:
            return str(int_val)  # No decimal scaling needed
        dp = int(math.log10(scale))  # Number of decimal places from scale
        formatted = f"{int_val / scale:.{dp}f}"
        # Remove trailing zeros after decimal point (e.g. "7.050000" -> "7.05")
        if "." in formatted:
            formatted = formatted.rstrip("0").rstrip(".")
        return formatted

    # Numeric-looking values: strip leading zeros (e.g. "07" -> "7", "001" -> "1")
    if raw.isdigit() and len(raw) > 1:
        return str(int(raw))

    return raw


# -------------------------------------------------------------------------
# Excel reading functions
# -------------------------------------------------------------------------

def read_mapping(mapping_file: Path) -> Dict[str, List[Tuple[str, str]]]:
    """Read the zone-to-RDS-column mapping from the correspondence Excel file.

    The mapping file has two sheets:
      - "catr"  -> CATR mappings
      - "Engag" -> ENGAG mappings

    Each sheet has columns: A = ESPF zone name, C = RDS target column name.
    Returns a dict with keys "CATR" and "ENGAG", each containing an ordered
    list of (zone_name, rds_column_name) tuples.
    """
    wb = load_workbook(mapping_file, read_only=True, data_only=True)
    try:
        result: Dict[str, List[Tuple[str, str]]] = {}
        for sheet_name, type_name in (("catr", "CATR"), ("Engag", "ENGAG")):
            ws = wb[sheet_name]
            pairs: List[Tuple[str, str]] = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                report = row[0]                          # Column A: ESPF zone name
                sbprr = row[2] if len(row) > 2 else None  # Column C: RDS target name
                if not report or not sbprr:
                    continue
                pairs.append((str(report).strip(), str(sbprr).strip()))
            result[type_name] = pairs
        return result
    finally:
        wb.close()


def read_column_order(mapping_file: Path) -> Dict[str, List[str]]:
    """Read output column order from the 'output_order' sheet in the mapping file.

    The sheet has columns: type, position, column_name.
    Returns a dict {type_name: [col1, col2, ...]} in the defined order.
    """
    wb = load_workbook(mapping_file, read_only=True, data_only=True)
    try:
        if "output_order" not in wb.sheetnames:
            return {}
        ws = wb["output_order"]
        result: Dict[str, List[Tuple[int, str]]] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            type_name = row[0] if len(row) > 0 else None
            position = row[1] if len(row) > 1 else None
            col_name = row[2] if len(row) > 2 else None
            if not type_name or not col_name or position is None:
                continue
            type_name = str(type_name).strip()
            if type_name not in result:
                result[type_name] = []
            result[type_name].append((int(position), str(col_name).strip()))
        # Sort by position and return column names only
        return {t: [col for _, col in sorted(pairs)] for t, pairs in result.items()}
    finally:
        wb.close()


def read_espf_specs(espf_structure_file: Path) -> Dict[str, List[FieldSpec]]:
    """Read the ESPF structure (field positions and lengths) from the Spins sheet.

    The Spins sheet in RUBA_BDF_FR.xlsx defines where each variable zone is
    located within the fixed-width ESPF line. Zones are grouped by type
    (CATR or ENGAG) indicated by the Spin code in column B.

    Returns a dict {"CATR": [FieldSpec, ...], "ENGAG": [FieldSpec, ...]}.
    """
    wb = load_workbook(espf_structure_file, read_only=True, data_only=True)
    try:
        ws = wb["Spins"]
        by_type: Dict[str, List[FieldSpec]] = {"CATR": [], "ENGAG": []}
        current_type: Optional[str] = None

        # Sheet columns (1-based): B=Spin code, D=Zone code, F=Position, H=Length.
        for row in ws.iter_rows(min_row=4, values_only=True):
            spin_code = row[1] if len(row) > 1 else None
            zone = row[3] if len(row) > 3 else None
            position = row[5] if len(row) > 5 else None
            length = row[7] if len(row) > 7 else None

            if spin_code:
                normalized = normalize_type(str(spin_code))
                current_type = normalized if normalized in by_type else None

            if current_type not in by_type:
                continue
            if not zone or position is None or length is None:
                continue

            try:
                pos_int = int(position)
                len_int = int(length)
            except (TypeError, ValueError):
                continue

            zone_name = str(zone).strip()
            if not zone_name:
                continue

            by_type[current_type].append(FieldSpec(
                zone=zone_name,
                position=pos_int,
                length=len_int,
            ))

        return by_type
    finally:
        wb.close()


def read_rds_output_columns(rds_structure_file: Path) -> Dict[str, List[str]]:
    """Read all RDS column names from the structuresTransformations sheet.

    This reads the full list of RDS columns defined for each dataset type.
    Note: the output column order is now determined by EXPECTED_COLUMNS,
    not by this function. This is still called for compatibility.
    """
    wb = load_workbook(rds_structure_file, read_only=True, data_only=True)
    try:
        ws = wb["structuresTransformations"]
        result: Dict[str, List[str]] = {"CATR": [], "ENGAG": []}
        current_dataset: Optional[str] = None

        for row in ws.iter_rows(min_row=2, values_only=True):
            dataset_code = row[0] if len(row) > 0 else None
            dataset_field = row[5] if len(row) > 5 else None

            if dataset_code:
                current_dataset = str(dataset_code).strip()

            if not current_dataset or not dataset_field:
                continue

            type_name = DATASET_TO_TYPE.get(current_dataset)
            if type_name is None:
                continue

            field_name = str(dataset_field).strip()
            if field_name and field_name not in result[type_name]:
                result[type_name].append(field_name)

        return result
    finally:
        wb.close()


# -------------------------------------------------------------------------
# Column ordering and conversion
# -------------------------------------------------------------------------

def build_output_columns(
    type_name: str,
    mapping_pairs: List[Tuple[str, str]],
    rds_columns: Optional[List[str]] = None,
) -> List[str]:
    """Build the ordered list of columns for the output CSV.

    Starts with the 7 BASE_COLUMNS, then appends the expected columns
    for the given type (CATR or ENGAG) as defined in EXPECTED_COLUMNS.
    Duplicates are skipped (e.g. 'currency' already in BASE won't repeat).
    """
    columns = list(BASE_COLUMNS)
    seen = set(columns)

    # Append type-specific columns in the expected order
    for col in EXPECTED_COLUMNS.get(type_name, []):
        if col not in seen:
            columns.append(col)
            seen.add(col)

    return columns


def _fmestd_min_length() -> int:
    """Minimum line length required to read the FMESTD field."""
    spec = COMMON_HEADER_ZONES["FMESTD"]
    return (spec.position - 1) + spec.length


def convert_file(
    input_sxx: Path,
    mapping: Dict[str, List[Tuple[str, str]]],
    specs: Dict[str, List[FieldSpec]],
    rds_columns_by_type: Dict[str, List[str]],
    output_dir: Path,
    periodicity: str,
    institution_code: str = "",
) -> Dict[str, Path]:
    """Main conversion: reads the ESPF file and writes RDS CSV files.

    For each line in the input:
      1. Extract the FMESTD field to determine the record type (CATR/ENGAG).
      2. Extract common header fields (date, datasource, currency, etc.).
      3. Extract variable zone fields using the mapping (zone -> RDS column).
      4. Format values (decimal scaling, sign handling, leading zeros).
      5. Collect all records grouped by type.

    After reading all lines, write one CSV per type with the expected column
    order. Columns without a defined mapping are filled with '<NoMapping>'.

    Returns a dict mapping type name to the output file path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # Build zone -> FieldSpec lookup per type for fast access during parsing
    spec_lookup: Dict[str, Dict[str, FieldSpec]] = {}
    for type_name, field_list in specs.items():
        spec_lookup[type_name] = {fs.zone: fs for fs in field_list}

    # Accumulate parsed records per type
    records: Dict[str, List[Dict[str, str]]] = {"CATR": [], "ENGAG": []}
    counters = {"CATR": 0, "ENGAG": 0}  # Incremental identifier per type

    # FMESTD zone spec — used to determine record type from each line
    fmestd_spec = COMMON_HEADER_ZONES["FMESTD"]

    # --- Phase 1: Parse all lines from the ESPF input file ---
    with input_sxx.open("r", encoding="latin-1") as f:
        for raw_line in f:
            line = raw_line.rstrip("\n")
            if not line.strip():
                continue  # Skip blank lines

            # Determine the record type from the FMESTD field (pos 163, len 5)
            fmestd = normalize_type(slice_value(line, fmestd_spec.position, fmestd_spec.length))
            if fmestd not in records:
                continue  # Skip lines that are neither CATR nor ENGAG

            type_name = fmestd
            dataset_code = TARGETS[type_name]
            counters[type_name] += 1
            identifier = str(counters[type_name])  # Sequential row number

            # Initialize the row with base field values
            row: Dict[str, str] = {
                "datasetCode": dataset_code,
                "reportingSessionPeriodicity": periodicity,
                "institutionCode": institution_code,
                "identifier": identifier,
                "commonDatasetRecordId": "",
            }

            # Extract common header fields (date, datasource, currency, etc.)
            # from their fixed positions in the ESPF line
            for zone_name, target_field in COMMON_HEADER_TARGETS.items():
                spec = COMMON_HEADER_ZONES[zone_name]
                row[target_field] = slice_value(line, spec.position, spec.length)

            # Extract variable zone fields using the mapping rules.
            # For each (zone, target) pair, look up the zone's position/length
            # in the spec, extract the raw value, and format it.
            zone_specs = spec_lookup[type_name]
            for zone, target in mapping[type_name]:
                field_spec = zone_specs.get(zone)
                if not field_spec:
                    continue  # Zone not defined in Spins — skip
                raw = slice_value(line, field_spec.position, field_spec.length)
                row[target] = format_value(raw, zone)

            # Ensure all base columns exist in the row (default to empty)
            for base_col in BASE_COLUMNS:
                row.setdefault(base_col, "")

            row["datasetCode"] = dataset_code
            if not row.get("reportingSessionPeriodicity"):
                row["reportingSessionPeriodicity"] = periodicity

            records[type_name].append(row)

    # --- Phase 2: Write CSV output files ---
    outputs: Dict[str, Path] = {}
    for type_name in ("CATR", "ENGAG"):
        if not records[type_name]:
            continue  # No records of this type found

        # Determine the output column list in the expected order
        columns = build_output_columns(
            type_name=type_name,
            mapping_pairs=mapping[type_name],
            rds_columns=rds_columns_by_type.get(type_name),
        )

        dataset_code = TARGETS[type_name]
        out_path = output_dir / f"{dataset_code}.csv"

        # Build the set of columns that have actual data (from base fields,
        # common header extraction, or the zone-to-column mapping).
        # Any column NOT in this set will be marked with "<NoMapping>".
        mapped_cols = set(BASE_COLUMNS)
        mapped_cols.update(COMMON_HEADER_TARGETS.values())
        for _, target in mapping[type_name]:
            mapped_cols.add(target)

        # Write the CSV: header row + one data row per ESPF record
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
            writer.writeheader()
            for row in records[type_name]:
                # For mapped columns: write the value (or empty if missing).
                # For unmapped columns: write "<NoMapping>" if no value exists.
                writer.writerow({
                    col: row.get(col, "") if col in mapped_cols else row.get(col, "") or "<NoMapping>"
                    for col in columns
                })

        outputs[type_name] = out_path

    return outputs


# -------------------------------------------------------------------------
# Analysis / diagnostic helpers
# -------------------------------------------------------------------------

def _build_spec_lookup(specs: Dict[str, List[FieldSpec]]) -> Dict[str, Dict[str, FieldSpec]]:
    """Convert the spec list to a dict keyed by zone name for fast lookup."""
    return {t: {fs.zone: fs for fs in fs_list} for t, fs_list in specs.items()}


def _mapped_zones_missing_in_specs(
    mapping: Dict[str, List[Tuple[str, str]]],
    spec_lookup: Dict[str, Dict[str, FieldSpec]],
) -> Dict[str, List[str]]:
    """Find zones referenced in the mapping but absent from the Spins structure.

    These zones cannot be extracted because their position/length is unknown.
    """
    missing: Dict[str, List[str]] = {"CATR": [], "ENGAG": []}
    for type_name in ("CATR", "ENGAG"):
        for zone, _ in mapping.get(type_name, []):
            if zone not in spec_lookup.get(type_name, {}):
                missing[type_name].append(zone)
    return missing


def _collect_line_stats(
    input_sxx: Path,
    mapping: Dict[str, List[Tuple[str, str]]],
    spec_lookup: Dict[str, Dict[str, FieldSpec]],
) -> Tuple[int, int, int, Counter[str], Counter[str], Counter[int], Dict[str, Counter[str]]]:
    type_counts: Counter[str] = Counter()
    unknown_type_counts: Counter[str] = Counter()
    line_length_counts: Counter[int] = Counter()
    non_empty_by_type_zone: Dict[str, Counter[str]] = {
        "CATR": Counter(),
        "ENGAG": Counter(),
    }

    total_lines = 0
    blank_lines = 0
    short_for_type = 0

    min_len = _fmestd_min_length()
    fmestd_spec = COMMON_HEADER_ZONES["FMESTD"]

    with input_sxx.open("r", encoding="latin-1") as f:
        for raw_line in f:
            total_lines += 1
            line = raw_line.rstrip("\n")
            if not line.strip():
                blank_lines += 1
                continue

            line_length_counts[len(line)] += 1

            if len(line) < min_len:
                short_for_type += 1
                continue

            fmestd = normalize_type(slice_value(line, fmestd_spec.position, fmestd_spec.length))
            if fmestd in ("CATR", "ENGAG"):
                type_counts[fmestd] += 1
                for zone, _ in mapping.get(fmestd, []):
                    fs = spec_lookup[fmestd].get(zone)
                    if not fs:
                        continue
                    value = slice_value(line, fs.position, fs.length)
                    if value:
                        non_empty_by_type_zone[fmestd][zone] += 1
            else:
                unknown_type_counts[fmestd or "<empty>"] += 1

    return (
        total_lines,
        blank_lines,
        short_for_type,
        type_counts,
        unknown_type_counts,
        line_length_counts,
        non_empty_by_type_zone,
    )


def _compute_zone_fill_rates(
    mapping: Dict[str, List[Tuple[str, str]]],
    type_counts: Counter[str],
    non_empty_by_type_zone: Dict[str, Counter[str]],
) -> Dict[str, List[Tuple[str, int, float]]]:
    """Compute the fill rate (% of lines with non-empty data) for each mapped zone."""
    zone_fill_rates: Dict[str, List[Tuple[str, int, float]]] = {"CATR": [], "ENGAG": []}
    for type_name in ("CATR", "ENGAG"):
        typed_total = type_counts[type_name]
        for zone, _ in mapping.get(type_name, []):
            present = non_empty_by_type_zone[type_name][zone]
            rate = (present / typed_total) if typed_total else 0.0
            zone_fill_rates[type_name].append((zone, present, rate))
    return zone_fill_rates


def analyze_espf_input(
    input_sxx: Path,
    mapping: Dict[str, List[Tuple[str, str]]],
    specs: Dict[str, List[FieldSpec]],
) -> Dict[str, object]:
    """Analyze a raw ESPF fixed-width file and report structure health."""
    spec_lookup = _build_spec_lookup(specs)
    missing_spec_zones = _mapped_zones_missing_in_specs(mapping, spec_lookup)
    (
        total_lines,
        blank_lines,
        short_for_type,
        type_counts,
        unknown_type_counts,
        line_length_counts,
        non_empty_by_type_zone,
    ) = _collect_line_stats(input_sxx, mapping, spec_lookup)

    dominant_lengths = line_length_counts.most_common(5)
    zone_fill_rates = _compute_zone_fill_rates(mapping, type_counts, non_empty_by_type_zone)

    return {
        "total_lines": total_lines,
        "blank_lines": blank_lines,
        "short_for_type": short_for_type,
        "type_counts": dict(type_counts),
        "unknown_type_counts": dict(unknown_type_counts),
        "dominant_lengths": dominant_lengths,
        "missing_spec_zones": missing_spec_zones,
        "zone_fill_rates": zone_fill_rates,
    }


def print_analysis_report(report: Dict[str, object]) -> None:
    print("ESPF input analysis")
    print(f"  total lines: {report['total_lines']}")
    print(f"  blank lines: {report['blank_lines']}")
    print(f"  too short for FMESTD detection (<167): {report['short_for_type']}")

    type_counts = report["type_counts"]
    print("  recognized types:")
    print(f"    CATR: {type_counts.get('CATR', 0)}")
    print(f"    ENGAG: {type_counts.get('ENGAG', 0)}")

    unknown_type_counts = report["unknown_type_counts"]
    if unknown_type_counts:
        print("  unknown FMESTD values:")
        for key, value in sorted(unknown_type_counts.items()):
            print(f"    {key}: {value}")
    else:
        print("  unknown FMESTD values: none")

    dominant_lengths = report["dominant_lengths"]
    if dominant_lengths:
        print("  dominant line lengths (top 5):")
        for length, count in dominant_lengths:
            print(f"    len={length}: {count}")

    missing_spec_zones = report["missing_spec_zones"]
    for type_name in ("CATR", "ENGAG"):
        missing = missing_spec_zones[type_name]
        if missing:
            joined = ", ".join(sorted(set(missing)))
            print(f"  {type_name} zones mapped but missing in Spins: {joined}")
        else:
            print(f"  {type_name} zones mapped but missing in Spins: none")

    zone_fill_rates = report["zone_fill_rates"]
    for type_name in ("CATR", "ENGAG"):
        rows = zone_fill_rates[type_name]
        if not rows:
            continue
        print(f"  {type_name} mapped zones non-empty rate (sample):")
        for zone, present, rate in sorted(rows, key=lambda x: x[2], reverse=True)[:8]:
            print(f"    {zone}: {present} ({rate:.1%})")


# -------------------------------------------------------------------------
# CLI argument parsing
# -------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert ESPF fixed-width input to RDS CSV for CATR and ENGAG.")
    parser.add_argument("--input-sxx", help="Path to ESPF fixed-width input file")
    parser.add_argument(
        "--mapping-file",
        default="2025-11-21_Correspondance_entrees_specifiques CATR & creditContracts_NW.xlsx",
        help="Path to mapping workbook",
    )
    parser.add_argument(
        "--espf-structure-file",
        default="RUBA_BDF_FR.xlsx",
        help="Path to ESPF structure workbook",
    )
    parser.add_argument(
        "--rds-structure-file",
        default="pure-transformedDataStructure-frbdfruba.xlsm",
        help="Path to RDS structure workbook",
    )
    parser.add_argument("--output-dir", default=None, help="Directory for generated CSV files")
    parser.add_argument(
        "--periodicity",
        default="monthly",
        help="Value for reportingSessionPeriodicity",
    )
    parser.add_argument(
        "--institution-code",
        default="RUBA",
        help="Value for institutionCode",
    )
    parser.add_argument(
        "--analyze",
        action="store_true",
        help="Analyze ESPF format and compatibility before conversion",
    )
    parser.add_argument(
        "--analyze-only",
        action="store_true",
        help="Only analyze ESPF format (no CSV generation)",
    )
    parser.add_argument(
        "--no-dialog",
        action="store_true",
        help="Disable file picker and require --input-sxx",
    )
    return parser.parse_args()


# -------------------------------------------------------------------------
# Interactive dialogs (tkinter)
# -------------------------------------------------------------------------

def pick_input_file() -> Optional[Path]:
    """Show a styled dialog for the user to select an ESPF input file.

    Displays a dialog with a path entry field and a Browse button that
    opens the native file picker. Returns the selected Path or None.
    """
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog
    except Exception:
        return None

    result: Dict[str, Optional[str]] = {"path": None}

    dialog = tk.Tk()
    dialog.title("ESPF to RDS Converter")
    dialog.resizable(False, False)
    dialog.configure(bg="#f0f0f0")

    style = ttk.Style(dialog)
    style.theme_use("vista")
    style.configure("OK.TButton", font=("Segoe UI", 10, "bold"), padding=(20, 6))
    style.configure("Cancel.TButton", font=("Segoe UI", 10), padding=(16, 6))
    style.configure("Browse.TButton", font=("Segoe UI", 10), padding=(10, 4))

    # Header banner
    header_frame = tk.Frame(dialog, bg="#1a3a5c")
    header_frame.pack(fill="x")
    tk.Label(
        header_frame, text="ESPF \u2192 RDS Conversion",
        font=("Segoe UI", 15, "bold"), fg="white", bg="#1a3a5c", anchor="w",
    ).pack(fill="x", padx=20, pady=(14, 0))
    tk.Label(
        header_frame, text="Step 1 \u2014 Select input file",
        font=("Segoe UI", 9), fg="#a0c0e0", bg="#1a3a5c", anchor="w",
    ).pack(fill="x", padx=20, pady=(0, 12))

    # Form area
    form_frame = ttk.Frame(dialog, padding=(24, 18, 24, 6))
    form_frame.pack(fill="x")

    ttk.Label(form_frame, text="ESPF input file (.sXX):", font=("Segoe UI", 11)).grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 6),
    )
    path_var = tk.StringVar(value="")
    path_entry = ttk.Entry(form_frame, textvariable=path_var, width=50, font=("Segoe UI", 10))
    path_entry.grid(row=1, column=0, pady=4, sticky="we")
    form_frame.columnconfigure(0, weight=1)

    def browse():
        file_path = filedialog.askopenfilename(
            title="Select ESPF input file",
            filetypes=[("ESPF files", "*.sXX *.txt"), ("All files", "*.*")],
        )
        if file_path:
            path_var.set(file_path)

    ttk.Button(form_frame, text="Browse...", command=browse, style="Browse.TButton").grid(
        row=1, column=1, padx=(8, 0), pady=4,
    )

    # Separator
    ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=20, pady=(12, 0))

    # Buttons
    btn_frame = ttk.Frame(dialog, padding=(24, 14, 24, 16))
    btn_frame.pack(fill="x")

    def on_ok(event=None):
        val = path_var.get().strip()
        if not val:
            return
        result["path"] = val
        dialog.destroy()

    def on_cancel(event=None):
        dialog.destroy()

    ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok, style="OK.TButton")
    ok_btn.pack(side="right", padx=(8, 0))
    ttk.Button(btn_frame, text="Cancel", command=on_cancel, style="Cancel.TButton").pack(side="right")

    # Disable OK when path is empty
    def _validate_path(*_args):
        if path_var.get().strip():
            ok_btn.state(["!disabled"])
        else:
            ok_btn.state(["disabled"])

    path_var.trace_add("write", _validate_path)
    _validate_path()

    dialog.bind("<Return>", on_ok)
    dialog.bind("<Escape>", on_cancel)
    path_entry.focus_set()

    # Center on screen
    dialog.update_idletasks()
    w = dialog.winfo_reqwidth()
    h = dialog.winfo_reqheight()
    x = (dialog.winfo_screenwidth() - w) // 2
    y = (dialog.winfo_screenheight() - h) // 2
    dialog.geometry(f"+{x}+{y}")

    dialog.mainloop()

    if result["path"] is None:
        return None
    return Path(result["path"])


def pick_parameters(default_institution: str = "RUBA", default_periodicity: str = "monthly") -> Tuple[str, str]:
    """Show a styled dialog for the user to enter institution code and periodicity.

    Returns a tuple (institution_code, periodicity).
    Periodicity is a combobox allowing selection from predefined values
    (monthly, quarterly, half-yearly, yearly) or free-text input.
    The OK button is disabled when the institution field is empty.
    """
    import tkinter as tk
    from tkinter import ttk

    result = {"institution": default_institution, "periodicity": default_periodicity}

    # --- Window setup ---
    dialog = tk.Tk()
    dialog.title("ESPF to RDS Converter")
    dialog.resizable(False, False)
    dialog.configure(bg="#f0f0f0")

    # Apply the Windows-native ttk theme for modern widget appearance
    style = ttk.Style(dialog)
    style.theme_use("vista")

    # Custom styles for the OK and Cancel buttons
    style.configure("OK.TButton", font=("Segoe UI", 10, "bold"), padding=(20, 6))
    style.configure("Cancel.TButton", font=("Segoe UI", 10), padding=(16, 6))

    # --- Header banner (dark blue background with white text) ---
    header_frame = tk.Frame(dialog, bg="#1a3a5c")
    header_frame.pack(fill="x")
    tk.Label(
        header_frame, text="ESPF \u2192 RDS Conversion",
        font=("Segoe UI", 15, "bold"), fg="white", bg="#1a3a5c", anchor="w",
    ).pack(fill="x", padx=20, pady=(14, 0))
    tk.Label(
        header_frame, text="Step 2 \u2014 Configure conversion parameters",
        font=("Segoe UI", 9), fg="#a0c0e0", bg="#1a3a5c", anchor="w",
    ).pack(fill="x", padx=20, pady=(0, 12))

    # --- Form area ---
    form_frame = ttk.Frame(dialog, padding=(24, 18, 24, 6))
    form_frame.pack(fill="x")

    # Institution field
    ttk.Label(form_frame, text="Institution code:", font=("Segoe UI", 11)).grid(
        row=0, column=0, padx=(0, 12), pady=10, sticky="e",
    )
    inst_var = tk.StringVar(value=default_institution)
    inst_entry = ttk.Entry(form_frame, textvariable=inst_var, width=32, font=("Segoe UI", 11))
    inst_entry.grid(row=0, column=1, pady=10, sticky="w")

    # Periodicity field
    ttk.Label(form_frame, text="Periodicity:", font=("Segoe UI", 11)).grid(
        row=1, column=0, padx=(0, 12), pady=10, sticky="e",
    )
    period_var = tk.StringVar(value=default_periodicity)
    period_choices = ["monthly", "quarterly", "half-yearly", "yearly"]
    period_combo = ttk.Combobox(
        form_frame, textvariable=period_var, values=period_choices,
        width=29, state="normal", font=("Segoe UI", 11),
    )
    period_combo.grid(row=1, column=1, pady=10, sticky="w")

    # --- Separator ---
    ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=20, pady=(8, 0))

    # --- Buttons ---
    btn_frame = ttk.Frame(dialog, padding=(24, 14, 24, 16))
    btn_frame.pack(fill="x")

    def on_ok(event=None):
        if not inst_var.get().strip():
            return  # Do nothing if institution is empty
        result["institution"] = inst_var.get().strip()
        result["periodicity"] = period_var.get().strip()
        dialog.destroy()

    def on_cancel(event=None):
        dialog.destroy()

    ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok, style="OK.TButton")
    ok_btn.pack(side="right", padx=(8, 0))
    cancel_btn = ttk.Button(btn_frame, text="Cancel", command=on_cancel, style="Cancel.TButton")
    cancel_btn.pack(side="right")

    # Disable OK button when institution field is empty
    def _validate_inst(*_args):
        if inst_var.get().strip():
            ok_btn.state(["!disabled"])
        else:
            ok_btn.state(["disabled"])

    inst_var.trace_add("write", _validate_inst)
    _validate_inst()  # Initial check

    # Keyboard bindings
    dialog.bind("<Return>", on_ok)
    dialog.bind("<Escape>", on_cancel)
    inst_entry.focus_set()

    # Center the window on screen
    dialog.update_idletasks()
    w = dialog.winfo_reqwidth()
    h = dialog.winfo_reqheight()
    x = (dialog.winfo_screenwidth() - w) // 2
    y = (dialog.winfo_screenheight() - h) // 2
    dialog.geometry(f"+{x}+{y}")

    dialog.mainloop()

    return result["institution"], result["periodicity"]


def pick_output_dir() -> Optional[Path]:
    """Show a styled dialog for the user to select the output directory.

    Displays a dialog with a path entry field and a Browse button that
    opens the native folder picker. Returns the selected Path or None.
    """
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog
    except Exception:
        return None

    result: Dict[str, Optional[str]] = {"path": None}

    dialog = tk.Tk()
    dialog.title("ESPF to RDS Converter")
    dialog.resizable(False, False)
    dialog.configure(bg="#f0f0f0")

    style = ttk.Style(dialog)
    style.theme_use("vista")
    style.configure("OK.TButton", font=("Segoe UI", 10, "bold"), padding=(20, 6))
    style.configure("Cancel.TButton", font=("Segoe UI", 10), padding=(16, 6))
    style.configure("Browse.TButton", font=("Segoe UI", 10), padding=(10, 4))

    # Header banner
    header_frame = tk.Frame(dialog, bg="#1a3a5c")
    header_frame.pack(fill="x")
    tk.Label(
        header_frame, text="ESPF \u2192 RDS Conversion",
        font=("Segoe UI", 15, "bold"), fg="white", bg="#1a3a5c", anchor="w",
    ).pack(fill="x", padx=20, pady=(14, 0))
    tk.Label(
        header_frame, text="Step 3 \u2014 Select output folder",
        font=("Segoe UI", 9), fg="#a0c0e0", bg="#1a3a5c", anchor="w",
    ).pack(fill="x", padx=20, pady=(0, 12))

    # Form area
    form_frame = ttk.Frame(dialog, padding=(24, 18, 24, 6))
    form_frame.pack(fill="x")

    ttk.Label(form_frame, text="Output directory:", font=("Segoe UI", 11)).grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 6),
    )
    path_var = tk.StringVar(value="")
    path_entry = ttk.Entry(form_frame, textvariable=path_var, width=50, font=("Segoe UI", 10))
    path_entry.grid(row=1, column=0, pady=4, sticky="we")
    form_frame.columnconfigure(0, weight=1)

    def browse():
        dir_path = filedialog.askdirectory(title="Select output directory")
        if dir_path:
            path_var.set(dir_path)

    ttk.Button(form_frame, text="Browse...", command=browse, style="Browse.TButton").grid(
        row=1, column=1, padx=(8, 0), pady=4,
    )

    # Separator
    ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=20, pady=(12, 0))

    # Buttons
    btn_frame = ttk.Frame(dialog, padding=(24, 14, 24, 16))
    btn_frame.pack(fill="x")

    def on_ok(event=None):
        val = path_var.get().strip()
        if not val:
            return
        result["path"] = val
        dialog.destroy()

    def on_cancel(event=None):
        dialog.destroy()

    ok_btn = ttk.Button(btn_frame, text="OK", command=on_ok, style="OK.TButton")
    ok_btn.pack(side="right", padx=(8, 0))
    ttk.Button(btn_frame, text="Cancel", command=on_cancel, style="Cancel.TButton").pack(side="right")

    # Disable OK when path is empty
    def _validate_path(*_args):
        if path_var.get().strip():
            ok_btn.state(["!disabled"])
        else:
            ok_btn.state(["disabled"])

    path_var.trace_add("write", _validate_path)
    _validate_path()

    dialog.bind("<Return>", on_ok)
    dialog.bind("<Escape>", on_cancel)
    path_entry.focus_set()

    # Center on screen
    dialog.update_idletasks()
    w = dialog.winfo_reqwidth()
    h = dialog.winfo_reqheight()
    x = (dialog.winfo_screenwidth() - w) // 2
    y = (dialog.winfo_screenheight() - h) // 2
    dialog.geometry(f"+{x}+{y}")

    dialog.mainloop()

    if result["path"] is None:
        return None
    return Path(result["path"])


# -------------------------------------------------------------------------
# Main entry point
# -------------------------------------------------------------------------

def main() -> None:
    """Main entry point: parse arguments, load config files, run conversion.

    Workflow:
      1. Determine input file (CLI arg or file picker dialog).
      2. Determine institution code and periodicity (dialog or CLI defaults).
      3. Determine output directory (CLI arg or folder picker dialog).
      4. Load mapping, ESPF specs, and RDS column definitions from Excel.
      5. Optionally run ESPF analysis (--analyze / --analyze-only).
      6. Run the conversion and write CSV output files.
    """
    args = parse_args()
    input_sxx: Optional[Path] = Path(args.input_sxx) if args.input_sxx else None

    # Step 1: Get input file — either from CLI or via file picker
    if input_sxx is None and not args.no_dialog:
        input_sxx = pick_input_file()

    if input_sxx is None:
        raise SystemExit("Aucun fichier ESPF fourni. Utilisez --input-sxx ou la boite de dialogue.")

    if not input_sxx.exists():
        raise SystemExit(f"Fichier ESPF introuvable: {input_sxx}")

    # Resolve paths to the three configuration Excel files
    mapping_file = Path(args.mapping_file)
    espf_structure_file = Path(args.espf_structure_file)
    rds_structure_file = Path(args.rds_structure_file)

    # Step 2: Get institution and periodicity — dialog or CLI
    if not args.no_dialog:
        institution, periodicity = pick_parameters(args.institution_code, args.periodicity)
    else:
        institution, periodicity = args.institution_code, args.periodicity

    # Step 3: Get output directory — CLI arg, dialog, or error
    if args.output_dir:
        output_dir = Path(args.output_dir)
    elif not args.no_dialog:
        picked = pick_output_dir()
        if picked is None:
            raise SystemExit("Aucun dossier de sortie selectionne.")
        output_dir = picked
    else:
        raise SystemExit("Aucun dossier de sortie fourni. Utilisez --output-dir.")

    # Step 4: Load configuration from Excel files
    mapping = read_mapping(mapping_file)                      # Zone -> RDS column mapping
    specs = read_espf_specs(espf_structure_file)               # ESPF field positions/lengths
    rds_columns_by_type = read_rds_output_columns(rds_structure_file)  # Full RDS column lists

    # Step 5: Optional analysis of the ESPF input file
    if args.analyze or args.analyze_only:
        report = analyze_espf_input(
            input_sxx=input_sxx,
            mapping=mapping,
            specs=specs,
        )
        print_analysis_report(report)

    if args.analyze_only:
        return  # Analysis only mode — skip CSV generation

    # Step 6: Run the conversion and write CSV output
    outputs = convert_file(
        input_sxx=input_sxx,
        mapping=mapping,
        specs=specs,
        rds_columns_by_type=rds_columns_by_type,
        output_dir=output_dir,
        periodicity=periodicity,
        institution_code=institution,
    )

    print("Conversion complete")
    for type_name, path in outputs.items():
        print(f"  {type_name}: {path}")


if __name__ == "__main__":
    main()
