"""
excel_to_csv.py
===============
Standalone Excel-to-CSV converter for regulatory submissions.

Usage (standalone):
    python excel_to_csv.py <input_folder_or_file> [--output <output_folder>]

Usage (as module):
    from excel_to_csv import convert_folder, convert_file
    csv_files = convert_folder(Path("excel_dir"), Path("csv_output_dir"))
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import openpyxl


# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────

def _format_cell_value(value):
    """Format cell value for CSV output.

    - Converts percentage strings ('54,78%' or '54.78%') to decimal (e.g. 0,5478).
    - Formats numeric values with dot as decimal separator.
    - Removes trailing '00:00:00' from date values.
    """
    if value is None:
        return ""

    s = str(value)

    # Handle dates (datetime objects or strings with 00:00:00)
    if "00:00:00" in s:
        return s.replace("00:00:00", "").strip()

    # Handle percentage strings
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.endswith('%'):
            try:
                num_str = stripped[:-1].replace(',', '.').replace('\xa0', '').replace(' ', '')
                num = float(num_str) / 100
                if num == int(num):
                    return str(int(num))
                return str(num)
            except ValueError:
                pass
        return value

    # Handle numeric values with dot as decimal separator
    if isinstance(value, float):
        value = round(value, 10)
        if value == int(value) and abs(value) < 1e15:
            return str(int(value))
        return str(value)

    return s


def _get_table_code(row):
    """Find table code cell in a row. Returns (column_index, table_code)."""
    for i in range(len(row)):
        if row[i] and str(row[i]).strip().lower() in ("table code", "table"):
            table_value = str(row[i + 1]).strip()
            if " - " in table_value:
                table_code = table_value.split(" - ")[0].strip()
            else:
                table_code = table_value
            return i, table_code
    return 0, ""


def _read_identification(wb):
    """Read identification sheet for institution/regulation/output/form codes."""
    inst_code = ""
    reg_code = ""
    regoutput_code = ""
    form_code = ""
    id_sheet = wb["identification"]
    for row in id_sheet.iter_rows(min_row=1, max_row=15, values_only=True):
        for i in range(len(row)):
            if row[i] and str(row[i]).strip().lower() == "institution code":
                inst_code = str(row[i + 1]).strip()
                break
            elif row[i] and str(row[i]).strip().lower() == "regulation code":
                reg_code = str(row[i + 1]).strip()
                break
            elif row[i] and str(row[i]).strip().lower() == "regulatory output code":
                regoutput_code = str(row[i + 1]).strip()
                break
            elif row[i] and str(row[i]).strip().lower() == "form code":
                form_code = str(row[i + 1]).strip()
                break
    return inst_code, reg_code, regoutput_code, form_code


def _read_table_characteristics(wb, sheetname):
    """Read table code, detect list-form, and collect z-axis values."""
    read_state = "readTableCode"
    table = wb[sheetname]
    table_code = ""
    table_column = 0
    z_axis = {}
    for row in table.iter_rows(min_row=1, values_only=True):
        match read_state:
            case "readTableCode":
                table_column, table_code = _get_table_code(row)
                if table_code:
                    read_state = "readzAxis"
            case "readzAxis":
                z_axis_code = row[table_column]
                if z_axis_code:
                    z_axis[str(z_axis_code)] = str(row[table_column + 1]).strip()
                else:
                    read_state = "readRows"
            case "readRows":
                if row[table_column]:
                    row_code = str(row[table_column])
                    is_list_form = row_code.strip() == "999"
                    return table_code, is_list_form, z_axis
    return "", False, {}


def _read_table_list(csv_path_stem, wb, sheetname, inst_code, reg_code, regoutput_code, form_code):
    """Read a list-form table and write to separate CSV files."""
    read_state = "readTableCode"
    table_code = ""
    table_column = 0
    row_start = 0
    col_nums_row = None
    z_axis = {}
    table = wb[sheetname]
    row_index = 0
    list_csv_file = None
    generated_files = []

    for row in table.iter_rows(min_row=1, values_only=True):
        match read_state:
            case "readTableCode":
                table_column, table_code = _get_table_code(row)
                if table_code:
                    read_state = "readzAxis"
            case "readzAxis":
                z_axis_code = row[table_column]
                if z_axis_code:
                    z_axis[str(z_axis_code)] = str(row[table_column + 1]).strip()
                else:
                    read_state = "readListRows"
            case "readListRows":
                if row[table_column]:
                    if row_start == 0:
                        csv_file_path = csv_path_stem + "_" + table_code + ".csv"
                        if os.path.exists(csv_file_path):
                            list_csv_file = open(csv_file_path, "a", newline="")
                        else:
                            list_csv_file = open(csv_file_path, "w", newline="")
                            list_csv_file.write(
                                "institutionCode;regulationCode;regulatoryOutputCode;formCode;tableCode"
                            )
                            for cellindex in range(table_column + 2, len(row)):
                                if col_nums_row[cellindex]:
                                    list_csv_file.write(f";c{col_nums_row[cellindex]}")
                            for z_code in z_axis:
                                code = str(z_code)
                                if "(" in code and ")" in code:
                                    code = code[code.index("(") + 1 : code.index(")")]
                                list_csv_file.write(f";{code}")
                            list_csv_file.write("\n")
                            generated_files.append(csv_file_path)
                        row_start = row_index
                    list_csv_file.write(
                        f"{inst_code};{reg_code};{regoutput_code};{form_code};{table_code}"
                    )
                    for cellindex in range(table_column + 2, len(row)):
                        cell_value = ""
                        if row[cellindex] is not None:
                            cell_value = row[cellindex]
                        row_code = str(row[table_column])
                        if row_code.strip():
                            list_csv_file.write(f";{_format_cell_value(cell_value)}")
                    if z_axis:
                        for z_code in z_axis:
                            list_csv_file.write(f";{z_axis[z_code]}")
                    list_csv_file.write("\n")
                elif row_start == 0:
                    col_nums_row = row
        row_index += 1

    if list_csv_file:
        list_csv_file.close()
    return generated_files


def _read_table(wb, sheetname, inst_code, reg_code, regoutput_code, form_code, z_axis_form, csv_file):
    """Read a regular (non-list) table and append rows to csv_file."""
    read_state = "readTableCode"
    table_code = ""
    table_column = 0
    row_start = 0
    col_nums_row = None
    z_axis = {}
    table = wb[sheetname]
    row_index = 0

    for row in table.iter_rows(min_row=1, values_only=True):
        match read_state:
            case "readTableCode":
                table_column, table_code = _get_table_code(row)
                if table_code:
                    read_state = "readzAxis"
            case "readzAxis":
                z_axis_code = row[table_column]
                if z_axis_code:
                    z_axis[str(z_axis_code)] = str(row[table_column + 1]).strip()
                else:
                    read_state = "readRows"
            case "readRows":
                if row[table_column]:
                    if row_start == 0:
                        row_start = row_index
                    for cellindex in range(table_column + 2, len(row)):
                        if row[cellindex] is not None:
                            row_code = row[table_column]
                            if row_code and str(row_code).strip():
                                csv_file.write(
                                    f"{inst_code};{reg_code};{regoutput_code};{form_code};"
                                    f"{table_code};r{row[table_column]};c{col_nums_row[cellindex]};"
                                    f"{_format_cell_value(row[cellindex])}"
                                )
                                if z_axis:
                                    for z_code in z_axis_form:
                                        if z_code in z_axis:
                                            csv_file.write(f";{z_axis[z_code]}")
                                        else:
                                            csv_file.write(";")
                                csv_file.write("\n")
                elif row_start == 0:
                    col_nums_row = row
        row_index += 1


# ─────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────

def convert_file(excel_path: Path, output_dir: Path | None = None) -> list[Path]:
    """Convert a single Excel file to CSV(s).

    Parameters
    ----------
    excel_path : Path to the .xlsx file.
    output_dir : Folder where CSVs will be written.  Defaults to same folder as the xlsx.

    Returns
    -------
    List of generated CSV file paths.
    """
    excel_path = Path(excel_path)
    if output_dir is None:
        output_dir = excel_path.parent
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_base = output_dir / excel_path.stem
    csv_main = str(csv_base) + ".csv"
    generated: list[Path] = []

    # Delete previously generated CSVs for this file
    for f in output_dir.iterdir():
        if f.name.startswith(excel_path.stem) and f.suffix == ".csv":
            f.unlink()

    synthesis_generated = False

    try:
        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    except Exception as e:
        print(f"  [ERROR] Cannot open {excel_path.name}: {e}")
        return []

    inst_code, reg_code, regoutput_code, form_code = _read_identification(wb)

    # Determine z-axis keys and table types
    z_axis_form: dict[str, str] = {}
    table_list: dict[str, bool] = {}
    for sheet in wb.sheetnames:
        if sheet.lower() == "identification":
            continue
        table_code, is_list_form, z_axis = _read_table_characteristics(wb, sheet)
        table_list[sheet] = is_list_form
        for z_code in z_axis:
            if z_code not in z_axis_form:
                z_axis_form[z_code] = ""

    with open(csv_main, "w", newline="") as csv_file:
        # Write header
        csv_file.write(
            "institutionCode;regulationCode;regulatoryOutputCode;formCode;"
            "tableCode;rowId;columnId;value"
        )
        for z_code in z_axis_form:
            code = str(z_code)
            if "(" in code and ")" in code:
                code = code[code.index("(") + 1 : code.index(")")]
            csv_file.write(f";{code}")
        csv_file.write("\n")

        for sheet in wb.sheetnames:
            if sheet.lower() == "identification":
                continue
            if table_list.get(sheet, False):
                extra = _read_table_list(
                    str(csv_base), wb, sheet,
                    inst_code, reg_code, regoutput_code, form_code,
                )
                generated.extend(Path(p) for p in extra)
            else:
                _read_table(
                    wb, sheet, inst_code, reg_code, regoutput_code, form_code,
                    z_axis_form, csv_file,
                )
                synthesis_generated = True

    wb.close()

    if synthesis_generated:
        generated.insert(0, Path(csv_main))
        print(f"  [OK] {excel_path.name} -> {Path(csv_main).name}")
    else:
        # No regular table data — remove the empty main CSV
        os.remove(csv_main)

    return generated


def convert_folder(input_dir: Path, output_dir: Path | None = None) -> list[Path]:
    """Convert all .xlsx files in a folder to CSVs.

    Parameters
    ----------
    input_dir  : Folder containing .xlsx files.
    output_dir : Folder where CSVs will be written.  Defaults to input_dir.

    Returns
    -------
    List of all generated CSV file paths.
    """
    input_dir = Path(input_dir)
    if output_dir is None:
        output_dir = input_dir
    output_dir = Path(output_dir)

    xlsx_files = sorted(f for f in input_dir.iterdir() if f.suffix.lower() == ".xlsx")
    if not xlsx_files:
        print(f"  [WARN] No .xlsx files found in {input_dir}")
        return []

    print(f"  Converting {len(xlsx_files)} Excel file(s) ...")
    all_generated: list[Path] = []
    for xf in xlsx_files:
        all_generated.extend(convert_file(xf, output_dir))

    print(f"  {len(all_generated)} CSV file(s) generated in {output_dir}")
    return all_generated


# ─────────────────────────────────────────────
#  Interactive CLI
# ─────────────────────────────────────────────

# ANSI helpers
def _enable_ansi():
    if sys.platform == "win32":
        os.system("")


class _C:
    B = "\033[1m"
    CY = "\033[96m"
    GR = "\033[92m"
    YE = "\033[93m"
    RD = "\033[91m"
    BL = "\033[94m"
    MG = "\033[95m"
    DM = "\033[2m"
    E = "\033[0m"


C = _C


def _banner():
    art = [
        " ___        _   _     _          ___ _____   __",
        "| __|_ ____| |_| |   | |_ ___   / __/ __\\ \\ / /",
        "| _|\\ \\ / _| / / |   |  _/ _ \\ | (__\\__ \\\\ V / ",
        "|___/_\\_\\__|_\\_\\_|    \\__\\___/  \\___|___/ \\_/  ",
    ]
    print()
    for line in art:
        print(f"  {C.B}{C.YE}{line}{C.E}")
    print()


def main():
    _enable_ansi()
    _banner()

    # If argument provided, use it directly
    if len(sys.argv) > 1 and sys.argv[1] != "--interactive":
        import argparse
        parser = argparse.ArgumentParser(
            description="Convert regulatory Excel files (.xlsx) to CSV"
        )
        parser.add_argument("input", help="Excel file or directory of .xlsx files")
        parser.add_argument("--output", "-o", default=None, help="Output directory")
        args = parser.parse_args()

        input_path = Path(args.input)
        if not input_path.exists():
            print(f"  {C.RD}Error: {input_path} does not exist{C.E}")
            return 1

        if args.output:
            output_path = Path(args.output)
        else:
            output_path = input_path.parent / "excel_converted" if input_path.is_dir() else input_path.parent / "excel_converted"

        output_path.mkdir(parents=True, exist_ok=True)

        if input_path.is_dir():
            results = convert_folder(input_path, output_path)
        else:
            results = convert_file(input_path, output_path)

        if not results:
            print(f"  {C.RD}No CSV files generated.{C.E}")
            return 1
        print(f"\n  {C.GR}Done — {len(results)} file(s) generated in {output_path}{C.E}")
        return 0

    # Interactive mode — open folder picker
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()

    sep = "━" * 50
    print(f"  {C.B}{C.BL}{sep}")
    print(f"  Excel → CSV Converter")
    print(f"  {sep}{C.E}\n")

    print(f"  {C.BL}→ Select the Excel file(s) to convert ...{C.E}")
    print(f"  {C.DM}(Opening file explorer...){C.E}\n")

    chosen = filedialog.askopenfilenames(
        title="Select Excel (.xlsx) files to convert",
        filetypes=[("Excel files", "*.xlsx")],
    )
    root.destroy()

    if not chosen:
        print(f"  {C.RD}✗ No file selected. Aborted.{C.E}")
        return 1

    selected_files = sorted(Path(f) for f in chosen if not Path(f).name.startswith("~$"))
    if not selected_files:
        print(f"  {C.RD}No valid .xlsx files selected. Aborted.{C.E}")
        return 1

    print(f"  {C.GR}✓ {len(selected_files)} file(s) selected:{C.E}\n")
    for i, f in enumerate(selected_files, 1):
        print(f"    {C.CY}[{i}]{C.E} {f.parent.name}/{f.name}")
    print(flush=True)

    # Output = excel_converted/ next to the first file's parent folder
    output_dir = selected_files[0].parent.parent / "excel_converted"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n  {C.BL}Output → {output_dir}{C.E}")
    print(f"  {C.DM}Converting {len(selected_files)} file(s)...{C.E}\n")

    results = []
    for xf in selected_files:
        results.extend(convert_file(xf, output_dir))

    if not results:
        print(f"\n  {C.RD}No CSV files generated.{C.E}")
        return 1

    print(f"\n  {C.B}{C.GR}{'═' * 50}")
    print(f"  Done — {len(results)} CSV file(s) generated")
    print(f"  Output : {output_dir}")
    print(f"  {'═' * 50}{C.E}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
