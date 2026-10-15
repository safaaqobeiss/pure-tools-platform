"""
launch_validation.py
====================
Interactive launcher wizard for the regulatory validation tool.
Guides the user step by step with file/folder dialog boxes.
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

# Add script folder to path for import
sys.path.insert(0, str(Path(__file__).resolve().parent))


# ─────────────────────────────────────────────
#  ANSI
# ─────────────────────────────────────────────
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


# ─────────────────────────────────────────────
#  Helpers UI
# ─────────────────────────────────────────────
_root = None


def _init_tk():
    global _root
    import tkinter as tk

    _root = tk.Tk()
    _root.withdraw()
    _root.attributes("-topmost", True)
    _root.update()


def _destroy_tk():
    global _root
    if _root:
        _root.destroy()
        _root = None


def _pick_file(title: str, filetypes: list[tuple[str, str]] | None = None) -> Path | None:
    from tkinter import filedialog

    if filetypes is None:
        filetypes = [("All files", "*.*")]
    print(f"  {C.DM}(Opening file explorer...){C.E}")
    _root.update()
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if not path:
        print(f"  {C.RD}✗ No file selected.{C.E}")
        return None
    p = Path(path)
    print(f"  {C.GR}✓ {p}{C.E}")
    return p


def _pick_folder(title: str) -> Path | None:
    from tkinter import filedialog

    print(f"  {C.DM}(Opening folder explorer...){C.E}")
    _root.update()
    path = filedialog.askdirectory(title=title)
    if not path:
        print(f"  {C.RD}✗ No folder selected.{C.E}")
        return None
    p = Path(path)
    print(f"  {C.GR}✓ {p}{C.E}")
    return p


def _ask_yn(prompt: str, default: bool = True) -> bool:
    hint = "[Y/n]" if default else "[y/N]"
    while True:
        r = input(f"  {C.YE}{prompt} {hint} : {C.E}").strip().lower()
        if not r:
            return default
        if r in ("y", "yes"):
            return True
        if r in ("n", "no"):
            return False


def _ask_choice(options: list[tuple[str, str]]) -> str:
    for i, (label, _) in enumerate(options, 1):
        print(f"    {C.CY}{i}.{C.E} {label}")
    while True:
        r = input(f"\n  {C.YE}Votre choix : {C.E}").strip()
        if r.isdigit() and 1 <= int(r) <= len(options):
            sel = options[int(r) - 1]
            print(f"  {C.GR}→ {sel[0]}{C.E}")
            return sel[1]


def _step(num: int, total: int, title: str):
    bar = "━" * 60
    print(f"\n{C.B}{C.BL}{bar}")
    print(f"  Step {num}/{total} — {title}")
    print(f"{bar}{C.E}\n")


def _abort(msg: str = "Aborted.") -> int:
    print(f"\n  {C.RD}{msg}{C.E}")
    _destroy_tk()
    return 1


# ─────────────────────────────────────────────
#  Banner
# ─────────────────────────────────────────────
def _banner():
    # Lengths: Check=17, Submission=41, Input=21
    L_CHK, L_SUB, L_INP = 17, 41, 21
    check = [
        "  ___ _           _   ",
        " / __| |_  ___ __| |__",
        "| (__| ' \\/ -_) _| / /",
        " \\___|_||_\\___\\__|_\\_\\",
        "                      ",
    ]
    submission = [
        " ___      _          _       _          ",
        "/ __|_  _| |__ _ __ (_)_____(_)___ _ _  ",
        "\\__ \\ || | '_ \\ '  \\| (_-<_-< / _ \\ ' \\ ",
        "|___/\\_,_|_.__/_|_|_|_/__/__/_\\___/_||_|",
        "                                         ",
    ]
    inp = [
        " ___                _   ",
        "|_ _|_ _  _ __ _  _| |_ ",
        " | || ' \\| '_ \\ || |  _|",
        "|___|_||_| .__/\\_,_|\\__|",
        "         |_|            ",
    ]
    print()
    for c, s, i in zip(check, submission, inp):
        print(f"  {C.B}{C.YE}{c}{C.MG}{s}{C.CY}{i}{C.E}")
    print()


# ─────────────────────────────────────────────
#  Wizard
# ─────────────────────────────────────────────
REGULATIONS = [
    "RUBA",
    "COREP",
    "RESOL",
    "FP",
    "P3DH",
    "MCONTRAN",
    "DBM",
]


def _scan_input_dir(root: Path) -> dict[str, dict[str, Path]]:
    """Scan an input directory for institution subfolders with csv_direct/ and/or excel/.

    Returns e.g. {'RUBA1': {'csv_direct': Path(...), 'excel': Path(...)}, ...}
    Also handles the case where root itself contains csv_direct/ or excel/ (single institution).
    """
    found: dict[str, dict[str, Path]] = {}
    # Check if root itself has csv_direct/ or excel/
    has_csv = (root / "csv_direct").is_dir()
    has_xl = (root / "excel").is_dir()
    if has_csv or has_xl:
        entry: dict[str, Path] = {}
        if has_csv:
            entry["csv_direct"] = root / "csv_direct"
        if has_xl:
            entry["excel"] = root / "excel"
        found[root.name] = entry
        return found
    # Otherwise scan subfolders
    for sub in sorted(root.iterdir()):
        if not sub.is_dir():
            continue
        entry = {}
        if (sub / "csv_direct").is_dir():
            entry["csv_direct"] = sub / "csv_direct"
        if (sub / "excel").is_dir():
            entry["excel"] = sub / "excel"
        if entry:
            found[sub.name] = entry
    return found


def _find_xbrl_files(expected_dir: Path) -> dict[str, tuple[Path, str]]:
    """Scan expected directory for submission files (.xbrl, .xml, .zip), index by institution prefix.

    Returns e.g. {'RUBA1': (Path('RUBA1_rubad_submission.xbrl'), 'rubad'), ...}
    """
    EXPECTED_EXTS = {".xbrl", ".xml", ".zip"}
    result: dict[str, tuple[Path, str]] = {}
    for f in sorted(expected_dir.iterdir()):
        if f.suffix.lower() in EXPECTED_EXTS and "submission" in f.name.lower():
            # Handle double extensions like .submission.zip
            stem = f.stem
            if stem.lower().endswith(".submission"):
                stem = stem[: -len(".submission")]
            parts = stem.split("_")
            if len(parts) >= 2:
                inst = parts[0]
                reg_output = parts[1]  # e.g. rubad, rubaj, mcontran, or "submission"
                if reg_output.lower() == "submission":
                    reg_output = "output"  # default when filename is INST_submission.ext
                # Keep the shortest-named file per institution (avoid _dataset variants)
                if inst not in result or len(f.name) < len(result[inst][0].name):
                    result[inst] = (f, reg_output)
            elif len(parts) == 1:
                # Just INST_submission.ext -> only inst, no reg_output in name
                inst = parts[0]
                if inst not in result:
                    result[inst] = (f, "output")
    return result


def _convert_excel(excel_dir: Path) -> Path | None:
    """Convert xlsx files from excel_dir to CSVs in a temp dir using the built-in converter."""
    import tempfile
    from excel_to_csv import convert_folder

    conv_tmp = Path(tempfile.mkdtemp(prefix="xlsx_conv_"))
    try:
        results = convert_folder(excel_dir, conv_tmp)
        if not results:
            print(f"  {C.RD}No CSV files generated from Excel conversion.{C.E}")
            return None
        return conv_tmp
    except Exception as e:
        print(f"  {C.RD}Conversion error: {e}{C.E}")
        return None


def _run_single(
    label: str,
    input_dir: Path,
    xbrl_file: Path,
    taxonomy: Path,
    output_dir: Path,
) -> dict:
    """Run one validation and return the summary dict."""
    bar = "━" * 60
    print(f"\n{C.B}{C.BL}{bar}")
    print(f"  {label}")
    print(f"{bar}{C.E}\n")

    from validate_regulation_unified import run

    rc = run(
        input_dir=input_dir,
        xbrl_file=xbrl_file,
        taxonomy_path=taxonomy,
        output_dir=output_dir,
    )
    summary = rc if isinstance(rc, dict) else {"RETURN_CODE": rc}
    ok = summary.get("OK", 0)
    diff = summary.get("DIFF", 0)
    miss = summary.get("MISSING_IN_OUTPUT", 0)
    extra = summary.get("CALCULATED_OR_EXTRA", 0)
    total = summary.get("TOTAL", 0)

    print(f"\n{C.B}{C.MG}{'═' * 60}")
    print(f"  {label}")
    print(f"  OK={ok}  DIFF={diff}  MISSING={miss}  EXTRA={extra}  TOTAL={total}")
    print(f"{'═' * 60}{C.E}")
    return summary


def main() -> int:
    _enable_ansi()
    _init_tk()
    _banner()

    TOTAL = 5

    # ══════════════════════════════════════════
    #  Step 1 — Choose regulation
    # ══════════════════════════════════════════
    _step(1, TOTAL, "Choose regulation")
    print(f"  Select the regulation to validate:\n")
    for i, name in enumerate(REGULATIONS, 1):
        print(f"    {C.CY}{i}.{C.E} {name}")
    print(f"    {C.CY}{len(REGULATIONS) + 1}.{C.E} Other (enter name)")

    while True:
        r = input(f"\n  {C.YE}Your choice : {C.E}").strip()
        if r.isdigit() and 1 <= int(r) <= len(REGULATIONS):
            reg_name = REGULATIONS[int(r) - 1]
            break
        if r.isdigit() and int(r) == len(REGULATIONS) + 1:
            reg_name = ""
            while not reg_name:
                reg_name = input(f"  {C.YE}Regulation name : {C.E}").strip().upper()
            break
    print(f"  {C.GR}✓ Regulation : {reg_name}{C.E}")

    # ══════════════════════════════════════════
    #  Step 2 — Input directory
    # ══════════════════════════════════════════
    _step(2, TOTAL, "Input directory")
    print(f"  {C.DM}Select the root folder containing institution subfolders")
    print(f"  (each with csv_direct/ and/or excel/ subfolders).{C.E}")
    print(f"  {C.BL}→ Select input directory ...{C.E}\n")
    input_root = _pick_folder(f"[{reg_name}] Input directory")
    if not input_root:
        return _abort("Input directory is required. Aborted.")

    institutions = _scan_input_dir(input_root)
    if not institutions:
        return _abort("No institution subfolders with csv_direct/ or excel/ found.")

    has_excel = any("excel" in fmt for fmt in institutions.values())
    has_csv = any("csv_direct" in fmt for fmt in institutions.values())

    print(f"\n  {C.GR}Found {len(institutions)} institution(s):{C.E}")
    for inst, fmts in institutions.items():
        tags = ", ".join(fmts.keys())
        print(f"    {C.CY}{inst}{C.E} → {tags}")

    # ══════════════════════════════════════════
    #  Step 3 — Expected directory
    # ══════════════════════════════════════════
    _step(3, TOTAL, "Expected submission directory")
    print(f"  {C.DM}Select the folder containing expected submission files (.xbrl, .xml, .zip).{C.E}")
    print(f"  {C.BL}→ Select expected directory ...{C.E}\n")
    expected_dir = _pick_folder(f"[{reg_name}] Expected submission directory")
    if not expected_dir:
        return _abort("Expected directory is required. Aborted.")

    xbrl_map = _find_xbrl_files(expected_dir)
    if not xbrl_map:
        return _abort(f"No *_submission.{{xbrl,xml,zip}} files found in {expected_dir}")

    print(f"\n  {C.GR}Found {len(xbrl_map)} expected file(s):{C.E}")
    for inst, (xf, reg_out) in xbrl_map.items():
        print(f"    {C.CY}{inst}{C.E} → {xf.name}  ({reg_out})")

    # Match institutions with expected files (smart matching with suffix fallback)
    # Three passes: direct match, suffix match, trailing-number fallback
    # (inst, format, input_dir, xbrl_file, reg_output_name)
    test_plan: list[tuple[str, str, Path, Path, str]] = []
    unmatched = []
    used_keys: set[str] = set()
    matched: dict[str, tuple[Path, str]] = {}

    # Pass 1: Direct match
    for inst in institutions:
        if inst in xbrl_map:
            matched[inst] = xbrl_map[inst]
            used_keys.add(inst)

    # Pass 2: Suffix match (e.g. input 'BE1' matches expected 'FINIFBE1')
    for inst in institutions:
        if inst in matched:
            continue
        upper = inst.upper()
        candidates = [k for k in xbrl_map if k not in used_keys and k.upper().endswith(upper)]
        if len(candidates) == 1:
            matched[inst] = xbrl_map[candidates[0]]
            used_keys.add(candidates[0])
        elif len(candidates) > 1:
            candidates.sort(key=len)
            matched[inst] = xbrl_map[candidates[0]]
            used_keys.add(candidates[0])

    # Pass 3: Trailing-number fallback for EU-type (e.g. 'EU1' -> 'FINIF1')
    for inst in institutions:
        if inst in matched:
            continue
        m = re.match(r'^[A-Za-z]+(\d+)$', inst)
        if m:
            num = m.group(1)
            num_cands = [
                k for k in xbrl_map
                if k not in used_keys and re.match(r'^[A-Za-z]+' + re.escape(num) + r'$', k)
            ]
            if len(num_cands) == 1:
                matched[inst] = xbrl_map[num_cands[0]]
                used_keys.add(num_cands[0])

    for inst, fmts in institutions.items():
        if inst not in matched:
            unmatched.append(inst)
            continue
        xbrl_file, reg_output = matched[inst]
        for fmt_name, fmt_dir in fmts.items():
            test_plan.append((inst, fmt_name, fmt_dir, xbrl_file, reg_output))

    if unmatched:
        print(f"\n  {C.YE}⚠  No matching expected file for: {', '.join(unmatched)}{C.E}")
    if not test_plan:
        return _abort("No matching institution/expected-file pairs found.")

    # ══════════════════════════════════════════
    #  Step 4 — Taxonomy
    # ══════════════════════════════════════════
    _step(4, TOTAL, "XBRL Taxonomy")
    print(f"  {C.DM}Expected file : xbrl_package.tar.gz{C.E}")
    print(f"  {C.BL}→ Select the taxonomy file ...{C.E}\n")
    taxonomy_path = _pick_file(
        f"[{reg_name}] Select xbrl_package.tar.gz",
        [("Taxonomy tar.gz", "*.gz"), ("All files", "*.*")],
    )
    if not taxonomy_path:
        return _abort("Taxonomy is required. Aborted.")



    # ══════════════════════════════════════════
    #  Step 5 — Output directory (Conversion XBRL/XML → CSV intégrée)
    # ══════════════════════════════════════════
    _step(5, TOTAL, "Output directory (XBRL/XML → CSV intégré)")

    print(f"\n  {C.BL}→ Select the output folder for reports ...{C.E}\n")
    output_base = _pick_folder(f"[{reg_name}] Output folder for reports")
    if not output_base:
        return _abort("Output folder is required. Aborted.")

    # ══════════════════════════════════════════
    #  Summary & confirmation
    # ══════════════════════════════════════════
    sep = "═" * 60
    print(f"\n{C.B}{C.CY}{sep}")
    print(f"  SUMMARY — {reg_name}")
    print(f"{sep}{C.E}\n")
    print(f"  Input root :     {input_root}")
    print(f"  Expected dir :   {expected_dir}")
    print(f"  Taxonomy :       {taxonomy_path}")
    print(f"  XBRL→CSV conv. : (intégré Python)")
    if has_excel:
        print(f"  Excel conv. :    (built-in)")
    print(f"  Output :         {output_base}")
    print(f"\n  {C.B}Tests to run ({len(test_plan)}):{C.E}")
    for inst, fmt, inp, xbrl, reg_out in test_plan:
        print(f"    {C.CY}{inst}{C.E} [{fmt}] → {reg_out}/{inst}  ({xbrl.name})")
    print()

    if not _ask_yn("Launch all validations?"):
        print(f"  {C.YE}Validation cancelled.{C.E}")
        _destroy_tk()
        return 0

    _destroy_tk()

    # ══════════════════════════════════════════
    #  Execution
    # ══════════════════════════════════════════
    results: dict[str, dict] = {}

    for inst, fmt, inp_dir, xbrl_file, reg_out in test_plan:
        label = f"{reg_name} / {reg_out} / {inst} — {fmt}"
        out_dir = output_base / reg_out / inst / fmt
        out_dir.mkdir(parents=True, exist_ok=True)

        if fmt == "excel":
            print(f"\n  {C.BL}Converting Excel → CSV for {inst}...{C.E}")
            converted = _convert_excel(inp_dir)
            if not converted:
                results[f"{inst}_{fmt}"] = {"ERROR": "Excel conversion failed"}
                continue
            print(f"  {C.GR}✓ Conversion done{C.E}")
            actual_input = converted
        else:
            actual_input = inp_dir

        results[f"{inst}_{fmt}"] = _run_single(
            label=label,
            input_dir=actual_input,
            xbrl_file=xbrl_file,
            taxonomy=taxonomy_path,
            output_dir=out_dir,
        )

    print()
    return 0


# ─────────────────────────────────────────────
#  Auto-test mode (hardcoded paths for quick re-runs)
# ─────────────────────────────────────────────
_TEST_BASE = Path(r"C:\Users\ooumas\Desktop\pure-regulatoryoutput-frbdfruba-bdfrubad-master")
_TEST_INPUT = _TEST_BASE / "e2e-test-data" / "input" / "load"
_TEST_EXPECTED = _TEST_BASE / "e2e-test-data" / "expected"
_TEST_TAXONOMY = _TEST_BASE / "xbrl" / "xbrl_package.tar.gz"
# _CONVERTER_JAR supprimé, conversion intégrée
# Excel converter is now built-in (excel_to_csv.py)
_OUTPUT_BASE = Path(__file__).resolve().parent / "user_reports" / "ruba" / "rubad"


def auto_test() -> int:
    """Run CSV + Excel tests for all institutions found in test data."""
    _enable_ansi()
    _banner()

    print(f"  {C.B}{C.CY}AUTO-TEST MODE — RUBA rubad{C.E}\n")

    institutions = _scan_input_dir(_TEST_INPUT)
    xbrl_map = _find_xbrl_files(_TEST_EXPECTED)
    results: dict[str, dict] = {}

    for inst, fmts in institutions.items():
        if inst not in xbrl_map:
            print(f"  {C.YE}⚠  No XBRL file for {inst} — skipping{C.E}")
            continue
        xbrl_file, reg_output = xbrl_map[inst]

        for fmt_name, fmt_dir in fmts.items():
            out_dir = _OUTPUT_BASE / reg_output / inst / fmt_name
            out_dir.mkdir(parents=True, exist_ok=True)

            if fmt_name == "excel":
                print(f"\n  {C.BL}Converting Excel → CSV for {inst}...{C.E}")
                converted = _convert_excel(fmt_dir)
                if not converted:
                    results[f"{inst}_{fmt_name}"] = {"ERROR": "Excel conversion failed"}
                    continue
                print(f"  {C.GR}✓ Conversion done{C.E}")
                actual_input = converted
            else:
                actual_input = fmt_dir

            results[f"{inst}_{fmt_name}"] = _run_single(
                label=f"{inst} — {fmt_name}",
                input_dir=actual_input,
                xbrl_file=xbrl_file,
                taxonomy=_TEST_TAXONOMY,
                # conversion intégrée, plus de converter_jar
                output_dir=out_dir,
            )

    # Final summary
    print(f"\n\n{C.B}{C.CY}{'═' * 60}")
    print(f"  FINAL SUMMARY — RUBA rubad auto-test")
    print(f"{'═' * 60}{C.E}\n")

    all_ok = True
    for name, s in results.items():
        if "ERROR" in s:
            status = f"{C.RD}ERROR{C.E}"
            all_ok = False
        elif s.get("DIFF", 0) > 0 or s.get("MISSING_IN_OUTPUT", 0) > 0:
            status = (
                f"{C.RD}FAIL{C.E}  OK={s['OK']}  DIFF={s['DIFF']}  "
                f"MISSING={s['MISSING_IN_OUTPUT']}"
            )
            all_ok = False
        else:
            status = f"{C.GR}PASS{C.E}  OK={s.get('OK', '?')}  TOTAL={s.get('TOTAL', '?')}"
        print(f"  {C.B}{name:25s}{C.E} {status}")

    print()
    if all_ok:
        print(f"  {C.B}{C.GR}All tests passed!{C.E}")
    else:
        print(f"  {C.B}{C.RD}Some tests have differences.{C.E}")
    print()

    return 0 if all_ok else 1


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--auto-test":
        raise SystemExit(auto_test())
    raise SystemExit(main())
