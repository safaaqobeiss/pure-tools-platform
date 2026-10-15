"""Wrappers to invoke each Python tool programmatically."""

import json
import os
import subprocess
import sys
from pathlib import Path

from app.services.tools_path import TOOLS_ROOT, require_dir


def run_convert_espf_to_rds(
    input_file: Path,
    output_dir: Path,
    mapping_file: Path | None = None,
    espf_structure_file: Path | None = None,
    rds_structure_file: Path | None = None,
    periodicity: str = "monthly",
    institution_code: str = "RUBA",
) -> dict:
    """Run the ESPF to RDS conversion tool."""
    tool_dir = require_dir(TOOLS_ROOT / "pure-convertespftords")
    script = tool_dir / "convert_espf_to_rds.py"

    if mapping_file is None:
        mapping_file = tool_dir / "2025-11-21_Correspondance_entrees_specifiques CATR & creditContracts_NW.xlsx"
    if espf_structure_file is None:
        espf_structure_file = tool_dir / "RUBA_BDF_FR.xlsx"
    if rds_structure_file is None:
        rds_structure_file = tool_dir / "pure-transformedDataStructure-frbdfruba.xlsm"

    cmd = [
        sys.executable, str(script),
        "--input-sxx", str(input_file),
        "--output-dir", str(output_dir),
        "--mapping-file", str(mapping_file),
        "--espf-structure-file", str(espf_structure_file),
        "--rds-structure-file", str(rds_structure_file),
        "--periodicity", periodicity,
        "--institution-code", institution_code,
        "--no-dialog",
    ]

    result = subprocess.run(
        cmd, capture_output=True, text=True, cwd=str(tool_dir), timeout=300,
    )

    if result.returncode != 0:
        raise RuntimeError(f"Conversion failed:\n{result.stderr}")

    output_files = [str(f) for f in output_dir.iterdir() if f.is_file()]
    return {"output_files": output_files, "stdout": result.stdout}


def run_forms_writer(
    mode: str,
    source_file: Path,
    output_dir: Path,
    params: dict,
    validation_rules_file: Path | None = None,
) -> dict:
    """Run the forms writer tool (EBA or Report mode)."""
    tool_dir = require_dir(TOOLS_ROOT / "pure-param-forms-writer")

    if mode == "eba":
        script = tool_dir / "eba" / "cli_runner.py"
        env_params = {
            "PURE_NO_DIALOG": "1",
            "PURE_DPM_SOURCE": str(source_file),
            "PURE_OUTPUT_DIR": str(output_dir),
            "PURE_FORM_CODE": params.get("form_code", ""),
            "PURE_FORM_GROUP_CODE": params.get("form_group_code", ""),
            "PURE_VERSION": params.get("version", ""),
            "PURE_VALIDITY_DATE": params.get("validity_starting_date", ""),
            "PURE_LINKED_DOMAIN": params.get("linked_domain", ""),
            "PURE_DATATYPE": params.get("datatype", "euebacorep.positiveOrNullAmount"),
            "PURE_ROUNDING_UNIT": params.get("rounding_unit", ""),
            "PURE_DEBUG": "0",
        }
        if validation_rules_file:
            env_params["PURE_VALIDATION_RULES"] = str(validation_rules_file)
    elif mode == "report":
        script = tool_dir / "report" / "cli_runner.py"
        env_params = {
            "PURE_NO_DIALOG": "1",
            "PURE_REPORT_FORM": str(source_file),
            "PURE_OUTPUT_DIR": str(output_dir),
            "PURE_FORM_GROUP_CODE": params.get("form_group_code", "schemaaXYZ"),
            "PURE_VERSION": params.get("version", "2026.1.0"),
            "PURE_VALIDITY_DATE": params.get("validity_starting_date", "01/01/2026"),
            "PURE_LINKED_DOMAIN": params.get("linked_domain", "bebnbschemaa"),
            "PURE_DATATYPE": params.get("datatype", "bebnbschemaa.genericPositiveAmount"),
            "PURE_ROUNDING_UNIT": params.get("rounding_unit", "thousands"),
            "PURE_CONVERT_ROWS": "1" if params.get("convert_row_ids") else "0",
            "PURE_DEBUG": "0",
        }
    else:
        raise ValueError(f"Mode formulaire inconnu: {mode}. Utilisez 'eba' ou 'report'.")

    env = {**os.environ, **env_params}
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=True, text=True, cwd=str(tool_dir), env=env, timeout=600,
    )

    if result.returncode != 0:
        raise RuntimeError(f"Forms generation failed:\n{result.stderr}")

    output_files = [str(f) for f in output_dir.iterdir() if f.is_file()]
    return {"output_files": output_files, "stdout": result.stdout}


def run_init_project(
    regulatory_code: str,
    project_type: str,
    git_project_name: str,
    output_dir: Path,
    import_mode: str = "manual",
    jira_pdf_file: Path | None = None,
    regulation_files: list[str] | None = None,
    output_files: list[str] | None = None,
    version: str = "2026.1.0",
) -> dict:
    """Run the project initialization tool."""
    tool_dir = require_dir(TOOLS_ROOT / "pure-param-initproject" / "regulation-init")
    script = tool_dir / "src" / "cli_runner.py"

    env_params = {
        "PURE_NO_DIALOG": "1",
        "PURE_OUTPUT_DIR": str(output_dir),
        "PURE_IMPORT_MODE": import_mode,
        "PURE_REGULATORY_CODE": regulatory_code,
        "PURE_PROJECT_TYPE": project_type,
        "PURE_GIT_PROJECT_NAME": git_project_name,
        "PURE_VERSION": version,
    }
    if import_mode == "jira_pdf" and jira_pdf_file:
        env_params["PURE_JIRA_PDF"] = str(jira_pdf_file)
    if regulation_files:
        env_params["PURE_REGULATION_FILES"] = json.dumps(regulation_files)
    if output_files:
        env_params["PURE_OUTPUT_FILES"] = json.dumps(output_files)

    env = {**os.environ, **env_params}
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=True, text=True, cwd=str(tool_dir), env=env, timeout=300,
    )

    if result.returncode != 0:
        raise RuntimeError(f"Project init failed:\n{result.stderr}")

    output_files_list = [str(f) for f in output_dir.rglob("*") if f.is_file()]
    return {"output_files": output_files_list, "stdout": result.stdout}


def run_pure_form_generator(
    dpm_source: Path,
    dict_file: Path,
    template_file: Path,
    output_dir: Path,
    params: dict,
) -> dict:
    """Run the PURE Form Generator pipeline (Generation + Enrichment + RowKey)."""
    tool_dir = require_dir(TOOLS_ROOT / "pure-param-forms-writer" / "new")
    script = tool_dir / "cli_runner.py"

    env_params = {
        "PURE_NO_DIALOG": "1",
        "PURE_DPM_SOURCE": str(dpm_source),
        "PURE_DICT_PATH": str(dict_file),
        "PURE_TEMPLATE_PATH": str(template_file),
        "PURE_OUTPUT_DIR": str(output_dir),
        "PURE_FORM_CODE": params.get("form_code", ""),
        "PURE_FORM_GROUP_CODE": params.get("form_group_code", ""),
        "PURE_VERSION": params.get("version", ""),
        "PURE_VALIDITY_DATE": params.get("validity_starting_date", ""),
        "PURE_LINKED_DOMAIN": params.get("linked_domain", ""),
        "PURE_DATATYPE": params.get("datatype", ""),
        "PURE_ROUNDING_UNIT": params.get("rounding_unit", ""),
        "PURE_AUTO_ENRICH": "1" if params.get("auto_enrich", True) else "0",
        "PURE_AUTO_ROWKEY": "1" if params.get("auto_rowkey", True) else "0",
        "PURE_DEBUG": "1" if params.get("debug", False) else "0",
    }

    env = {**os.environ, **env_params}
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=True, text=True, cwd=str(tool_dir), env=env,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"PURE Form Generator failed:\n{result.stderr}\n{result.stdout}"
        )

    output_files = [str(f) for f in output_dir.rglob("*") if f.is_file()]
    return {"output_files": output_files, "stdout": result.stdout}