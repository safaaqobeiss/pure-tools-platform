"""
Dedicated router for the Check Submission Input validation tool.

Provides an endpoint to validate input data against XBRL submissions
using the batch mode of validate_regulation_unified.py.
"""
import json
import logging
import shutil
import sys
import tempfile
import traceback
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from pydantic import BaseModel
from sqlalchemy.orm import Session

from app.database import get_db
from app.middleware.auth_middleware import get_current_user
from app.models.tool import Tool
from app.models.tool_execution import ToolExecution
from app.models.user import User
from app.services.storage_service import create_execution_dirs

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/check-submission", tags=["CheckSubmission"])

from app.services.tools_path import TOOLS_ROOT  # noqa: E402
CHECK_SUBMISSION_ROOT = TOOLS_ROOT / "pure-param-check-submission-input"


# ── Schemas ──────────────────────────────────────────────

class CaseSummary(BaseModel):
    case_id: str
    OK: int = 0
    DIFF: int = 0
    MISSING_IN_OUTPUT: int = 0
    CALCULATED_OR_EXTRA: int = 0
    TOTAL: int = 0


class ValidateResponse(BaseModel):
    execution_id: str
    status: str
    summary: dict = {}
    cases_summary: list[CaseSummary] = []
    output_files: list[str] = []
    error_message: str | None = None


# ── Validate (batch mode) ───────────────────────────────

@router.post("/validate", response_model=ValidateResponse)
async def validate_submission(
    regulation: str = Form(...),
    input_files: list[UploadFile] = File(...),
    expected_files: list[UploadFile] = File(...),
    taxonomy_file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """
    Validate input data against expected XBRL submissions (batch mode).

    Expects:
    - regulation: Regulation name (e.g. RUBA, COREP, P3DH…)
    - input_files: Files from the input/load directory (with relative paths
      like INST1/csv_direct/file.csv preserved via webkitRelativePath)
    - expected_files: Expected submission files (.xbrl, .zip, .xml)
    - taxonomy_file: Taxonomy package (xbrl_package.tar.gz)
    """

    # Find the tool
    tool = db.query(Tool).filter(Tool.script_path == "check_submission").first()
    if not tool:
        raise HTTPException(status_code=404, detail="Check submission tool not found in DB")

    # Create execution record
    params = {"regulation": regulation}
    execution = ToolExecution(
        user_id=current_user.id,
        tool_id=tool.id,
        status="running",
        parameters=json.dumps(params, ensure_ascii=False),
    )
    db.add(execution)
    db.commit()
    db.refresh(execution)

    _, output_dir = create_execution_dirs(current_user.id, execution.id)
    execution.output_path = str(output_dir)

    tmp_dir = Path(tempfile.mkdtemp())
    load_dir = tmp_dir / "input" / "load"
    expected_dir = tmp_dir / "expected"
    load_dir.mkdir(parents=True)
    expected_dir.mkdir(parents=True)

    try:
        # ── Save input files preserving directory structure ──
        # The filename from the browser carries the webkitRelativePath,
        # e.g. "load/INST1/csv_direct/file.csv" or "INST1/csv_direct/file.csv"
        for f in input_files:
            raw_name = f.filename or ""
            # Strip the root folder name from webkitRelativePath
            parts = raw_name.replace("\\", "/").split("/")
            # Remove first element if it's the uploaded root folder name
            if len(parts) > 1:
                rel = "/".join(parts[1:])
            else:
                rel = parts[0]
            file_path = load_dir / rel
            file_path.parent.mkdir(parents=True, exist_ok=True)
            with open(file_path, "wb") as fh:
                content = await f.read()
                fh.write(content)

        # ── Save expected submission files ──
        for f in expected_files:
            raw_name = f.filename or ""
            # Only keep the file name itself (strip any folder prefix)
            name = raw_name.replace("\\", "/").split("/")[-1]
            file_path = expected_dir / name
            with open(file_path, "wb") as fh:
                content = await f.read()
                fh.write(content)

        # ── Save taxonomy ──
        taxonomy_path = tmp_dir / (taxonomy_file.filename or "xbrl_package.tar.gz")
        with open(taxonomy_path, "wb") as fh:
            content = await taxonomy_file.read()
            fh.write(content)

        # ── Run batch validation via CLI subprocess ──
        # This avoids import conflicts and keeps the tool isolated
        script = CHECK_SUBMISSION_ROOT / "validate_regulation_unified.py"
        cmd = [
            sys.executable,
            str(script),
            "--batch",
            "--load-dir", str(load_dir),
            "--expected-dir", str(expected_dir),
            "--taxonomy", str(taxonomy_path),
            "--output-dir", str(output_dir),
        ]

        logger.info(f"Running validation: {' '.join(cmd)}")

        proc = await _run_subprocess(cmd, cwd=str(CHECK_SUBMISSION_ROOT))

        # ── Collect output files ──
        output_files = [
            str(f.relative_to(output_dir))
            for f in output_dir.rglob("*")
            if f.is_file()
        ]

        # ── Parse summary from stdout ──
        global_summary, cases_summary = _parse_batch_output(proc.stdout or "")

        if proc.returncode == 0:
            execution.status = "completed"
            db.commit()
            return ValidateResponse(
                execution_id=execution.id,
                status="completed",
                summary=global_summary,
                cases_summary=cases_summary,
                output_files=output_files,
            )
        else:
            error_msg = (proc.stderr or proc.stdout or "Unknown error")[-2000:]
            execution.status = "failed"
            execution.error_message = error_msg
            db.commit()
            return ValidateResponse(
                execution_id=execution.id,
                status="failed",
                summary=global_summary,
                cases_summary=cases_summary,
                output_files=output_files,
                error_message=error_msg,
            )

    except Exception as e:
        logger.error(f"Validation failed: {e}\n{traceback.format_exc()}")
        execution.status = "failed"
        execution.error_message = str(e)
        db.commit()

        return ValidateResponse(
            execution_id=execution.id,
            status="failed",
            output_files=[],
            error_message=str(e),
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ── Helpers ──────────────────────────────────────────────

import asyncio
import subprocess as _subprocess


async def _run_subprocess(cmd: list[str], cwd: str | None = None):
    """Run a subprocess asynchronously."""
    import os as _os

    env = {**_os.environ, "PYTHONIOENCODING": "utf-8"}
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(
        None,
        lambda: _subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=cwd,
            timeout=600,  # 10 min max
            env=env,
        ),
    )


import re as _re


def _parse_batch_output(stdout: str) -> tuple[dict, list[CaseSummary]]:
    """
    Parse the batch stdout to extract per-case and global summaries.

    Expected format lines:
      CASE_ID              OK=   12 | DIFF=    0 | MISSING=    0 | EXTRA=    0 | TOTAL=   12
      GLOBAL               OK=   24 | DIFF=    0 | MISSING=    0 | EXTRA=    0 | TOTAL=   24
    """
    cases: list[CaseSummary] = []
    global_summary: dict = {}

    pattern = _re.compile(
        r"^\s*(\S+)\s+OK=\s*(\d+)\s*\|\s*DIFF=\s*(\d+)\s*\|\s*MISSING=\s*(\d+)\s*\|\s*EXTRA=\s*(\d+)\s*\|\s*TOTAL=\s*(\d+)",
        _re.MULTILINE,
    )

    for m in pattern.finditer(stdout):
        case_id = m.group(1)
        ok = int(m.group(2))
        diff = int(m.group(3))
        missing = int(m.group(4))
        extra = int(m.group(5))
        total = int(m.group(6))

        if case_id == "GLOBAL":
            global_summary = {
                "OK": ok,
                "DIFF": diff,
                "MISSING_IN_OUTPUT": missing,
                "CALCULATED_OR_EXTRA": extra,
                "TOTAL": total,
            }
        else:
            cases.append(CaseSummary(
                case_id=case_id,
                OK=ok,
                DIFF=diff,
                MISSING_IN_OUTPUT=missing,
                CALCULATED_OR_EXTRA=extra,
                TOTAL=total,
            ))

    return global_summary, cases
