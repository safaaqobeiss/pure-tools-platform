"""
Dedicated router for the PURE Form Generator tool.
Provides endpoints for:
  - Listing available templates
  - Generating forms (Generation + Enrichment + RowKey pipeline)
"""

import json
import logging
import threading
import traceback
from pathlib import Path

from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    HTTPException,
    UploadFile,
    status,
)
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from app.database import get_db, SessionLocal
from app.middleware.auth_middleware import get_current_user
from app.models.tool import Tool
from app.models.tool_execution import ToolExecution
from app.models.user import User
from app.services.storage_service import create_execution_dirs
from app.services.tool_wrappers import run_pure_form_generator

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/forms-generator", tags=["FormsGenerator"])

from app.services.tools_path import TOOLS_ROOT  # noqa: E402
TEMPLATE_DIR = TOOLS_ROOT / "pure-param-forms-writer" / "template"
DEFAULT_TEMPLATE = TEMPLATE_DIR / "pure-form-regulatoryOutputCode-formCode-(shortName).xlsm"


@router.get("/template-info")
def get_template_info(current_user: User = Depends(get_current_user)):
    """Return information about the default template availability."""
    return {
        "template_available": DEFAULT_TEMPLATE.is_file(),
        "template_name": DEFAULT_TEMPLATE.name if DEFAULT_TEMPLATE.is_file() else None,
    }


def _run_generation_in_background(
    execution_id: str,
    dpm_path: Path,
    dict_path: Path,
    template_file: Path,
    output_dir: Path,
    params: dict,
):
    """Run the heavy generation process in a background thread."""
    db = SessionLocal()
    try:
        execution = db.query(ToolExecution).filter(ToolExecution.id == execution_id).first()
        if not execution:
            return

        result = run_pure_form_generator(
            dpm_source=dpm_path,
            dict_file=dict_path,
            template_file=template_file,
            output_dir=output_dir,
            params=params,
        )

        execution.status = "completed"
        db.commit()
        logger.info(f"Form generation {execution_id} completed successfully.")

    except Exception as e:
        logger.error(f"Form generation {execution_id} failed: {e}\n{traceback.format_exc()}")
        execution = db.query(ToolExecution).filter(ToolExecution.id == execution_id).first()
        if execution:
            execution.status = "failed"
            execution.error_message = str(e)
            db.commit()
    finally:
        db.close()


@router.post("/generate")
async def generate_forms(
    parameters: str = Form("{}"),
    dpm_file: UploadFile = File(...),
    dict_file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """
    Generate PURE forms from a DPM source + dictionary.
    Launches generation in background and returns immediately with execution_id.

    Parameters (JSON):
      - form_code: str
      - form_group_code: str (optional)
      - version: str (optional)
      - validity_starting_date: str (optional)
      - linked_domain: str (optional)
      - datatype: str (optional)
      - rounding_unit: str (optional)
      - auto_enrich: bool (default true)
      - auto_rowkey: bool (default true)
    """
    # Parse parameters
    try:
        params = json.loads(parameters)
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid parameters JSON")

    # Validate template exists
    if not DEFAULT_TEMPLATE.is_file():
        raise HTTPException(
            status_code=500,
            detail="Template XLSM not found on server",
        )

    # Find the tool in DB (script_path = "forms_writer")
    tool = db.query(Tool).filter(Tool.script_path == "forms_writer").first()
    if not tool:
        raise HTTPException(
            status_code=404,
            detail="Tool 'forms_writer' not found in database",
        )

    # Create execution record
    execution = ToolExecution(
        user_id=current_user.id,
        tool_id=tool.id,
        status="running",
        parameters=parameters,
    )
    db.add(execution)
    db.commit()
    db.refresh(execution)

    # Create storage dirs
    input_dir, output_dir = create_execution_dirs(current_user.id, execution.id)
    execution.input_path = str(input_dir)
    execution.output_path = str(output_dir)
    db.commit()

    # Save uploaded files
    dpm_path = input_dir / dpm_file.filename
    with open(dpm_path, "wb") as f:
        f.write(await dpm_file.read())

    dict_path = input_dir / dict_file.filename
    with open(dict_path, "wb") as f:
        f.write(await dict_file.read())

    # Launch generation in background thread
    thread = threading.Thread(
        target=_run_generation_in_background,
        args=(execution.id, dpm_path, dict_path, DEFAULT_TEMPLATE, output_dir, params),
        daemon=True,
    )
    thread.start()

    # Return immediately with execution ID
    return {
        "id": execution.id,
        "status": "running",
        "output_files": [],
        "stdout": "",
    }


@router.get("/{execution_id}/status")
def get_generation_status(
    execution_id: str,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Poll the status of a form generation execution."""
    execution = (
        db.query(ToolExecution)
        .filter(
            ToolExecution.id == execution_id,
            ToolExecution.user_id == current_user.id,
        )
        .first()
    )
    if not execution:
        raise HTTPException(status_code=404, detail="Execution not found")

    output_files = []
    if execution.status == "completed" and execution.output_path:
        output_dir = Path(execution.output_path)
        if output_dir.exists():
            output_files = [f.name for f in output_dir.rglob("*") if f.is_file()]

    return {
        "id": execution.id,
        "status": execution.status,
        "output_files": output_files,
        "error_message": execution.error_message or "",
    }


@router.get("/{execution_id}/download")
def download_results(
    execution_id: str,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Download generated forms as a ZIP file."""
    import zipfile
    from io import BytesIO

    execution = (
        db.query(ToolExecution)
        .filter(
            ToolExecution.id == execution_id,
            ToolExecution.user_id == current_user.id,
        )
        .first()
    )
    if not execution:
        raise HTTPException(status_code=404, detail="Execution not found")

    if execution.status != "completed":
        raise HTTPException(status_code=400, detail="Execution not completed")

    output_dir = Path(execution.output_path)
    if not output_dir.exists():
        raise HTTPException(status_code=404, detail="Output files not found")

    output_files = [f for f in output_dir.rglob("*") if f.is_file()]
    if not output_files:
        raise HTTPException(status_code=404, detail="No output files")

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in output_files:
            arcname = fpath.relative_to(output_dir)
            zf.write(fpath, arcname)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="pure-forms-{execution_id[:8]}.zip"'
        },
    )
