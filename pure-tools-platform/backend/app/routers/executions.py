import json
import logging
import traceback
import zipfile
from io import BytesIO
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile, status
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.middleware.auth_middleware import get_current_user
from app.models.tool import Tool
from app.models.tool_execution import ToolExecution
from app.models.user import User
from app.schemas.execution import ExecutionListResponse, ExecutionResponse
from app.services.storage_service import create_execution_dirs
from app.services.tool_wrappers import (
    run_convert_espf_to_rds,
    run_forms_writer,
    run_init_project,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/executions", tags=["Executions"])


@router.post("/", response_model=ExecutionResponse)
async def execute_tool(
    tool_id: str = Form(...),
    parameters: str = Form("{}"),
    files: list[UploadFile] = File(default=[]),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Upload files, execute a tool, and return the execution record."""
    # Validate tool exists
    tool = db.query(Tool).filter(Tool.id == tool_id).first()
    if not tool:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Tool not found")

    # Parse parameters
    try:
        params = json.loads(parameters)
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid parameters JSON")

    # Create execution record
    execution = ToolExecution(
        user_id=current_user.id,
        tool_id=tool_id,
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

    # Save uploaded files
    saved_files = []
    for upload_file in files:
        file_path = input_dir / upload_file.filename
        with open(file_path, "wb") as f:
            content = await upload_file.read()
            f.write(content)
        saved_files.append(file_path)

    # Execute the tool
    try:
        tool_key = tool.script_path  # We use script_path as tool identifier

        if tool_key == "convert_espf_to_rds":
            if not saved_files:
                raise ValueError("Un fichier ESPF (.sXX) est requis")
            result = run_convert_espf_to_rds(
                input_file=saved_files[0],
                output_dir=output_dir,
                periodicity=params.get("periodicity", "monthly"),
                institution_code=params.get("institution_code", "RUBA"),
            )

        elif tool_key == "forms_writer":
            if not saved_files:
                raise ValueError("Un fichier source (.xlsx ou .xlsm) est requis")
            mode = params.get("mode", "eba")
            result = run_forms_writer(
                mode=mode,
                source_file=saved_files[0],
                output_dir=output_dir,
                params=params,
                validation_rules_file=saved_files[1] if len(saved_files) > 1 else None,
            )

        elif tool_key == "init_project":
            import_mode = params.get("import_mode", "manual")
            jira_pdf = saved_files[0] if import_mode == "jira_pdf" and saved_files else None
            result = run_init_project(
                regulatory_code=params.get("regulatory_code", ""),
                project_type=params.get("project_type", "Both"),
                git_project_name=params.get("git_project_name", ""),
                output_dir=output_dir,
                import_mode=import_mode,
                jira_pdf_file=jira_pdf,
                regulation_files=params.get("regulation_files"),
                output_files=params.get("output_files"),
                version=params.get("version", "2026.1.0"),
            )

        else:
            raise ValueError(f"Outil inconnu: {tool_key}")

        execution.status = "completed"
        db.commit()
        db.refresh(execution)

    except Exception as e:
        logger.error(f"Tool execution failed: {e}\n{traceback.format_exc()}")
        execution.status = "failed"
        execution.error_message = str(e)
        db.commit()
        db.refresh(execution)

    return execution


@router.get("/", response_model=list[ExecutionListResponse])
def list_executions(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """List executions. Admins see all; regular users see only their own."""
    query = (
        db.query(ToolExecution, Tool.name, Tool.script_path, User.full_name, User.email)
        .join(Tool, ToolExecution.tool_id == Tool.id)
        .join(User, ToolExecution.user_id == User.id)
    )
    if current_user.role != "admin":
        query = query.filter(ToolExecution.user_id == current_user.id)
    executions = query.order_by(ToolExecution.created_at.desc()).all()
    return [
        ExecutionListResponse(
            id=ex.id,
            tool_id=ex.tool_id,
            tool_name=tool_name,
            script_path=script_path,
            status=ex.status,
            created_at=ex.created_at,
            error_message=ex.error_message,
            user_name=user_full_name,
            user_email=user_email,
        )
        for ex, tool_name, script_path, user_full_name, user_email in executions
    ]


@router.get("/notifications/recent", response_model=list[ExecutionListResponse])
def get_recent_completed(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Return executions completed/failed in the last 60 seconds (for polling notifications)."""
    from datetime import datetime, timedelta, timezone

    cutoff = datetime.now(timezone.utc) - timedelta(seconds=60)
    executions = (
        db.query(ToolExecution, Tool.name)
        .join(Tool, ToolExecution.tool_id == Tool.id)
        .filter(
            ToolExecution.user_id == current_user.id,
            ToolExecution.status.in_(["completed", "failed"]),
            ToolExecution.created_at >= cutoff,
        )
        .order_by(ToolExecution.created_at.desc())
        .all()
    )
    return [
        ExecutionListResponse(
            id=ex.id,
            tool_id=ex.tool_id,
            tool_name=tool_name,
            status=ex.status,
            created_at=ex.created_at,
            error_message=ex.error_message,
        )
        for ex, tool_name in executions
    ]


@router.get("/{execution_id}", response_model=ExecutionResponse)
def get_execution(
    execution_id: str,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Get details of a specific execution."""
    execution = (
        db.query(ToolExecution)
        .filter(ToolExecution.id == execution_id, ToolExecution.user_id == current_user.id)
        .first()
    )
    if not execution:
        raise HTTPException(status_code=404, detail="Execution not found")
    return execution


@router.get("/{execution_id}/download")
def download_results(
    execution_id: str,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Download execution results as a zip file."""
    execution = (
        db.query(ToolExecution)
        .filter(ToolExecution.id == execution_id, ToolExecution.user_id == current_user.id)
        .first()
    )
    if not execution:
        raise HTTPException(status_code=404, detail="Execution not found")

    if execution.status != "completed":
        raise HTTPException(status_code=400, detail="Execution is not completed")

    output_dir = Path(execution.output_path)
    if not output_dir.exists():
        raise HTTPException(status_code=404, detail="Output files not found")

    output_files = list(output_dir.rglob("*"))
    output_files = [f for f in output_files if f.is_file()]

    if not output_files:
        raise HTTPException(status_code=404, detail="No output files generated")

    # If single file, return it directly
    if len(output_files) == 1:
        file_path = output_files[0]
        return StreamingResponse(
            open(file_path, "rb"),
            media_type="application/octet-stream",
            headers={"Content-Disposition": f'attachment; filename="{file_path.name}"'},
        )

    # Multiple files: zip them preserving directory structure
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in output_files:
            arcname = str(file_path.relative_to(output_dir))
            zf.write(file_path, arcname)
    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="results_{execution_id[:8]}.zip"'},
    )
