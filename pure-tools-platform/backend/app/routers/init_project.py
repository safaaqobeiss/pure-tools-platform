"""
Dedicated router for the Init Project tool.

Provides endpoints to:
- Parse a Jira PDF and return extracted data
- List reference template files (for manual wizard)
- Generate files (with full params)
- Push generated results to GitLab
"""
import json
import logging
import shutil
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

router = APIRouter(prefix="/api/init-project", tags=["InitProject"])

from app.services.tools_path import TOOLS_ROOT  # noqa: E402
INITPROJECT_ROOT = TOOLS_ROOT / "pure-param-initproject" / "regulation-init" / "src"


# ── Schemas ──────────────────────────────────────────────

class JiraParsedResponse(BaseModel):
    regulation_code: str = ""
    project_type: str = "both"
    git_project_name: str = ""
    version: str = ""
    validity_date: str = ""
    authority_name: str = ""
    ticket_key: str = ""
    title: str = ""
    regulation_project_name: str = ""
    output_project_name: str = ""
    git_project_url: str = ""
    output_code: str = ""
    output_name: str = ""
    output_description: str = ""
    output_periodicity: str = ""
    output_type: str = ""
    common_collection: str = ""
    submission_format: str = ""
    container_type: str = ""
    regulation_identification: dict = {}
    dictionary_identification: dict = {}
    transformation_rules_identification: dict = {}
    transformed_data_identification: dict = {}
    ui_dashboard_identification: dict = {}
    output_identification: dict = {}
    output_structure_identification: dict = {}
    file_mappings: dict = {}


class ReferenceFile(BaseModel):
    name: str
    extension: str


class ReferenceFilesResponse(BaseModel):
    regulation_files: list[ReferenceFile]
    output_files: list[ReferenceFile]


class GenerateRequest(BaseModel):
    regulatory_code: str
    project_type: str = "both"
    git_project_name: str = ""
    version: str = "2026.1.0"
    regulation_file_map: dict[str, str] = {}
    output_file_map: dict[str, str] = {}
    regulation_identification: dict = {}
    dictionary_identification: dict = {}
    transformation_rules_identification: dict = {}
    transformed_data_identification: dict = {}
    ui_dashboard_identification: dict = {}
    output_identification: dict = {}
    output_structure_identification: dict = {}


class GenerateResponse(BaseModel):
    execution_id: str
    status: str
    output_files: list[str]
    error_message: str | None = None


class GitLabPushRequest(BaseModel):
    execution_id: str
    branch: str = "main"
    token: str
    regulation_project_name: str = ""
    output_project_name: str = ""


class GitLabPushResponse(BaseModel):
    success: bool
    message: str


# ── Parse Jira PDF ──────────────────────────────────────

@router.post("/parse-jira", response_model=JiraParsedResponse)
async def parse_jira_pdf_endpoint(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    """Upload a Jira PDF and return the extracted regulation data."""
    import sys
    import tempfile

    # Save the uploaded PDF to a temp file
    tmp_dir = Path(tempfile.mkdtemp())
    pdf_path = tmp_dir / file.filename
    with open(pdf_path, "wb") as f:
        content = await file.read()
        f.write(content)

    try:
        # Import the jira_parser from the tool's source
        sys.path.insert(0, str(INITPROJECT_ROOT))
        from jira_parser import parse_jira_file, jira_data_to_collected

        jira_data = parse_jira_file(str(pdf_path))
        collected = jira_data_to_collected(jira_data)

        return JiraParsedResponse(
            regulation_code=collected.get("regulation_code", ""),
            project_type=collected.get("project_type", "both"),
            git_project_name=collected.get("git_project_name", ""),
            version=jira_data.version or "",
            validity_date=jira_data.validity_date or "",
            authority_name=jira_data.authority_name or "",
            ticket_key=jira_data.ticket_key or "",
            title=jira_data.title or "",
            regulation_project_name=jira_data.regulation_project_name or collected.get("regulation_project_name", ""),
            output_project_name=jira_data.output_project_name or collected.get("output_project_name", ""),
            git_project_url=jira_data.git_project_url or "",
            output_code=jira_data.output_code or "",
            output_name=jira_data.output_name or "",
            output_description=jira_data.output_description or "",
            output_periodicity=jira_data.output_periodicity or "",
            output_type=jira_data.output_type or "",
            common_collection=jira_data.common_collection or "",
            submission_format=jira_data.submission_format or "",
            container_type=jira_data.container_type or "",
            regulation_identification=collected.get("regulation_identification", {}),
            dictionary_identification=collected.get("dictionary_identification", {}),
            transformation_rules_identification=collected.get("transformation_rules_identification", {}),
            transformed_data_identification=collected.get("transformed_data_identification", {}),
            ui_dashboard_identification=collected.get("ui_dashboard_identification", {}),
            output_identification=collected.get("output_identification", {}),
            output_structure_identification=collected.get("output_structure_identification", {}),
            file_mappings=jira_data.file_mappings or {},
        )
    except Exception as e:
        logger.error(f"Jira parse error: {e}\n{traceback.format_exc()}")
        raise HTTPException(status_code=400, detail=f"Erreur lors du parsing du PDF: {e}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        # Clean up sys.path
        if str(INITPROJECT_ROOT) in sys.path:
            sys.path.remove(str(INITPROJECT_ROOT))


# ── List Reference Files ────────────────────────────────

@router.get("/reference-files", response_model=ReferenceFilesResponse)
def list_reference_files(
    current_user: User = Depends(get_current_user),
):
    """List available template files for regulation and output scopes."""
    regulation_dir = INITPROJECT_ROOT / "regulation"
    output_dir = INITPROJECT_ROOT / "Output"

    reg_files = []
    if regulation_dir.exists():
        for f in sorted(regulation_dir.iterdir()):
            if f.is_file():
                reg_files.append(ReferenceFile(name=f.name, extension=f.suffix))

    out_files = []
    if output_dir.exists():
        for f in sorted(output_dir.iterdir()):
            if f.is_file():
                out_files.append(ReferenceFile(name=f.name, extension=f.suffix))

    return ReferenceFilesResponse(
        regulation_files=reg_files,
        output_files=out_files,
    )


# ── Generate ────────────────────────────────────────────

@router.post("/generate", response_model=GenerateResponse)
async def generate_project(
    request: GenerateRequest,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Generate the project files using the full configuration."""
    import sys

    # Find the init_project tool
    tool = db.query(Tool).filter(Tool.script_path == "init_project").first()
    if not tool:
        raise HTTPException(status_code=404, detail="Init project tool not found in DB")

    # Create execution record
    execution = ToolExecution(
        user_id=current_user.id,
        tool_id=tool.id,
        status="running",
        parameters=json.dumps(request.model_dump(), ensure_ascii=False),
    )
    db.add(execution)
    db.commit()
    db.refresh(execution)

    _, output_dir = create_execution_dirs(current_user.id, execution.id)
    execution.output_path = str(output_dir)

    try:
        sys.path.insert(0, str(INITPROJECT_ROOT))
        from config import AppConfig
        from models import ProjectInput, RegulationInfo
        from project_initializer import ProjectInitializer

        config = AppConfig(
            output_base_dir=output_dir,
            project_base_name=request.git_project_name or request.regulatory_code.lower(),
            project_type=request.project_type,
            regulation_folder_name="",
            output_folder_name="",
            regulation_file_map=request.regulation_file_map,
            output_file_map=request.output_file_map,
            regulation_identification=request.regulation_identification,
            dictionary_identification=request.dictionary_identification,
            output_identification=request.output_identification,
            transformation_rules_identification=request.transformation_rules_identification,
            transformed_data_identification=request.transformed_data_identification,
            ui_dashboard_identification=request.ui_dashboard_identification,
            output_structure_identification=request.output_structure_identification,
        )

        regulation = RegulationInfo(
            code=request.regulatory_code,
            version=request.version,
            name="",
            authority="",
            frequency="",
            first_reporting_date="",
        )

        project_input = ProjectInput(regulation=regulation)
        initializer = ProjectInitializer(config, project_input)
        result_dir = initializer.run()

        # Move generated files from nested subdir to output_dir root
        if result_dir.resolve() != output_dir.resolve():
            for p in result_dir.rglob("*"):
                if p.is_file():
                    rel = p.relative_to(result_dir)
                    dest = output_dir / rel
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(p), str(dest))
            # Remove the now-empty nested directory
            shutil.rmtree(result_dir, ignore_errors=True)

        output_files = [
            str(f.relative_to(output_dir)) for f in output_dir.rglob("*") if f.is_file()
        ]

        execution.status = "completed"
        db.commit()
        db.refresh(execution)

        return GenerateResponse(
            execution_id=execution.id,
            status="completed",
            output_files=output_files,
        )

    except Exception as e:
        logger.error(f"Generation failed: {e}\n{traceback.format_exc()}")
        execution.status = "failed"
        execution.error_message = str(e)
        db.commit()
        db.refresh(execution)

        return GenerateResponse(
            execution_id=execution.id,
            status="failed",
            output_files=[],
            error_message=str(e),
        )
    finally:
        if str(INITPROJECT_ROOT) in sys.path:
            sys.path.remove(str(INITPROJECT_ROOT))


# ── Jira PDF Generate (combined parse + generate) ──────

@router.post("/generate-from-jira", response_model=GenerateResponse)
async def generate_from_jira(
    file: UploadFile = File(...),
    parameters: str = Form("{}"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Parse a Jira PDF and generate the project in one step."""
    import sys
    import tempfile

    params = json.loads(parameters)

    # Save PDF
    tmp_dir = Path(tempfile.mkdtemp())
    pdf_path = tmp_dir / file.filename
    with open(pdf_path, "wb") as f:
        content = await file.read()
        f.write(content)

    # Find tool
    tool = db.query(Tool).filter(Tool.script_path == "init_project").first()
    if not tool:
        raise HTTPException(status_code=404, detail="Init project tool not found")

    # Create execution
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

    try:
        sys.path.insert(0, str(INITPROJECT_ROOT))
        from jira_parser import parse_jira_file, jira_data_to_collected
        from config import AppConfig
        from models import ProjectInput, RegulationInfo
        from project_initializer import ProjectInitializer

        # Parse
        jira_data = parse_jira_file(str(pdf_path))
        collected = jira_data_to_collected(jira_data)

        code = collected.get("regulation_code", "")
        project_type = collected.get("project_type", "both")
        git_name = collected.get("git_project_name", "") or code.lower()

        # Build file maps from reference dirs
        import re
        def sanitize_project_name(value: str) -> str:
            cleaned = re.sub(r"[^a-zA-Z0-9_-]+", "-", (value or "").strip().lower())
            return cleaned.strip("-_")

        def sanitize_file_name(value: str, extension: str) -> str:
            base = re.sub(r"[^a-zA-Z0-9._()\-]+", "-", (value or "").strip())
            base = base.strip(" .")
            if not base:
                return ""
            if not base.lower().endswith(extension.lower()):
                base = f"{base}{extension}"
            return base

        def tokenized_default_name(file_name: str, regulation_code: str) -> str:
            lower_code = sanitize_project_name(regulation_code)
            tokenized = file_name
            replacements = {
                "regulatorCode": "REGULATION_LOWER_CODE",
                "regulationCode": "REGULATION_LOWER_CODE",
                "regulatoryDomainCode": "REGULATION_LOWER_CODE",
                "regulatoryOutputCode": "REGULATORY_OUTPUT_CODE",
                "outputStructureCode": "OUTPUT_STRUCTURE_CODE",
                "formCode": "FORM_CODE",
                "(shortName)": "SHORT_NAME",
            }
            for old, new in replacements.items():
                tokenized = tokenized.replace(old, new)
            tokenized = tokenized.replace("REGULATION_LOWER_CODE", lower_code or "regulation")
            return tokenized

        def resolve_jira_file_name(ref_name, regulation_code, jira_mappings):
            for jira_src, jira_dst in jira_mappings.items():
                jira_prefix = jira_src.rsplit("-", 1)[0] if "-" in jira_src else jira_src
                ref_prefix = ref_name.rsplit("-", 1)[0] if "-" in ref_name else ref_name
                if jira_prefix.lower() == ref_prefix.lower():
                    return jira_dst
            return tokenized_default_name(ref_name, regulation_code)

        jira_mappings = collected.get("file_mappings", {}) or jira_data.file_mappings or {}
        reg_file_map = {}
        out_file_map = {}

        if project_type in ("both", "regulation"):
            ref_dir = INITPROJECT_ROOT / "regulation"
            if ref_dir.exists():
                for f in sorted(ref_dir.iterdir()):
                    if f.is_file():
                        target = resolve_jira_file_name(f.name, code, jira_mappings)
                        final = sanitize_file_name(target, f.suffix)
                        if final:
                            reg_file_map[f.name] = final

        if project_type in ("both", "output"):
            ref_dir = INITPROJECT_ROOT / "Output"
            if ref_dir.exists():
                for f in sorted(ref_dir.iterdir()):
                    if f.is_file():
                        target = resolve_jira_file_name(f.name, code, jira_mappings)
                        final = sanitize_file_name(target, f.suffix)
                        if final:
                            out_file_map[f.name] = final

        git_name = sanitize_project_name(git_name)

        config = AppConfig(
            output_base_dir=output_dir,
            project_base_name=git_name,
            project_type=project_type,
            regulation_folder_name="",
            output_folder_name="",
            regulation_file_map=reg_file_map,
            output_file_map=out_file_map,
            regulation_identification=collected.get("regulation_identification", {}),
            dictionary_identification=collected.get("dictionary_identification", {}),
            output_identification=collected.get("output_identification", {}),
            transformation_rules_identification=collected.get("transformation_rules_identification", {}),
            transformed_data_identification=collected.get("transformed_data_identification", {}),
            ui_dashboard_identification=collected.get("ui_dashboard_identification", {}),
            output_structure_identification=collected.get("output_structure_identification", {}),
        )

        regulation = RegulationInfo(
            code=code,
            version=jira_data.version or params.get("version", ""),
            name="",
            authority=jira_data.authority_name or "",
            frequency="",
            first_reporting_date="",
        )

        project_input = ProjectInput(regulation=regulation)
        initializer = ProjectInitializer(config, project_input)
        result_dir = initializer.run()

        # Move generated files from nested subdir to output_dir root
        if result_dir.resolve() != output_dir.resolve():
            for p in result_dir.rglob("*"):
                if p.is_file():
                    rel = p.relative_to(result_dir)
                    dest = output_dir / rel
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(p), str(dest))
            # Remove the now-empty nested directory
            shutil.rmtree(result_dir, ignore_errors=True)

        output_files = [
            str(f.relative_to(output_dir)) for f in output_dir.rglob("*") if f.is_file()
        ]

        execution.status = "completed"
        db.commit()
        db.refresh(execution)

        return GenerateResponse(
            execution_id=execution.id,
            status="completed",
            output_files=output_files,
        )

    except Exception as e:
        logger.error(f"Jira generation failed: {e}\n{traceback.format_exc()}")
        execution.status = "failed"
        execution.error_message = str(e)
        db.commit()
        db.refresh(execution)
        return GenerateResponse(
            execution_id=execution.id,
            status="failed",
            output_files=[],
            error_message=str(e),
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        if str(INITPROJECT_ROOT) in sys.path:
            sys.path.remove(str(INITPROJECT_ROOT))


# ── GitLab Push ─────────────────────────────────────────

@router.post("/push-gitlab", response_model=GitLabPushResponse)
def push_to_gitlab(
    request: GitLabPushRequest,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """Push generated files to GitLab."""
    import sys

    execution = (
        db.query(ToolExecution)
        .filter(ToolExecution.id == request.execution_id, ToolExecution.user_id == current_user.id)
        .first()
    )
    if not execution or execution.status != "completed":
        raise HTTPException(status_code=400, detail="Execution not found or not completed")

    output_dir = Path(execution.output_path)
    if not output_dir.exists():
        raise HTTPException(status_code=400, detail="Output directory not found")

    try:
        sys.path.insert(0, str(INITPROJECT_ROOT))
        from config import GitLabConfig
        from gitlab_publisher import GitLabPublisher

        gitlab_cfg = GitLabConfig(private_token=request.token)
        publisher = GitLabPublisher(gitlab_cfg)
        username = publisher.get_authenticated_username()

        params = json.loads(execution.parameters) if execution.parameters else {}
        project_type = params.get("project_type", "both")
        reg_code = params.get("regulatory_code", "")

        messages = []

        if project_type in ("both", "regulation") and request.regulation_project_name:
            reg_project = request.regulation_project_name
            if "/" not in reg_project:
                reg_project = f"{username}/{reg_project}"
            publisher.ensure_project_exists(reg_project, username, create_if_missing=True)
            reg_dir = output_dir / "regulation"
            if not reg_dir.exists():
                # Try finding it in a subdirectory
                for d in output_dir.iterdir():
                    if d.is_dir():
                        sub_reg = d / "regulation"
                        if sub_reg.exists():
                            reg_dir = sub_reg
                            break
            if reg_dir.exists():
                publisher.publish_directory(
                    reg_project, reg_dir, request.branch,
                    f"chore: init {reg_code} regulation",
                )
                messages.append(f"Regulation pushed to {reg_project}")

        if project_type in ("both", "output") and request.output_project_name:
            out_project = request.output_project_name
            if "/" not in out_project:
                out_project = f"{username}/{out_project}"
            publisher.ensure_project_exists(out_project, username, create_if_missing=True)
            out_dir = output_dir / "output"
            if not out_dir.exists():
                for d in output_dir.iterdir():
                    if d.is_dir():
                        sub_out = d / "output"
                        if sub_out.exists():
                            out_dir = sub_out
                            break
            if out_dir.exists():
                publisher.publish_directory(
                    out_project, out_dir, request.branch,
                    f"chore: init {reg_code} output",
                )
                messages.append(f"Output pushed to {out_project}")

        return GitLabPushResponse(success=True, message="; ".join(messages) or "Push completed")

    except Exception as e:
        logger.error(f"GitLab push failed: {e}\n{traceback.format_exc()}")
        return GitLabPushResponse(success=False, message=str(e))
    finally:
        if str(INITPROJECT_ROOT) in sys.path:
            sys.path.remove(str(INITPROJECT_ROOT))
