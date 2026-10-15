"""Storage service for managing uploaded and generated files."""

import os
import shutil
import uuid
from pathlib import Path

STORAGE_ROOT = Path(__file__).resolve().parent.parent.parent / "storage"
UPLOADS_DIR = STORAGE_ROOT / "uploads"
OUTPUTS_DIR = STORAGE_ROOT / "outputs"


def ensure_dirs():
    """Create storage directories if they don't exist."""
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)


def create_execution_dirs(user_id: int, execution_id: str) -> tuple[Path, Path]:
    """Create input/output directories for an execution."""
    input_dir = UPLOADS_DIR / str(user_id) / execution_id / "input"
    output_dir = OUTPUTS_DIR / str(user_id) / execution_id / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return input_dir, output_dir


def cleanup_execution(user_id: int, execution_id: str):
    """Remove all files for an execution."""
    upload_path = UPLOADS_DIR / str(user_id) / execution_id
    output_path = OUTPUTS_DIR / str(user_id) / execution_id
    if upload_path.exists():
        shutil.rmtree(upload_path)
    if output_path.exists():
        shutil.rmtree(output_path)
