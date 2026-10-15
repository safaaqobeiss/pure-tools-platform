"""Résolution du dossier racine des outils (TOOLS_ROOT).

Ordre de recherche :
  1. variable TOOLS_ROOT (environnement ou fichier backend/.env)
  2. détection automatique autour du dossier du projet
  3. ancien chemin c:\\projects\\Tools (en dernier recours)
"""
import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)

# Dossiers attendus à l'intérieur de TOOLS_ROOT
TOOL_FOLDERS = (
    "pure-convertespftords",
    "pure-param-forms-writer",
    "pure-param-initproject",
    "pure-param-check-submission-input",
)

# app/services/tools_path.py -> parents[2] = dossier backend
BACKEND_DIR = Path(__file__).resolve().parents[2]


def _read_dotenv(key: str) -> str | None:
    """Lit une valeur dans backend/.env sans dépendance externe."""
    env_file = BACKEND_DIR / ".env"
    if not env_file.is_file():
        return None
    for line in env_file.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        if k.strip() == key:
            return v.strip().strip('"').strip("'")
    return None


def _has_tools(path: Path) -> bool:
    try:
        return path.is_dir() and any((path / name).is_dir() for name in TOOL_FOLDERS)
    except OSError:
        return False


def _auto_detect() -> Path | None:
    bases = [BACKEND_DIR, *BACKEND_DIR.parents][:4]
    for base in bases:
        if _has_tools(base):
            return base
        try:
            children = [c for c in base.iterdir() if c.is_dir()]
        except OSError:
            continue
        for child in children:
            if _has_tools(child):
                return child
    return None


def resolve_tools_root() -> Path:
    explicit = os.environ.get("TOOLS_ROOT") or _read_dotenv("TOOLS_ROOT")
    if explicit:
        return Path(explicit)

    detected = _auto_detect()
    if detected:
        return detected

    legacy = Path(r"c:\projects\Tools")
    logger.warning(
        "Dossier des outils introuvable. Ajoutez TOOLS_ROOT=... dans backend/.env "
        "(valeur actuelle utilisée : %s)",
        legacy,
    )
    return legacy


TOOLS_ROOT: Path = resolve_tools_root()
logger.info("TOOLS_ROOT = %s", TOOLS_ROOT)


def require_dir(path: Path) -> Path:
    """Retourne `path` s'il existe, sinon lève une erreur lisible."""
    if not path.is_dir():
        raise FileNotFoundError(
            f"Dossier de l'outil introuvable : {path}\n"
            f"TOOLS_ROOT actuel : {TOOLS_ROOT}\n"
            "Corrigez TOOLS_ROOT dans backend/.env "
            "(le dossier qui contient pure-convertespftords, pure-param-forms-writer, etc.)."
        )
    return path