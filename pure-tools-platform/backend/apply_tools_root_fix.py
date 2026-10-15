"""Remplace le chemin en dur c:\\projects\\Tools dans les 3 routeurs.

Usage (depuis le dossier backend) :
    python apply_tools_root_fix.py
Une copie .bak de chaque fichier modifié est créée.
"""
import re
import shutil
import sys
from pathlib import Path

ROUTERS = ["init_project.py", "forms_generator.py", "check_submission.py"]
PATTERN = re.compile(
    r'^TOOLS_ROOT\s*=\s*Path\(\s*r?["\']c:\\+projects\\+Tools["\']\s*\)[ \t]*$',
    re.IGNORECASE | re.MULTILINE,
)
NEW_LINE = "from app.services.tools_path import TOOLS_ROOT  # noqa: E402"

base = Path(__file__).resolve().parent / "app" / "routers"
if not base.is_dir():
    sys.exit(f"Dossier introuvable : {base}\nLancez ce script depuis le dossier backend.")

for name in ROUTERS:
    path = base / name
    if not path.is_file():
        print(f"[!] {name} : fichier absent, ignoré")
        continue
    text = path.read_text(encoding="utf-8")
    if NEW_LINE in text:
        print(f"[=] {name} : déjà corrigé")
        continue
    new_text, n = PATTERN.subn(NEW_LINE, text)
    if n == 0:
        print(f"[!] {name} : ligne TOOLS_ROOT non trouvée, à modifier à la main")
        continue
    shutil.copy2(path, path.with_suffix(".py.bak"))
    path.write_text(new_text, encoding="utf-8")
    print(f"[OK] {name} : corrigé")