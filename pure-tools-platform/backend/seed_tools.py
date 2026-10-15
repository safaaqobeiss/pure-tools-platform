"""Seed script to insert the tools into the database."""

import sys
from pathlib import Path

# Add backend root to path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.database import SessionLocal, engine, Base
from app.models.user import User  # noqa: F401 - register model
from app.models.tool import Tool
from app.models.tool_execution import ToolExecution  # noqa: F401 - register model

TOOLS = [
    {
        "name": "Convertisseur ESPF vers RDS",
        "description": "Convertit les fichiers ESPF (format fixe .sXX) en fichiers CSV conformes au format RDS. Supporte les datasets CATR et ENGAG.",
        "icon": "SwapOutlined",
        "category": "Conversion",
        "script_path": "convert_espf_to_rds",
    },
    {
        "name": "Initialiseur de projet",
        "description": "Initialise les fichiers d'un projet Git réglementaire (dictionnaire, régulation, structure de données, dashboard, output). Supporte l'import depuis un ticket Jira PDF.",
        "icon": "ProjectOutlined",
        "category": "Projet",
        "script_path": "init_project",
    },
    {
        "name": "Générateur de formulaires PURE",
        "description": "Génère des formulaires PURE Excel (.xlsm) à partir de spécifications DPM (EBA) ou de formulaires Report (BNB). Supporte les feeding rules et la conversion de RowID.",
        "icon": "FileExcelOutlined",
        "category": "Formulaires",
        "script_path": "forms_writer",
    },
    {
        "name": "Validateur Input vs Submission",
        "description": "Compare les données d'entrée (CSV/Excel) aux soumissions XBRL attendues au niveau cellule. Supporte XBRL-XML, XBRL-CSV ZIP et OneGate XML. Génère des rapports colorés Excel avec statuts OK/DIFF/MISSING.",
        "icon": "SafetyCertificateOutlined",
        "category": "Validation",
        "script_path": "check_submission",
    },
]


def seed():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        # Remove old tools that no longer exist
        old_tools = db.query(Tool).filter(
            Tool.script_path.in_(["forms_from_dpm_eba", "forms_from_report"])
        ).all()
        for old in old_tools:
            print(f"  Removed old tool: {old.name}")
            db.delete(old)

        for tool_data in TOOLS:
            existing = db.query(Tool).filter(Tool.script_path == tool_data["script_path"]).first()
            if existing:
                # Update existing tool
                for key, value in tool_data.items():
                    setattr(existing, key, value)
                print(f"  Updated: {tool_data['name']}")
            else:
                tool = Tool(**tool_data)
                db.add(tool)
                print(f"  Added: {tool_data['name']}")
        db.commit()
        print("Seed completed successfully!")
    finally:
        db.close()


if __name__ == "__main__":
    seed()
