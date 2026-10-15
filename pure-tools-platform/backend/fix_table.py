"""Drop and recreate tool_executions table with correct FK types."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.database import engine
from sqlalchemy import text

with engine.connect() as conn:
    conn.execute(text("IF OBJECT_ID('tool_executions', 'U') IS NOT NULL DROP TABLE tool_executions"))
    conn.commit()
    print("Dropped tool_executions table")

# Now recreate with correct types
from app.database import Base
from app.models.user import User  # noqa
from app.models.tool import Tool  # noqa
from app.models.tool_execution import ToolExecution  # noqa

Base.metadata.create_all(bind=engine)
print("Recreated tool_executions table with BIGINT user_id")
