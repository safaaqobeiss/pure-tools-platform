"""Add role and is_active columns to users table."""
from app.database import engine
from sqlalchemy import text

with engine.connect() as conn:
    conn.execute(text("ALTER TABLE users ADD role VARCHAR(20) NOT NULL DEFAULT 'user'"))
    conn.execute(text("ALTER TABLE users ADD is_active BIT NOT NULL DEFAULT 1"))
    conn.commit()
    print("OK - columns added")
