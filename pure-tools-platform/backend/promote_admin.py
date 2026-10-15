from app.database import engine
from sqlalchemy import text

with engine.connect() as conn:
    # Promote all existing users to admin (first users)
    result = conn.execute(text("SELECT id, email FROM users"))
    for row in result:
        print(f"  User {row[0]}: {row[1]}")
    conn.execute(text("UPDATE users SET role='admin' WHERE id=(SELECT MIN(id) FROM users)"))
    conn.commit()
    print("First user promoted to admin")
