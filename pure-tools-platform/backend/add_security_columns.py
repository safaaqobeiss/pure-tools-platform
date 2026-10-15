"""Add security_question and security_answer columns to users table."""
import pymssql
from urllib.parse import quote_plus

conn = pymssql.connect(
    server="localhost",
    port=1433,
    user="app_user",
    password="AppUser@2026#Secure",
    database="mydb",
)
cursor = conn.cursor()

for col, col_type in [
    ("security_question", "NVARCHAR(255)"),
    ("security_answer", "NVARCHAR(255)"),
]:
    try:
        cursor.execute(f"ALTER TABLE users ADD {col} {col_type} NULL")
        print(f"Column {col} added.")
    except Exception as e:
        if "Column names in each table must be unique" in str(e):
            print(f"Column {col} already exists.")
        else:
            raise

conn.commit()
conn.close()
print("Done.")
