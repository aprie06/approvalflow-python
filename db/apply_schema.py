"""
db/apply_schema.py

Applies db/schema.sql to the database pointed at by DATABASE_URL, using the
already-required psycopg2/SQLAlchemy dependencies directly. Avoids requiring
a separate psql CLI installation, which is unnecessary friction for anyone
trying to run this project locally.

Usage:
    export DATABASE_URL="postgresql://approvalflow:approvalflow_local@localhost:5432/approvalflow_dev"
    python3 -m db.apply_schema
"""

from pathlib import Path

from sqlalchemy import text

from db.models import get_engine

SCHEMA_PATH = Path(__file__).parent / "schema.sql"


def main():
    sql = SCHEMA_PATH.read_text()
    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(text(sql))
    print(f"Applied {SCHEMA_PATH.name} successfully.")


if __name__ == "__main__":
    main()
