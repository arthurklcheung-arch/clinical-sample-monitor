"""
Clinical Sample Monitor - Database Initialization Script
Run this once to create the SQLite master database.
"""

import sqlite3
import os

# Database location
DB_PATH = os.path.join(os.path.dirname(__file__), "clinical_samples.db")
SCHEMA_PATH = os.path.join(os.path.dirname(__file__), "schema.sql")


def init_database():
    print("=" * 55)
    print("  Clinical Sample Monitor - Database Initializer")
    print("=" * 55)

    if os.path.exists(DB_PATH):
        confirm = input(f"\n⚠️  Database already exists at:\n   {DB_PATH}\n\nRe-initialize? This will NOT delete existing data. (y/n): ")
        if confirm.lower() != "y":
            print("Aborted.")
            return

    print(f"\n📂 Creating database at:\n   {DB_PATH}\n")

    with open(SCHEMA_PATH, "r") as f:
        schema_sql = f.read()

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        cursor.executescript(schema_sql)
        conn.commit()
        print("✅ Tables created:")
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
        tables = cursor.fetchall()
        for t in tables:
            print(f"   - {t[0]}")

        print("\n✅ Indexes created successfully")

        # Verify seed data
        cursor.execute("SELECT service_type_id, service_name FROM service_types;")
        services = cursor.fetchall()
        print("\n✅ Default service types seeded:")
        for s in services:
            print(f"   - {s[0]}: {s[1]}")

        print("\n🎉 Database initialized successfully!")
        print(f"\n📍 Location: {DB_PATH}")

    except Exception as e:
        conn.rollback()
        print(f"\n❌ Error: {e}")
        raise
    finally:
        conn.close()


def verify_database():
    """Quick health check on the database."""
    if not os.path.exists(DB_PATH):
        print("❌ Database not found. Run init_database() first.")
        return False

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
    tables = [t[0] for t in cursor.fetchall()]
    conn.close()

    expected = [
        "billing", "client_billing_rates", "clients",
        "projects", "qc_metrics", "run_metrics",
        "sample_status_tracking", "samples",
        "sequencing_runs", "service_types", "snp_markers"
    ]

    missing = [t for t in expected if t not in tables]
    if missing:
        print(f"⚠️  Missing tables: {missing}")
        return False

    print(f"✅ Database OK — {len(tables)} tables found")
    return True


if __name__ == "__main__":
    init_database()
    print()
    verify_database()
