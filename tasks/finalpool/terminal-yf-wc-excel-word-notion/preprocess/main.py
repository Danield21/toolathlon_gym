"""Preprocess for terminal-yf-wc-excel-word-notion.
Clears notion. YF and WC are read-only. Injects noise data."""
import argparse
import json
import os
import uuid
import glob as globmod

import psycopg2

def _db_config():
    """Build the PostgreSQL connection config.

    The MCP servers the agent calls (local_servers/*/pg_adapter.py) read the
    PG_HOST / PG_PORT / PG_DATABASE / PG_USER / PG_PASSWORD env-var family,
    while this preprocess (and evaluation) historically read the PGHOST family.
    Prefer the PGHOST family when it is present, otherwise fall back to the
    PG_* family so judge and agent keep reading the same database when only one
    family is exported, then to the local seeded defaults.
    """
    host = os.environ["PGHOST"] if "PGHOST" in os.environ else os.environ.get("PG_HOST", "localhost")
    port = int(os.environ["PGPORT"] if "PGPORT" in os.environ else os.environ.get("PG_PORT", "5432"))
    dbname = os.environ["PGDATABASE"] if "PGDATABASE" in os.environ else os.environ.get("PG_DATABASE", "toolathlon_gym")
    user = os.environ["PGUSER"] if "PGUSER" in os.environ else os.environ.get("PG_USER", "eigent")
    password = os.environ["PGPASSWORD"] if "PGPASSWORD" in os.environ else os.environ.get("PG_PASSWORD", "camel")
    return {"host": host, "port": port, "dbname": dbname, "user": user, "password": password}


DB_CONFIG = _db_config()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Clear notion
        cur.execute("DELETE FROM notion.comments")
        cur.execute("DELETE FROM notion.blocks")
        cur.execute("DELETE FROM notion.pages")
        cur.execute("DELETE FROM notion.databases")
        print("[preprocess] Cleared notion data.")

        # Inject noise notion data
        noise_db_id = str(uuid.uuid4())
        cur.execute("""
            INSERT INTO notion.databases (id, title, description, properties, parent, created_time, last_edited_time)
            VALUES (%s, %s::jsonb, %s::jsonb, %s::jsonb, %s::jsonb, NOW(), NOW())
        """, (noise_db_id,
              json.dumps([{"type": "text", "text": {"content": "Old Product Catalog"}}]),
              json.dumps([{"type": "text", "text": {"content": "Deprecated product list"}}]),
              json.dumps({"Name": {"id": "title", "type": "title", "title": {}}}),
              json.dumps({"type": "workspace", "workspace": True})))

        for i in range(2):
            page_id = str(uuid.uuid4())
            cur.execute("""
                INSERT INTO notion.pages (id, parent, properties, created_time, last_edited_time)
                VALUES (%s, %s::jsonb, %s::jsonb, NOW(), NOW())
            """, (page_id,
                  json.dumps({"type": "database_id", "database_id": noise_db_id}),
                  json.dumps({"Name": {"title": [{"text": {"content": f"Archived Product {i+1}"}}]}})))
        print("[preprocess] Injected noise notion data.")

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Commodity_Impact_Analysis.xlsx", "Pricing_Strategy_Memo.docx", "correlation_analysis.py"]:
            for f in globmod.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
