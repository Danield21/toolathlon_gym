"""
Preprocess for fetch-terminal-data-pipeline task.

1. Clears gsheet schema.
"""

import os
import argparse
import asyncio

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_gsheet(conn):
    """Clear all Google Sheet data."""
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.spreadsheets")
    conn.commit()
    print("[preprocess] Cleared gsheet schema.")


def clean_residual_outputs(agent_workspace):
    """Remove any stale pipeline artifacts so the agent starts fresh."""
    if not agent_workspace or not os.path.isdir(agent_workspace):
        return
    residuals = [
        "clean_data.py",
        "cleaned_sales.json",
        "cleaned_inventory.json",
        "Data_Pipeline_Report.xlsx",
    ]
    for name in residuals:
        p = os.path.join(agent_workspace, name)
        if os.path.isfile(p):
            try:
                os.remove(p)
                print(f"[preprocess] Removed residual: {name}")
            except OSError as e:
                print(f"[preprocess] Warning: could not remove {name}: {e}")
    # Also reset memory/memory.json to an empty JSON object if it exists with stale content
    mem_dir = os.path.join(agent_workspace, "memory")
    mem_file = os.path.join(mem_dir, "memory.json")
    try:
        os.makedirs(mem_dir, exist_ok=True)
        with open(mem_file, "w") as f:
            f.write("{}")
        print("[preprocess] Reset memory/memory.json to empty object")
    except OSError as e:
        print(f"[preprocess] Warning: could not reset memory.json: {e}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clear residual output files so the task starts clean
    clean_residual_outputs(args.agent_workspace)

    # Clear Google Sheets schema
    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_gsheet(conn)
    finally:
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
