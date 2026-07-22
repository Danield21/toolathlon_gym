"""
Preprocess for howtocook-notion-recipe-kb task.
- Clears Notion schema (comments, blocks, pages, databases)
- Ensures memory directory and file exist in agent workspace

Prerequisites:
  - PostgreSQL toolathlon_gym database running on localhost:5432
"""
import argparse
import json
import os

import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clear Notion tables
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()
    cur.execute("DELETE FROM notion.comments")
    cur.execute("DELETE FROM notion.blocks")
    cur.execute("DELETE FROM notion.pages")
    cur.execute("DELETE FROM notion.databases")
    conn.commit()
    cur.close()
    conn.close()
    print("Cleared Notion schema (comments, blocks, pages, databases)")

    # Ensure memory directory and file exist
    if args.agent_workspace:
        mem_dir = os.path.join(args.agent_workspace, "memory")
        os.makedirs(mem_dir, exist_ok=True)
        mem_file = os.path.join(mem_dir, "memory.json")
        if not os.path.exists(mem_file):
            with open(mem_file, "w") as f:
                json.dump({"entities": [], "relations": []}, f)
        print(f"Memory file ensured at {mem_file}")

        # Cleanup stale Recipe_Comparison.xlsx (avoid agent seeing previous run output)
        # Only cleanup if workspace is NOT the groundtruth_workspace (safety)
        ws_abs = os.path.abspath(args.agent_workspace)
        if "groundtruth" not in ws_abs.lower():
            stale_xlsx = os.path.join(args.agent_workspace, "Recipe_Comparison.xlsx")
            if os.path.exists(stale_xlsx):
                try:
                    os.remove(stale_xlsx)
                    print(f"Removed stale {stale_xlsx}")
                except Exception as e:
                    print(f"Could not remove stale xlsx: {e}")

    print("Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
