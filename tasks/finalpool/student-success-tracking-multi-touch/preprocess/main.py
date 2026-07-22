#!/usr/bin/env python3
"""Preprocess script for task setup"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import os

def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clean writable schemas for deterministic V2 runs.
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=os.environ.get("PGHOST", "localhost"), port=5432,
            dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
            user="eigent", password="camel",
        )
        conn.autocommit = True
        cur = conn.cursor()
        # Notion: clear pages/blocks/databases
        for stmt in (
            "DELETE FROM notion.blocks",
            "DELETE FROM notion.pages",
            "DELETE FROM notion.databases",
            "DELETE FROM gcal.events",
            "DELETE FROM email.messages WHERE folder_id IN (SELECT id FROM email.folders WHERE LOWER(name) IN ('sent','drafts'))",
        ):
            try:
                cur.execute(stmt)
            except Exception:
                pass
        cur.close()
        conn.close()
    except Exception as e:
        print(f"WARN: preprocess cleanup skipped: {e}")

    # Copy initial_workspace files to agent_workspace if provided
    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        initial_dir = Path(__file__).resolve().parent.parent / "initial_workspace"
        if initial_dir.exists():
            for src in initial_dir.iterdir():
                dst = agent_ws / src.name
                if src.is_file() and not dst.exists():
                    try:
                        shutil.copy2(src, dst)
                    except Exception:
                        pass

    print("Preprocess completed successfully")

if __name__ == "__main__":
    main()
