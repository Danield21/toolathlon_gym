#!/usr/bin/env python3
"""Preprocess script for arxiv-research-workflow-pipeline.

Clears writable schemas (notion, gcal) so the agent's outputs can be
verified deterministically. Read-only data sources (arxiv) untouched.
"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import os


def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"), port=5432,
        dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
        user="eigent", password="camel",
    )


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clear writable schemas
    try:
        conn = get_conn()
        cur = conn.cursor()
        # Clear notion
        for tbl in ("notion.blocks", "notion.pages", "notion.databases"):
            try:
                cur.execute(f"DELETE FROM {tbl}")
            except Exception:
                conn.rollback()
                continue
        cur.execute("DELETE FROM gcal.events")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Preprocess DB cleanup skipped: {e}")

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        # Copy initial_workspace if it exists
        init_ws = Path(__file__).parent.parent / "initial_workspace"
        if init_ws.exists():
            for f in init_ws.iterdir():
                target = agent_ws / f.name
                if not target.exists():
                    if f.is_file():
                        shutil.copy2(f, target)
                    elif f.is_dir():
                        shutil.copytree(f, target)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
