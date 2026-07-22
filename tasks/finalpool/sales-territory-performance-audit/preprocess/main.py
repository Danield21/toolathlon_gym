#!/usr/bin/env python3
"""Preprocess script for sales-territory-performance-audit."""

import os
import shutil
from argparse import ArgumentParser
from pathlib import Path

import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
}

TASK_ROOT = Path(__file__).resolve().parent.parent


def clear_writable_schemas():
    conn = psycopg2.connect(**DB); cur = conn.cursor()
    # Full cleanup to prevent cross-task contamination — partial WHERE LIKE
    # 'stpa-%' clauses would not match agent-generated record IDs.
    for tbl in [
        "email.messages",
        "email.attachments",
        "email.sent_log",
        "email.drafts",
        "gcal.events",
    ]:
        try:
            cur.execute(f"DELETE FROM {tbl}")
        except Exception:
            conn.rollback()
            continue
        conn.commit()
    cur.close(); conn.close()


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_writable_schemas()

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        src = TASK_ROOT / "initial_workspace"
        for item in src.iterdir():
            dst = agent_ws / item.name
            if item.is_file() and not dst.exists():
                shutil.copy(item, dst)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
