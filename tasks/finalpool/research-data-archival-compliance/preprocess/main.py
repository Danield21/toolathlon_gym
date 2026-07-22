#!/usr/bin/env python3
"""Preprocess script for research-data-archival-compliance.

Clears writable schemas (email + gcal). The agent is expected to:
  - Build an archival_registry from the raw datasets in initial_workspace/raw_datasets
  - Author a compliance certificate document
  - Send a compliance verification email to research_leadership@institution.org
  - Schedule a research data compliance review meeting
"""

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


def get_conn():
    return psycopg2.connect(**DB)


def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    # Full cleanup to prevent cross-task contamination — partial WHERE LIKE
    # 'rda-%' clauses would leave agent-created records untouched.
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


def stage_raw_datasets(agent_workspace):
    """Copy fixture raw_datasets/ from initial_workspace into the agent workspace
    so the agent can inventory them."""
    src = TASK_ROOT / "initial_workspace" / "raw_datasets"
    if not agent_workspace or not src.exists():
        return
    dst = Path(agent_workspace) / "raw_datasets"
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()

    if args.agent_workspace:
        Path(args.agent_workspace).mkdir(parents=True, exist_ok=True)
        stage_raw_datasets(args.agent_workspace)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
