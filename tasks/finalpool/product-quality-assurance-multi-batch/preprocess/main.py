#!/usr/bin/env python3
"""Preprocess script for task setup"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import os


def cleanup_writable_schemas():
    """Clean stale gcal/email/gsheet rows that may falsely satisfy verifier ILIKE checks."""
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=os.environ.get("PGHOST", "localhost"), port=5432,
            dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
            user="eigent", password="camel",
        )
        cur = conn.cursor()
        cur.execute(
            """DELETE FROM email.messages
               WHERE subject ILIKE %s OR subject ILIKE %s
                  OR subject ILIKE %s OR body_text ILIKE %s""",
            ('%quality%', '%defect%', '%batch%', '%quality%'))
        cur.execute(
            """DELETE FROM gcal.events
               WHERE summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s""",
            ('%quality%', '%review%', '%production%'))
        cur.execute(
            """DELETE FROM gsheet.spreadsheets
               WHERE title ILIKE %s OR title ILIKE %s OR title ILIKE %s""",
            ('%quality%', '%qa%', '%tracking%'))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"  [WARN] cleanup_writable_schemas: {e}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clean stale rows so evaluator ILIKE checks aren't pre-satisfied.
    cleanup_writable_schemas()

    # Copy initial_workspace files to agent_workspace if provided
    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
