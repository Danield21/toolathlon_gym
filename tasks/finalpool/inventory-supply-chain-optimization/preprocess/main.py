#!/usr/bin/env python3
"""Preprocess script for task setup"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import os


def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"),
        port=int(os.environ.get("PGPORT", "5432")),
        dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
        user="eigent", password="camel",
    )


def main():
    # WC + Snowflake are read-only; gcal/email/gsheet are writable and shared.
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clear writable schemas to ensure V2 determinism for inventory task.
    # Scope DELETE to inventory/supply-chain matching titles to avoid
    # disrupting unrelated concurrent tests (per QA finding).
    try:
        conn = get_conn()
        cur = conn.cursor()
        # Targeted gcal cleanup - only inventory/supply/procurement-themed events
        cur.execute(
            """DELETE FROM gcal.events
               WHERE summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s
                  OR summary ILIKE %s OR summary ILIKE %s""",
            ('%inventory%', '%supply%', '%procurement%',
             '%reorder%', '%purchase order%'))
        # Don't wipe email.messages globally - other tasks share it.
        # Only remove inventory/procurement-flavored test emails.
        cur.execute(
            "DELETE FROM email.messages WHERE subject ILIKE %s OR subject ILIKE %s OR subject ILIKE %s",
            ('%inventory%', '%procurement%', '%purchase order%'))
        # Clear inventory-related google sheets
        try:
            cur.execute(
                "DELETE FROM gsheet.spreadsheets WHERE title ILIKE %s OR title ILIKE %s OR title ILIKE %s",
                ('%inventory%', '%procurement%', '%tracking%'))
        except Exception:
            pass
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Preprocess DB cleanup skipped: {e}")

    # Copy initial_workspace files to agent_workspace if provided
    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
