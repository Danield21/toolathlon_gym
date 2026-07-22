#!/usr/bin/env python3
"""Preprocess script for customer-satisfaction-analytics-framework.

Cleans stale email and gcal rows whose subject/summary matches our verifier
ILIKE patterns so prior task runs cannot falsely satisfy the Phase 6 checks.

Snowflake / WooCommerce data is read-only and pre-injected upstream.
"""

from argparse import ArgumentParser
from pathlib import Path
import os


def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"), port=5432,
        dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
        user="eigent", password="camel",
    )


def cleanup_phase6_artifacts():
    """Delete stale email/gcal rows matching verifier ILIKE patterns."""
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """DELETE FROM email.messages
               WHERE subject ILIKE %s OR subject ILIKE %s OR subject ILIKE %s
                  OR body_text ILIKE %s OR body_text ILIKE %s""",
            ('%satisfaction%', '%customer experience%', '%findings%',
             '%satisfaction%', '%customer experience%'))
        cur.execute(
            """DELETE FROM gcal.events
               WHERE summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s""",
            ('%satisfaction%', '%review%', '%customer%', '%feedback%'))
        conn.commit()
        conn.close()
        print("[preprocess] Cleaned stale email/gcal rows for Phase 6 checks.")
    except Exception as e:
        print(f"[preprocess] Cleanup skipped (DB not available): {e}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    cleanup_phase6_artifacts()

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
