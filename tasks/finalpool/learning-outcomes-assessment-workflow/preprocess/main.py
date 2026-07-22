#!/usr/bin/env python3
"""Preprocess script for task setup"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import os


def cleanup_writable_schemas():
    """Clean stale data from writable schemas matching task ILIKE patterns.

    Verifier queries:
      - gform.forms ILIKE %learning%/%course%/%student%
      - gcal.events ILIKE %curriculum%/%review%/%learning%
      - email.messages subject/body ILIKE %learning%/%curriculum%/%assessment%/%outcome%
    Stale data from prior tasks could falsely pass these checks.
    """
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=os.environ.get("PGHOST", "localhost"), port=5432,
            dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
            user="eigent", password="camel",
        )
        cur = conn.cursor()
        cur.execute(
            """DELETE FROM gform.forms
               WHERE title ILIKE %s OR title ILIKE %s OR title ILIKE %s""",
            ('%learning%', '%course%', '%student%'))
        cur.execute(
            """DELETE FROM gcal.events
               WHERE summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s""",
            ('%curriculum%', '%review%', '%learning%'))
        cur.execute(
            """DELETE FROM email.messages
               WHERE subject ILIKE %s OR subject ILIKE %s OR subject ILIKE %s
                  OR body_text ILIKE %s""",
            ('%learning%', '%curriculum%', '%assessment%', '%outcome%'))
        conn.commit()
        cur.close()
        conn.close()
        print("Cleaned stale learning/curriculum/assessment data from gform.forms + gcal.events + email.messages")
    except Exception as e:
        print(f"Warning: cleanup_writable_schemas skipped: {e}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clean writable schemas to avoid stale data false positives
    cleanup_writable_schemas()

    # Copy initial_workspace files to agent_workspace if provided
    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)

    print("Preprocess completed successfully")

if __name__ == "__main__":
    main()
