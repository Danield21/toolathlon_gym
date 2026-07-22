#!/usr/bin/env python3
"""Preprocess script for academic-conference-planning-coordination.

Clears writable schemas (email, gcal) so the agent's outgoing emails
and calendar events can be verified post-hoc. Read-only data sources
(google_forms, notion) are not cleared here.
"""

from argparse import ArgumentParser
from pathlib import Path
import os
import shutil


def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"),
        port=5432,
        dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
        user="eigent",
        password="camel",
    )


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clear gcal events so we can detect agent-created events
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM gcal.events")
        # Don't clear email.messages globally - other tasks share it.
        # Just remove old conference-related test emails.
        cur.execute(
            "DELETE FROM email.messages WHERE subject ILIKE %s OR subject ILIKE %s",
            ('%conference%', '%paper accept%'))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Preprocess DB cleanup skipped: {e}")

    # Set up agent workspace
    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        # Copy initial_workspace files if they exist
        init_ws = Path(__file__).parent.parent / "initial_workspace"
        if init_ws.exists():
            for f in init_ws.iterdir():
                target = agent_ws / f.name
                if not target.exists():
                    if f.is_file():
                        shutil.copy2(f, target)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
