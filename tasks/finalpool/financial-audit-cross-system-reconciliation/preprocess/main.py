#!/usr/bin/env python3
"""Preprocess script for financial-audit-cross-system-reconciliation.

Clears writable schemas (gcal, audit-related email entries) so the agent's
deliverables can be verified deterministically. WC/Snowflake are read-only.
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

    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM gcal.events")
        cur.execute(
            "DELETE FROM email.messages WHERE subject ILIKE %s OR subject ILIKE %s",
            ('%audit%', '%reconcil%'))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Preprocess DB cleanup skipped: {e}")

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
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
