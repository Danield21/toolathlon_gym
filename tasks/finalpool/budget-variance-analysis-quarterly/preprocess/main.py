#!/usr/bin/env python3
"""Preprocess: clear writable gcal/email tables for clean state.
Snowflake budget data is read-only. Workspace files come from initial_workspace.
"""

from argparse import ArgumentParser
import os
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
}


def clear_writable():
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM gcal.events")
            cur.execute("DELETE FROM email.attachments")
            try:
                cur.execute("DELETE FROM email.sent_log")
            except Exception:
                pass
            cur.execute("DELETE FROM email.messages")
            try:
                cur.execute("DELETE FROM email.drafts")
            except Exception:
                pass
        conn.commit()
        conn.close()
        print("[preprocess] Cleared gcal.events and email tables.")
    except Exception as e:
        print(f"[preprocess] Warning: clear failed: {e}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_writable()
    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
