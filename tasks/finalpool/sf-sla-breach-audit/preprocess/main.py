"""Preprocess: clear email data for clean state."""
import os
import argparse
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": int(os.environ.get("PGPORT", "5432")), "dbname": "toolathlon_gym", "user": "eigent", "password": "camel"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    # Clear messages/drafts/attachments only — KEEP folders and account_config
    # so the emails MCP retains its INBOX folder and account configuration.
    try:
        cur.execute("DELETE FROM email.sent_log")
    except Exception:
        pass
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    cur.execute("DELETE FROM email.attachments")
    # Ensure INBOX folder exists for MCP
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    if cur.fetchone() is None:
        cur.execute("INSERT INTO email.folders (name) VALUES ('INBOX')")
    conn.commit()
    cur.close()
    conn.close()
    print("Email messages/drafts/attachments cleared; folders/account_config preserved.")


if __name__ == "__main__":
    main()
