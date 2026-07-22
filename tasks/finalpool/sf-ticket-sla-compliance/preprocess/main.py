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
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    cur.execute("DELETE FROM email.attachments")
    # NOTE: Do NOT delete email.folders or email.account_config
    # These are needed by the email MCP to send mail. Clearing them breaks the tool.
    conn.commit()
    cur.close()
    conn.close()
    print("Email schema cleared.")


if __name__ == "__main__":
    main()
