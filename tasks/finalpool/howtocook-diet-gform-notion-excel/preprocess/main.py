"""
Preprocess for howtocook-diet-gform-notion-excel task.
- Clears gform, notion, and email schemas so agent starts fresh.
"""
import os
import argparse
import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
}


def clear_schemas(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gform.questions")
        cur.execute("DELETE FROM gform.forms")
        # notion.blocks has no FK to notion.pages, so deleting pages alone would
        # orphan blocks; delete blocks first to keep runs clean.
        cur.execute("DELETE FROM notion.blocks")
        cur.execute("DELETE FROM notion.pages")
        # email.sent_log has a FK (sent_log_message_id_fkey, NO ON DELETE
        # CASCADE) pointing at email.messages; email.attachments cascades but we
        # clear it too for hygiene. Order matters: referencing rows must be
        # deleted BEFORE messages, otherwise a fresh seed (which ships one
        # sent_log row pointing at a message) makes DELETE FROM email.messages
        # raise a ForeignKeyViolation and abort preprocess before the model runs.
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
    conn.commit()
    print("[preprocess] Cleared gform, notion, and email schemas")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_schemas(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
