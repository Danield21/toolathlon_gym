"""
Preprocess script for canvas-quiz-analysis-email task.

Canvas is read-only. This script clears email tables,
injects a small amount of noise email into the coordinator inbox so the
agent must filter to the correct recipient/subject, and removes any stale
Quiz_Performance.xlsx in the workspace for idempotency.
"""
import json
import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        print("[preprocess] Clearing email data...")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        cur.execute("DELETE FROM email.drafts")
        cur.execute("DELETE FROM email.folders")
        print("[preprocess] Email data cleared.")

        # Re-insert standard email folders
        print("[preprocess] Inserting email folders...")
        cur.execute("""
            INSERT INTO email.folders (id, name, delimiter, flags, message_count, unread_count)
            VALUES (1, 'INBOX', '/', '{}', 0, 0),
                   (2, 'Sent', '/', '{}', 0, 0),
                   (3, 'Drafts', '/', '{}', 0, 0),
                   (4, 'Trash', '/', '{}', 0, 0)
            ON CONFLICT (id) DO NOTHING
        """)
        print("[preprocess] Email folders inserted.")

        # Inject unrelated noise emails into coordinator@university.edu inbox so the
        # agent must filter to produce the correct report email (not forwarding noise).
        noise = [
            ("Weekly All-Hands Agenda", "staff-comm@university.edu",
             "coordinator@university.edu",
             "Agenda for the weekly all-hands meeting attached."),
            ("IT Security Reminder", "it-security@university.edu",
             "coordinator@university.edu",
             "Reminder: rotate your password every 90 days."),
            ("Cafeteria Menu Update", "dining@university.edu",
             "coordinator@university.edu",
             "New seasonal menu available next week."),
        ]
        for subj, sender, recipient, body in noise:
            cur.execute(
                "INSERT INTO email.messages "
                "(folder_id, subject, from_addr, to_addr, body_text, created_at) "
                "VALUES (1, %s, %s, %s::jsonb, %s, NOW())",
                (subj, sender, json.dumps([recipient]), body),
            )
        print(f"[preprocess] Injected {len(noise)} noise emails.")

        conn.commit()

        # Remove any stale Quiz_Performance.xlsx from workspace for idempotency
        if args.agent_workspace:
            stale = os.path.join(args.agent_workspace, "Quiz_Performance.xlsx")
            if os.path.isfile(stale):
                try:
                    os.remove(stale)
                    print(f"[preprocess] Removed stale {stale}")
                except Exception as e:
                    print(f"[preprocess] Could not remove {stale}: {e}")

        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
