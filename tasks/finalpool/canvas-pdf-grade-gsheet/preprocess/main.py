"""
Preprocess script for canvas-pdf-grade-gsheet task.

Canvas is read-only, so no changes there.
This script:
1. Clears email data (messages, attachments, sent_log, drafts)
2. Clears Google Sheet data (cells, sheets, spreadsheets, folders, permissions)
3. Injects noise emails and noise gsheet to exercise agent filtering.
4. Removes any stale semester_grade_report.xlsx from agent_workspace for idempotency.
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


def clear_emails(cur):
    """Clear all email data except folder structure and account config."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def clear_gsheet(cur):
    """Clear all Google Sheet data."""
    print("[preprocess] Clearing Google Sheet data...")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.permissions")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.spreadsheets")
    cur.execute("DELETE FROM gsheet.folders")
    print("[preprocess] Google Sheet data cleared.")


def inject_noise_emails(cur):
    """Inject unrelated noise emails so agent must filter to the right recipients/subjects."""
    print("[preprocess] Injecting noise emails...")
    # Look up INBOX folder
    cur.execute("SELECT id FROM email.folders WHERE name='INBOX' LIMIT 1")
    row = cur.fetchone()
    if not row:
        print("[preprocess] INBOX folder not found — skipping noise emails")
        return
    inbox_id = row[0]

    noise = [
        ("IT Maintenance Window Scheduled", "it-ops@openuniversity.ac.uk",
         "registrar@openuniversity.ac.uk",
         "Campus IT will perform scheduled maintenance this weekend."),
        ("Library Renewal Notice", "library@openuniversity.ac.uk",
         "registrar@openuniversity.ac.uk",
         "Your borrowed items are due next Monday."),
        ("Fall 2014 Enrollment Reminder", "admissions@openuniversity.ac.uk",
         "registrar@openuniversity.ac.uk",
         "Fall 2014 registration opens in three weeks."),
    ]
    for subj, sender, recipient, body in noise:
        cur.execute(
            "INSERT INTO email.messages "
            "(folder_id, subject, from_addr, to_addr, body_text, created_at) "
            "VALUES (%s, %s, %s, %s::jsonb, %s, NOW())",
            (inbox_id, subj, sender, json.dumps([recipient]), body),
        )
    print(f"[preprocess] Injected {len(noise)} noise emails.")


def inject_noise_gsheet(cur):
    """Inject an unrelated gsheet so agent must not accidentally target it."""
    print("[preprocess] Injecting noise gsheet...")
    cur.execute(
        "INSERT INTO gsheet.spreadsheets (id, title, created_at, updated_at) "
        "VALUES (%s, %s, NOW(), NOW()) ON CONFLICT (id) DO NOTHING",
        ("noise_ss_qa_01", "Last Quarter Course Registration Report"),
    )
    print("[preprocess] Injected 1 noise spreadsheet.")


def cleanup_workspace(agent_workspace):
    """Remove any stale semester_grade_report.xlsx to guarantee idempotency."""
    if not agent_workspace:
        return
    stale = os.path.join(agent_workspace, "semester_grade_report.xlsx")
    if os.path.isfile(stale):
        try:
            os.remove(stale)
            print(f"[preprocess] Removed stale {stale}")
        except Exception as e:
            print(f"[preprocess] Could not remove {stale}: {e}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_emails(cur)
        clear_gsheet(cur)
        inject_noise_emails(cur)
        inject_noise_gsheet(cur)
        conn.commit()
        cleanup_workspace(args.agent_workspace)
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
