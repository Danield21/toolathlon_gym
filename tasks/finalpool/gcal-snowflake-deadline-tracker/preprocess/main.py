"""
Preprocess script for gcal-snowflake-deadline-tracker task.

Snowflake data is read-only. This script:
1. Clears Google Calendar events
2. Clears email data
3. Clears Notion data
"""

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


def clear_gcal(cur):
    """Clear all Google Calendar events."""
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")


def clear_emails(cur):
    """Clear all email data."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def clear_notion(cur):
    """Clear all Notion data."""
    print("[preprocess] Clearing Notion data...")
    cur.execute("DELETE FROM notion.comments")
    cur.execute("DELETE FROM notion.blocks")
    cur.execute("DELETE FROM notion.pages")
    cur.execute("DELETE FROM notion.databases")
    cur.execute("DELETE FROM notion.users")
    print("[preprocess] Notion data cleared.")


def clear_gsheet(cur):
    """Clear Google Sheets data."""
    print("[preprocess] Clearing Google Sheets data...")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.sheets")
    try:
        cur.execute("DELETE FROM gsheet.permissions")
    except Exception:
        pass
    cur.execute("DELETE FROM gsheet.spreadsheets")
    print("[preprocess] Google Sheets data cleared.")


def inject_noise(cur):
    """Inject noise notion pages, gcal events, and emails (3-5 each) so agent must filter."""
    print("[preprocess] Injecting noise notion/gcal/email ...")
    # Noise gcal events (different topics, should not be deleted)
    noise_events = [
        ("noise-gcal-1", "Marketing Planning Sync", "Unrelated to SLA or support.",
         "2026-03-06 10:00:00+00", "2026-03-06 11:00:00+00"),
        ("noise-gcal-2", "All Hands Meeting", "Company update, no ticket content.",
         "2026-03-08 15:00:00+00", "2026-03-08 16:00:00+00"),
        ("noise-gcal-3", "Lunch with Vendor", "Friendly meal, no deliverables.",
         "2026-03-05 12:00:00+00", "2026-03-05 13:30:00+00"),
    ]
    for ev in noise_events:
        cur.execute(
            """INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status)
               VALUES (%s, %s, %s, %s, %s, 'confirmed')""",
            ev,
        )

    # Noise notion pages (unrelated titles)
    noise_pages = [
        ("noise-page-1", '{"title":{"title":[{"plain_text":"Q1 OKR Retrospective Notes"}]}}'),
        ("noise-page-2", '{"title":{"title":[{"plain_text":"Office Move Checklist 2026"}]}}'),
        ("noise-page-3", '{"title":{"title":[{"plain_text":"Hiring Pipeline Snapshot"}]}}'),
    ]
    for pid, props in noise_pages:
        cur.execute(
            """INSERT INTO notion.pages (id, object, archived, in_trash, properties)
               VALUES (%s, 'page', false, false, %s::jsonb)""",
            (pid, props),
        )

    # Noise emails in INBOX (unrelated recipients/subjects, should remain untouched)
    cur.execute("SELECT id FROM email.folders WHERE name IN ('INBOX','Inbox') ORDER BY id LIMIT 1")
    row = cur.fetchone()
    if row:
        inbox_id = row[0]
        noise_emails = [
            ("noise-email-1", "FYI: Coffee machine repair", "facilities@company.com",
             '["team@company.com"]', "The repair window is scheduled for this Friday."),
            ("noise-email-2", "Reminder: Security training due", "security@company.com",
             '["staff@company.com"]', "Please complete training module 3 by end of month."),
            ("noise-email-3", "Newsletter: Engineering blog roundup", "eng-news@company.com",
             '["subscribers@company.com"]', "Top posts of the week - unrelated to SLA."),
        ]
        for (mid, subj, frm, to_j, body) in noise_emails:
            cur.execute(
                """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                   VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s, false)""",
                (inbox_id, mid, subj, frm, to_j, body),
            )
    print("[preprocess] Noise injected (3 gcal + 3 notion + 3 email).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
        clear_emails(cur)
        clear_notion(cur)
        clear_gsheet(cur)
        inject_noise(cur)
        conn.commit()
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
