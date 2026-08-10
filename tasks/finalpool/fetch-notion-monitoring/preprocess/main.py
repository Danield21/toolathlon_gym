"""
Preprocess script for fetch-notion-monitoring task.

This script:
1. Clears Notion, Google Calendar, and email data
2. Injects a small set of unrelated Notion pages and Google Calendar events so the
   agent must filter to the correct service/incident artifacts.
"""

import argparse
import asyncio
import json
import os
import uuid

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_notion(cur):
    """Clear all Notion data."""
    print("[preprocess] Clearing Notion data...")
    cur.execute("DELETE FROM notion.comments")
    cur.execute("DELETE FROM notion.blocks")
    cur.execute("DELETE FROM notion.pages")
    cur.execute("DELETE FROM notion.databases")
    cur.execute("DELETE FROM notion.users")
    print("[preprocess] Notion data cleared.")


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
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] Email data cleared.")


def inject_noise_notion(cur):
    """Inject unrelated Notion pages (onboarding, benefits) for filtering test."""
    print("[preprocess] Injecting noise Notion pages...")
    noise_pages = [
        ("Onboarding Checklist", "General onboarding reference page."),
        ("Company Handbook", "Company policies and benefits overview."),
        ("Q1 OKR Draft", "Team OKRs draft for Q1 planning."),
    ]
    for title, desc in noise_pages:
        page_id = str(uuid.uuid4())
        props = {
            "title": {
                "title": [{"type": "text", "text": {"content": title},
                           "plain_text": title}]
            }
        }
        cur.execute(
            "INSERT INTO notion.pages "
            "(id, object, created_time, last_edited_time, archived, in_trash, properties, url) "
            "VALUES (%s, 'page', NOW(), NOW(), false, false, %s::jsonb, %s)",
            (page_id, json.dumps(props), f"https://www.notion.so/{page_id}"),
        )
    print(f"[preprocess] Injected {len(noise_pages)} noise Notion pages.")


def inject_noise_gcal(cur):
    """Inject unrelated calendar events so the agent must filter to the right ones."""
    print("[preprocess] Injecting noise Google Calendar events...")
    noise_events = [
        ("Weekly All-Hands", "Standard weekly sync",
         "2026-03-04 10:00:00+00", "2026-03-04 11:00:00+00"),
        ("Department Training Session", "Non-incident-related training",
         "2026-03-04 14:00:00+00", "2026-03-04 15:30:00+00"),
        ("Q1 Budget Review", "Finance review meeting",
         "2026-03-09 15:00:00+00", "2026-03-09 16:00:00+00"),
    ]
    for summary, desc, start, end in noise_events:
        event_id = str(uuid.uuid4())
        cur.execute(
            "INSERT INTO gcal.events "
            "(id, summary, description, start_datetime, end_datetime, status, created, updated) "
            "VALUES (%s, %s, %s, %s::timestamptz, %s::timestamptz, 'confirmed', NOW(), NOW())",
            (event_id, summary, desc, start, end),
        )
    print(f"[preprocess] Injected {len(noise_events)} noise calendar events.")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_notion(cur)
        clear_gcal(cur)
        clear_emails(cur)
        inject_noise_notion(cur)
        inject_noise_gcal(cur)
        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
