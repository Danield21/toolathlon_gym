"""
Preprocess script for quarterly-sales-restock-review task.

Prerequisites:
- PostgreSQL running at localhost:5432 with toolathlon_gym database
- Snowflake (sf/sf_data) and WooCommerce (wc) schemas are read-only and pre-populated
- Email and Notion schemas are writable

This script:
1. Clears email data (messages, attachments, sent_log)
2. Clears notion data (blocks, comments, pages, databases, users)
"""

import os
import argparse
import json
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_emails(cur):
    """Clear all email data except folder structure."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def clear_notion(cur):
    """Clear all notion data."""
    print("[preprocess] Clearing notion data...")
    cur.execute("DELETE FROM notion.comments")
    cur.execute("DELETE FROM notion.blocks")
    cur.execute("DELETE FROM notion.pages")
    cur.execute("DELETE FROM notion.databases")
    cur.execute("DELETE FROM notion.users")
    print("[preprocess] Notion data cleared.")


def inject_noise_emails(cur):
    """Insert unrelated inbound emails to test agent filtering."""
    print("[preprocess] Injecting noise emails...")
    noise = [
        (
            "Q3 Marketing Campaign Metrics",
            "marketing-analytics@mcp.com",
            '["ops-manager@mcp.com"]',
            "FYI on Q3 metrics - no action needed.",
        ),
        (
            "Reminder: Annual Security Training",
            "security@mcp.com",
            '["ops-manager@mcp.com"]',
            "Please complete the annual security training module by month end.",
        ),
        (
            "Supplier Trade Show Invitation",
            "events@partners.com",
            '["ops-manager@mcp.com"]',
            "You are invited to the annual trade show. Unrelated to restock requests.",
        ),
    ]
    for subj, from_addr, to_addr, body in noise:
        cur.execute(
            "INSERT INTO email.messages (folder_id, subject, from_addr, to_addr, body_text) "
            "VALUES (%s, %s, %s, %s::jsonb, %s)",
            (3100, subj, from_addr, to_addr, body),
        )
    print(f"[preprocess] Injected {len(noise)} noise emails.")


def inject_noise_notion(cur):
    """Insert unrelated Notion pages to test agent filtering.

    Only inserts pages (no blocks) - simple noise pages unrelated to Q4 review.
    """
    print("[preprocess] Injecting noise Notion pages...")
    import uuid

    def make_title_prop(text):
        return json.dumps({
            "title": {
                "title": [
                    {"type": "text", "text": {"content": text}}
                ]
            }
        })

    noise_titles = [
        "Team Offsite Planning 2026",
        "Onboarding Docs (Draft)",
        "Weekly Standup Notes - Week 10",
    ]
    for t in noise_titles:
        cur.execute(
            "INSERT INTO notion.pages (id, properties, archived) VALUES (%s, %s::jsonb, false)",
            (str(uuid.uuid4()), make_title_prop(t)),
        )
    print(f"[preprocess] Injected {len(noise_titles)} noise Notion pages.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_emails(cur)
        clear_notion(cur)
        inject_noise_emails(cur)
        inject_noise_notion(cur)
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
