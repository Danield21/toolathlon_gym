"""
Preprocess for ppt-woocommerce-sales-review task.

Clears email schema data so the evaluation can verify only emails
sent during the task. WooCommerce and Yahoo Finance are read-only.
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


def clear_email_data():
    """Delete all rows from email tables so evaluation starts clean."""
    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = True
    cur = conn.cursor()

    # Order matters due to foreign keys
    cur.execute("DELETE FROM email.attachments;")
    cur.execute("DELETE FROM email.sent_log;")
    cur.execute("DELETE FROM email.drafts;")
    cur.execute("DELETE FROM email.messages;")

    # Reset message counts on folders
    cur.execute("UPDATE email.folders SET message_count = 0, unread_count = 0;")

    # Inject noise emails in INBOX to test filtering.
    # Look up the actual INBOX folder id from the table (don't hardcode).
    cur.execute("SELECT id FROM email.folders WHERE name ILIKE 'INBOX' LIMIT 1")
    inbox_row = cur.fetchone()
    if not inbox_row:
        # Fallback: try any folder
        cur.execute("SELECT id FROM email.folders ORDER BY id LIMIT 1")
        inbox_row = cur.fetchone()
    inbox_id = inbox_row[0] if inbox_row else 1

    # Subject-level filter in evaluation is '%monthly sales review%', so these
    # should not match that filter.
    noise_emails = [
        (
            "Weekly Supplier Shipment Update",
            "supplier@partners.com",
            '["operations@company.com"]',
            "Shipment #4821 departed on schedule. No delays expected.",
        ),
        (
            "IT Notice: VPN Maintenance Window",
            "it-ops@company.com",
            '["all-staff@company.com"]',
            "VPN will be unavailable Sunday 01:00-03:00 UTC for updates.",
        ),
        (
            "Customer Feedback Summary - Unrelated",
            "cx-analytics@company.com",
            '["cx-team@company.com"]',
            "Top complaints this week: slow packaging, delivery timing.",
        ),
    ]
    for subj, from_addr, to_addr, body in noise_emails:
        cur.execute(
            "INSERT INTO email.messages (folder_id, subject, from_addr, to_addr, body_text) "
            "VALUES (%s, %s, %s, %s::jsonb, %s)",
            (inbox_id, subj, from_addr, to_addr, body),
        )

    conn.close()
    print(f"Cleared email data. Injected {len(noise_emails)} noise emails.")


def cleanup_agent_workspace(agent_workspace):
    """Remove stale outputs from previous runs in the agent workspace."""
    if not agent_workspace:
        return
    targets = [
        "Monthly_Sales_Review.xlsx",
        "Sales_Review_Presentation.pptx",
    ]
    import pathlib
    ws = pathlib.Path(agent_workspace)
    if not ws.exists():
        return
    for name in targets:
        p = ws / name
        if p.exists():
            try:
                p.unlink()
                print(f"Removed stale {p}")
            except Exception as ex:
                print(f"Could not remove {p}: {ex}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_email_data()
    cleanup_agent_workspace(args.agent_workspace)
    print("Preprocess complete.")


if __name__ == "__main__":
    main()
