"""
Preprocess script for gsheet-woocommerce-inventory task.

WooCommerce and Yahoo Finance are read-only, so no changes there.
This script clears writable schemas (gsheet, email) to ensure
a clean environment for the agent.
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


def clear_gsheet(cur):
    """Clear all Google Sheets data, respecting FK constraints."""
    print("[preprocess] Clearing Google Sheets data...")
    cur.execute("DELETE FROM gsheet.cells")
    cur.execute("DELETE FROM gsheet.permissions")
    cur.execute("DELETE FROM gsheet.sheets")
    cur.execute("DELETE FROM gsheet.spreadsheets")
    print("[preprocess] Cleared Google Sheets data.")


def clear_email(cur):
    """Clear email data."""
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Cleared email data.")


def inject_noise(cur):
    """Inject noise emails (non-warehouse topics) and noise spreadsheets (not Inventory Dashboard)."""
    print("[preprocess] Injecting noise emails and spreadsheets ...")
    # Noise emails - should NOT be deleted and should NOT be confused with low stock alerts
    cur.execute("SELECT id FROM email.folders WHERE name IN ('INBOX','Inbox') ORDER BY id LIMIT 1")
    row = cur.fetchone()
    if row:
        inbox_id = row[0]
        noise_emails = [
            ("noise-email-w-1", "Monthly All Hands Agenda", "ceo@company.com",
             '["staff@company.com"]', "Agenda for the monthly company meeting - unrelated to inventory."),
            ("noise-email-w-2", "Office holiday schedule", "hr@company.com",
             '["staff@company.com"]', "Upcoming office closures - please plan accordingly."),
            ("noise-email-w-3", "Cafeteria menu update", "kitchen@company.com",
             '["staff@company.com"]', "New lunch items added - enjoy!"),
            ("noise-email-w-4", "Vendor newsletter", "vendor@third-party.com",
             '["accounts@company.com"]', "General vendor update - not low stock related."),
        ]
        for (mid, subj, frm, to_j, body) in noise_emails:
            cur.execute(
                """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                   VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s, false)""",
                (inbox_id, mid, subj, frm, to_j, body),
            )
    # Noise spreadsheets (without 'inventory' in title)
    noise_ss = [
        ("noise-ss-w-1", "2026 Holiday Calendar", "Days", [["Date", "Holiday"], ["2026-01-01", "New Year"]]),
        ("noise-ss-w-2", "Team Off-site Agenda", "Topics", [["Time", "Topic"], ["9 AM", "Kickoff"]]),
        ("noise-ss-w-3", "Weekly Engineering Standup Notes", "Notes", [["Sprint", "Lead"], ["S1", "Alex"]]),
    ]
    nid = 991001
    for (ss_id, ss_title, sh_title, cells) in noise_ss:
        cur.execute("INSERT INTO gsheet.spreadsheets (id, title) VALUES (%s, %s)", (ss_id, ss_title))
        cur.execute(
            "INSERT INTO gsheet.sheets (id, spreadsheet_id, title, index, row_count, column_count) VALUES (%s, %s, %s, %s, %s, %s)",
            (nid, ss_id, sh_title, 0, max(len(cells), 3), max(len(cells[0]) if cells else 2, 2)),
        )
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                cur.execute(
                    "INSERT INTO gsheet.cells (spreadsheet_id, sheet_id, row_index, col_index, value) VALUES (%s, %s, %s, %s, %s)",
                    (ss_id, nid, ri, ci, str(v)),
                )
        nid += 1
    print("[preprocess] Noise injected (4 emails + 3 gsheets).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--launch_time", required=False, help="Launch time")
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = True
    cur = conn.cursor()

    try:
        clear_gsheet(cur)
        clear_email(cur)
        inject_noise(cur)
        print("[preprocess] Done. Writable schemas cleared and noise injected.")
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
