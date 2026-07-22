"""
Preprocess script for gsheet-pdf-financial-report task.

Snowflake is read-only, so no changes there.
This script clears writable schemas (gsheet) to ensure
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


def inject_noise_spreadsheets(cur):
    """Inject 3-4 noise spreadsheets with unrelated titles. Should remain untouched."""
    print("[preprocess] Injecting noise spreadsheets ...")
    noise = [
        ("noise-ss-1", "HR Headcount Tracker", "HR Overview", [["Department", "Count"], ["Eng", "45"], ["Sales", "22"]]),
        ("noise-ss-2", "Marketing Campaign Calendar Q1", "Campaigns", [["Campaign", "Start"], ["Winter Promo", "2026-01-15"]]),
        ("noise-ss-3", "Office Supplies Budget 2026", "Items", [["Item", "Cost"], ["Pens", "50"], ["Paper", "200"]]),
    ]
    next_sheet_id = 990001
    for ss_id, ss_title, sh_title, cells in noise:
        cur.execute("INSERT INTO gsheet.spreadsheets (id, title) VALUES (%s, %s)", (ss_id, ss_title))
        cur.execute(
            "INSERT INTO gsheet.sheets (id, spreadsheet_id, title, index, row_count, column_count) VALUES (%s, %s, %s, %s, %s, %s)",
            (next_sheet_id, ss_id, sh_title, 0, max(len(cells), 5), max((len(cells[0]) if cells else 2), 2)),
        )
        for ri, row in enumerate(cells):
            for ci, v in enumerate(row):
                cur.execute(
                    "INSERT INTO gsheet.cells (spreadsheet_id, sheet_id, row_index, col_index, value) VALUES (%s, %s, %s, %s, %s)",
                    (ss_id, next_sheet_id, ri, ci, str(v)),
                )
        next_sheet_id += 1
    print("[preprocess] Noise spreadsheets injected (3).")


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
        inject_noise_spreadsheets(cur)
        print("[preprocess] Done. Writable schemas cleared and noise injected.")
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
