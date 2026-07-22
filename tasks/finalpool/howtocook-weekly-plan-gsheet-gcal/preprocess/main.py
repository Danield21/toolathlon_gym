"""
Preprocess for howtocook-weekly-plan-gsheet-gcal task.
- Clears gsheet, gcal, and email schemas so agent starts fresh.
"""
import os
import argparse
import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_schemas(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.spreadsheets")
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Cleared gsheet, gcal, and email schemas")


def inject_noise(conn):
    """Inject noise gcal events, emails, and spreadsheets for filter testing."""
    with conn.cursor() as cur:
        noise_events = [
            ("noise-gcal-wp-1", "Team Retrospective Q1", "Unrelated work meeting.",
             "2026-03-11 14:00:00+00", "2026-03-11 15:00:00+00"),
            ("noise-gcal-wp-2", "Dentist Appointment", "Personal appointment.",
             "2026-03-10 10:00:00+00", "2026-03-10 11:00:00+00"),
            ("noise-gcal-wp-3", "Book Club Meetup", "Monthly book discussion.",
             "2026-03-13 19:00:00+00", "2026-03-13 21:00:00+00"),
        ]
        for ev in noise_events:
            cur.execute(
                """INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status)
                   VALUES (%s, %s, %s, %s, %s, 'confirmed')""",
                ev,
            )
        # Noise emails
        cur.execute("SELECT id FROM email.folders WHERE name IN ('INBOX','Inbox') ORDER BY id LIMIT 1")
        r = cur.fetchone()
        if r:
            inbox_id = r[0]
            noise_emails = [
                ("noise-em-wp-1", "Utility bill reminder", "billing@utility.com",
                 '["resident@home.com"]', "Reminder for monthly bill."),
                ("noise-em-wp-2", "Gym schedule update", "gym@local.com",
                 '["member@home.com"]', "New class times for next week."),
                ("noise-em-wp-3", "Library book due", "library@city.gov",
                 '["reader@home.com"]', "Return by Friday."),
            ]
            for (mid, subj, frm, to_j, body) in noise_emails:
                cur.execute(
                    """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                       VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s, false)""",
                    (inbox_id, mid, subj, frm, to_j, body),
                )
        # Noise spreadsheet
        noise_ss = [
            ("noise-ss-wp-1", "Home Chore Schedule", "Chores", [["Task","Day"], ["Vacuum","Sat"]]),
            ("noise-ss-wp-2", "Kids Homework Tracker", "Work", [["Subject","Due"], ["Math","Mon"]]),
            ("noise-ss-wp-3", "Garden Watering Plan", "Plants", [["Plant","Water"], ["Rose","Daily"]]),
        ]
        nid = 992001
        for ss_id, ss_title, sh_title, cells in noise_ss:
            cur.execute("INSERT INTO gsheet.spreadsheets (id, title) VALUES (%s, %s)", (ss_id, ss_title))
            cur.execute(
                "INSERT INTO gsheet.sheets (id, spreadsheet_id, title, index, row_count, column_count) VALUES (%s, %s, %s, %s, %s, %s)",
                (nid, ss_id, sh_title, 0, 5, 2),
            )
            for ri, row in enumerate(cells):
                for ci, v in enumerate(row):
                    cur.execute(
                        "INSERT INTO gsheet.cells (spreadsheet_id, sheet_id, row_index, col_index, value) VALUES (%s, %s, %s, %s, %s)",
                        (ss_id, nid, ri, ci, str(v)),
                    )
            nid += 1
    conn.commit()
    print("[preprocess] Noise injected (3 gcal + 3 email + 3 gsheet).")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_schemas(conn)
        inject_noise(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
