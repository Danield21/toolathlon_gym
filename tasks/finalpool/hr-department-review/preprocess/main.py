"""
Preprocess script for hr-department-review task.

Clears email data (attachments, sent_log, messages) in FK order.
Snowflake HR_ANALYTICS is read-only and not touched.
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("Cleared email data")

        # Inject noise emails (unrelated inbound emails) to test filtering
        # Place them into Inbox (id=3100) so they don't interfere with sent detection.
        noise_emails = [
            {
                "subject": "Quarterly Town Hall Invitation",
                "from_addr": "ceo@company.com",
                "to_addr": '["all-staff@company.com"]',
                "body_text": (
                    "Hello team, please join us for the quarterly town hall next "
                    "Friday. We will discuss updates across sales, marketing, and "
                    "product. Agenda attached."
                ),
            },
            {
                "subject": "Office Maintenance Notice",
                "from_addr": "facilities@company.com",
                "to_addr": '["all-staff@company.com"]',
                "body_text": (
                    "The HVAC system on the 4th floor will be serviced Saturday. "
                    "Please remove any perishable items from your desks."
                ),
            },
            {
                "subject": "Holiday Schedule Reminder",
                "from_addr": "hr-announcements@company.com",
                "to_addr": '["all-staff@company.com"]',
                "body_text": (
                    "Reminder: the office will be closed on observed holidays "
                    "next month. No department review actions are required here."
                ),
            },
            {
                "subject": "Security Training Module Available",
                "from_addr": "security@company.com",
                "to_addr": '["all-staff@company.com"]',
                "body_text": (
                    "Please complete the new annual security training in the LMS. "
                    "Not related to department performance review."
                ),
            },
        ]
        for e in noise_emails:
            cur.execute(
                """
                INSERT INTO email.messages (folder_id, subject, from_addr, to_addr, body_text)
                VALUES (%s, %s, %s, %s::jsonb, %s)
                """,
                (3100, e["subject"], e["from_addr"], e["to_addr"], e["body_text"]),
            )
        conn.commit()
        print(f"Injected {len(noise_emails)} noise emails into Inbox")
    except Exception as e:
        conn.rollback()
        print(f"Error clearing email data: {e}")
        raise
    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    main()
