"""
Preprocess for fetch-yf-macro-portfolio-excel-gcal-email task.

Clears gcal and email data. Extracts mock_pages.tar.gz and starts HTTP server on port 30211.
"""
import argparse
import asyncio
import os
import shutil
import tarfile

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_db(cur):
    """Clear gcal events and email data."""
    print("[preprocess] Clearing gcal events and email data...")
    cur.execute("DELETE FROM gcal.events")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    try:
        cur.execute("DELETE FROM email.drafts")
    except Exception:
        pass
    print("[preprocess] DB cleared.")


def inject_noise(cur):
    """Inject noise gcal events and email messages to test filtering capability."""
    print("[preprocess] Injecting noise events/emails...")
    # Noise gcal events (different dates/topics, should NOT be removed)
    noise_events = [
        ("noise-evt-1", "Weekly Team Standup", "Regular team sync, no portfolio business.",
         "2026-03-25 09:00:00+00", "2026-03-25 09:30:00+00"),
        ("noise-evt-2", "Compliance Training 2026", "Annual training for all staff.",
         "2026-03-28 13:00:00+00", "2026-03-28 14:30:00+00"),
        ("noise-evt-3", "Client Dinner - Acme Corp", "Hosting Acme CFO, unrelated to holdings.",
         "2026-04-02 19:00:00+00", "2026-04-02 21:00:00+00"),
        ("noise-evt-4", "Office IT Maintenance", "Server reboot, no meetings this window.",
         "2026-03-31 22:00:00+00", "2026-03-31 23:00:00+00"),
    ]
    for ev in noise_events:
        cur.execute(
            """INSERT INTO gcal.events (id, summary, description, start_datetime, end_datetime, status)
               VALUES (%s, %s, %s, %s, %s, 'confirmed')""",
            ev,
        )
    # Noise emails in INBOX (different recipients/subjects, should NOT be sent again or removed)
    cur.execute("SELECT id FROM email.folders WHERE name='INBOX' OR name='Inbox' ORDER BY id LIMIT 1")
    row = cur.fetchone()
    inbox_id = row[0] if row else None
    if inbox_id is not None:
        noise_emails = [
            ("noise-msg-1", "Newsletter: Weekly Market Digest", "newsletter@marketdigest.com",
             '["subscriber@firm.com"]', "Generic market commentary, unrelated to our five holdings."),
            ("noise-msg-2", "Reminder: Expense Reports Due", "hr@firm.com",
             '["staff@firm.com"]', "Please submit Q1 expense reports by Friday."),
            ("noise-msg-3", "HR: Benefits Enrollment Window", "benefits@firm.com",
             '["staff@firm.com"]', "Open enrollment runs through end of month."),
            ("noise-msg-4", "Vendor Update: Bloomberg Terminal", "support@bloomberg.com",
             '["it@firm.com"]', "Scheduled maintenance notice."),
        ]
        for (mid, subj, frm, to_j, body) in noise_emails:
            cur.execute(
                """INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
                   VALUES (%s, %s, %s, %s, %s::jsonb, NOW(), %s, false)""",
                (inbox_id, mid, subj, frm, to_j, body),
            )
    print("[preprocess] Noise injected (4 gcal + 4 emails).")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30211."""
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30211

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory {mock_dir} "
        f"> {mock_dir}/server.log 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{port}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        clear_db(cur)
        inject_noise(cur)
        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] DB error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    await setup_mock_server()
    print("[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
