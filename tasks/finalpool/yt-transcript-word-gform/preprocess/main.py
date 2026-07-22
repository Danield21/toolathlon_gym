"""
Preprocess for yt-transcript-word-gform task.

Clears gform tables, email tables for clean state.
The Afrobeat transcript is already in youtube.transcripts (READ-ONLY).
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

FORM_TITLE = "Afrobeat Mix Feedback"


def clear_tables(conn):
    with conn.cursor() as cur:
        # Clear matching gform data - responses cascade from forms via FK
        cur.execute("""
            DELETE FROM gform.responses WHERE form_id IN (
                SELECT id FROM gform.forms WHERE title ILIKE %s
            )
        """, (f"%{FORM_TITLE}%",))
        cur.execute("""
            DELETE FROM gform.questions WHERE form_id IN (
                SELECT id FROM gform.forms WHERE title ILIKE %s
            )
        """, (f"%{FORM_TITLE}%",))
        cur.execute("DELETE FROM gform.forms WHERE title ILIKE %s",
                    (f"%{FORM_TITLE}%",))
        # Clear email tables (must clear sent_log before messages due to FK)
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        try:
            cur.execute("DELETE FROM email.drafts")
        except Exception:
            pass
    conn.commit()
    print("[preprocess] Cleared gform and email tables.")


def verify_transcript(conn):
    """Verify the Afrobeat transcript exists and is non-trivial (READ-ONLY check).

    Reverse validation: assert that the video ID referenced in task.md (7ZQzGq32kAY)
    resolves to a real row in youtube.transcripts, and that the transcript content
    has sufficient length to be analyzable. Ensures the task target is solvable.
    """
    with conn.cursor() as cur:
        cur.execute("""
            SELECT video_id, title, LENGTH(content) as content_len
            FROM youtube.transcripts
            WHERE video_id = '7ZQzGq32kAY'
        """)
        row = cur.fetchone()
        if not row:
            print("[preprocess] ERROR: Afrobeat transcript for 7ZQzGq32kAY not found in youtube.transcripts!")
            raise RuntimeError("Required transcript missing: 7ZQzGq32kAY")
        video_id, title, content_len = row
        assert video_id == '7ZQzGq32kAY', f"video_id mismatch: {video_id}"
        assert content_len and content_len > 100, f"transcript too short: length={content_len}"
        print(f"[preprocess] Transcript verified: video_id={video_id}, title={title}, length={content_len}")


def ensure_email_folder(conn):
    """Ensure INBOX and Sent folders exist so sent_log logic and outbound mail work."""
    with conn.cursor() as cur:
        for fname in ("INBOX", "Sent"):
            cur.execute("SELECT id FROM email.folders WHERE name = %s LIMIT 1", (fname,))
            row = cur.fetchone()
            if not row:
                cur.execute("INSERT INTO email.folders (name) VALUES (%s) RETURNING id", (fname,))
        conn.commit()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        clear_tables(conn)
        verify_transcript(conn)
        ensure_email_folder(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
