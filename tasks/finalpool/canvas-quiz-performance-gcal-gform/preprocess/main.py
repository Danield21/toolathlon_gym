"""
Preprocess script for canvas-quiz-performance-gcal-gform task.

Canvas is read-only, so no changes there.
This script:
1. Exports the quiz submission data for the Spring 2014 Creative Computing &
   Culture course (course_id=8) into quiz_submissions.csv in the agent workspace
   (so the agent always has the data file, even if initial_workspace sync is skipped).
2. Clears Google Forms data
3. Clears Google Calendar events
4. Clears email data
"""

import csv
import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
}

QUIZ_COURSE_ID = 8  # Creative Computing & Culture (Spring 2014)


def export_quiz_submissions(cur, agent_workspace):
    """Write quiz_submissions.csv for the target course into the agent workspace."""
    if not agent_workspace:
        print("[preprocess] No agent_workspace provided; skipping quiz_submissions.csv export.")
        return
    os.makedirs(agent_workspace, exist_ok=True)
    # Quiz metadata for the course (id, title, points_possible).
    cur.execute(
        "SELECT id, title, points_possible FROM canvas.quizzes "
        "WHERE course_id = %s ORDER BY id",
        (QUIZ_COURSE_ID,),
    )
    quiz_rows = cur.fetchall()
    if not quiz_rows:
        print("[preprocess] WARNING: no quizzes found for course_id=%s; "
              "quiz_submissions.csv will be empty." % QUIZ_COURSE_ID)
        quiz_rows = []
    meta = {}
    for qid, title, points in quiz_rows:
        meta[int(qid)] = (str(title).strip(), float(points) if points is not None else 0.0)

    out_path = os.path.join(agent_workspace, "quiz_submissions.csv")
    written = 0
    with open(out_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["quiz_title", "points_possible", "score"])
        if meta:
            cur.execute(
                "SELECT quiz_id, score FROM canvas.quiz_submissions "
                "WHERE quiz_id = ANY(%s) ORDER BY quiz_id",
                (list(meta.keys()),),
            )
            for quiz_id, score in cur.fetchall():
                if quiz_id is None or score is None:
                    continue
                title, points = meta.get(int(quiz_id), (str(quiz_id), 0.0))
                w.writerow([title, points, float(score)])
                written += 1
    print(f"[preprocess] Exported {written} quiz submission rows -> {out_path}")


def clear_gform(cur):
    print("[preprocess] Clearing Google Forms data...")
    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    print("[preprocess] Google Forms data cleared.")


def clear_gcal(cur):
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")


def clear_emails(cur):
    print("[preprocess] Clearing email data...")
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    cur.execute("DELETE FROM email.drafts")
    print("[preprocess] Email data cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        export_quiz_submissions(cur, args.agent_workspace)
        clear_gform(cur)
        clear_gcal(cur)
        clear_emails(cur)
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
