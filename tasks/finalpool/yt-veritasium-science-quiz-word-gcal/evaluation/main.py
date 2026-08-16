"""
Evaluation for yt-veritasium-science-quiz-word-gcal task.

Checks:
1. Veritasium_Science_Quiz.docx exists with title and >=8 video sections (each Heading 1)
2. GCal has 4 "Science Study Session" events on Saturdays Mar 14/21/28, Apr 4 2026 10:00-12:00
3. GSheet "Veritasium Study Tracker" with Videos (8-10 rows) and Questions (>=24 rows)
4. Email to studygroup@university.edu with subject + body referencing quiz/sessions/tracker
"""
import os
import sys
import json
from argparse import ArgumentParser

import psycopg2
from zoneinfo import ZoneInfo

# ──────────────────────────────────────────────────────────────────────────
# EVALUATION GROUND TRUTH SPEC (gcal tz root-fix v3, case-study 2026-08-13)
# Source of truth: docs/task.md line 11: "Saturdays at 10:00 to 12:00" for the
# four weekly Science Study Session events (Mar 14/21/28, Apr 4 2026). The
# Calendar MCP writes events as UTC instants into the DB (TIMESTAMPTZ), so all
# comparisons anchor to America/New_York. Never use bare `st.hour` /
# `st.date()` on rows read from `gcal.events`: psycopg2 returns them in the
# PG session TimeZone (case-study: compute node default was Asia/Shanghai,
# masking ET wall-clock). Use `utils.evaluation.gcal_helpers` instead
# (session-tz-independent, DST-aware).
# ──────────────────────────────────────────────────────────────────────────
EXPECTED_TIMEZONE = ZoneInfo("America/New_York")
from utils.evaluation.gcal_helpers import get_zone_components  # noqa: E402

PASS_COUNT = 0
FAIL_COUNT = 0

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_word_doc(agent_workspace):
    print("\n=== Check 1-3: Veritasium_Science_Quiz.docx ===")
    docx_path = os.path.join(agent_workspace, "Veritasium_Science_Quiz.docx")
    if not os.path.exists(docx_path):
        record("Veritasium_Science_Quiz.docx exists", False, f"Not found at {docx_path}")
        return
    record("Veritasium_Science_Quiz.docx exists", True)

    try:
        import docx
        doc = docx.Document(docx_path)
    except Exception as e:
        record("Word doc readable", False, str(e))
        return

    paras = doc.paragraphs
    headings_l1 = [p.text for p in paras if p.style and p.style.name == "Heading 1"]
    all_text = "\n".join(p.text for p in paras)
    text_lower = all_text.lower()

    # Title at top
    record("Word doc has title 'Veritasium Science Quiz - Q1 2026'",
           "veritasium science quiz" in text_lower and "q1 2026" in text_lower,
           "Title not found")

    # 8-10 video headings (Heading 1). Title might be Title style (not H1) or extra H1.
    # Allow 8-12 total (videos + optional notes section).
    record("Word doc has 8-12 Heading 1 paragraphs (video sections)",
           8 <= len(headings_l1) <= 12,
           f"Found {len(headings_l1)} Heading 1 paragraphs")

    # Scientific keywords (>=5 distinct)
    science_keywords = ["physics", "math", "science", "engineering", "biology",
                        "chemistry", "quantum", "gravity", "speed", "light",
                        "energy", "wave", "probability", "paradox", "experiment",
                        "atom", "particle", "electron", "force", "motion"]
    found = [kw for kw in science_keywords if kw in text_lower]
    record("Word doc contains >=5 distinct scientific keywords",
           len(found) >= 5,
           f"Found: {found}")

    # 3 questions per video — count Q1:/Q2:/Q3: prefixes (structured)
    import re
    q_prefix_count = len(re.findall(r"\bQ[123]\s*:", all_text, re.IGNORECASE))
    expected_min_q = 8 * 3  # at least 8 videos * 3 questions
    record(f"Word doc has >= {expected_min_q} structured questions (Q1:/Q2:/Q3:)",
           q_prefix_count >= expected_min_q,
           f"Found {q_prefix_count} Qn: prefixes")

    # Difficulty rating mentioned (1-5 scale)
    has_difficulty = any(f"difficulty" in p.text.lower() or f"rating" in p.text.lower() for p in paras) or "difficulty" in text_lower
    record("Word doc mentions difficulty/rating",
           has_difficulty,
           "Not found")


def check_gcal():
    print("\n=== Check 4: GCal Science Study Sessions ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        # Read raw datetimes (with tz) — accept either UTC 10:00 or stored-tz 10:00
        cur.execute("""
            SELECT summary, start_datetime, end_datetime
            FROM gcal.events
            WHERE LOWER(summary) LIKE '%science study session%'
              AND start_datetime::date >= '2026-03-14'
              AND start_datetime::date <= '2026-04-04'
            ORDER BY start_datetime
        """)
        events = cur.fetchall()
        cur.close()
        conn.close()

        record("GCal has 4 'Science Study Session' events (Mar 14, 21, 28, Apr 4)",
               len(events) == 4,
               f"Found {len(events)}: {[(e[0], str(e[1])) for e in events]}")

        if len(events) >= 1:
            expected_dates = {"2026-03-14", "2026-03-21", "2026-03-28", "2026-04-04"}
            actual_dates = set()
            for summary, st, et in events:
                if st:
                    # gcal.events.start_datetime is TIMESTAMPTZ (UTC instant).
                    # Use the session-tz-independent helper to extract the date
                    # in EXPECTED_TIMEZONE (America/New_York per task.md).
                    # Bare st.date() silently uses the PG session tz.
                    ev_date, _, _ = get_zone_components(st, EXPECTED_TIMEZONE)
                    if ev_date is not None:
                        actual_dates.add(ev_date.isoformat())
                    # Each event 2 hours. Duration is an instant delta, so it
                    # is session-tz-independent (no helper needed).
                    if et:
                        duration = (et - st).total_seconds() / 60
                        record(f"Event on {ev_date} duration is 120 min (10:00-12:00)",
                               abs(duration - 120) <= 5,
                               f"Duration: {duration} min")
            missing = expected_dates - actual_dates
            record("All 4 expected Saturday dates present",
                   not missing,
                   f"Missing: {sorted(missing)}")
    except Exception as e:
        record("GCal check", False, str(e))


def check_gsheet():
    print("\n=== Check 5: GSheet 'Veritasium Study Tracker' ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        all_ss = cur.fetchall()
        target_title = "veritasium study tracker"
        ss_id = None
        for sid, t in all_ss:
            if t and target_title in t.strip().lower():
                ss_id = sid
                break

        record(f"GSheet '{target_title}' exists",
               ss_id is not None,
               f"Titles: {[t for _, t in all_ss]}")
        if ss_id is None:
            cur.close(); conn.close()
            return

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
        subs = cur.fetchall()

        videos_id = None
        questions_id = None
        for sid, t in subs:
            tnorm = (t or "").strip().lower()
            if tnorm == "videos":
                videos_id = sid
            if tnorm == "questions":
                questions_id = sid

        record("Worksheet 'Videos' exists",
               videos_id is not None,
               f"Worksheets: {[t for _, t in subs]}")
        record("Worksheet 'Questions' exists",
               questions_id is not None,
               f"Worksheets: {[t for _, t in subs]}")

        # Videos sheet: header + 8-10 data rows
        if videos_id is not None:
            cur.execute("""SELECT row_index, col_index, COALESCE(formatted_value, value)
                           FROM gsheet.cells WHERE spreadsheet_id = %s AND sheet_id = %s""",
                        (ss_id, videos_id))
            cells = cur.fetchall()
            grid = {}
            for r, c, v in cells:
                grid.setdefault(r, {})[c] = v
            data_rows = sorted(r for r in grid if r > 0)
            record("Videos sheet has 8-10 data rows",
                   8 <= len(data_rows) <= 10,
                   f"Found {len(data_rows)}")
            # Validate header columns
            h = grid.get(0, {})
            h_cols = [str(h.get(i, "") or "").strip().lower().replace(" ", "_") for i in sorted(h.keys())]
            expected = ["video_id", "title", "topic", "duration_min", "difficulty", "study_status"]
            for i, eh in enumerate(expected):
                ah = h_cols[i] if i < len(h_cols) else "MISSING"
                record(f"Videos header[{i}] = '{eh}'", ah == eh, f"Got '{ah}'")

            # Each Study_Status should be 'Pending'
            status_col = expected.index("study_status")
            statuses = [str(grid[r].get(status_col, "")).strip().lower() for r in data_rows]
            n_pending = sum(1 for s in statuses if s == "pending")
            record("All Study_Status values are 'Pending'",
                   n_pending == len(data_rows),
                   f"Got {n_pending}/{len(data_rows)} pending")

        # Questions sheet: header + >=24 (8*3) rows
        if questions_id is not None:
            cur.execute("""SELECT row_index, col_index, COALESCE(formatted_value, value)
                           FROM gsheet.cells WHERE spreadsheet_id = %s AND sheet_id = %s""",
                        (ss_id, questions_id))
            cells = cur.fetchall()
            grid = {}
            for r, c, v in cells:
                grid.setdefault(r, {})[c] = v
            data_rows = sorted(r for r in grid if r > 0)
            record("Questions sheet has >= 24 data rows (8 videos * 3 questions)",
                   len(data_rows) >= 24,
                   f"Found {len(data_rows)}")
            # Header check
            h = grid.get(0, {})
            h_cols = [str(h.get(i, "") or "").strip().lower().replace(" ", "_") for i in sorted(h.keys())]
            expected = ["video_title", "question_number", "question_text", "difficulty"]
            for i, eh in enumerate(expected):
                ah = h_cols[i] if i < len(h_cols) else "MISSING"
                record(f"Questions header[{i}] = '{eh}'", ah == eh, f"Got '{ah}'")

        cur.close()
        conn.close()
    except Exception as e:
        record("GSheet check", False, str(e))


def check_email():
    print("\n=== Check 6: Email to studygroup@university.edu ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("""
            SELECT subject, to_addr, COALESCE(body_text, body_html, '')
            FROM email.messages
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()

        target_to = "studygroup@university.edu"
        match = None
        for subj, to_addr, body in rows:
            tos = []
            if isinstance(to_addr, list):
                tos = [str(t).lower() for t in to_addr]
            elif isinstance(to_addr, str):
                try:
                    p = json.loads(to_addr)
                    tos = [str(t).lower() for t in p] if isinstance(p, list) else [to_addr.lower()]
                except Exception:
                    tos = [to_addr.lower()]
            if any(target_to in t for t in tos):
                match = (subj, body)
                break

        record(f"Email to {target_to} exists", match is not None, "")
        if match:
            subj_lower = (match[0] or "").strip().lower()
            body_lower = (match[1] or "").lower()
            # Subject must reference quiz document AND study schedule/sessions
            record("Email subject references quiz AND study schedule/sessions",
                   "quiz" in subj_lower and ("study schedule" in subj_lower or "sessions" in subj_lower),
                   f"Subject: {match[0]}")
            # Body mentions tracker / spreadsheet / 4 sessions
            record("Email body mentions tracker / spreadsheet",
                   "tracker" in body_lower or "spreadsheet" in body_lower,
                   "")
            record("Email body mentions 4 study sessions",
                   "4" in body_lower or "four" in body_lower or "march" in body_lower,
                   "")
    except Exception as e:
        record("Email check", False, str(e))


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    print(f"Running evaluation for yt-veritasium-science-quiz-word-gcal")
    print(f"Agent workspace: {agent_workspace}")

    check_word_doc(agent_workspace)
    check_gcal()
    check_gsheet()
    check_email()

    all_passed = FAIL_COUNT == 0
    summary = f"Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}"
    print(f"\n{'='*40}")
    print(f"Result: {'PASS' if all_passed else 'FAIL'} - {summary}")

    if res_log_file:
        with open(res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "all_passed": all_passed}, f)

    return all_passed, summary


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    success, message = run_evaluation(
        args.agent_workspace, args.groundtruth_workspace,
        args.launch_time, args.res_log_file
    )
    print(message)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
