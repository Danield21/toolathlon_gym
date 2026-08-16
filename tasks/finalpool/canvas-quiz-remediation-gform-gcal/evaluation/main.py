"""Evaluation for canvas-quiz-remediation-gform-gcal."""
import argparse
import json
import os
import re
import sys
from datetime import datetime, timedelta

import psycopg2

# ──────────────────────────────────────────────────────────────────────────
# EVALUATION GROUND TRUTH SPEC (gcal tz root-fix v3, case-study 2026-08-13)
# Source of truth: docs/task.md. gcal.events.start_datetime is TIMESTAMPTZ
# (UTC instant). psycopg2 returns it in the PG session TimeZone (compute node
# default Asia/Shanghai, masking 15:00 ET as 03:00+08 next day). Always use
# utils.evaluation.gcal_helpers; never bare start_dt.hour / .date().
# ──────────────────────────────────────────────────────────────────────────
# task.md line 7: "Each session should start at 3:00 PM" (canvas Fall 2013J
# finance course → America/New_York ET)
EXPECTED_TIMEZONE = "America/New_York"

from utils.evaluation.gcal_helpers import get_zone_components  # noqa: E402

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=int(os.environ.get("PGPORT", "5432")), dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0


def _compute_underperforming_quizzes():
    """Compute underperforming quizzes (avg score < 80) for course 16 dynamically.

    Falls back to a verified static list if DB query fails (so eval still runs).
    """
    fallback = ["CMA 34880", "CMA 34881", "CMA 34882", "CMA 34883", "CMA 34884"]
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute(
            "SELECT q.title FROM canvas.quizzes q "
            "JOIN canvas.quiz_submissions s ON s.quiz_id = q.id "
            "WHERE q.course_id = 16 "
            "GROUP BY q.title HAVING AVG(s.score) < 80 "
            "ORDER BY q.title"
        )
        rows = [r[0] for r in cur.fetchall() if r and r[0]]
        conn.close()
        if rows:
            return rows
    except Exception:
        pass
    return fallback


# Quizzes for course 16 with avg score below 80 (computed dynamically)
UNDERPERFORMING_QUIZZES = _compute_underperforming_quizzes()


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def check_gform():
    """Check Google Form creation."""
    print("\n=== Checking Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    check("At least one form created", len(forms) >= 1, f"Found {len(forms)} forms")

    # Match exact title 'Finance Quiz Self-Assessment' (case-insensitive, trim spaces)
    form_id = None
    matched_title = None
    for fid, title in forms:
        norm_title = (title or "").strip().lower()
        if norm_title == "finance quiz self-assessment":
            form_id = fid
            matched_title = title
            break

    check(
        "Form titled 'Finance Quiz Self-Assessment' found (exact match)",
        form_id is not None,
        f"Found titles: {[f[1] for f in forms]}",
    )

    if form_id:
        cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (form_id,))
        q_count = cur.fetchone()[0]
        check("Form has exactly 4 questions", q_count == 4,
              f"Found {q_count} questions")

    conn.close()


def _expected_session_dates(launch_time_str):
    """Return list of 5 expected datetimes (Mon-Fri 15:00) of the week after launch."""
    if not launch_time_str:
        return None
    try:
        # Accept multiple formats; we only need the date part
        dt = datetime.fromisoformat(launch_time_str.replace("Z", "+00:00"))
    except Exception:
        try:
            dt = datetime.strptime(launch_time_str.split("T")[0], "%Y-%m-%d")
        except Exception:
            return None
    # The Monday of the week following the launch
    # Find the next Monday strictly after the launch_date's week start
    launch_date = dt.date()
    # Next Monday after launch_date (i.e. Monday of NEXT week)
    days_until_next_monday = (7 - launch_date.weekday()) % 7
    if days_until_next_monday == 0:
        days_until_next_monday = 7
    next_monday = launch_date + timedelta(days=days_until_next_monday)
    return [next_monday + timedelta(days=i) for i in range(5)]


def check_gcal(launch_time_str=None):
    """Check remediation calendar events."""
    print("\n=== Checking Google Calendar Events ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, summary, start_datetime, end_datetime FROM gcal.events ORDER BY start_datetime")
    events = cur.fetchall()
    check("At least 1 calendar event created", len(events) >= 1,
          f"Found {len(events)} events")

    quiz_nums = [q.split()[-1] for q in UNDERPERFORMING_QUIZZES]
    # Tighten: event summary must contain 'remediation' AND at least one underperforming quiz number
    remediation_events = [
        e for e in events
        if "remediation" in (e[1] or "").lower()
        and any(qn in (e[1] or "") for qn in quiz_nums)
    ]
    check(
        "Exactly 5 remediation session events",
        len(remediation_events) == len(UNDERPERFORMING_QUIZZES),
        f"Expected {len(UNDERPERFORMING_QUIZZES)}, found {len(remediation_events)}: {[e[1] for e in remediation_events]}",
    )

    # Quiz titles must each appear in at least one event title
    for quiz_title in UNDERPERFORMING_QUIZZES:
        quiz_num = quiz_title.split()[-1]
        found = any(quiz_num in (e[1] or "") for e in remediation_events)
        check(
            f"Event for quiz {quiz_title} present",
            found,
            f"Looking for '{quiz_num}' in event summaries: {[e[1] for e in remediation_events]}",
        )

    # Each remediation event must start at 15:00 local and last 1 hour
    for ev_id, summary, start_dt, end_dt in remediation_events:
        if start_dt is None or end_dt is None:
            check(f"Event '{summary}' has start/end timestamps", False, f"start={start_dt} end={end_dt}")
            continue
        # gcal.events.start_datetime is TIMESTAMPTZ; convert to EXPECTED_TIMEZONE
        # via session-tz-independent helper. Bare start_dt.hour silently compares
        # against the PG session tz (case-study 2026-08-13).
        ev_date, ev_hour, ev_minute = get_zone_components(start_dt, EXPECTED_TIMEZONE)
        check(
            f"Event '{summary}' starts at 15:00",
            ev_hour == 15 and ev_minute == 0,
            f"start_datetime={start_dt}",
        )
        duration = end_dt - start_dt
        check(
            f"Event '{summary}' duration is exactly 1 hour",
            duration == timedelta(hours=1),
            f"duration={duration}",
        )

    # Scheduling: first event on next Monday after launch, sequential days
    expected_dates = _expected_session_dates(launch_time_str)
    if expected_dates is not None and len(remediation_events) >= 1:
        sorted_remed = sorted(remediation_events, key=lambda e: e[2])
        actual_dates = [get_zone_components(e[2], EXPECTED_TIMEZONE)[0] for e in sorted_remed[: len(expected_dates)]]
        # First session: Monday of week after launch
        check(
            f"First remediation session on Monday {expected_dates[0]}",
            len(actual_dates) >= 1 and actual_dates[0] == expected_dates[0],
            f"actual first date={actual_dates[0] if actual_dates else None}",
        )
        # Each subsequent event one day apart (Mon, Tue, Wed, Thu, Fri)
        check(
            "Remediation events scheduled on consecutive days starting Monday",
            actual_dates == expected_dates[: len(actual_dates)],
            f"expected={expected_dates}, actual={actual_dates}",
        )

    conn.close()


def check_emails():
    """Check that remediation email was sent."""
    print("\n=== Checking Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()
    conn.close()

    def find_email_for_recipient(recipient):
        for subj, from_addr, to_addr, body in all_emails:
            if to_addr:
                recipients = []
                if isinstance(to_addr, list):
                    recipients = [str(r).strip().lower() for r in to_addr]
                elif isinstance(to_addr, str):
                    try:
                        parsed = json.loads(to_addr)
                        if isinstance(parsed, list):
                            recipients = [str(r).strip().lower() for r in parsed]
                        else:
                            recipients = [str(to_addr).strip().lower()]
                    except (json.JSONDecodeError, TypeError):
                        recipients = [str(to_addr).strip().lower()]
                if recipient.lower() in recipients:
                    return subj, from_addr, to_addr, body
        return None

    result = find_email_for_recipient("remediation@financeou.example.com")
    check("Email sent to remediation@financeou.example.com", result is not None,
          f"Total emails: {len(all_emails)}")

    if result:
        subj, from_addr, to_addr, body = result
        subj_lower = (subj or "").lower()
        check(
            "Email subject is 'Remediation Study Sessions - Foundations of Finance (Fall 2013)'",
            "remediation study sessions" in subj_lower
            and "foundations of finance" in subj_lower
            and "fall 2013" in subj_lower,
            f"Subject: {subj}",
        )
        check("Email from tutor@financeou.example.com",
              "tutor@financeou.example.com" in (from_addr or "").lower(),
              f"From: {from_addr}")
        body_lower = (body or "").lower()
        # Body must mention all 5 underperforming quiz titles (e.g. CMA 34880)
        for quiz_title in UNDERPERFORMING_QUIZZES:
            quiz_num = quiz_title.split()[-1]
            check(
                f"Email body mentions quiz {quiz_title}",
                quiz_num in body_lower,
                f"Looking for '{quiz_num}' in body",
            )
        # Mention number of sessions (5) somewhere in body, as a standalone token
        n_sessions = len(UNDERPERFORMING_QUIZZES)
        check(
            f"Email body mentions number of sessions ({n_sessions})",
            re.search(r"\b" + str(n_sessions) + r"\b", body_lower) is not None,
            f"Expected standalone {n_sessions} to appear (number of sessions)",
        )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("CANVAS QUIZ REMEDIATION GFORM GCAL - EVALUATION")
    print("=" * 70)

    check_gform()
    check_gcal(args.launch_time)
    check_emails()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
