"""
Evaluation for howtocook-weekly-gsheet-gcal task.

Checks:
1. GSheet "Weekly Meal Plan" exists in gsheet.spreadsheets
2. GSheet has 21 data rows (7 days x 3 meals)
3. GCal has at least 7 dinner prep events in April 2026
4. Email sent to meal_planning@service.com

DB access reads the same env vars the agent's MCP servers use (PGHOST/PGPORT/
PGDATABASE/PGUSER/PGPASSWORD) so that in concurrent worker-DB mode the
evaluator queries the same database the agent wrote into.
"""
import json
import os
import sys
from argparse import ArgumentParser

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
}

PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_gsheet():
    print("\n=== Check 1: Google Sheet Weekly Meal Plan ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except psycopg2.Error as e:
        record("'Weekly Meal Plan' spreadsheet exists", False,
               f"DB connection error: {e}")
        return
    cur = conn.cursor()

    try:
        # Gather ALL candidate spreadsheets. In the homogeneous 2-agent mode the
        # two agents may each create one; we pick the most complete one below so
        # an incomplete duplicate cannot shadow a fully-populated spreadsheet.
        # The title match tolerates the exact name "Weekly Meal Plan" as well as
        # underscore/case variants (e.g. "Weekly_Meal_Plan"); preprocess clears
        # gsheet.spreadsheets before each task, so only this task's sheets exist.
        cur.execute("""
            SELECT id, title FROM gsheet.spreadsheets
            WHERE title ILIKE '%meal plan%'
               OR title ILIKE '%weekly%meal%'
               OR title ILIKE '%meal%plan%'
            ORDER BY created_at DESC
        """)
        spreadsheets = cur.fetchall()
    except psycopg2.Error as e:
        record("'Weekly Meal Plan' spreadsheet exists", False,
               f"DB error: {e}")
        cur.close()
        conn.close()
        return

    if not spreadsheets:
        record("'Weekly Meal Plan' spreadsheet exists", False,
               "No spreadsheet found with 'meal plan' in title")
        cur.close()
        conn.close()
        return

    record("'Weekly Meal Plan' spreadsheet exists", True,
           f"Found {len(spreadsheets)} candidate spreadsheet(s)")

    # Pick the spreadsheet with the most non-empty data rows.
    best_id, best_title, best_rows = None, None, -1
    for sid, title in spreadsheets:
        try:
            cur.execute("""
                SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
                WHERE spreadsheet_id = %s AND row_index > 0
                AND value IS NOT NULL AND value != ''
            """, (sid,))
            rows = cur.fetchone()[0]
        except psycopg2.Error as e:
            record("GSheet has at least 21 data rows", False,
                   f"DB error counting rows: {e}")
            cur.close()
            conn.close()
            return
        if rows > best_rows:
            best_id, best_title, best_rows = sid, title, rows

    spreadsheet_id, title, data_row_count = best_id, best_title, best_rows

    record("Spreadsheet title contains 'meal' or 'weekly'",
           "meal" in title.lower() or "weekly" in title.lower(),
           f"Title: {title}")

    record("GSheet has at least 21 data rows", data_row_count >= 21,
           f"Found {data_row_count} non-header rows with data (most complete spreadsheet)")

    # Check meal types: the task requires Breakfast / Lunch / Dinner all present.
    try:
        cur.execute("""
            SELECT DISTINCT LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND row_index > 0
            AND LOWER(value) IN ('breakfast', 'lunch', 'dinner')
        """, (spreadsheet_id,))
        meal_types = sorted({r[0] for r in cur.fetchall()})
    except psycopg2.Error as e:
        record("GSheet contains Breakfast, Lunch, Dinner meal types", False,
               f"DB error: {e}")
        cur.close()
        conn.close()
        return
    record("GSheet contains Breakfast, Lunch, Dinner meal types",
           len(meal_types) >= 3,
           f"Found meal types: {meal_types}")

    cur.close()
    conn.close()


def check_gcal():
    print("\n=== Check 2: Google Calendar Dinner Prep Events ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except psycopg2.Error as e:
        record("At least 7 dinner prep events Apr 7-13 2026", False,
               f"DB connection error: {e}")
        return
    cur = conn.cursor()

    # Window bounds are written with an explicit UTC offset so the comparison is
    # independent of the session timezone. The window spans every possible
    # wall-clock reading of April 7-13 2026 across real-world timezones; since
    # preprocess clears gcal.events, only this task's events live in the table.
    try:
        cur.execute("""
            SELECT summary, start_datetime,
                   (start_datetime AT TIME ZONE 'UTC') AS start_utc,
                   end_datetime
            FROM gcal.events
            WHERE start_datetime >= '2026-04-06T00:00:00+00'
              AND start_datetime < '2026-04-15T00:00:00+00'
              AND summary ILIKE '%dinner%'
            ORDER BY start_datetime
        """)
        events = cur.fetchall()
    except psycopg2.Error as e:
        record("At least 7 dinner prep events Apr 7-13 2026", False,
               f"DB error: {e}")
        cur.close()
        conn.close()
        return

    cur.close()
    conn.close()

    record("At least 7 dinner prep events Apr 7-13 2026", len(events) >= 7,
           f"Found {len(events)} events in Apr 7-13")

    if events:
        # At least 7 events must start at 18:00. The task does not specify a
        # timezone, so a correct agent may legitimately write "18:00" as UTC,
        # as its local time (naive), or with an explicit offset such as +08:00
        # (Beijing). Because the deployment Postgres session timezone is UTC,
        # only reading the session wall-clock or the UTC wall-clock would
        # wrongly reject the +08:00 interpretation (18:00+08 == 10:00 UTC).
        # Accept the absolute instants that read as 18:00 in the two plausible
        # reference timezones -- UTC (hour 18) and Asia/Shanghai (hour 10) --
        # plus the session-timezone wall-clock reading. This is independent of
        # the deployment's session timezone.
        hour_ok = 0
        for summary, start_dt, start_utc, end_dt in events:
            if start_dt is None or start_utc is None:
                continue
            if start_dt.hour == 18 or start_utc.hour in (18, 10):
                hour_ok += 1
        record("At least 7 dinner prep events start at 18:00", hour_ok >= 7,
               f"Found {hour_ok} events starting at 18:00 "
               f"(session-TZ wall clock, or 18:00 UTC, or 18:00 +08:00/Beijing)")

        # 7 distinct calendar days (UTC date is stable across timezone choices).
        dates = set(e[2].date() for e in events if e[2])
        record("Events on 7 distinct days", len(dates) >= 7,
               f"Distinct dates (UTC): {sorted(dates)}")


def check_email():
    print("\n=== Check 3: Email to meal_planning@service.com ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
    except psycopg2.Error as e:
        record("Email sent to meal_planning@service.com", False,
               f"DB connection error: {e}")
        return
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT subject, from_addr, to_addr, body_text
            FROM email.messages
            ORDER BY id DESC
            LIMIT 500
        """)
        messages = cur.fetchall()
    except psycopg2.Error as e:
        record("Email sent to meal_planning@service.com", False,
               f"DB error: {e}")
        cur.close()
        conn.close()
        return

    cur.close()
    conn.close()

    matching = None
    for subject, from_addr, to_addr, body_text in messages:
        to_str = ""
        if isinstance(to_addr, list):
            to_str = " ".join(str(r).lower() for r in to_addr)
        elif isinstance(to_addr, str):
            try:
                parsed = json.loads(to_addr)
                to_str = " ".join(str(r).lower() for r in parsed) if isinstance(parsed, list) else str(to_addr).lower()
            except Exception:
                to_str = str(to_addr).lower()
        if "meal_planning@service.com" in to_str:
            matching = (subject, from_addr, to_addr, body_text)
            break

    record("Email sent to meal_planning@service.com", matching is not None,
           f"Messages found: {len(messages)}")

    if matching:
        subject, _, _, body_text = matching
        all_text = ((subject or "") + " " + (body_text or "")).lower()
        has_meal_content = (
            "meal plan" in all_text or "weekly" in all_text or
            "breakfast" in all_text or "dinner" in all_text
        )
        record("Email mentions meal plan content", has_meal_content,
               f"Subject: {subject}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_gcal()
    check_email()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
