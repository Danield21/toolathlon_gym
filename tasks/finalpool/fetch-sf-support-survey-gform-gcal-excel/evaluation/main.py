"""
Evaluation script for fetch-sf-support-survey-gform-gcal-excel task.

Checks:
1. Support_Satisfaction_Analysis.xlsx with 4 sheets and correct data
2. Google Form for ongoing feedback
3. Calendar events for 4 quarterly reviews (runtime_only by default;
   becomes blocking if agent populated relevant events)
"""

import argparse
import json
import os
import sys

import openpyxl
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_ONLY_FAIL = 0


def record(name, passed, detail="", runtime_only=False):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_ONLY_FAIL
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        if runtime_only:
            RUNTIME_ONLY_FAIL += 1
        msg = f": {detail[:300]}" if detail else ""
        suffix = " (runtime-only)" if runtime_only else ""
        print(f"  [FAIL] {name}{suffix}{msg}")


def num_close(a, b, tol=0.5):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_contains(haystack, needle):
    if haystack is None or needle is None:
        return False
    return needle.strip().lower() in str(haystack).strip().lower()


def check_excel(agent_workspace):
    """Check Support_Satisfaction_Analysis.xlsx."""
    print("\n=== Checking Excel Output ===")

    fpath = os.path.join(agent_workspace, "Support_Satisfaction_Analysis.xlsx")
    if not os.path.isfile(fpath):
        record("Excel file exists", False, f"Not found: {fpath}")
        return False

    record("Excel file exists", True)

    try:
        wb = openpyxl.load_workbook(fpath, data_only=True)
    except Exception as e:
        record("Excel file readable", False, str(e))
        return False

    all_ok = True

    # --- Sheet 1: Survey Results ---
    survey_sheet = None
    for name in wb.sheetnames:
        if "survey" in name.lower() and "summary" not in name.lower():
            survey_sheet = name
            break
    if not survey_sheet:
        record("Survey Results sheet exists", False, f"Sheets: {wb.sheetnames}")
        all_ok = False
    else:
        record("Survey Results sheet exists", True)
        ws = wb[survey_sheet]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = rows[1:] if len(rows) > 1 else []
        ok = len(data_rows) == 20
        record("Survey Results has 20 rows", ok, f"Found {len(data_rows)}")
        if not ok:
            all_ok = False

    # --- Sheet 2: Survey Summary ---
    summary_sheet = None
    for name in wb.sheetnames:
        if "summary" in name.lower():
            summary_sheet = name
            break
    if not summary_sheet:
        record("Survey Summary sheet exists", False, f"Sheets: {wb.sheetnames}")
        all_ok = False
    else:
        record("Survey Summary sheet exists", True)
        ws = wb[summary_sheet]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = rows[1:] if len(rows) > 1 else []

        for row in data_rows:
            if row and row[0]:
                metric = str(row[0]).strip().lower()
                val = row[1]
                if "total_respondents" in metric or "total" in metric and "respondent" in metric:
                    ok = num_close(val, 20, tol=0)
                    record("Total respondents = 20", ok, f"Got {val}")
                    if not ok:
                        all_ok = False
                elif "avg_overall" in metric or ("overall" in metric and "satisfaction" in metric):
                    ok = num_close(val, 3.25, tol=0.3)
                    record("Avg overall satisfaction ~3.25", ok, f"Got {val}")
                    if not ok:
                        all_ok = False
                elif "lowest" in metric:
                    ok = str_contains(val, "low")
                    record("Lowest rated priority is Low", ok, f"Got {val}")
                    if not ok:
                        all_ok = False
                elif "highest" in metric:
                    ok = str_contains(val, "high")
                    record("Highest rated priority is High", ok, f"Got {val}")
                    if not ok:
                        all_ok = False

    # --- Sheet 3: Ticket System Comparison ---
    comp_sheet = None
    for name in wb.sheetnames:
        if "ticket" in name.lower() or "comparison" in name.lower():
            comp_sheet = name
            break
    if not comp_sheet:
        record("Ticket System Comparison sheet exists", False, f"Sheets: {wb.sheetnames}")
        all_ok = False
    else:
        record("Ticket System Comparison sheet exists", True)
        ws = wb[comp_sheet]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = rows[1:] if len(rows) > 1 else []
        ok = len(data_rows) == 3
        record("Ticket Comparison has 3 rows", ok, f"Found {len(data_rows)}")
        if not ok:
            all_ok = False

        # Check ticket avg response hours for High priority
        for row in data_rows:
            if row and str_contains(row[0], "high"):
                # Ticket avg response hours should be ~6.23
                found = False
                for cell in row[1:]:
                    if num_close(cell, 6.23, tol=1.0):
                        found = True
                        break
                record("High priority ticket response ~6.23 hrs", found,
                       f"Row: {str(row)[:200]}")
                if not found:
                    all_ok = False

    # --- Sheet 4: Improvement Areas ---
    # task.md says: "Include rows for each metric (Response Time, Resolution
    # Quality, Agent Professionalism) where the survey average is below 4.0.
    # The target score should be 4.5 for all metrics. The gap is the target
    # minus the current score."
    # GT data: Response Time=3.10, Resolution Quality=3.55 (both below 4.0).
    # Agent Professionalism avg = 4.0 (NOT below) so it should NOT be present.
    # → Expect exactly 2 rows: Response Time and Resolution Quality.
    imp_sheet = None
    for name in wb.sheetnames:
        if "improvement" in name.lower():
            imp_sheet = name
            break
    if not imp_sheet:
        record("Improvement Areas sheet exists", False, f"Sheets: {wb.sheetnames}")
        all_ok = False
    else:
        record("Improvement Areas sheet exists", True)
        ws = wb[imp_sheet]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if r and r[0]]
        ok_count = len(data_rows) == 2
        record("Improvement Areas has exactly 2 rows", ok_count,
               f"Found {len(data_rows)} (expected 2: Response Time + Resolution Quality)")
        if not ok_count:
            all_ok = False

        # Build lookup by area name
        by_area = {}
        for r in data_rows:
            if r and r[0]:
                by_area[str(r[0]).strip().lower()] = r

        # Expected areas + scores per GT
        expected_imp = {
            "response time": {"current": 3.10, "target": 4.5, "gap": 1.40},
            "resolution quality": {"current": 3.55, "target": 4.5, "gap": 0.95},
        }

        for area_key, exp in expected_imp.items():
            row = None
            for k, v in by_area.items():
                if area_key in k:
                    row = v
                    break
            if row is None:
                record(f"Improvement Area '{area_key}' present", False,
                       f"Got areas: {list(by_area.keys())}")
                all_ok = False
                continue
            record(f"Improvement Area '{area_key}' present", True)
            # row format: (Area, Current_Score, Target_Score, Gap)
            if len(row) >= 4:
                cur_score = row[1]
                tgt_score = row[2]
                gap = row[3]
                ok = num_close(cur_score, exp["current"], 0.05)
                record(f"'{area_key}' Current_Score ~{exp['current']}",
                       ok, f"Got {cur_score}")
                if not ok:
                    all_ok = False
                # Target_Score must be 4.5 — task.md is explicit
                ok_tgt = num_close(tgt_score, 4.5, 0.01)
                record(f"'{area_key}' Target_Score = 4.5",
                       ok_tgt, f"Got {tgt_score}")
                if not ok_tgt:
                    all_ok = False
                # Gap = Target - Current
                ok_gap = num_close(gap, exp["gap"], 0.05)
                record(f"'{area_key}' Gap ~{exp['gap']}",
                       ok_gap, f"Got {gap}")
                if not ok_gap:
                    all_ok = False
            else:
                record(f"'{area_key}' has 4 columns", False,
                       f"Row has {len(row)} columns: {row}")
                all_ok = False

        # Reject Agent Professionalism (avg = 4.0, NOT below 4.0)
        has_prof = any("professionalism" in k or "agent" in k for k in by_area.keys())
        record("Agent Professionalism NOT in improvement (avg = 4.0)",
               not has_prof, f"Got areas: {list(by_area.keys())}")
        if has_prof:
            all_ok = False

    wb.close()
    return all_ok


def check_gform():
    """Check Google Form for ongoing feedback."""
    print("\n=== Checking Google Form ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT id, title, description FROM gform.forms")
        forms = cur.fetchall()

        found_form = False
        form_id = None
        for fid, title, desc in forms:
            title_lower = (title or "").lower()
            if ("support" in title_lower or "feedback" in title_lower or
                    "customer" in title_lower or "satisfaction" in title_lower):
                if "employee" not in title_lower:  # Skip noise form
                    found_form = True
                    form_id = fid
                    break

        record("Customer feedback form exists", found_form,
               f"Found forms: {[(t, d[:50] if d else '') for _, t, d in forms]}",
               runtime_only=True)

        all_ok = found_form

        if form_id:
            cur.execute("SELECT title, question_type FROM gform.questions WHERE form_id = %s", (form_id,))
            questions = cur.fetchall()
            q_count = len(questions)
            ok = q_count >= 4
            record(f"Form has >= 4 questions", ok, f"Found {q_count}", runtime_only=True)
            if not ok:
                all_ok = False

            # Check for specific question types
            q_titles = " ".join((t or "").lower() for t, _ in questions)
            has_satisfaction = "satisfaction" in q_titles or "overall" in q_titles or "rating" in q_titles
            record("Has satisfaction question", has_satisfaction, f"Q titles: {q_titles[:200]}",
                   runtime_only=True)
            if not has_satisfaction:
                all_ok = False

            has_comment = any("text" in (qt or "").lower() or "paragraph" in (qt or "").lower()
                              for _, qt in questions)
            # Also check if there's a question about comments
            has_comment = has_comment or "comment" in q_titles or "feedback" in q_titles
            record("Has comments/text question", has_comment, runtime_only=True)
            if not has_comment:
                all_ok = False

        cur.close()
        conn.close()
        return all_ok

    except Exception as e:
        record("Google Form DB accessible", False, str(e), runtime_only=True)
        return False


def check_calendar():
    """Check calendar events for 4 quarterly review meetings.

    Pattern: runtime_only by default (won't block on V1 GT-only test). BUT
    when the agent has populated 'Support Satisfaction Review' events, all
    quality checks (count, dates, durations) are blocking.
    """
    print("\n=== Checking Google Calendar ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT summary, description, start_datetime, end_datetime FROM gcal.events")
        events = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Calendar DB accessible", False, str(e), runtime_only=True)
        return False

    # Detect agent population: any event with 'support' / 'satisfaction'
    # and 'review' or quarter marker
    relevant = []
    for s, d, st, et in events:
        sl = (s or "").lower()
        if ("support" in sl or "satisfaction" in sl) and "review" in sl:
            relevant.append((s, d, st, et))
        elif "satisfaction review" in sl:
            relevant.append((s, d, st, et))
    agent_populated = len(relevant) > 0
    is_runtime_only = not agent_populated

    quarters_found = set()
    quarter_events = []
    for summary, description, start_dt, end_dt in events:
        summary_lower = (summary or "").lower()
        if "support" in summary_lower or "satisfaction" in summary_lower or "review" in summary_lower:
            for q in ["q1", "q2", "q3", "q4"]:
                if q in summary_lower:
                    quarters_found.add(q)
                    quarter_events.append((q, summary, start_dt, end_dt))

    ok = len(quarters_found) >= 4
    record("All 4 quarterly review events found", ok,
           f"Found quarters: {quarters_found}",
           runtime_only=is_runtime_only)

    # Per-event verifications: when agent populated, BLOCKING
    if agent_populated:
        # Expected dates: 2026-03-15, 2026-06-15, 2026-09-15, 2026-12-15
        expected_dates_by_q = {
            "q1": "2026-03-15",
            "q2": "2026-06-15",
            "q3": "2026-09-15",
            "q4": "2026-12-15",
        }
        for q, summary, start_dt, end_dt in quarter_events:
            if start_dt and end_dt:
                duration_min = (end_dt - start_dt).total_seconds() / 60
                record(f"Event {q} duration 90 min (10:00-11:30)",
                       abs(duration_min - 90) <= 10,
                       f"Got {duration_min} min",
                       runtime_only=False)
                # Date check
                exp_date = expected_dates_by_q.get(q)
                if exp_date:
                    actual_date = str(start_dt)[:10]
                    record(f"Event {q} date = {exp_date}",
                           actual_date == exp_date,
                           f"Got {actual_date}",
                           runtime_only=False)
    else:
        # Not populated → keep all per-event checks as runtime_only
        for q, summary, start_dt, end_dt in quarter_events:
            if start_dt and end_dt:
                duration_min = (end_dt - start_dt).total_seconds() / 60
                record(f"Event {q} duration 90 min (10:00-11:30)",
                       abs(duration_min - 90) <= 10,
                       f"Got {duration_min} min",
                       runtime_only=True)

    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    excel_ok = check_excel(args.agent_workspace)
    gform_ok = check_gform()
    cal_ok = check_calendar()

    print(f"\n=== SUMMARY ===")
    print(f"  Excel:    {'PASS' if excel_ok else 'FAIL'}")
    print(f"  GForm:    {'PASS' if gform_ok else 'FAIL'}")
    print(f"  Calendar: {'PASS' if cal_ok else 'FAIL'}")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT} (runtime-only fails: {RUNTIME_ONLY_FAIL})")

    blocking_fail = FAIL_COUNT - RUNTIME_ONLY_FAIL
    overall = blocking_fail == 0
    print(f"  Blocking failures: {blocking_fail}")
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
