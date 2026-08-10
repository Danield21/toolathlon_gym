"""
Evaluation script for canvas-assignment-calendar task.

Checks:
1. Text file (assignment_schedule.txt) matches groundtruth
2. Google Calendar events created for each assignment
3. Email sent with consolidated schedule
"""

import argparse
import json
import os
import sys
from datetime import timezone

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
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


def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


def str_match(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return str(a).strip().lower() == str(b).strip().lower()


# ============================================================================
# Check 1: assignment_schedule.txt
# ============================================================================

def check_text_file(agent_workspace, groundtruth_workspace):
    print("\n=== Checking assignment_schedule.txt ===")

    agent_file = os.path.join(agent_workspace, "assignment_schedule.txt")
    gt_file = os.path.join(groundtruth_workspace, "assignment_schedule.txt")

    if not os.path.isfile(agent_file):
        record("Text file exists", False, f"Not found: {agent_file}")
        return False
    record("Text file exists", True)

    if not os.path.isfile(gt_file):
        record("Groundtruth text file exists", False, f"Not found: {gt_file}")
        return False

    with open(agent_file) as f:
        agent_lines = [l.strip() for l in f.readlines() if l.strip()]
    with open(gt_file) as f:
        gt_lines = [l.strip() for l in f.readlines() if l.strip()]

    record("Line count matches", len(agent_lines) == len(gt_lines),
           f"Expected {len(gt_lines)}, got {len(agent_lines)}")

    # Check that each groundtruth line appears in agent output
    matched = 0
    for gt_line in gt_lines:
        gt_parts = [p.strip().lower() for p in gt_line.split(" - ")]
        found = False
        for agent_line in agent_lines:
            agent_parts = [p.strip().lower() for p in agent_line.split(" - ")]
            if len(agent_parts) >= 3 and len(gt_parts) >= 3:
                if (agent_parts[0] == gt_parts[0] and
                    agent_parts[1] == gt_parts[1] and
                    agent_parts[2] == gt_parts[2]):
                    found = True
                    break
        if found:
            matched += 1

    # Strict 100% match required
    record(f"Assignment lines matched ({matched}/{len(gt_lines)})",
           matched == len(gt_lines),
           f"Matched {matched} of {len(gt_lines)}")

    # Check sort order: by course code alpha, then by date asc within each course
    sorted_lines = sorted(agent_lines, key=lambda l: (
        (l.split(" - ")[0].strip() if " - " in l else l),
        (l.split(" - ")[-1].strip() if " - " in l else l),
    ))
    record("assignment_schedule.txt sorted by course then date",
           agent_lines == sorted_lines,
           f"Order mismatch (first 3 actual: {agent_lines[:3]})")

    return matched == len(gt_lines) and agent_lines == sorted_lines


# ============================================================================
# Check 2: Google Calendar events
# ============================================================================

def _fetch_expected_assignments():
    """Return list of (course_code, course_name, assignment_name, due_date_str) for Fall 2014."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT c.course_code, c.name, a.name, a.due_at
        FROM canvas.assignments a
        JOIN canvas.courses c ON c.id = a.course_id
        WHERE c.course_code LIKE %s AND a.due_at IS NOT NULL
        ORDER BY c.course_code, a.due_at
    """, ('%2014J',))
    out = []
    for code, cname, aname, due in cur.fetchall():
        # task.md: use the UTC calendar date of due_at (no local-tz conversion).
        # Render in UTC explicitly so the expected date is deterministic and
        # independent of the PG session TimeZone.
        try:
            if getattr(due, "tzinfo", None) is None:
                due = due.replace(tzinfo=timezone.utc)
            d = due.astimezone(timezone.utc).strftime("%Y-%m-%d")
        except Exception:
            d = str(due)[:10]
        out.append((code, cname, aname, d))
    cur.close()
    conn.close()
    return out


def check_gcal():
    print("\n=== Checking Google Calendar ===")

    expected_assignments = _fetch_expected_assignments()
    expected_per_code = {}
    for code, _, aname, _ in expected_assignments:
        expected_per_code[code] = expected_per_code.get(code, 0) + 1
    expected_total = len(expected_assignments)

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime
        FROM gcal.events
        ORDER BY summary
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_gcal] Found {len(events)} calendar events. Expected total {expected_total}.")

    record(f"Calendar event count == {expected_total}", len(events) == expected_total,
           f"Found {len(events)}")

    spot_checks = sorted(expected_per_code.keys())

    all_ok = (len(events) == expected_total)
    for code in spot_checks:
        code_events = [e for e in events if code.lower() in (e[0] or "").lower()]
        exp_n = expected_per_code[code]
        record(f"gcal: {exp_n} events for {code}",
               len(code_events) == exp_n,
               f"Got {len(code_events)} for {code}")
        if len(code_events) != exp_n:
            all_ok = False

    # Verify summary format and description format on each event
    summary_ok = 0
    desc_ok = 0
    for ev in events:
        summary, description, start_dt, _ = ev
        summary_str = str(summary or "")
        desc_str = str(description or "")
        # summary should contain a 2014J code and " - "
        if any(code.lower() in summary_str.lower() for code in spot_checks) and " - " in summary_str:
            summary_ok += 1
        # description should mention 'Assignment for' and 'Due date'
        if "assignment for" in desc_str.lower() and "due date" in desc_str.lower():
            desc_ok += 1
    if events:
        record("gcal: summaries follow '[Code] - [Name]' format",
               summary_ok == len(events),
               f"{summary_ok}/{len(events)} match")
        record("gcal: descriptions include 'Assignment for' and 'Due date'",
               desc_ok == len(events),
               f"{desc_ok}/{len(events)} match")
        if summary_ok != len(events) or desc_ok != len(events):
            all_ok = False

    return all_ok


# ============================================================================
# Check 3: Email sent
# ============================================================================

def check_emails():
    print("\n=== Checking Emails ===")

    expected_assignments = _fetch_expected_assignments()
    expected_total = len(expected_assignments)
    course_codes = sorted({a[0] for a in expected_assignments})

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_emails] Found {len(all_emails)} total emails.")

    record("At least 1 email sent", len(all_emails) >= 1,
           f"Found {len(all_emails)}")

    all_ok = True
    found_schedule = False

    for subject, from_addr, to_addr, body_text in all_emails:
        subject_lower = (subject or "").lower()
        if "assignment schedule" in subject_lower or "fall 2014" in subject_lower:
            found_schedule = True

            # Sender address is fixed by the email environment; match on
            # recipient + subject only.
            to_str = str(to_addr or "").lower()
            record("email: to students",
                   "students@openuniversity.ac.uk" in to_str,
                   f"To: {to_addr}")

            body_lower = (body_text or "").lower()
            # Check ALL course codes appear (not just AAA and GGG)
            missing_codes = [c for c in course_codes if c.lower() not in body_lower]
            record(f"email: body mentions all {len(course_codes)} course codes",
                   not missing_codes,
                   f"Missing: {missing_codes}")
            if missing_codes:
                all_ok = False
            # Total count - dynamic from DB
            record(f"email: body mentions total count {expected_total}",
                   str(expected_total) in (body_text or ""),
                   f"Missing total count of {expected_total}")
            if str(expected_total) not in (body_text or ""):
                all_ok = False
            break

    record("email: schedule email found", found_schedule,
           "No email with 'assignment schedule' in subject")
    if not found_schedule:
        all_ok = False

    return all_ok


# ============================================================================
# Main
# ============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    text_ok = check_text_file(args.agent_workspace, args.groundtruth_workspace)
    gcal_ok = check_gcal()
    email_ok = check_emails()

    all_passed = text_ok and gcal_ok and email_ok and FAIL_COUNT == 0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Overall: {'PASS' if all_passed else 'FAIL'}")

    if args.res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "success": all_passed,
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
