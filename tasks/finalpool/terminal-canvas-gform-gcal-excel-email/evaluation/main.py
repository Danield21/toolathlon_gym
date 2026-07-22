"""Evaluation for terminal-canvas-gform-gcal-excel-email."""
import argparse
import json
import os
import sys

import openpyxl
import psycopg2

def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
          user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0

# Quizzes below 75 avg from courses 7 and 11 — derived from DB at module load.
# Falls back to hardcoded snapshot if DB unreachable so eval still runs.
def _fetch_below_75_quizzes():
    fallback = [
        "CMA 24295", "CMA 24298",
        "CMA 25341", "CMA 25343", "CMA 25344",
        "CMA 25345", "CMA 25346", "CMA 25347",
    ]
    try:
        import psycopg2 as _ps
        conn = _ps.connect(**DB)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT q.title FROM canvas.quizzes q
            JOIN canvas.quiz_submissions s ON s.quiz_id = q.id
            WHERE q.course_id IN (7, 11) AND s.score IS NOT NULL
            GROUP BY q.title
            HAVING AVG(s.score) < 75
            ORDER BY q.title
            """
        )
        rows = cur.fetchall()
        cur.close()
        conn.close()
        if rows:
            return [r[0] for r in rows]
        return fallback
    except Exception:
        return fallback


BELOW_75_QUIZZES = _fetch_below_75_quizzes()


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        d = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {name}{d}")


def safe_float(val, default=None):
    try:
        if val is None:
            return default
        return float(str(val).replace(",", "").replace("%", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return default


def check_excel(ws_path):
    """Check Quiz_Performance_Report.xlsx."""
    print("\n=== Checking Excel ===")
    path = os.path.join(ws_path, "Quiz_Performance_Report.xlsx")
    if not os.path.isfile(path):
        check("Excel file exists", False, f"Not found: {path}")
        return
    check("Excel file exists", True)

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        check("Excel readable", False, str(e))
        return

    # Sheet 1: Quiz_Performance
    sheet_names_lower = {s.lower(): s for s in wb.sheetnames}
    qp_name = None
    for candidate in ["quiz_performance", "quiz performance"]:
        if candidate in sheet_names_lower:
            qp_name = sheet_names_lower[candidate]
            break
    if qp_name is None:
        check("Quiz_Performance sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        check("Quiz_Performance sheet exists", True)
        ws = wb[qp_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if r and r[0] is not None]
        # Query dynamic quiz count from Canvas DB
        try:
            conn = psycopg2.connect(**DB)
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM canvas.quizzes WHERE course_id IN (7, 11)")
            expected_quiz_count = cur.fetchone()[0]
            cur.close(); conn.close()
        except Exception:
            expected_quiz_count = 11
        check(f"Quiz_Performance has {expected_quiz_count} rows",
              len(data_rows) == expected_quiz_count,
              f"Found {len(data_rows)} data rows")

        # Query dynamic quiz avg scores + pass rates from Canvas DB
        try:
            conn = psycopg2.connect(**DB)
            cur = conn.cursor()
            cur.execute("""
                SELECT q.id,
                       q.title,
                       AVG(qs.score),
                       100.0 * SUM(CASE WHEN qs.score >= 70 THEN 1 ELSE 0 END) / COUNT(*)
                FROM canvas.quizzes q
                JOIN canvas.quiz_submissions qs ON q.id = qs.quiz_id
                WHERE q.course_id IN (7, 11)
                GROUP BY q.id, q.title
            """)
            quiz_stats = {}
            for qid, title, avg, pr in cur.fetchall():
                quiz_stats[str(qid)] = {
                    "title": title,
                    "avg": float(avg),
                    "pass_rate": float(pr),
                }
            cur.close(); conn.close()
        except Exception as e:
            print(f"  [INFO] Could not fetch quiz stats from DB: {e}")
            quiz_stats = {}

        # Check ALL quizzes from DB against agent's data
        a_by_qid = {}
        for r in data_rows:
            try:
                qid = int(r[2]) if r[2] is not None else None
            except (TypeError, ValueError):
                qid = None
            if qid is not None:
                a_by_qid[str(qid)] = r

        for qid, stat in quiz_stats.items():
            r = a_by_qid.get(qid)
            if r is None:
                check(f"Quiz {qid} ({stat['title']}) present", False,
                      f"Not found in Excel")
                continue
            avg = safe_float(r[4])
            pr = safe_float(r[5])
            needs = str(r[6]).strip().lower() if len(r) > 6 and r[6] else ""
            check(f"Quiz {qid} ({stat['title']}) Avg_Score ~{stat['avg']:.2f}",
                  avg is not None and abs(avg - stat["avg"]) < 1.0,
                  f"Got {avg}")
            check(f"Quiz {qid} Pass_Rate ~{stat['pass_rate']:.2f}",
                  pr is not None and abs(pr - stat["pass_rate"]) < 2.0,
                  f"Got {pr}")
            expected_needs = "yes" if stat["avg"] < 75 else "no"
            # Tighten: exact match (case-insensitive trim) instead of 'in' substring
            check(f"Quiz {qid} Needs_Review = {expected_needs}",
                  needs == expected_needs,
                  f"Got '{r[6]}' for avg {stat['avg']:.2f}")

    # Sheet 2: Feedback_Summary
    fb_name = None
    for candidate in ["feedback_summary", "feedback summary"]:
        if candidate in sheet_names_lower:
            fb_name = sheet_names_lower[candidate]
            break
    if fb_name is None:
        check("Feedback_Summary sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        check("Feedback_Summary sheet exists", True)
        ws2 = wb[fb_name]
        rows2 = list(ws2.iter_rows(values_only=True))
        data_rows2 = [r for r in rows2[1:] if r and r[0] is not None]
        check("Feedback_Summary has 5 rows", len(data_rows2) == 5,
              f"Found {len(data_rows2)} rows")
        # Validate question types per task spec
        # Q1: free text, Q2: multiple choice, Q3: multiple choice, Q4: scale, Q5: free text
        expected_types = {
            1: ["short_answer", "text", "paragraph", "free_text"],
            2: ["multiple_choice", "radio", "choice"],
            3: ["multiple_choice", "radio", "choice"],
            4: ["scale", "linear_scale", "rating"],
            5: ["short_answer", "text", "paragraph", "free_text"],
        }
        for r in data_rows2:
            try:
                qn = int(r[0])
            except (TypeError, ValueError):
                continue
            qtype = str(r[2] or "").strip().lower().replace(" ", "_")
            if qn in expected_types:
                ok = any(t in qtype for t in expected_types[qn])
                check(f"Feedback_Summary Q{qn} type matches expected",
                      ok,
                      f"Got '{r[2]}', expected one of {expected_types[qn]}")
        # Q1 should mention 'challenging' or 'topic'
        q1_text = next((str(r[1] or "").lower() for r in data_rows2 if r[0] == 1), "")
        check("Q1 mentions challenging course topics",
              "challeng" in q1_text or "topic" in q1_text or "difficult" in q1_text,
              f"Q1 text: {q1_text}")

    # Sheet 3: Remediation_Schedule
    rs_name = None
    for candidate in ["remediation_schedule", "remediation schedule"]:
        if candidate in sheet_names_lower:
            rs_name = sheet_names_lower[candidate]
            break
    if rs_name is None:
        check("Remediation_Schedule sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        check("Remediation_Schedule sheet exists", True)
        ws3 = wb[rs_name]
        rows3 = list(ws3.iter_rows(values_only=True))
        data_rows3 = [r for r in rows3[1:] if r and r[0] is not None]
        check("Remediation_Schedule has 8 rows", len(data_rows3) == 8,
              f"Found {len(data_rows3)} rows")
        # Expected dates: 2026-03-16 .. 2026-03-23 (one per day) at 14:00, 1 hour
        expected_dates = [f"2026-03-{16+i:02d}" for i in range(8)]
        seen_dates = set()
        for r in data_rows3:
            d = str(r[1] or "")
            t = str(r[2] or "")
            dur = safe_float(r[3])
            for ed in expected_dates:
                if ed in d:
                    seen_dates.add(ed)
            check(f"Remediation row '{r[0]}' Start_Time = 14:00",
                  "14:00" in t or "14" in t.split(":")[0:1] or t.startswith("14"),
                  f"Got '{r[2]}'")
            check(f"Remediation row '{r[0]}' Duration_Hours = 1",
                  dur is not None and abs(dur - 1.0) < 0.05,
                  f"Got {r[3]}")
        check("All 8 expected session dates (Mar 16-23) present",
              len(seen_dates) == 8,
              f"Got dates {sorted(seen_dates)}, expected {expected_dates}")

    wb.close()


def check_gform():
    """Check Google Form creation."""
    print("\n=== Checking Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()

    # Find the form by title that mentions academic+performance+self-assessment
    form_id = None
    for fid, title in forms:
        t = (title or "").lower()
        if ("academic" in t and "performance" in t and ("self" in t or "assess" in t)):
            form_id = fid
            break
    if form_id is None:
        # Fallback to broader match
        for fid, title in forms:
            t = (title or "").lower()
            if "performance" in t and "assess" in t:
                form_id = fid
                break

    check("Assessment feedback form 'Academic Performance Self-Assessment' created",
          form_id is not None,
          f"Forms: {[f[1] for f in forms]}")

    if form_id:
        cur.execute("""
            SELECT id, title, question_type, config FROM gform.questions
            WHERE form_id = %s ORDER BY position
        """, (form_id,))
        questions = cur.fetchall()
        check("Form has exactly 5 questions", len(questions) == 5,
              f"Found {len(questions)}")

        if len(questions) >= 5:
            types = [(q[2] or "").upper() for q in questions]
            titles = [(q[1] or "").lower() for q in questions]

            def _get_choices(cfg):
                if cfg is None:
                    return []
                if isinstance(cfg, str):
                    try:
                        cfg = json.loads(cfg)
                    except (TypeError, ValueError):
                        return []
                if isinstance(cfg, dict):
                    return [str(c).strip().lower() for c in (cfg.get("choices") or cfg.get("options") or [])]
                return []

            # Q1: free text about challenging topics
            q1_ok = any(
                ("challeng" in t or "topic" in t or "difficult" in t)
                and types[i] in ("TEXT", "SHORT_ANSWER", "PARAGRAPH")
                for i, t in enumerate(titles)
            )
            check("Q1: free text question about challenging topics", q1_ok,
                  f"Questions: {list(zip(titles, types))}")

            # Q2: multiple choice with hours options
            q2_ok = False
            for i, q in enumerate(questions):
                t = titles[i]
                qtype = types[i]
                if "hour" in t and "stud" in t and qtype in ("RADIO", "MULTIPLE_CHOICE", "CHOICE"):
                    choices = _get_choices(q[3])
                    expected_choices = ["less than 3", "3 to 5", "5 to 8", "more than 8"]
                    matched = sum(1 for ec in expected_choices
                                  if any(ec in c for c in choices))
                    if matched >= 3:
                        q2_ok = True
                        break
            check("Q2: multiple choice for study hours with 4 specific options",
                  q2_ok,
                  f"Questions+choices: {[(t, types[i], _get_choices(questions[i][3])) for i, t in enumerate(titles)]}")

            # Q3: multiple choice for learning format
            q3_ok = False
            for i, q in enumerate(questions):
                t = titles[i]
                qtype = types[i]
                if ("learn" in t or "format" in t) and qtype in ("RADIO", "MULTIPLE_CHOICE", "CHOICE"):
                    choices = _get_choices(q[3])
                    expected_choices = ["lectures", "hands-on labs", "group projects", "self-paced online"]
                    matched = sum(1 for ec in expected_choices
                                  if any(ec in c or c in ec for c in choices))
                    if matched >= 3:
                        q3_ok = True
                        break
            check("Q3: multiple choice for learning format with 4 specific options",
                  q3_ok)

            # Q4: scale 1-5
            q4_ok = False
            for i, q in enumerate(questions):
                t = titles[i]
                qtype = types[i]
                if "difficult" in t or "rate" in t or "scale" in t:
                    if qtype in ("SCALE", "LINEAR_SCALE", "RATING", "STAR_RATING"):
                        q4_ok = True
                        break
                    cfg = q[3]
                    if isinstance(cfg, str):
                        try:
                            cfg = json.loads(cfg)
                        except (TypeError, ValueError):
                            cfg = {}
                    if isinstance(cfg, dict):
                        lo = cfg.get("low") or cfg.get("min")
                        hi = cfg.get("high") or cfg.get("max")
                        if lo is not None and hi is not None and int(lo) == 1 and int(hi) == 5:
                            q4_ok = True
                            break
            check("Q4: scale 1-5 for course difficulty", q4_ok)

            # Q5: free text suggestions
            q5_ok = any(
                ("suggest" in t or "improv" in t)
                and types[i] in ("TEXT", "SHORT_ANSWER", "PARAGRAPH")
                for i, t in enumerate(titles)
            )
            check("Q5: free text question about suggestions", q5_ok)

    conn.close()


def check_gcal():
    """Check calendar events."""
    print("\n=== Checking Calendar Events ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT summary, start_datetime FROM gcal.events ORDER BY start_datetime")
    events = cur.fetchall()

    # Filter to "Quiz Review:" events
    review_events = [e for e in events
                     if "quiz review" in (e[0] or "").lower()]
    check("Exactly 8 'Quiz Review:' events created",
          len(review_events) == 8,
          f"Found {len(review_events)} Quiz Review events: {[e[0] for e in review_events]}")

    # Reference all 8 below-75 quiz titles
    found_quizzes = []
    for quiz_title in BELOW_75_QUIZZES:
        quiz_num = quiz_title.split()[-1]
        if any(quiz_num in (e[0] or "") for e in review_events):
            found_quizzes.append(quiz_title)
    check(f"Review events reference all {len(BELOW_75_QUIZZES)} below-75 quizzes",
          len(found_quizzes) == len(BELOW_75_QUIZZES),
          f"Referenced quizzes: {found_quizzes}; missing: {set(BELOW_75_QUIZZES) - set(found_quizzes)}")

    # Validate dates: 2026-03-16 ... 2026-03-23
    expected_dates = [f"2026-03-{16+i:02d}" for i in range(len(BELOW_75_QUIZZES))]
    seen_dates = set()
    for _, sd in review_events:
        s = str(sd or "")
        for d in expected_dates:
            if d in s:
                seen_dates.add(d)
    check("All 8 expected calendar dates (2026-03-16 to 2026-03-23) present",
          len(seen_dates) == len(expected_dates),
          f"Got {sorted(seen_dates)}, expected {expected_dates}")

    # Check 2 PM start time
    correct_time = 0
    for _, sd in review_events:
        s = str(sd or "")
        # accept local 14:00 or UTC 18:00/19:00 (NY DST -4 / -5)
        if "14:00" in s or "18:00" in s or "19:00" in s:
            correct_time += 1
    check("All 8 review events start at 2:00 PM equivalent",
          correct_time >= len(review_events) and len(review_events) == 8,
          f"Got {correct_time}/{len(review_events)} with valid start time")

    conn.close()


def check_email():
    """Check email sent."""
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()

    target_email = None
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
            if "faculty@assessment.example.com" in recipients:
                target_email = (subj, from_addr, to_addr, body)
                break

    check("Email sent to faculty@assessment.example.com", target_email is not None,
          f"Total emails: {len(all_emails)}")

    if target_email:
        subj, from_addr, to_addr, body = target_email
        sl = (subj or "").lower()
        # Subject: "Quiz Performance Analysis - Cross-Course Review"
        # Normalized exact match (whitespace collapsed, lowercase)
        import re as _re_subj
        _norm_subj = _re_subj.sub(r"\s+", " ", sl).strip()
        _expected_subj = "quiz performance analysis - cross-course review"
        check("Email subject 'Quiz Performance Analysis - Cross-Course Review'",
              _norm_subj == _expected_subj,
              f"Subject: '{subj}' (normalized: '{_norm_subj}', expected: '{_expected_subj}')")
        check("Email from coordinator@assessment.example.com",
              "coordinator@assessment.example.com" in (from_addr or "").lower(),
              f"From: {from_addr}")
        body_lower = (body or "").lower()
        check("Email body mentions remediation or review",
              "remediation" in body_lower or "review" in body_lower or "flagged" in body_lower,
              "Expected remediation/review content in body")
        check("Email body mentions survey or feedback",
              "survey" in body_lower or "feedback" in body_lower,
              "Expected survey/feedback mention in body")
        # Body should include total quizzes count (11) and flagged count (8)
        # Tighten: word boundary so '18'/'811' don't false-positive
        import re as _re_body
        check("Email body mentions total quizzes (11) word-boundary",
              _re_body.search(r"\b11\b", body or "") is not None,
              f"Body sample: {body_lower[:300]}")
        check("Email body mentions flagged count (8) word-boundary",
              _re_body.search(r"\b8\b", body or "") is not None,
              f"Body sample: {body_lower[:300]}")
        # Body should include the program average (around 73-74)
        # Loose: any number in the 70-80 range
        import re as _re
        avg_nums = _re.findall(r'\b(7[0-9]\.\d{1,2}|7[0-9])\b', body or "")
        check("Email body includes program average score (in 70-80 range)",
              len(avg_nums) >= 1,
              f"Numbers found: {avg_nums}; body sample: {(body or '')[:300]}")

    conn.close()


def check_reverse_validation(workspace):
    print("\n=== Reverse Validation ===")
    # Check that noise emails are not in agent output (using assessment.example.com domain noise)
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("""
            SELECT to_addr FROM email.messages
            WHERE from_addr = 'coordinator@assessment.example.com'
        """)
        sent_emails = cur.fetchall()
        noise_recipients = [
            # Mismatched domains
            "all-staff@university.edu", "faculty@university.edu",
            "all@university.edu", "researchers@university.edu",
            # Same domain but wrong recipient
            "admin@assessment.example.com", "it@assessment.example.com",
            "facilities@assessment.example.com", "grants@assessment.example.com",
        ]
        for email_row in sent_emails:
            to_str = str(email_row[0]).lower()
            for noise in noise_recipients:
                if noise in to_str:
                    check("No email sent to noise recipients", False,
                          f"Sent to noise recipient: {noise}")
                    cur.close(); conn.close()
                    return
        check("No email sent to noise recipients", True)
        cur.close(); conn.close()
    except Exception as e:
        check("Reverse validation", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("TERMINAL-CANVAS-GFORM-GCAL-EXCEL-EMAIL - EVALUATION")
    print("=" * 70)

    check_excel(args.agent_workspace)
    check_gform()
    check_gcal()
    check_email()
    check_reverse_validation(args.agent_workspace)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
