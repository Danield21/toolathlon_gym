"""Evaluation for terminal-sf-canvas-word-gform-email.
Checks:
1. Training_Effectiveness_Report.docx content
2. Google Form "Training Feedback Survey" with 5 questions
3. Emails to hr_director and training_team
4. Script files exist (training_matches.py, effectiveness_analysis.py, survey_analysis.py)
5. JSON output files exist
"""
import argparse
import json
import os
import sys

import psycopg2
from docx import Document

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
          user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_ONLY_FAIL = 0


def check(name, condition, detail="", runtime_only=False):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_ONLY_FAIL
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        if runtime_only:
            RUNTIME_ONLY_FAIL += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def num_close(a, b, tol=2.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_values():
    """Query DB for expected values."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Course avg scores
    cur.execute("""
        SELECT a.course_id, ROUND(AVG(s.score)::numeric, 2)
        FROM canvas.assignments a
        JOIN canvas.submissions s ON s.assignment_id = a.id
        WHERE a.course_id IN (9, 10) AND s.score IS NOT NULL
        GROUP BY a.course_id
    """)
    course_avgs = {int(r[0]): float(r[1]) for r in cur.fetchall()}

    # Enrollment counts
    cur.execute("""
        SELECT course_id, COUNT(DISTINCT user_id)
        FROM canvas.enrollments
        WHERE course_id IN (9, 10) AND type='StudentEnrollment'
        GROUP BY course_id
    """)
    enrollments = {int(r[0]): r[1] for r in cur.fetchall()}

    # SF dept ratings
    cur.execute("""
        SELECT "DEPARTMENT", ROUND(AVG("PERFORMANCE_RATING")::numeric, 2)
        FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
        WHERE "DEPARTMENT" IN ('Engineering', 'R&D')
        GROUP BY "DEPARTMENT"
    """)
    dept_ratings = {r[0]: float(r[1]) for r in cur.fetchall()}

    cur.close()
    conn.close()

    c9_avg = course_avgs.get(9, 69.59)
    c10_avg = course_avgs.get(10, 71.53)
    eng_rating = dept_ratings.get("Engineering", 3.21)
    rnd_rating = dept_ratings.get("R&D", 3.20)
    eng_impr = eng_rating - 3.00
    rnd_impr = rnd_rating - 2.95
    avg_impr = (eng_impr + rnd_impr) / 2
    overall_avg_score = (c9_avg + c10_avg) / 2

    return {
        "c9_avg": c9_avg, "c10_avg": c10_avg,
        "eng_enrolled": enrollments.get(9, 1938),
        "rnd_enrolled": enrollments.get(10, 1803),
        "eng_rating": eng_rating, "rnd_rating": rnd_rating,
        "eng_impr": eng_impr, "rnd_impr": rnd_impr,
        "avg_impr": avg_impr,
        "overall_avg_score": overall_avg_score,
    }


def check_word(workspace):
    """Check Training_Effectiveness_Report.docx."""
    print("\n=== Check 1: Word Document ===")
    path = os.path.join(workspace, "Training_Effectiveness_Report.docx")
    if not os.path.exists(path):
        check("Word document exists", False, f"Not found: {path}")
        return
    check("Word document exists", True)

    doc = Document(path)
    full_text = " ".join(p.text for p in doc.paragraphs).lower()

    check("Has Executive Summary section",
          "executive summary" in full_text)
    check("Has Methodology section",
          "methodology" in full_text)
    check("Has Performance Impact section",
          "performance impact" in full_text or "impact analysis" in full_text)
    check("Has Survey Findings section",
          "survey findings" in full_text or "survey" in full_text)
    check("Has ROI section",
          "roi" in full_text or "return on investment" in full_text)
    check("Has Recommendations section",
          "recommendation" in full_text)
    check("Mentions Engineering department",
          "engineering" in full_text)
    check("Mentions R&D department",
          "r&d" in full_text or "r & d" in full_text)
    check("Mentions Data-Driven Design",
          "data-driven design" in full_text or "data driven design" in full_text)
    check("Has substantial content", len(full_text) > 500, f"Length: {len(full_text)}")

    ev = get_expected_values()

    # Check that key numbers appear in text
    check("Mentions course 9 avg score",
          str(round(ev["c9_avg"], 1)) in full_text or str(round(ev["c9_avg"], 2)) in full_text
          or str(int(round(ev["c9_avg"]))) in full_text,
          f"Expected ~{ev['c9_avg']:.2f}")
    check("Mentions course 10 avg score",
          str(round(ev["c10_avg"], 1)) in full_text or str(round(ev["c10_avg"], 2)) in full_text
          or str(int(round(ev["c10_avg"]))) in full_text,
          f"Expected ~{ev['c10_avg']:.2f}")

    # Check conditional recommendation
    if ev["avg_impr"] < 0.15:
        check("Recommends restructuring (improvement < 0.15)",
              "restructur" in full_text,
              f"Avg improvement: {ev['avg_impr']:.2f}")
    else:
        check("Recommends expanding (improvement >= 0.15)",
              "expand" in full_text,
              f"Avg improvement: {ev['avg_impr']:.2f}")

    # Survey-derived analysis must match seeded responses (V2 trap)
    survey_ev = _get_survey_expected()
    avg_sat = survey_ev["avg_sat"]
    rec_pct = survey_ev["recommend_pct"]
    pref = survey_ev["preferred_format"]

    # Accept rounded forms: e.g. 3.87 / 3.9 / 3.86
    sat_candidates = {
        f"{avg_sat:.2f}",
        f"{avg_sat:.1f}",
        f"{round(avg_sat, 1):.1f}",
    }
    check("Word doc reports correct avg satisfaction",
          any(c in full_text for c in sat_candidates),
          f"Expected avg ~{avg_sat:.2f}; candidates {sorted(sat_candidates)}")

    # Recommendation rate: accept either 73.3 or 73 or 11 of 15 phrasing
    rec_candidates = {
        f"{rec_pct:.1f}",
        f"{int(round(rec_pct))}",
        f"{survey_ev['recommend_yes']} ",
    }
    rec_ok = any(c in full_text for c in rec_candidates)
    check("Word doc reports correct recommendation stat",
          rec_ok,
          f"Expected ~{rec_pct:.1f}% / {survey_ev['recommend_yes']} respondents")

    # Most preferred format must match seeded mode
    check(f"Word doc reports preferred format '{pref}'",
          pref in full_text,
          f"Expected '{pref}' to appear in survey findings")


def check_gform():
    """Check Google Form creation.

    Note: preprocess pre-seeds a form titled exactly 'Training Feedback Survey'
    with 5 questions and 15 responses. The agent is expected to ANALYZE these
    responses (not necessarily create a new form). We resolve the form by
    exact title match; if multiple, we pick the one with 5 questions and the
    most responses (the seeded one).
    """
    print("\n=== Check 2: Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()

    # 1) Exact title match
    candidate_ids = []
    for fid, title in forms:
        if (title or "").strip().lower() == "training feedback survey":
            candidate_ids.append(fid)

    form_id = None
    if candidate_ids:
        # Pick the candidate with exactly 5 questions AND the most responses.
        best = None  # (resp_count, fid)
        for fid in candidate_ids:
            cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (fid,))
            q_count = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM gform.responses WHERE form_id = %s", (fid,))
            r_count = cur.fetchone()[0]
            if q_count == 5 and (best is None or r_count > best[0]):
                best = (r_count, fid)
        if best is None:
            # Fall back: any candidate with 5 questions (or just the first)
            for fid in candidate_ids:
                cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (fid,))
                if cur.fetchone()[0] == 5:
                    form_id = fid
                    break
            if form_id is None:
                form_id = candidate_ids[0]
        else:
            form_id = best[1]

    check("Training Feedback Survey form exists (exact title)", form_id is not None,
          f"Forms: {[f[1] for f in forms]}", runtime_only=True)

    if form_id:
        cur.execute("SELECT title, question_type FROM gform.questions WHERE form_id = %s ORDER BY position",
                    (form_id,))
        questions = cur.fetchall()
        check("Form has exactly 5 questions", len(questions) == 5,
              f"Found {len(questions)}", runtime_only=True)

        if len(questions) >= 5:
            q_titles = [q[0].lower() for q in questions]
            check("Q1 about satisfaction",
                  any("satisfaction" in t or "rating" in t for t in q_titles[:2]),
                  f"Q titles: {q_titles}", runtime_only=True)
            check("Q4 about recommendation",
                  any("recommend" in t for t in q_titles),
                  f"Q titles: {q_titles}", runtime_only=True)
            check("Q5 about format",
                  any("format" in t for t in q_titles),
                  f"Q titles: {q_titles}", runtime_only=True)

        # Check responses exist
        cur.execute("SELECT COUNT(*) FROM gform.responses WHERE form_id = %s", (form_id,))
        resp_count = cur.fetchone()[0]
        check("Form has responses", resp_count >= 15,
              f"Found {resp_count} responses", runtime_only=True)

    cur.close()
    conn.close()


def _get_survey_expected():
    """Compute expected survey-derived stats from seeded responses.

    Resolves the 'Training Feedback Survey' form (exact title, 5 questions,
    most responses) and computes:
      - avg satisfaction (rounded to 2 decimals)
      - recommendation Yes count + percentage
      - most preferred training format

    Falls back to seeded values used by preprocess if anything goes wrong.
    """
    fallback = {
        "avg_sat": 3.87,
        "recommend_yes": 11,
        "recommend_pct": 73.3,
        "preferred_format": "hybrid",
    }
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("SELECT id FROM gform.forms WHERE LOWER(title) = 'training feedback survey'")
        rows = cur.fetchall()
        if not rows:
            cur.close()
            conn.close()
            return fallback

        # Pick form with 5 questions + most responses
        best = None
        for (fid,) in rows:
            cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (fid,))
            q_count = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM gform.responses WHERE form_id = %s", (fid,))
            r_count = cur.fetchone()[0]
            if q_count == 5 and (best is None or r_count > best[0]):
                best = (r_count, fid)
        if best is None:
            cur.close()
            conn.close()
            return fallback
        form_id = best[1]

        # Identify question IDs by title prefix
        cur.execute("SELECT id, title FROM gform.questions WHERE form_id = %s ORDER BY position",
                    (form_id,))
        qrows = cur.fetchall()
        if len(qrows) < 5:
            cur.close()
            conn.close()
            return fallback
        sat_qid = qrows[0][0]
        rec_qid = qrows[3][0]
        fmt_qid = qrows[4][0]

        cur.execute("SELECT answers FROM gform.responses WHERE form_id = %s", (form_id,))
        sat_vals = []
        rec_vals = []
        fmt_vals = []
        for (ans,) in cur.fetchall():
            try:
                if isinstance(ans, str):
                    a = json.loads(ans)
                else:
                    a = ans or {}
            except Exception:
                continue
            try:
                sat_vals.append(float(a.get(sat_qid)))
            except (TypeError, ValueError):
                pass
            r = a.get(rec_qid)
            if isinstance(r, str):
                rec_vals.append(r.strip().lower())
            f = a.get(fmt_qid)
            if isinstance(f, str):
                fmt_vals.append(f.strip().lower())

        cur.close()
        conn.close()

        if not sat_vals or not rec_vals or not fmt_vals:
            return fallback

        avg = round(sum(sat_vals) / len(sat_vals), 2)
        yes = sum(1 for r in rec_vals if r == "yes")
        pct = round(yes / len(rec_vals) * 100, 1)
        # Most preferred format
        from collections import Counter
        fmt = Counter(fmt_vals).most_common(1)[0][0]
        return {
            "avg_sat": avg,
            "recommend_yes": yes,
            "recommend_pct": pct,
            "preferred_format": fmt,
        }
    except Exception as e:
        print(f"[fallback] _get_survey_expected: {e}")
        return fallback


def check_emails():
    """Check emails to hr_director and training_team."""
    print("\n=== Check 3: Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()

    # Check email to hr_director
    hr_email = None
    training_email = None
    for subj, from_addr, to_addr, body in all_emails:
        to_str = str(to_addr).lower() if to_addr else ""
        if "hr_director" in to_str:
            hr_email = (subj, from_addr, to_addr, body)
        if "training_team" in to_str:
            training_email = (subj, from_addr, to_addr, body)

    check("Email sent to hr_director@company.com", hr_email is not None,
          f"Total emails: {len(all_emails)}", runtime_only=True)
    if hr_email:
        subj, from_addr, to_addr, body = hr_email
        # Stricter subject check per task: "Training Effectiveness Analysis - Key Findings"
        subj_l = (subj or "").lower()
        check("HR email subject is 'Training Effectiveness Analysis - Key Findings'",
              "training effectiveness" in subj_l and "key findings" in subj_l,
              f"Subject: {subj}", runtime_only=True)
        check("HR email from training_analytics@company.com",
              "training_analytics" in (from_addr or "").lower(),
              f"From: {from_addr}", runtime_only=True)
        body_lower = (body or "").lower()
        check("HR email mentions performance rating",
              ("performance" in body_lower) and ("rating" in body_lower or "improvement" in body_lower),
              "Expected performance/rating in body", runtime_only=True)

    check("Email sent to training_team@company.com", training_email is not None,
          f"Total emails: {len(all_emails)}", runtime_only=True)
    if training_email:
        subj, from_addr, to_addr, body = training_email
        subj_l = (subj or "").lower()
        check("Training email subject is 'Training Feedback Survey Results'",
              "training feedback" in subj_l and "results" in subj_l,
              f"Subject: {subj}", runtime_only=True)
        body_lower = (body or "").lower()
        check("Training email mentions satisfaction score",
              "satisfaction" in body_lower or "rating" in body_lower,
              "Expected satisfaction mention", runtime_only=True)
        check("Training email mentions format / recommend",
              "format" in body_lower or "recommend" in body_lower,
              "Expected format/recommend mention", runtime_only=True)

    cur.close()
    conn.close()


def check_scripts(workspace):
    """Check that required scripts exist."""
    print("\n=== Check 4: Scripts ===")
    for script in ["training_matches.py", "effectiveness_analysis.py", "survey_analysis.py"]:
        path = os.path.join(workspace, script)
        check(f"{script} exists", os.path.exists(path), f"Not found: {path}",
              runtime_only=True)


def check_json_outputs(workspace):
    """Check JSON output files."""
    print("\n=== Check 5: JSON Outputs ===")
    for jfile in ["training_matches.json", "effectiveness_analysis.json", "survey_results.json"]:
        path = os.path.join(workspace, jfile)
        if not os.path.exists(path):
            check(f"{jfile} exists", False, f"Not found: {path}", runtime_only=True)
            continue
        check(f"{jfile} exists", True)
        try:
            with open(path) as f:
                data = json.load(f)
            check(f"{jfile} is valid JSON", True)
            check(f"{jfile} is non-empty", len(data) > 0, "Empty JSON")
        except json.JSONDecodeError as e:
            check(f"{jfile} is valid JSON", False, str(e))

    # Check effectiveness_analysis.json content
    ea_path = os.path.join(workspace, "effectiveness_analysis.json")
    if os.path.exists(ea_path):
        try:
            with open(ea_path) as f:
                ea = json.load(f)
            ea_str = json.dumps(ea).lower()
            check("effectiveness_analysis mentions Engineering or R&D",
                  "engineering" in ea_str or "r&d" in ea_str or "r_d" in ea_str,
                  "Expected department names")
        except Exception:
            pass


def check_reverse_validation():
    """Verify noise data not misused."""
    print("\n=== Reverse Validation ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # No emails to noise recipients
    cur.execute("""
        SELECT to_addr FROM email.messages
        WHERE from_addr ILIKE '%%training_analytics%%'
    """)
    sent = cur.fetchall()
    noise_addrs = ["all@company.com", "managers@company.com", "leadership@company.com"]
    for row in sent:
        to_str = str(row[0]).lower()
        for noise in noise_addrs:
            if noise in to_str:
                check("No emails sent to noise recipients", False, f"Sent to {noise}")
                cur.close()
                conn.close()
                return
    check("No emails sent to noise recipients", True)

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("TERMINAL-SF-CANVAS-WORD-GFORM-EMAIL - EVALUATION")
    print("=" * 70)

    check_word(args.agent_workspace)
    check_gform()
    check_emails()
    check_scripts(args.agent_workspace)
    check_json_outputs(args.agent_workspace)
    check_reverse_validation()

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    print(f"\nOverall: {PASS_COUNT}/{total} ({accuracy:.1f}%)")

    result = {"total_passed": PASS_COUNT, "total_checks": total, "accuracy": accuracy}
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # Overall: gate on non-runtime-only FAIL_COUNT == 0
    non_runtime_fail = FAIL_COUNT - RUNTIME_ONLY_FAIL
    print(f"Non-runtime fails: {non_runtime_fail}, runtime-only fails: {RUNTIME_ONLY_FAIL}")
    sys.exit(0 if non_runtime_fail == 0 else 1)


if __name__ == "__main__":
    main()
