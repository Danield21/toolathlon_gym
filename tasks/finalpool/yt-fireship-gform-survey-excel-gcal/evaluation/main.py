"""
Evaluation for yt-fireship-gform-survey-excel-gcal task.

Checks:
1. Community_Report.xlsx exists with "Top_Videos" sheet having >= 6 data rows
2. Engagement_Rate column exists in Top_Videos or Engagement_Analysis has rates
3. Engagement_Analysis sheet exists with >= 3 rows
4. GForm created with title containing "Survey", "Fireship", or "Community"
5. GForm has >= 4 questions
6. GCal has new event in April 2026 with "Community" or "Standup" in summary (not the noise Q&A at 16:00)
7. Email sent to community@devclub.io
"""
import os
import sys
import json
from argparse import ArgumentParser

import psycopg2
import openpyxl

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
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
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=1.0):
    try: return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError): return False


def str_match(a, b):
    if a is None or b is None: return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def check_excel(agent_workspace, groundtruth_workspace="."):
    print("\n=== Check 1-3: Community_Report.xlsx ===")

    xlsx_path = os.path.join(agent_workspace, "Community_Report.xlsx")
    if not os.path.exists(xlsx_path):
        record("Community_Report.xlsx exists", False, f"Not found at {xlsx_path}")
        record("Top_Videos sheet has >= 6 data rows", False, "File missing")
        record("Engagement_Rate column or analysis exists", False, "File missing")
        record("Engagement_Analysis sheet has >= 3 rows", False, "File missing")
        return
    record("Community_Report.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path)
    except Exception as e:
        record("Excel readable", False, str(e))
        return

    sheet_names_lower = {s.lower(): s for s in wb.sheetnames}

    # Top_Videos sheet
    top_key = next((sheet_names_lower[k] for k in sheet_names_lower
                    if "video" in k or "top" in k), None)
    if not top_key:
        record("Top_Videos sheet has >= 6 data rows", False,
               f"No Top_Videos sheet. Sheets: {wb.sheetnames}")
        record("Engagement_Rate column or analysis exists", False, "Sheet missing")
    else:
        ws = wb[top_key]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)] if rows else []
        record("Top_Videos sheet has 8 data rows", len(data_rows) == 8,
               f"Found {len(data_rows)} data rows, expected 8")
        if rows:
            headers = [str(c).strip().lower() if c else "" for c in rows[0]]
            has_engagement = any("engagement" in h or "rate" in h for h in headers)
            record("Engagement_Rate column exists in Top_Videos", has_engagement,
                   f"Headers: {rows[0]}")
        else:
            record("Engagement_Rate column exists in Top_Videos", False, "Sheet empty")

    # Engagement_Analysis sheet
    eng_key = next((sheet_names_lower[k] for k in sheet_names_lower
                    if "engagement" in k or "analysis" in k), None)
    if not eng_key:
        record("Engagement_Analysis sheet has >= 3 rows", False,
               f"No Engagement_Analysis sheet. Sheets: {wb.sheetnames}")
    else:
        ws2 = wb[eng_key]
        rows2 = list(ws2.iter_rows(values_only=True))
        data_rows2 = [r for r in rows2[1:] if any(c for c in r)] if rows2 else []
        record("Engagement_Analysis sheet has >= 4 rows", len(data_rows2) >= 4,
               f"Found {len(data_rows2)} data rows, expected at least 4")

    # --- Groundtruth XLSX value comparison ---
    # Match rows by Video_ID (Top_Videos) or Topic (Engagement_Analysis) rather
    # than by positional index — agents may legitimately sort with different
    # tie-breakers. Skip subjective columns (Topic_Tags) since the agent and GT
    # may pick different tags for multi-topic videos.
    gt_path = os.path.join(groundtruth_workspace, "Community_Report.xlsx")
    if os.path.isfile(gt_path):
        import openpyxl as opx
        gt_wb = opx.load_workbook(gt_path, data_only=True)
        try:
            a_wb = opx.load_workbook(os.path.join(agent_workspace, "Community_Report.xlsx"), data_only=True)
        except Exception:
            a_wb = None
        if a_wb:
            for gt_sname in gt_wb.sheetnames:
                gt_ws = gt_wb[gt_sname]
                a_ws = None
                for asn in a_wb.sheetnames:
                    if asn.strip().lower() == gt_sname.strip().lower():
                        a_ws = a_wb[asn]; break
                if a_ws is None:
                    record(f"GT sheet '{gt_sname}' exists in agent xlsx", False, f"Available: {a_wb.sheetnames}")
                    continue
                gt_headers = [str(c.value).strip().lower() if c.value else "" for c in next(gt_ws.iter_rows(max_row=1))]
                a_headers = [str(c.value).strip().lower() if c.value else "" for c in next(a_ws.iter_rows(max_row=1))]
                gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
                a_rows = [r for r in a_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
                record(f"GT '{gt_sname}' row count", len(a_rows) == len(gt_rows),
                       f"Expected {len(gt_rows)}, got {len(a_rows)}")
                # Determine PK column for stable row matching
                pk_candidates = ["video_id", "topic"]
                pk_col = None
                for cand in pk_candidates:
                    if cand in gt_headers and cand in a_headers:
                        pk_col = cand
                        break
                if pk_col is None:
                    # No stable key — skip per-row comparison gracefully
                    record(f"GT '{gt_sname}' has stable PK column", False,
                           f"No video_id / topic header found: {gt_headers}")
                    continue
                gt_pk_idx = gt_headers.index(pk_col)
                a_pk_idx = a_headers.index(pk_col)
                # Skip subjective columns (Topic_Tags) when comparing values
                skip_headers = {"topic_tags", "topic tags"}
                # Build agent lookup keyed on PK (lowercased trimmed)
                a_lookup = {}
                for r in a_rows:
                    if r and len(r) > a_pk_idx and r[a_pk_idx] is not None:
                        a_lookup[str(r[a_pk_idx]).strip().lower()] = r
                for gt_row in gt_rows:
                    if not gt_row or len(gt_row) <= gt_pk_idx or gt_row[gt_pk_idx] is None:
                        continue
                    pk_val = str(gt_row[gt_pk_idx]).strip().lower()
                    a_row = a_lookup.get(pk_val)
                    found = a_row is not None
                    record(f"GT '{gt_sname}' pk={pk_val} present", found)
                    if not found:
                        continue
                    ok = True
                    fail_detail = ""
                    for ci, gh in enumerate(gt_headers):
                        if ci >= len(gt_row):
                            continue
                        gv = gt_row[ci]
                        if gv is None or gh in skip_headers:
                            continue
                        # Locate agent column by header name
                        if gh not in a_headers:
                            continue
                        a_ci = a_headers.index(gh)
                        if a_ci >= len(a_row):
                            continue
                        av = a_row[a_ci]
                        if isinstance(gv, (int, float)):
                            # Engagement_Rate / averages: 5% relative or 0.1 abs
                            if "rate" in gh or "engagement" in gh:
                                tol = max(abs(gv) * 0.05, 0.1)
                            else:
                                tol = max(abs(gv) * 0.01, 0.5)
                            if not num_close(av, gv, tol):
                                ok = False
                                fail_detail = f"col '{gh}': gt={gv}, agent={av}"
                                break
                        else:
                            if not str_match(av, gv):
                                ok = False
                                fail_detail = f"col '{gh}': gt={gv!r}, agent={av!r}"
                                break
                    record(f"GT '{gt_sname}' pk={pk_val} values", ok, fail_detail)
            a_wb.close()
        gt_wb.close()


def check_gform():
    print("\n=== Check 4-5: Google Form ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        with conn.cursor() as cur:
            cur.execute("""
                SELECT f.id, f.title FROM gform.forms f
                WHERE f.title ILIKE '%Fireship Community Preference Survey%'
                   OR (f.title ILIKE '%Fireship%' AND f.title ILIKE '%Community%' AND f.title ILIKE '%Survey%')
            """)
            forms = cur.fetchall()
            if not forms:
                record("GForm 'Fireship Community Preference Survey' exists", False,
                       "No matching form found")
                record("GForm has 5 questions", False, "Form missing")
                conn.close()
                return
            record("GForm 'Fireship Community Preference Survey' exists", True,
                   f"Found: {[f[1] for f in forms]}")

            form_id = forms[0][0]
            cur.execute("""
                SELECT title FROM gform.questions WHERE form_id = %s ORDER BY id
            """, (form_id,))
            q_titles = [r[0] for r in cur.fetchall()]
            record("GForm has 5 questions", len(q_titles) == 5,
                   f"Found {len(q_titles)} questions, expected 5")

            # Validate question wording — each required question must be present.
            # Use distinctive keyword pairs to tolerate phrasing variation.
            required_keywords = [
                ("topic", "interest"),       # Q1: Which Fireship topic interests you most?
                ("often", "watch"),          # Q2: How often do you watch Fireship?
                ("format", "prefer"),        # Q3: What format do you prefer?
                ("role",),                   # Q4: Your current role
                ("topic", "next"),           # Q5: What topic should Fireship cover next?
            ]
            joined = " || ".join((q or "").lower() for q in q_titles)
            for kw_set in required_keywords:
                ok = any(all(kw in (q or "").lower() for kw in kw_set) for q in q_titles)
                record(
                    f"GForm has question matching keywords {kw_set}",
                    ok,
                    f"Question titles: {q_titles}",
                )
        conn.close()
    except Exception as e:
        record("GForm check", False, str(e))


def check_gcal():
    print("\n=== Check 6: GCal Community Standup in April 2026 ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        with conn.cursor() as cur:
            cur.execute("""
                SELECT summary, start_datetime, end_datetime FROM gcal.events
                WHERE (summary ILIKE '%Community%' OR summary ILIKE '%Standup%')
                  AND start_datetime >= '2026-04-01'
                  AND start_datetime < '2026-05-01'
                  AND summary NOT ILIKE '%Q&A%'
                ORDER BY start_datetime
            """)
            events = cur.fetchall()
        conn.close()
        record("GCal has new Community/Standup event in April 2026 (not Q&A noise)",
               len(events) > 0, f"Found {len(events)} events")

        # Find an event matching exact title 'Community Standup' on 2026-04-01 18:00-19:00.
        target_evt = None
        for s, sd, ed in events:
            sl = (s or "").lower()
            if "community" in sl and "standup" in sl:
                if sd and sd.date().isoformat() == "2026-04-01":
                    target_evt = (s, sd, ed)
                    break
        record(
            "GCal 'Community Standup' on 2026-04-01 exists",
            target_evt is not None,
            f"Candidates: {[(e[0], e[1]) for e in events]}",
        )
        if target_evt:
            s, sd, ed = target_evt
            time_ok = sd.hour == 18 and sd.minute == 0
            if ed is not None:
                time_ok = time_ok and (ed.hour == 19 and ed.minute == 0)
            record(
                "Community Standup time is 18:00-19:00",
                time_ok,
                f"start={sd}, end={ed}",
            )
    except Exception as e:
        record("GCal check", False, str(e))


def check_email():
    print("\n=== Check 7: Email sent to community@devclub.io ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        with conn.cursor() as cur:
            cur.execute("""
                SELECT subject, body_text FROM email.messages
                WHERE to_addr::text ILIKE '%community@devclub.io%'
                  AND from_addr != 'community@devclub.io'
                ORDER BY id DESC
            """)
            rows = cur.fetchall()
        conn.close()
        record("Email sent to community@devclub.io", len(rows) > 0, f"Found {len(rows)}")
        if rows:
            # Subject must reference engagement / monthly report
            subj_ok = False
            for s, _ in rows:
                sl = (s or "").lower()
                if "engagement" in sl or ("monthly" in sl and "report" in sl):
                    subj_ok = True; break
            record("Email subject references engagement/monthly report", subj_ok,
                   f"Subjects: {[r[0] for r in rows]}")
            # Body must mention survey AND standup explicitly
            body_ok = False
            for _, b in rows:
                bl = (b or "").lower()
                if ("survey" in bl) and ("standup" in bl):
                    body_ok = True; break
            record("Email body mentions survey + standup", body_ok,
                   f"First body excerpt: {(rows[0][1] or '')[:200]}")
    except Exception as e:
        record("Email check", False, str(e))


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    print(f"Running evaluation for yt-fireship-gform-survey-excel-gcal")
    print(f"Agent workspace: {agent_workspace}")

    check_excel(agent_workspace, groundtruth_workspace)
    check_gform()
    check_gcal()
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
