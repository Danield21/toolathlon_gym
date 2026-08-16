"""
Evaluation for yt-engagement-analysis-gsheet-word task.

Checks:
1. GSheet 'Fireship Engagement Analysis' exists with Engagement_Data sheet
2. Engagement_Data has correct columns and at least 100 rows (all 104 Fireship videos)
3. Monthly_Summary sheet exists with correct columns
4. Engagement_Rate_Pct values are computed correctly (spot check)
5. Word document Engagement_Analysis_Report.docx exists with required sections
6. Email sent to research@company.com with correct subject
"""
import json
import os
import sys
from argparse import ArgumentParser

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
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


def check_gsheet():
    print("\n=== Check 1: Google Sheet 'Fireship Engagement Analysis' ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE title ILIKE '%%fireship engagement%%'
        ORDER BY id DESC LIMIT 1
    """)
    ss = cur.fetchone()
    record("Spreadsheet 'Fireship Engagement Analysis' exists", ss is not None,
           "No matching spreadsheet found")

    if not ss:
        cur.close()
        conn.close()
        return

    ss_id = ss[0]

    # Check Engagement_Data sheet
    cur.execute("""
        SELECT id, title FROM gsheet.sheets
        WHERE spreadsheet_id = %s AND title ILIKE '%%engagement_data%%'
    """, (ss_id,))
    eng_sheet = cur.fetchone()
    record("Engagement_Data sheet exists", eng_sheet is not None,
           f"Sheets in spreadsheet: {ss_id}")

    if eng_sheet:
        sheet_id = eng_sheet[0]
        cur.execute("SELECT COUNT(DISTINCT row_index) FROM gsheet.cells WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 1", (ss_id, sheet_id,))
        unique_rows = cur.fetchone()[0]
        # Query expected video count dynamically (104 Fireship videos)
        try:
            cur.execute("SELECT COUNT(*) FROM youtube.videos WHERE channel_title ILIKE '%fireship%'")
            expected_count = cur.fetchone()[0]
        except Exception:
            expected_count = 104
        # Allow exactly the expected count (no off-by-one tolerance for 'all videos')
        record(
            f"Engagement_Data has exactly {expected_count} data rows (all Fireship videos)",
            unique_rows == expected_count,
            f"Found {unique_rows} unique data rows; expected {expected_count}",
        )

        # Check column headers - all 7 required
        cur.execute("""
            SELECT col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index = 1
            ORDER BY col_index
        """, (ss_id, sheet_id,))
        headers = [r[1] for r in cur.fetchall()]
        headers_lower = [h.strip().lower() if h else "" for h in headers]
        for required_col in [
            "video_id",
            "title",
            "published_month",
            "duration_sec",
            "view_count",
            "like_count",
            "engagement_rate_pct",
        ]:
            ok = any(required_col == h for h in headers_lower)
            record(
                f"Engagement_Data has '{required_col}' column",
                ok,
                f"Headers: {headers}",
            )

        # Validate per-video engagement rate (compare top-3 by er)
        try:
            cur.execute(
                """
                SELECT video_id, title, view_count, like_count,
                       (like_count::numeric / NULLIF(view_count,0) * 100) AS er
                FROM youtube.videos
                WHERE channel_title ILIKE '%fireship%'
                ORDER BY er DESC NULLS LAST
                LIMIT 3
                """
            )
            top3 = cur.fetchall()
        except Exception:
            top3 = []

        # Build a row map: row_index -> {col_idx: value}
        cur.execute(
            """
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 1
            """,
            (ss_id, sheet_id),
        )
        cell_map = {}
        for ri, ci, val in cur.fetchall():
            cell_map.setdefault(ri, {})[ci] = val

        # Map column name -> col index
        header_to_idx = {h: i for i, h in enumerate(headers_lower) if h}
        vid_col = header_to_idx.get("video_id")
        er_col = header_to_idx.get("engagement_rate_pct")

        if vid_col is not None and er_col is not None:
            # Build video_id -> er
            agent_er = {}
            for ri, cells in cell_map.items():
                vid = cells.get(vid_col)
                er = cells.get(er_col)
                if vid is not None:
                    agent_er[str(vid).strip()] = er

            for vid, title, vc, lc, er in top3:
                got = agent_er.get(str(vid).strip())
                expected = float(er)
                try:
                    got_f = float(str(got).replace("%", "").strip())
                except (TypeError, ValueError, AttributeError):
                    got_f = None
                ok = got_f is not None and abs(got_f - expected) <= 0.05
                record(
                    f"Engagement_Data video '{title[:40]}' engagement_rate ~{expected:.3f}",
                    ok,
                    f"got {got}",
                )

            # Validate Engagement_Data is sorted descending by Engagement_Rate_Pct.
            sorted_rows = sorted(cell_map.items(), key=lambda kv: kv[0])
            er_values = []
            for _ri, cells in sorted_rows:
                try:
                    er_v = float(str(cells.get(er_col)).replace("%", "").strip())
                    er_values.append(er_v)
                except (TypeError, ValueError, AttributeError):
                    pass
            sort_ok = len(er_values) >= 2 and all(
                er_values[i] >= er_values[i + 1] - 1e-6 for i in range(len(er_values) - 1)
            )
            record(
                "Engagement_Data sorted descending by Engagement_Rate_Pct",
                sort_ok,
                f"First few values: {er_values[:5]}",
            )

    # Check Monthly_Summary sheet
    cur.execute("""
        SELECT id, title FROM gsheet.sheets
        WHERE spreadsheet_id = %s AND title ILIKE '%%monthly_summary%%'
    """, (ss_id,))
    monthly_sheet = cur.fetchone()
    record("Monthly_Summary sheet exists", monthly_sheet is not None)

    if monthly_sheet:
        ms_id = monthly_sheet[0]
        cur.execute("SELECT COUNT(DISTINCT row_index) FROM gsheet.cells WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 1", (ss_id, ms_id,))
        month_rows = cur.fetchone()[0]
        # Query expected month count dynamically
        try:
            cur.execute(
                """
                SELECT COUNT(DISTINCT TO_CHAR(published_at, 'YYYY-MM'))
                FROM youtube.videos
                WHERE channel_title ILIKE '%fireship%'
                """
            )
            expected_months = cur.fetchone()[0]
        except Exception:
            expected_months = 10
        # Strict: must match exactly
        record(
            f"Monthly_Summary has exactly {expected_months} distinct month rows",
            month_rows == expected_months,
            f"Found {month_rows} month rows; expected {expected_months}",
        )

        # Verify Monthly_Summary is sorted ascending by Month.
        cur.execute(
            """
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index = 1
            ORDER BY col_index
            """,
            (ss_id, ms_id),
        )
        ms_headers = [r[2] for r in cur.fetchall()]
        ms_headers_lower = [h.strip().lower() if h else "" for h in ms_headers]
        ms_h_to_idx = {h: i for i, h in enumerate(ms_headers_lower) if h}
        month_col = ms_h_to_idx.get("month")
        if month_col is not None:
            cur.execute(
                """
                SELECT row_index, col_index, value FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 1 AND col_index = %s
                ORDER BY row_index
                """,
                (ss_id, ms_id, month_col),
            )
            month_vals = [str(r[2]).strip() for r in cur.fetchall() if r[2]]
            ms_sort_ok = (
                len(month_vals) >= 2
                and all(month_vals[i] <= month_vals[i + 1] for i in range(len(month_vals) - 1))
            )
            record(
                "Monthly_Summary sorted ascending by Month",
                ms_sort_ok,
                f"Months: {month_vals[:5]}",
            )

    cur.close()
    conn.close()


def check_word(agent_workspace):
    print("\n=== Check 2: Engagement_Analysis_Report.docx ===")
    doc_path = os.path.join(agent_workspace, "Engagement_Analysis_Report.docx")
    if not os.path.exists(doc_path):
        record("Engagement_Analysis_Report.docx exists", False, f"Not found at {doc_path}")
        return
    record("Engagement_Analysis_Report.docx exists", True)

    try:
        import docx
        doc = docx.Document(doc_path)
    except Exception as e:
        record("Word document readable", False, str(e))
        return
    record("Word document readable", True)

    full_text = "\n".join(p.text for p in doc.paragraphs).lower()

    required_sections = ["overview", "methodology", "key findings", "monthly trends", "conclusions"]
    for s in required_sections:
        record(f"Document has section '{s}'", s in full_text)

    # Methodology must explain engagement rate formula. Task: like_count / view_count * 100
    has_methodology_formula = (
        ("like_count" in full_text or "like count" in full_text)
        and ("view_count" in full_text or "view count" in full_text or "views" in full_text)
    )
    record("Methodology section explains engagement rate formula (likes/views)", has_methodology_formula)

    # Key findings should mention top-3 videos by engagement (dynamic query)
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute(
            """
            SELECT title FROM youtube.videos
            WHERE channel_title ILIKE '%fireship%'
            ORDER BY (like_count::numeric / NULLIF(view_count,0)) DESC NULLS LAST
            LIMIT 3
            """
        )
        top3_titles = [r[0] for r in cur.fetchall()]
        cur.close()
        conn.close()
    except Exception:
        top3_titles = []

    # Each of top 3 must be referenced. Use a few distinctive content words from the
    # title; tolerant of unicode ellipsis vs '...' differences.
    def _title_match(title, hay):
        # Build content words longer than 3 chars, removing punctuation.
        import re as _re
        words = [w for w in _re.split(r"\W+", title.lower()) if len(w) > 3]
        if not words:
            return False
        # Need at least 2 distinctive words present
        count = sum(1 for w in words if w in hay)
        return count >= 2

    found_titles = sum(1 for t in top3_titles if _title_match(t, full_text))
    record(
        f"Key Findings mentions top 3 videos (matched {found_titles}/3)",
        found_titles >= 3,
        f"Top titles: {top3_titles}",
    )

    has_percentage = "%" in full_text or "percent" in full_text
    record("Document contains engagement rate percentages", has_percentage)


def check_email():
    print("\n=== Check 3: Email to research@company.com ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT to_addr, subject, body_text FROM email.messages
        WHERE to_addr::text ILIKE '%research@company.com%'
        AND subject ILIKE '%engagement%'
        ORDER BY id DESC LIMIT 5
    """)
    emails = cur.fetchall()
    cur.close()
    conn.close()

    record("Email to research@company.com with engagement subject exists",
           len(emails) > 0, "No matching email found")

    if emails:
        to_addr, subject, body = emails[0]
        record("Email subject contains 'Fireship YouTube Engagement Analysis'",
               "fireship" in subject.lower() and "engagement" in subject.lower(),
               f"Subject: {subject}")

        body_lower = (body or "").lower()
        # Body must mention total video count near 104 (dynamic-aware)
        try:
            conn2 = psycopg2.connect(**DB_CONFIG)
            cur2 = conn2.cursor()
            cur2.execute("SELECT COUNT(*) FROM youtube.videos WHERE channel_title ILIKE '%fireship%'")
            expected_count = cur2.fetchone()[0]
            cur2.close()
            conn2.close()
        except Exception:
            expected_count = 104
        has_count = str(expected_count) in body_lower
        record(
            f"Email body mentions total video count {expected_count}",
            has_count,
            f"Body excerpt: {body_lower[:200]}",
        )

        # Body must reference at least one top engaging video by name (dynamic)
        try:
            conn2 = psycopg2.connect(**DB_CONFIG)
            cur2 = conn2.cursor()
            cur2.execute(
                """
                SELECT title FROM youtube.videos
                WHERE channel_title ILIKE '%fireship%'
                ORDER BY (like_count::numeric / NULLIF(view_count,0)) DESC NULLS LAST
                LIMIT 3
                """
            )
            top_titles = [r[0] for r in cur2.fetchall()]
            cur2.close()
            conn2.close()
        except Exception:
            top_titles = []

        def _body_title_match(title, hay):
            import re as _re
            words = [w for w in _re.split(r"\W+", title.lower()) if len(w) > 3]
            if not words:
                return False
            count = sum(1 for w in words if w in hay)
            return count >= 2

        ref_count = sum(1 for t in top_titles if _body_title_match(t, body_lower))
        record(
            "Email body references at least one top engaging video by name",
            ref_count >= 1,
            f"Top titles: {top_titles}, body excerpt: {body_lower[:200]}",
        )


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_word(args.agent_workspace)
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
