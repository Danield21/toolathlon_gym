"""Evaluation for sf-support-csat-gform-gsheet-email.

Checks:
1. Google Sheet "Support Center Performance Dashboard" with SLA_Compliance and Summary sheets
2. Google Forms "Customer Support Satisfaction Survey" with 4 questions
3. Email to support-management@company.example.com
"""
import argparse
import json
import os
import re
import sys

import psycopg2

DB = dict(
    host=os.environ.get("PGHOST", "localhost"),
    port=int(os.environ.get("PGPORT", "5432")),
    dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
    user=os.environ.get("PGUSER", "eigent"),
    password=os.environ.get("PGPASSWORD", "camel"),
)

PASS_COUNT = 0
FAIL_COUNT = 0

# Actual SLA data from DB
SLA_DATA = {
    "high":   {"total": 6466,  "compliant": 778,  "rate": 12.03, "csat": 3.26},
    "medium": {"total": 15774, "compliant": 1645, "rate": 10.43, "csat": 3.26},
    "low":    {"total": 9348,  "compliant": 4204, "rate": 44.97, "csat": 3.25},
}


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def _to_float(v):
    """Robustly parse a value to float.

    Handles None / str / int / float. Strips percent signs, currency symbols,
    thousands separators and surrounding whitespace. A cell whose raw value is
    a literal formula like "=12.03" is parsed after stripping the leading '='.
    Returns None when the value cannot be parsed numerically.
    """
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        if s.startswith("="):
            s = s[1:].strip()
        s = s.replace("%", "").replace(",", "").replace("$", "")
        s = s.replace("¥", "").replace("€", "").replace("#", "")
        s = s.replace(" ", " ").strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _num_close(a, b, tol=1.0):
    """Compare two values numerically within tol, with a string fallback.

    When both sides parse as numbers use |a - b| <= tol. When either side is
    not a number, fall back to a case-insensitive string comparison.
    """
    fa = _to_float(a)
    fb = _to_float(b)
    if fa is not None and fb is not None:
        return abs(fa - fb) <= tol
    return str(a).strip().lower() == str(b).strip().lower()


def num_close(a, b, tol=1.0):
    return _num_close(a, b, tol)


def _expected_from_address():
    """Derive the expected sender email from the task's email_config.json.

    The emails MCP has no 'from' parameter: the sender is always the account
    configured in email_config.json (task_dir/email_config.json). Read it so
    the check follows the actual configured identity. Returns None when the
    file is absent/unreadable/empty. When None the sender check is skipped
    (passes) rather than FAILing, because the sender is config-controlled, not
    model-controlled: a missing config file makes the MCP fall back to a
    default account that no correct model can override. This mirrors the
    sibling sf-hr-salary-benchmark evaluator convention.
    """
    cfg_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "..", "email_config.json"
    )
    try:
        with open(cfg_path, encoding="utf-8") as f:
            cfg = json.load(f)
        acct = cfg[0] if isinstance(cfg, list) else cfg
        email = (acct or {}).get("email") or ""
        if email:
            return email.strip().lower()
    except Exception:
        pass
    return None


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE title ILIKE '%support%' AND title ILIKE '%performance%'
    """)
    sheets = cur.fetchall()
    check("Support Center Performance Dashboard spreadsheet exists", len(sheets) >= 1,
          f"Found: {[s[1] for s in sheets]}")

    if not sheets:
        cur.close()
        conn.close()
        return False

    ss_id = sheets[0][0]

    # Check sheets exist
    cur.execute("SELECT title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
    sheet_tabs = [r[0] for r in cur.fetchall()]

    has_sla = any("sla" in t.lower() for t in sheet_tabs)
    has_summary = any("summary" in t.lower() for t in sheet_tabs)
    check("Has SLA_Compliance sheet", has_sla, f"Tabs: {sheet_tabs}")
    check("Has Summary sheet", has_summary, f"Tabs: {sheet_tabs}")
    # Fetch Summary sheet cells if present and verify Best/Worst priority references
    if has_summary:
        cur.execute("""
            SELECT c.value FROM gsheet.cells c
            JOIN gsheet.sheets s ON c.spreadsheet_id = s.spreadsheet_id AND c.sheet_id = s.id
            WHERE c.spreadsheet_id = %s AND s.title ILIKE '%%summary%%'
        """, (ss_id,))
        sum_cells = [str(r[0]).lower() for r in cur.fetchall() if r[0] is not None]
        sum_text = " ".join(sum_cells)
        # task.md: Best_Priority = priority with HIGHEST CSAT; Worst_SLA_Priority = priority with LOWEST compliance rate
        # Dynamically compute from DB (with fallback constants)
        try:
            ck_conn = psycopg2.connect(**DB)
            ck_cur = ck_conn.cursor()
            ck_cur.execute('''
                WITH m AS (
                    SELECT LOWER("PRIORITY") AS priority, "RESPONSE_TIME_HOURS", "CUSTOMER_SATISFACTION"
                    FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
                    WHERE "CUSTOMER_SATISFACTION" IS NOT NULL
                )
                SELECT priority,
                       AVG("CUSTOMER_SATISFACTION") AS avg_csat,
                       AVG(CASE
                           WHEN (priority='high' AND "RESPONSE_TIME_HOURS"<=4) OR
                                (priority='medium' AND "RESPONSE_TIME_HOURS"<=8) OR
                                (priority='low' AND "RESPONSE_TIME_HOURS"<=24)
                           THEN 100.0 ELSE 0.0 END) AS compliance_rate
                FROM m GROUP BY priority
            ''')
            stats = {row[0]: (float(row[1]), float(row[2])) for row in ck_cur.fetchall()}
            ck_cur.close(); ck_conn.close()
            best_priority = max(stats, key=lambda p: round(stats[p][0], 2))
            worst_sla_priority = min(stats, key=lambda p: stats[p][1])
        except Exception as e:
            # Fallback: based on SLA_DATA constants
            stats = {k: (v['csat'], v['rate']) for k, v in SLA_DATA.items()}
            best_priority = max(stats, key=lambda p: stats[p][0])
            worst_sla_priority = min(stats, key=lambda p: stats[p][1])
        # Tied CSAT (e.g., high and medium both 3.26): accept either of the top-tied priorities
        top_csat = round(max(s[0] for s in stats.values()), 2)
        best_candidates = [p for p, (c, _) in stats.items() if round(c, 2) == top_csat]
        check(
            f"Summary Best_Priority matches task.md rule (highest CSAT, one of {best_candidates})",
            any(p in sum_text for p in best_candidates),
            f"Summary: {sum_text[:200]}"
        )
        check(
            f"Summary Worst_SLA_Priority matches task.md rule (lowest compliance rate: {worst_sla_priority})",
            worst_sla_priority in sum_text,
            f"Summary: {sum_text[:200]}"
        )

    # Fetch cells with row/col structure for SLA_Compliance sheet
    cur.execute("""
        SELECT s.title, c.row_index, c.col_index, c.value
        FROM gsheet.cells c
        JOIN gsheet.sheets s ON c.spreadsheet_id = s.spreadsheet_id AND c.sheet_id = s.id
        WHERE c.spreadsheet_id = %s
    """, (ss_id,))
    all_cells = cur.fetchall()
    cells = [str(r[3]) for r in all_cells if r[3] is not None]
    all_vals = " ".join(cells).lower()

    check("Sheet contains 'High' priority data", "high" in all_vals, "Not found")
    check("Sheet contains 'Medium' priority data", "medium" in all_vals, "Not found")
    check("Sheet contains 'Low' priority data", "low" in all_vals, "Not found")

    # Check numeric compliance data appears (tolerate thousand separators like "6,466")
    norm_vals = all_vals.replace(",", "").replace(" ", "")
    check("Sheet contains ticket counts", any(v in norm_vals for v in ["6466", "15774", "9348"]),
          "Ticket counts not found")

    # Structural validation: each priority row should have compliance rate close to SLA_DATA
    sla_rows = {}
    for title, row_idx, col_idx, value in all_cells:
        if title and "sla" in title.lower() and value is not None:
            sla_rows.setdefault(row_idx, {})[col_idx] = value
    matched_pairs = 0
    for row_idx, cols in sla_rows.items():
        row_text = " ".join(str(v).lower() for v in cols.values() if v is not None)
        for prio, info in SLA_DATA.items():
            if prio in row_text:
                # Look for the compliance rate in the row (tolerate % signs,
                # thousand separators and literal formulas)
                for v in cols.values():
                    if _num_close(v, info["rate"], tol=1.0):
                        matched_pairs += 1
                        break
                break
    check("SLA_Compliance sheet rows match expected rates (>=2 of 3 priorities)",
          matched_pairs >= 2,
          f"Matched {matched_pairs}/3 priority-rate pairs")

    cur.close()
    conn.close()
    return has_sla and has_summary


def check_gform():
    print("\n=== Checking Google Forms ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, title FROM gform.forms
        WHERE title ILIKE '%customer%support%' OR title ILIKE '%support%satisfaction%'
    """)
    forms = cur.fetchall()
    check("Customer Support Satisfaction Survey form exists", len(forms) >= 1,
          f"Found: {[f[1] for f in forms]}")

    if forms:
        form_id = forms[0][0]
        cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (form_id,))
        q_count = cur.fetchone()[0]
        # >= 4 rather than == 4: in a multi-agent swarm a duplicate question may
        # be created by a second sub-agent; a correct 4-question form must pass.
        check("Form has at least 4 questions", q_count >= 4, f"Got {q_count}")
        # Check that at least one question has Yes/No/Partially choice options
        cur.execute("""
            SELECT q.id, q.question_type, q.config
            FROM gform.questions q
            WHERE q.form_id = %s
        """, (form_id,))
        questions = cur.fetchall()
        found_yes_no_part = False
        for q_id, q_type, config in questions:
            choices_str = str(config).lower() if config is not None else ""
            if "yes" in choices_str and "no" in choices_str and ("partial" in choices_str or "partly" in choices_str):
                found_yes_no_part = True
                break
        check("Form has question with Yes/No/Partially options", found_yes_no_part,
              "No question found with Yes/No/Partially choices")

    cur.close()
    conn.close()


def check_email():
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE subject ILIKE '%support%' AND subject ILIKE '%performance%'
        ORDER BY date DESC
    """)
    emails = cur.fetchall()
    check("Support performance email exists", len(emails) >= 1, f"Found {len(emails)}")

    if emails:
        e = emails[0]
        to_str = str(e[2])
        check("Email to support-management@company.example.com",
              "support-management@company.example.com" in to_str.lower(), f"to: {to_str}")
        expected_from = _expected_from_address()
        if expected_from:
            check(f"Email from {expected_from}",
                  expected_from in (e[1] or "").lower(), f"from: {e[1]}")
        else:
            check("Email from configured sender address", True,
                  "email_config.json not found; sender not verified")
        body = (e[3] or "").lower()
        check("Email body mentions SLA or compliance",
              any(kw in body for kw in ["sla", "compliance", "csat", "satisfaction", "high", "medium", "low"]),
              "Body missing key terms")
        # At least one specific SLA rate should be referenced. Extract all
        # numbers from the body and match against the expected rates with a
        # small tolerance, so "12%", "12.0%", "12.03%" and "12.03 percent" all count.
        rate_vals = [info["rate"] for info in SLA_DATA.values()]
        body_nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", body)]
        rate_hit = any(abs(n - r) <= 0.5 for n in body_nums for r in rate_vals)
        check("Email body mentions at least one SLA rate value", rate_hit,
              f"Expected one of rates: {rate_vals}")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    gsheet_ok = check_gsheet()
    check_gform()
    check_email()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
