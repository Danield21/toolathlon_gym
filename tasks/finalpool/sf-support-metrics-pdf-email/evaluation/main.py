"""Evaluation for sf-support-metrics-pdf-email."""
import argparse
import os
import re
import sys

import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "toolathlon_gym", "user": "eigent", "password": "camel"}


def extract_pdf_text(path):
    """Extract text from PDF using available libraries."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(path)
        text = ""
        for page in doc:
            text += page.get_text() or ""
        doc.close()
        return text
    except ImportError:
        pass
    try:
        from pypdf import PdfReader
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    except ImportError:
        pass
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() or ""
        return text
    except ImportError:
        pass
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except ImportError:
        pass
    with open(path, "rb") as f:
        return f.read().decode("latin-1", errors="ignore")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    all_errors = []

    # --- Check PDF ---
    agent_pdf = os.path.join(args.agent_workspace, "Support_Metrics.pdf")
    if not os.path.exists(agent_pdf):
        print(f"FAIL: Agent output not found: {agent_pdf}")
        sys.exit(1)

    print("  Checking Support_Metrics.pdf...")
    text = extract_pdf_text(agent_pdf).lower()

    # Check title
    if "support metrics report" not in text:
        all_errors.append("PDF missing title 'Support Metrics Report'")

    # Check sections
    if "overview" not in text:
        all_errors.append("PDF missing 'Overview' section")
    if "issue type" not in text:
        all_errors.append("PDF missing 'Issue Type' section")
    if "priority" not in text:
        all_errors.append("PDF missing 'Priority' section")

    # Validate data against DB
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT "ISSUE_TYPE", COUNT(*) as cnt,
          ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2) as avg_resp,
          ROUND(AVG("CUSTOMER_SATISFACTION"), 2) as avg_csat
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        GROUP BY "ISSUE_TYPE"
        ORDER BY cnt DESC
    """)
    issue_types = cur.fetchall()

    cur.execute("""
        SELECT "PRIORITY", COUNT(*) as cnt,
          ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2) as avg_resp,
          ROUND(AVG("CUSTOMER_SATISFACTION"), 2) as avg_csat
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        GROUP BY "PRIORITY"
        ORDER BY "PRIORITY"
    """)
    priorities = cur.fetchall()
    conn.close()

    total_tickets = sum(r[1] for r in issue_types)

    # Check issue types present (derive from DB to avoid stale hardcoded list)
    db_issue_types = [str(r[0]).strip().lower() for r in issue_types if r[0]]
    for issue in db_issue_types:
        if issue not in text:
            all_errors.append(f"PDF missing issue type: {issue}")

    # Check priorities present
    for p in ["high", "medium", "low"]:
        if p not in text:
            all_errors.append(f"PDF missing priority: {p}")

    # Helper: word-boundary regex match for an integer to avoid 12 matching 120
    def has_int(t, n):
        return re.search(r"(?<![\d.,])" + re.escape(str(n)) + r"(?![\d.,])", t) is not None

    # Helper: numeric tolerance match for a 2-decimal float in PDF text. Accept
    # any nearby formatting (e.g. 14.63, 14.6, 14.625) within tolerance.
    def has_float(t, val, tol=0.05):
        target = float(val)
        # find every numeric token in text and check at least one is within tol
        for m in re.finditer(r"\d+\.\d+", t):
            try:
                if abs(float(m.group(0)) - target) <= tol:
                    return True
            except Exception:
                continue
        return False

    # Check total ticket count appears
    if not has_int(text, total_tickets):
        all_errors.append(f"PDF missing total ticket count: {total_tickets}")

    # Check issue type ticket counts AND avg response AND avg satisfaction
    # appear in PDF text. Use word-boundary regex to avoid substring overlap
    # (e.g. '12' matching inside '120').
    for r in issue_types:
        issue_name, cnt, avg_resp, avg_csat = r
        if not has_int(text, cnt):
            all_errors.append(f"PDF missing ticket count {cnt} for {issue_name}")
        if avg_resp is not None and not has_float(text, avg_resp):
            all_errors.append(
                f"PDF missing avg response time {avg_resp} for {issue_name}"
            )
        if avg_csat is not None and not has_float(text, avg_csat):
            all_errors.append(
                f"PDF missing avg satisfaction {avg_csat} for {issue_name}"
            )

    # Check priority ticket counts AND metrics
    for r in priorities:
        priority_name, cnt, avg_resp, avg_csat = r
        if not has_int(text, cnt):
            all_errors.append(f"PDF missing ticket count {cnt} for priority {priority_name}")
        if avg_resp is not None and not has_float(text, avg_resp):
            all_errors.append(
                f"PDF missing avg response time {avg_resp} for priority {priority_name}"
            )
        if avg_csat is not None and not has_float(text, avg_csat):
            all_errors.append(
                f"PDF missing avg satisfaction {avg_csat} for priority {priority_name}"
            )

    if not all_errors:
        print("    PASS")

    # --- BLOCKING: Check email in DB ---
    print("  Checking email (blocking - required by task)...")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        # Required: to=manager@supportcenter.com, subject "Monthly Support Metrics Report"
        cur.execute("""
            SELECT id, subject, to_addr, body_text, from_addr
            FROM email.messages
            WHERE LOWER(to_addr::text) LIKE '%manager@supportcenter.com%'
              AND LOWER(subject) LIKE '%monthly%support%metrics%report%'
        """)
        emails = cur.fetchall()
        if not emails:
            all_errors.append("Email not found: required to=manager@supportcenter.com with subject 'Monthly Support Metrics Report'")
        else:
            # Validate body mentions total tickets and average satisfaction
            email_body = (emails[0][3] or "").lower()
            # word-boundary regex to avoid substring overlap
            tt_str = str(total_tickets)
            tt_with_sep = f"{total_tickets:,}"
            email_has_total = (
                re.search(r"(?<![\d.,])" + re.escape(tt_str) + r"(?![\d.,])", email_body) is not None
                or re.search(r"(?<![\d.,])" + re.escape(tt_with_sep) + r"(?![\d.,])", email_body) is not None
            )
            if not email_has_total:
                all_errors.append(f"Email body missing total ticket count: {total_tickets}")
            # Compute overall avg satisfaction
            cur.execute("""
                SELECT ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2)
                FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
            """)
            overall_csat = cur.fetchone()[0]
            # accept the rounded value within a small tolerance against any
            # numeric token in the body
            target = float(overall_csat)
            email_has_csat = False
            for m in re.finditer(r"\d+\.\d+", email_body):
                try:
                    if abs(float(m.group(0)) - target) <= 0.05:
                        email_has_csat = True
                        break
                except Exception:
                    continue
            if not email_has_csat:
                all_errors.append(f"Email body missing avg satisfaction (~{overall_csat})")
        conn.close()
    except Exception as e:
        all_errors.append(f"Email DB check error: {e}")

    if all_errors:
        print(f"\n=== RESULT: FAIL ({len(all_errors)} errors) ===")
        for e in all_errors[:15]:
            print(f"  {e}")
        sys.exit(1)
    else:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)


if __name__ == "__main__":
    main()
