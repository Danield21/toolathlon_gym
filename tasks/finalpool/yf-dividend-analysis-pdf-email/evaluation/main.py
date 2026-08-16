"""
Evaluation for yf-dividend-analysis-pdf-email task.

Checks:
1. PDF file Dividend_Analysis.pdf exists with correct data
2. Email sent (BLOCKING — required by task)
"""

import argparse
import json
import os
import sys

import psycopg2

def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0

EXPECTED_SYMBOLS = ["JNJ", "JPM", "XOM", "GOOGL"]

# Expected dividend yields (from DB: dividendYield field * 100)
# XOM: 2.75 (275%), JPM: 2.0 (200%), JNJ: 2.12 (212%), GOOGL: 0.28 (28%)
# These are the raw values from the DB
EXPECTED_YIELDS = {
    "XOM": 2.75,
    "JNJ": 2.12,
    "JPM": 2.0,
    "GOOGL": 0.28,
}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_pdf(agent_workspace):
    """Check the PDF file exists and has key content."""
    print("\n=== Checking Dividend_Analysis.pdf ===")

    pdf_path = os.path.join(agent_workspace, "Dividend_Analysis.pdf")
    if not os.path.isfile(pdf_path):
        record("PDF file exists", False, f"Not found: {pdf_path}")
        return False
    record("PDF file exists", True)

    size = os.path.getsize(pdf_path)
    record("PDF file size reasonable", size > 1024, f"Size: {size} bytes")

    # Try to extract text
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""
    except ImportError:
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as p:
                text = ""
                for page in p.pages:
                    text += page.extract_text() or ""
        except ImportError:
            print("  [WARN] No PDF reader available. Checking existence only.")
            return True

    text_lower = text.lower()

    # Check title
    record("PDF contains Dividend Analysis title",
           "dividend" in text_lower and "analysis" in text_lower,
           "Expected 'Dividend Analysis' in title")

    # Check all symbols appear
    symbols_found = sum(1 for s in EXPECTED_SYMBOLS if s in text)
    record("PDF lists all 4 stock symbols", symbols_found == 4,
           f"Found {symbols_found}/4 symbols")

    # Check dividend yield data appears (look for percentage values)
    record("PDF mentions dividend yield",
           "yield" in text_lower or "dividend" in text_lower,
           "Expected dividend yield data")

    # Check sectors appear
    sectors_found = sum(1 for s in ["Healthcare", "Financial", "Energy", "Technology"]
                       if s.lower() in text_lower)
    record("PDF mentions at least 3 sectors", sectors_found >= 3,
           f"Found {sectors_found}/4 sector names")

    # Check recommendation section
    record("PDF has recommendation section",
           "recommendation" in text_lower,
           "Expected recommendation section")

    # Check ranking - XOM should be identified as highest yield
    record("PDF identifies highest yield stock",
           "xom" in text_lower and ("highest" in text_lower or "1." in text),
           "Expected XOM as highest yield")

    # Check at least 2 expected yield numerical values appear.
    # YF dividendYield values in this DB are decimals (0.0275, 0.0212, 0.020, 0.0028).
    # Per task.md "multiplied by 100" — so PDF should show 2.75, 2.12, 2.00, 0.28.
    # However the GT PDF in this dataset uses raw YF value*100 again (275.00, 212.00).
    # Accept either convention.
    yield_strings_v1 = ["2.75", "2.12", "2.00", "0.28"]
    yield_strings_v2 = ["275.00", "212.00", "200.00", "28.00"]
    hits_v1 = sum(1 for y in yield_strings_v1 if y in text)
    hits_v2 = sum(1 for y in yield_strings_v2 if y in text)
    record("PDF cites >= 2 expected yield values",
           hits_v1 >= 2 or hits_v2 >= 2,
           f"v1 hits {hits_v1}/{yield_strings_v1}; v2 hits {hits_v2}/{yield_strings_v2}")

    return True


def check_email():
    """Check email sent — BLOCKING (task explicitly requires the email)."""
    print("\n=== Checking Email ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("""
            SELECT subject, from_addr, to_addr, body_text
            FROM email.messages
        """)
        emails = cur.fetchall()
        cur.close()
        conn.close()

        record("Email exists in DB", len(emails) >= 1, f"Found {len(emails)} emails")

        matched = None
        for subject, from_addr, to_addr, body_text in emails:
            subject_lower = (subject or "").lower()
            if "dividend analysis summary" in subject_lower:
                matched = (subject, from_addr, to_addr, body_text)
                break

        record("Email subject 'Dividend Analysis Summary - March 2026' present",
               matched is not None,
               "subject must contain 'Dividend Analysis Summary'")

        if matched:
            subject, from_addr, to_addr, body_text = matched
            from_str = str(from_addr or "").lower()
            record("Email from analyst@investteam.com",
                   "analyst@investteam.com" in from_str,
                   f"From: {from_addr}")
            to_str = str(to_addr or "").lower()
            record("Email to team@investteam.com",
                   "team@investteam.com" in to_str,
                   f"To: {to_addr}")

            sl = (subject or "").lower()
            record("Email subject mentions 'March 2026'",
                   "march 2026" in sl,
                   f"Subject: {subject}")

            body = body_text or ""
            symbols_in_body = sum(1 for s in EXPECTED_SYMBOLS if s in body)
            record("Email body lists all 4 symbols",
                   symbols_in_body == 4,
                   f"Found {symbols_in_body}/4 symbols in body")

            body_l = body.lower()
            record("Email body mentions 'yield'",
                   "yield" in body_l,
                   f"body sample: {body_l[:200]}")
            record("Email body mentions 'average'",
                   "average" in body_l or "avg" in body_l,
                   f"body sample: {body_l[:200]}")

    except Exception as e:
        record("Email check ran", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("YF DIVIDEND ANALYSIS PDF EMAIL - EVALUATION")
    print("=" * 70)

    check_pdf(args.agent_workspace)
    check_email()  # Non-blocking

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    if args.res_log_file:
        result = {"passed": PASS_COUNT, "failed": FAIL_COUNT, "success": overall}
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
