"""Evaluation for wc-customer-loyalty-word-gsheet.

Blocking checks: Customer_Loyalty.docx (Word document content) AND Google Sheet.
"""
import argparse
import os
import re
import sys

try:
    from docx import Document
except ImportError:
    Document = None

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
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


def norm(s):
    return str(s or "").strip().lower().replace("_", " ")


def safe_float(v):
    try:
        if v is None: return None
        return float(str(v).replace(",", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return None


def _fetch_expected_tiers():
    """Fetch real expected tier counts and totals from DB."""
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("""
            SELECT
              COUNT(*) FILTER (WHERE orders_count >= 1) AS active_count,
              COALESCE(SUM(total_spent::float) FILTER (WHERE orders_count >= 1), 0) AS total_rev,
              COUNT(*) FILTER (WHERE orders_count >= 1 AND total_spent >= 1000) AS gold,
              COUNT(*) FILTER (WHERE orders_count >= 1 AND total_spent >= 300 AND total_spent < 1000) AS silver,
              COUNT(*) FILTER (WHERE orders_count >= 1 AND total_spent < 300) AS bronze,
              COALESCE(SUM(total_spent::float) FILTER (WHERE orders_count >= 1 AND total_spent >= 1000), 0) AS gold_total,
              COALESCE(SUM(total_spent::float) FILTER (WHERE orders_count >= 1 AND total_spent >= 300 AND total_spent < 1000), 0) AS silver_total,
              COALESCE(SUM(total_spent::float) FILTER (WHERE orders_count >= 1 AND total_spent < 300), 0) AS bronze_total
            FROM wc.customers
        """)
        r = cur.fetchone()
        cur.execute("""
            SELECT first_name, last_name, email, total_spent::float, orders_count
            FROM wc.customers WHERE orders_count >= 1 ORDER BY total_spent DESC LIMIT 5
        """)
        top5 = [{"name": f"{a} {b}", "email": c, "total_spent": d, "orders": e} for a, b, c, d, e in cur.fetchall()]
        conn.close()
        return {
            "active_count": int(r[0]),
            "total_revenue": round(float(r[1]), 2),
            "gold": int(r[2]),
            "silver": int(r[3]),
            "bronze": int(r[4]),
            "gold_total": round(float(r[5]), 2),
            "silver_total": round(float(r[6]), 2),
            "bronze_total": round(float(r[7]), 2),
            "top5": top5,
        }
    except Exception as e:
        print(f"[WARN] expected fetch: {e}")
        return None


def check_word(agent_workspace, expected):
    """Check Word document structure and content against groundtruth."""
    print("\n=== Checking Word Document ===")

    agent_path = os.path.join(agent_workspace, "Customer_Loyalty.docx")
    if not os.path.isfile(agent_path):
        record("Word file exists", False, f"Not found: {agent_path}")
        return
    record("Word file exists", True)

    if Document is None:
        record("python-docx installed", False, "Cannot import docx")
        return

    doc = Document(agent_path)
    full_text = "\n".join(p.text for p in doc.paragraphs)
    full_text_lower = full_text.lower()

    # Check title
    record("Has 'Customer Loyalty Analysis Report' title",
           "customer loyalty analysis report" in full_text_lower,
           "Title not found")
    # Check date
    record("Mentions '2026-03-06' report date",
           "2026-03-06" in full_text,
           "Date not found")
    # Section headings
    for sec in ["tier summary", "gold tier", "silver tier", "bronze tier"]:
        record(f"Has '{sec}' section", sec in full_text_lower, "section not found")

    # Use regex for exact total active customers count (label-bound)
    if expected:
        ac = expected["active_count"]
        # Look for "Total Active Customers: <ac>" pattern
        m = re.search(r"total\s+active\s+customers[^0-9]*(\d+)", full_text_lower)
        record(f"Total Active Customers count == {ac}",
               m is not None and int(m.group(1)) == ac,
               f"matched: {m.group(0) if m else 'no match'}")

        def find_tier_count(tier_name, text_lower):
            """Find a tier's customer count tolerantly. Strategy:
            1. Locate every occurrence of '<tier> tier' in the text.
            2. For each occurrence, scan a 200-char window after it for any
               integer followed within ~40 chars by 'customer'.
            3. Also accept patterns where 'customer' precedes the number
               (e.g., 'gold tier customers: 25').
            Returns the first integer found, or None.
            """
            for m_tier in re.finditer(rf"{tier_name}\s+tier", text_lower):
                start = m_tier.start()
                window = text_lower[start:start + 250]
                # Pattern A: <tier> tier ... <num> ... customer
                for m in re.finditer(r"(\d[\d,]*)\s*[^\d]{0,40}?customers?\b", window):
                    try:
                        return int(m.group(1).replace(",", ""))
                    except ValueError:
                        continue
                # Pattern B: <tier> tier ... customer ... <num>
                for m in re.finditer(r"customers?\b[^\d]{0,40}?(\d[\d,]*)", window):
                    try:
                        return int(m.group(1).replace(",", ""))
                    except ValueError:
                        continue
                # Pattern C: <tier> tier ... <num> (no 'customer' in window — last resort)
                m_num = re.search(r"\b(\d[\d,]*)\b", window[len(f"{tier_name} tier"):])
                if m_num:
                    try:
                        return int(m_num.group(1).replace(",", ""))
                    except ValueError:
                        pass
            return None

        gold_n = find_tier_count("gold", full_text_lower)
        record(f"Gold tier count == {expected['gold']}",
               gold_n == expected["gold"],
               f"got: {gold_n}")
        silver_n = find_tier_count("silver", full_text_lower)
        record(f"Silver tier count == {expected['silver']}",
               silver_n == expected["silver"],
               f"got: {silver_n}")
        bronze_n = find_tier_count("bronze", full_text_lower)
        record(f"Bronze tier count == {expected['bronze']}",
               bronze_n == expected["bronze"],
               f"got: {bronze_n}")

        # Top 5 customers all present (by name & email)
        top_hits = 0
        for c in expected["top5"]:
            if c["email"].lower() in full_text_lower or c["name"].lower() in full_text_lower:
                top_hits += 1
        record(f"All 5 top spenders mentioned",
               top_hits >= 5, f"hits={top_hits}")

        # Top spender's full identity (name + email)
        top = expected["top5"][0]
        record(f"Top spender '{top['name']}' present",
               top["name"].lower() in full_text_lower,
               f"text[:300]: {full_text_lower[:300]}")
        record(f"Top spender email present",
               top["email"].lower() in full_text_lower)

        # Check total revenue mentioned (rough match, allowing $ and commas)
        rev = expected["total_revenue"]
        # Search for the integer part of the value
        rev_int = int(rev)
        record(f"Total revenue ${rev:.2f} present",
               f"{rev:.2f}" in full_text or f"{rev_int}" in full_text,
               f"expected {rev:.2f}")


def check_gsheet(expected):
    """Check Google Sheet 'Customer Loyalty Dashboard' / 'Tier Data' (BLOCKING)."""
    print("\n=== Checking Google Sheet (blocking) ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
    except Exception as e:
        record("DB connect", False, str(e))
        return

    try:
        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()
        target = None
        for sid, title in spreadsheets:
            if norm(title) == norm("Customer Loyalty Dashboard"):
                target = sid
                break
        record("'Customer Loyalty Dashboard' spreadsheet exists",
               target is not None,
               f"sheets: {[t for _,t in spreadsheets]}")

        if target:
            cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id=%s", (target,))
            sheets = cur.fetchall()
            sheet_norm = {norm(s[1]): s for s in sheets}
            td = sheet_norm.get(norm("Tier Data"))
            record("'Tier Data' sheet exists", td is not None,
                   f"sheets: {[s[1] for s in sheets]}")
            if td:
                # Header check
                cur.execute("""
                    SELECT value FROM gsheet.cells
                    WHERE spreadsheet_id=%s AND sheet_id=%s AND row_index=0
                    ORDER BY col_index
                """, (target, td[0]))
                headers = [norm(r[0]) for r in cur.fetchall()]
                for h in ["customer id", "name", "email", "orders", "total spent", "avg order value", "tier"]:
                    record(f"Tier Data has '{h}' header", h in headers,
                           f"got: {headers}")
                # Row count == active_count
                cur.execute("""
                    SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
                    WHERE spreadsheet_id=%s AND sheet_id=%s AND row_index>0
                """, (target, td[0]))
                rc = cur.fetchone()[0]
                if expected:
                    record(f"Tier Data has {expected['active_count']} rows",
                           rc == expected["active_count"], f"got {rc}")
                # Top spender appears
                cur.execute("""
                    SELECT LOWER(value) FROM gsheet.cells
                    WHERE spreadsheet_id=%s AND sheet_id=%s
                """, (target, td[0]))
                joined = " ".join(r[0] for r in cur.fetchall() if r[0])
                if expected:
                    top = expected["top5"][0]
                    # Tighten: require BOTH email AND name to ensure Email column populated
                    name_present = top["name"].lower() in joined
                    email_present = top["email"].lower() in joined
                    record(f"Tier Data mentions top spender (name + email)",
                           name_present and email_present,
                           f"name_present={name_present} email_present={email_present} sample: {joined[:200]}")
        cur.close()
        conn.close()
    except Exception as e:
        record("GSheet check", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = _fetch_expected_tiers()
    check_word(args.agent_workspace, expected)
    check_gsheet(expected)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
