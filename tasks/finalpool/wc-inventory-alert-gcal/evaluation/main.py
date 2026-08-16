"""
Evaluation script for wc-inventory-alert-gcal task.

Checks:
1. Text file (out_of_stock_report.txt) matches groundtruth
2. Google Calendar events created for out-of-stock products
3. Email sent with out-of-stock summary
"""

import argparse
import json
import os
import sys

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

OOS_PRODUCT_IDS = [20, 39, 45, 71, 79]


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


# ============================================================================
# Check 1: out_of_stock_report.txt
# ============================================================================

def check_text_file(agent_workspace, groundtruth_workspace):
    print("\n=== Checking out_of_stock_report.txt ===")

    agent_file = os.path.join(agent_workspace, "out_of_stock_report.txt")
    gt_file = os.path.join(groundtruth_workspace, "out_of_stock_report.txt")

    if not os.path.isfile(agent_file):
        record("Text file exists", False, f"Not found: {agent_file}")
        return False
    record("Text file exists", True)

    with open(agent_file) as f:
        agent_lines = [l.strip() for l in f.readlines() if l.strip()]
    with open(gt_file) as f:
        gt_lines = [l.strip() for l in f.readlines() if l.strip()]

    record("Line count matches (5)", len(agent_lines) == len(gt_lines),
           f"Expected {len(gt_lines)}, got {len(agent_lines)}")

    # Check each product ID appears as the FIRST field (split by ' - ')
    matched = 0
    sku_matched = 0
    cat_matched = 0
    for i, pid in enumerate(OOS_PRODUCT_IDS):
        # Find agent line where first ' - '-split field equals str(pid)
        line_for_pid = None
        for l in agent_lines:
            parts = l.split(" - ")
            if parts and parts[0].strip() == str(pid):
                line_for_pid = l
                break
        record(f"Report: product {pid} listed (first field exact)",
               line_for_pid is not None,
               f"ID {pid} not found as first field. Agent first fields: {[l.split(' - ')[0].strip() for l in agent_lines]}")
        if line_for_pid is None:
            continue
        matched += 1

        # Compare SKU (3rd field) and Category (4th field) against GT
        gt_line = next((g for g in gt_lines if g.split(" - ")[0].strip() == str(pid)), None)
        if gt_line is None:
            continue
        a_parts = [p.strip() for p in line_for_pid.split(" - ")]
        g_parts = [p.strip() for p in gt_line.split(" - ")]
        # Length should match (4 fields: ID, Name, SKU, Category)
        if len(a_parts) >= 3 and len(g_parts) >= 3:
            if a_parts[2] == g_parts[2]:
                sku_matched += 1
            else:
                record(f"Report: product {pid} SKU matches GT", False,
                       f"Got '{a_parts[2]}', expected '{g_parts[2]}'")
        if len(a_parts) >= 4 and len(g_parts) >= 4:
            if a_parts[3].lower() == g_parts[3].lower():
                cat_matched += 1
            else:
                record(f"Report: product {pid} Category matches GT", False,
                       f"Got '{a_parts[3]}', expected '{g_parts[3]}'")

    record(f"Report: SKU matches for all products",
           sku_matched == len(OOS_PRODUCT_IDS),
           f"Matched {sku_matched}/{len(OOS_PRODUCT_IDS)}")
    record(f"Report: Category matches for all products",
           cat_matched == len(OOS_PRODUCT_IDS),
           f"Matched {cat_matched}/{len(OOS_PRODUCT_IDS)}")

    # Verify sort order ascending by ID
    try:
        ids_in_order = [int(l.split(" - ")[0].strip()) for l in agent_lines]
        record("Report: rows sorted by product ID ascending",
               ids_in_order == sorted(ids_in_order),
               f"Got order: {ids_in_order}")
    except (ValueError, IndexError):
        record("Report: rows sorted by product ID ascending", False,
               "Could not parse IDs from first field")

    return matched == len(OOS_PRODUCT_IDS) and sku_matched == len(OOS_PRODUCT_IDS) and cat_matched == len(OOS_PRODUCT_IDS)


# ============================================================================
# Check 2: Google Calendar events
# ============================================================================

def check_gcal():
    print("\n=== Checking Google Calendar ===")

    import re as _re
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, description, start_datetime
        FROM gcal.events
        ORDER BY summary
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_gcal] Found {len(events)} calendar events.")

    # Check events have "Restock:" in title
    restock_events = [e for e in events if "restock" in (e[0] or "").lower()]
    # Tighten: exactly 5 restock events
    record("Exactly 5 'Restock:' events created", len(restock_events) == 5,
           f"Found {len(restock_events)} restock events: {[e[0] for e in restock_events]}")

    # Check date is March 10, 2026
    march_10_events = []
    for summary, description, start_dt in restock_events:
        if start_dt:
            date_str = start_dt.strftime("%Y-%m-%d")
            if date_str == "2026-03-10":
                march_10_events.append((summary, description, start_dt))

    record("All 5 restock events on 2026-03-10",
           len(march_10_events) == 5 and len(restock_events) == 5,
           f"Found {len(march_10_events)} on March 10 (out of {len(restock_events)} restock)")

    # Check descriptions mention each product ID via regex word-boundary
    all_ok = True
    for pid in OOS_PRODUCT_IDS:
        # Match "Product ID: 20" or "Product ID: 20." or "ID: 20\b" patterns
        pattern = _re.compile(r"(?:product\s*id|id)[:\s]*" + str(pid) + r"\b", _re.IGNORECASE)
        matched = [e for e in restock_events if pattern.search(e[1] or "")]
        record(f"gcal: event description references 'Product ID: {pid}'",
               len(matched) >= 1,
               f"Product {pid} not in any event description (regex match)")
        if not matched:
            all_ok = False
            continue
        # Description must include 'SKU:' and 'Category:' and indicate out of stock
        e = matched[0]
        desc = (e[1] or "").lower()
        sku_ok = "sku:" in desc
        cat_ok = "category:" in desc
        oos_ok = "out of stock" in desc
        record(f"  gcal {pid} description has SKU/Category/Out-of-stock fields",
               sku_ok and cat_ok and oos_ok,
               f"sku={sku_ok}, cat={cat_ok}, oos={oos_ok}; sample: {desc[:200]}")
        if not (sku_ok and cat_ok and oos_ok):
            all_ok = False

    return len(restock_events) == 5 and len(march_10_events) == 5 and all_ok


# ============================================================================
# Check 3: Email
# ============================================================================

def check_emails():
    print("\n=== Checking Emails ===")

    import re as _re
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
    record("At least 1 email sent", len(all_emails) >= 1, f"Found {len(all_emails)}")

    found = False
    all_ok = True
    for subject, from_addr, to_addr, body_text in all_emails:
        subject_lower = (subject or "").lower()
        # Subject must include 'out of stock alert' AND 'march 2026'
        if "out of stock alert" in subject_lower and ("march 2026" in subject_lower or "march" in subject_lower):
            found = True

            from_str = str(from_addr or "").lower()
            ok_from = "purchasing@ourstore.com" in from_str
            record("email: from purchasing@ourstore.com",
                   ok_from, f"From: {from_addr}")
            if not ok_from:
                all_ok = False

            to_str = str(to_addr or "").lower()
            ok_to = "warehouse@ourstore.com" in to_str
            record("email: to warehouse@ourstore.com",
                   ok_to, f"To: {to_addr}")
            if not ok_to:
                all_ok = False

            body = body_text or ""
            for pid in OOS_PRODUCT_IDS:
                # Match "ID 20:" or "ID 20." pattern (word-boundary), per task body format
                pattern = _re.compile(r"\bid\s*" + str(pid) + r"[:\s]", _re.IGNORECASE)
                ok_pid = pattern.search(body) is not None
                record(f"email: body mentions 'ID {pid}:' format",
                       ok_pid,
                       f"Pattern not found. Body sample: {body[:300]}")
                if not ok_pid:
                    all_ok = False

            # Total count line
            ok_total = _re.search(r"total\s+out[\s\-]+of[\s\-]+stock\s+products\s*:\s*5\b",
                                  body, _re.IGNORECASE) is not None
            record("email: body has 'Total out-of-stock products: 5'",
                   ok_total,
                   f"Body sample: {body[:300]}")
            if not ok_total:
                all_ok = False
            break

    record("email: 'Out of Stock Alert - March 2026' email found", found,
           f"No email with required subject. Subjects: {[(e[0] or '') for e in all_emails][:5]}")

    return found and all_ok


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

    all_passed = text_ok and gcal_ok and email_ok

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
