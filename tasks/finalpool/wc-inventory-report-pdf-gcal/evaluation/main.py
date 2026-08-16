"""
Evaluation for wc-inventory-report-pdf-gcal task.

Checks:
1. PDF file Inventory_Audit.pdf exists and contains expected data
2. Google Calendar event created (BLOCKING)
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime

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

# Known out-of-stock product IDs
OOS_IDS = [20, 39, 45, 71, 79]
# Known low-stock product IDs (qty 1-5)
LOW_IDS = [5, 7, 15, 19, 21, 22, 27, 28, 30, 31, 32, 50, 52, 53, 54, 61, 63, 70, 77, 82]

# Expected counts
EXPECTED_TOTAL = 82
EXPECTED_OOS = 5
EXPECTED_LOW = 20
EXPECTED_FLAGGED = 25


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def _word_boundary(pid, text):
    """Check if pid appears with proper boundaries.

    PDF text extraction may glue an ID to the next word with no separator
    (e.g. "71" followed by "30kg" becomes "7130kg" in extracted text).
    To avoid that false negative we accept a match when:
    - left side is non-digit (or start of string),
    - right side is non-digit (or end of string).
    The right side may be a letter (e.g. "71...30kg" => "7130kg" still has digits, fail);
    so additionally we look for the SKU which is unique per product.
    """
    pat = r'(?<!\d)' + str(pid) + r'(?!\d)'
    return re.search(pat, text) is not None


# Map product id -> SKU substring for fallback matching when PDF text extraction
# glues numbers (e.g. "71" + "30kg" -> "7130kg")
PRODUCT_SKU_HINTS = {
    20: "INFINITY-JBL-GLIDE",
    39: "BOXTUDIO-LIGHTBOX",
    45: "TYPEC-EARPHONE-FOR",
    71: "30KG-DIGITAL-SCALE",
    79: "SILENCER-PANELS",
    5: "THEGIFTKART-FLIP",
    7: "BOAT-DUAL-PORT",
    15: "POPIO-TEMPERED-GLASS",
    19: "THEGIFTKART-SHOCKPRO",
    21: "ROBUSTRION-SMART-TRI",
    22: "CASEOLOGY-BY-SPIGEN",
    27: "MI-EXTENDED-WARRANTY",
    28: "DELL-S2721HNMGREY",
    30: "ARTIS-65WATT",
    31: "ELOIES-TRIPOD",
    32: "TELESIN-BATTERY",
    50: "XTOUCH-XTBC01",
    52: "SAMSUNG-1M-09CM",
    53: "INVANTER-2022-MODEL",
    54: "LG-139-CM",
    61: "DLX-HMT-WOMEN",
    63: "PFN-ANALOG-DW",
    70: "RENEWED-HAVELLS",
    77: "SHOPNET-WIRELESS",
    82: "TRONICA-SUPER-KING",
}


def _id_appears(pid, text):
    """Returns True if the product id appears, OR its SKU prefix appears.

    PDF extraction sometimes glues numbers across columns; accepting either
    a word-boundary numeric match or a SKU-substring match avoids false negatives
    while still rejecting empty/wrong PDFs (since SKU substrings are unique).
    """
    if _word_boundary(pid, text):
        return True
    sku_hint = PRODUCT_SKU_HINTS.get(pid)
    if sku_hint:
        return sku_hint.lower() in text.lower()
    return False


def check_pdf(agent_workspace):
    """Check the PDF file exists and has key content."""
    print("\n=== Checking Inventory_Audit.pdf ===")

    pdf_path = os.path.join(agent_workspace, "Inventory_Audit.pdf")
    if not os.path.isfile(pdf_path):
        record("PDF file exists", False, f"Not found: {pdf_path}")
        # Add structural failures since file is missing
        for label in ["PDF file size reasonable", "PDF contains title", "PDF contains date",
                      "PDF total products count correct", "PDF out of stock count correct",
                      "PDF low stock count correct", "PDF total flagged count correct",
                      "PDF lists all 5 OOS product IDs", "PDF lists all 20 low-stock product IDs",
                      "PDF mentions out of stock", "PDF mentions low stock"]:
            record(label, False, "PDF missing")
        return False
    record("PDF file exists", True)

    # Check file size is reasonable (at least 1KB)
    size = os.path.getsize(pdf_path)
    record("PDF file size reasonable", size > 1024, f"Size: {size} bytes")

    # Try to extract text
    text = ""
    extraction_ok = False
    try:
        import PyPDF2
        try:
            with open(pdf_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    text += page.extract_text() or ""
            extraction_ok = True
        except Exception as e:
            print(f"  [WARN] PyPDF2 read error: {e}")
    except ImportError:
        pass
    if not extraction_ok:
        try:
            import pdfplumber
            try:
                with pdfplumber.open(pdf_path) as p:
                    for page in p.pages:
                        text += page.extract_text() or ""
                extraction_ok = True
            except Exception as e:
                print(f"  [WARN] pdfplumber read error: {e}")
        except ImportError:
            pass
    if not extraction_ok:
        record("PDF parseable as PDF", False, "Could not parse PDF content")
        # Add subordinate failures
        for label in ["PDF contains title", "PDF contains date",
                      "PDF total products count correct", "PDF out of stock count correct",
                      "PDF low stock count correct", "PDF total flagged count correct",
                      "PDF mentions out of stock", "PDF mentions low stock",
                      "PDF lists all 5 OOS product IDs",
                      "PDF lists all 20 low-stock product IDs"]:
            record(label, False, "PDF unparseable")
        return False
    record("PDF parseable as PDF", True)

    text_lower = text.lower()

    # Check title
    record("PDF contains title",
           "inventory" in text_lower and "audit" in text_lower,
           "Expected 'Inventory Audit' in PDF")

    # Check date
    record("PDF contains date",
           ("march 6, 2026" in text_lower or "march 6 2026" in text_lower or
            "2026-03-06" in text_lower or "03/06/2026" in text_lower or
            "3/6/2026" in text_lower or "06 march 2026" in text_lower or
            "06 mar 2026" in text_lower),
           "Expected March 6, 2026 date")

    # Check summary section -- exact numeric counts
    # Total products: 82
    record("PDF total products count correct",
           re.search(r'total\s*products?\s*[:\-]?\s*' + str(EXPECTED_TOTAL), text_lower) is not None
           or re.search(r'total\s*[:\-]?\s*' + str(EXPECTED_TOTAL) + r'\b', text_lower) is not None
           or re.search(r'\b' + str(EXPECTED_TOTAL) + r'\s*products?\b', text_lower) is not None,
           f"Expected total products = {EXPECTED_TOTAL}")

    # Out of stock count: 5
    record("PDF out of stock count correct",
           re.search(r'out[\s\-]of[\s\-]stock\s*[:\-]?\s*' + str(EXPECTED_OOS) + r'\b', text_lower) is not None
           or re.search(r'\b' + str(EXPECTED_OOS) + r'\s+out[\s\-]of[\s\-]stock', text_lower) is not None,
           f"Expected OOS count = {EXPECTED_OOS}")

    # Low stock count: 20
    record("PDF low stock count correct",
           re.search(r'low\s*stock[^\n]*' + str(EXPECTED_LOW) + r'\b', text_lower) is not None
           or re.search(r'\b' + str(EXPECTED_LOW) + r'\s+low\s*stock', text_lower) is not None,
           f"Expected low stock count = {EXPECTED_LOW}")

    # Total flagged: 25
    record("PDF total flagged count correct",
           re.search(r'total\s*flagged\s*[:\-]?\s*' + str(EXPECTED_FLAGGED) + r'\b', text_lower) is not None
           or re.search(r'flagged\s*[:\-]?\s*' + str(EXPECTED_FLAGGED) + r'\b', text_lower) is not None,
           f"Expected total flagged = {EXPECTED_FLAGGED}")

    record("PDF mentions out of stock", "out of stock" in text_lower,
           "Expected 'out of stock' text")

    record("PDF mentions low stock", "low stock" in text_lower,
           "Expected 'low stock' text")

    # Check ALL out-of-stock product IDs appear (id + word-boundary OR SKU substring)
    found_oos = sum(1 for pid in OOS_IDS if _id_appears(pid, text))
    record(f"PDF lists all 5 OOS product IDs",
           found_oos == len(OOS_IDS),
           f"Found {found_oos}/{len(OOS_IDS)} OOS product IDs")

    # Check ALL low-stock product IDs appear (id + word-boundary OR SKU substring)
    found_low = sum(1 for pid in LOW_IDS if _id_appears(pid, text))
    record(f"PDF lists all 20 low-stock product IDs",
           found_low == len(LOW_IDS),
           f"Found {found_low}/{len(LOW_IDS)} low-stock product IDs")

    return True


def check_gcal():
    """Check Google Calendar event - BLOCKING."""
    print("\n=== Checking Google Calendar (blocking) ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("""
            SELECT summary, description, start_datetime, end_datetime
            FROM gcal.events
        """)
        events = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("GCal DB query OK", False, f"DB error: {e}")
        return

    # Find restock-related events — title must equal or start with 'Restock Planning Meeting'
    matches = []
    for ev in events:
        summary = (ev[0] or "")
        description = (ev[1] or "")
        start = ev[2]
        end = ev[3]
        s_norm = summary.strip().lower()
        if s_norm == "restock planning meeting" or s_norm.startswith("restock planning meeting"):
            matches.append((summary, description, start, end))

    record("Restock Planning Meeting event exists (title equal/starts with)",
           len(matches) >= 1,
           f"Found {len(matches)} events with title 'Restock Planning Meeting'")

    if not matches:
        # Add subordinate failures
        record("Event datetime correct (2026-03-12 10:00-11:00)", False, "no event")
        record("Event description mentions OOS count near 'out of stock' phrase", False, "no event")
        record("Event description mentions low count near 'low stock' phrase", False, "no event")
        return

    # Pick the first matching event
    summary, description, start, end = matches[0]

    # Verify datetime: March 12, 2026 from 10:00 to 11:00
    def _to_str(dt):
        try:
            if isinstance(dt, str):
                return dt
            return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return str(dt)

    start_s = _to_str(start)
    end_s = _to_str(end)
    dt_ok = (
        ("2026-03-12" in start_s and "10:00" in start_s) and
        ("2026-03-12" in end_s and "11:00" in end_s)
    )
    record("Event datetime correct (2026-03-12 10:00-11:00)",
           dt_ok,
           f"start={start_s}, end={end_s}")

    # Description should mention counts: 5 OOS, 20 low — use word-boundary regex to avoid '5' substring leakage
    desc_l = description.lower()
    oos_pattern = re.compile(
        r"\b" + str(EXPECTED_OOS) + r"\b[^\n]{0,40}\bout\s+of\s+stock\b",
        re.IGNORECASE | re.DOTALL,
    )
    low_pattern = re.compile(
        r"\b" + str(EXPECTED_LOW) + r"\b[^\n]{0,40}\blow\s+stock\b",
        re.IGNORECASE | re.DOTALL,
    )
    desc_has_oos = bool(oos_pattern.search(description)) or (
        re.search(r"\b" + str(EXPECTED_OOS) + r"\b", description) is not None
        and "out of stock" in desc_l
    )
    desc_has_low = bool(low_pattern.search(description)) or (
        re.search(r"\b" + str(EXPECTED_LOW) + r"\b", description) is not None
        and "low stock" in desc_l
    )
    record("Event description mentions OOS count near 'out of stock' phrase",
           desc_has_oos,
           f"description: {description[:300]}")
    record("Event description mentions low count near 'low stock' phrase",
           desc_has_low,
           f"description: {description[:300]}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_pdf(args.agent_workspace)
    check_gcal()

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
