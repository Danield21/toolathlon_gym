"""Evaluation for terminal-yf-wc-excel-word-notion.
Checks:
1. Commodity_Impact_Analysis.xlsx with 4 sheets and correct numeric values
2. Pricing_Strategy_Memo.docx with required sections
3. Notion "Market Research Dashboard" database with 2 entries
4. correlation_analysis.py script exists
"""
import argparse
import json
import os
import re
import sys

import openpyxl
import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def num_close(a, b, tol):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _fetch_yf_expected():
    """Fetch expected YF stock stats from DB."""
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        out = {}
        for sym in ("GC=F", "AMZN"):
            cur.execute("""
                SELECT AVG(close)::float, STDDEV(close)::float,
                       (SELECT close::float FROM yf.stock_prices WHERE symbol=%s ORDER BY date ASC LIMIT 1),
                       (SELECT close::float FROM yf.stock_prices WHERE symbol=%s ORDER BY date DESC LIMIT 1)
                FROM yf.stock_prices WHERE symbol=%s
            """, (sym, sym, sym))
            avg, stddev, first, last = cur.fetchone()
            pct = (last - first) / first * 100 if first else 0
            out[sym] = {"avg_price": avg, "volatility": stddev, "pct_change": pct}
        cur.close()
        conn.close()
        return out
    except Exception as e:
        print(f"[fallback] YF DB fetch error: {e}")
        return None


YF_FALLBACK = {
    "GC=F": {"avg_price": 3163.01, "volatility": 796.12, "pct_change": 138.73},
    "AMZN": {"avg_price": 206.45, "volatility": 21.87, "pct_change": 26.18},
}


def check_excel(workspace):
    print("\n=== Check 1: Commodity_Impact_Analysis.xlsx ===")
    path = os.path.join(workspace, "Commodity_Impact_Analysis.xlsx")
    if not os.path.exists(path):
        check("Excel file exists", False, f"Not found at {path}")
        # Subordinate failures
        for label in ["Stock_Trends sheet present", "Product_Margins sheet present",
                      "Correlation_Analysis sheet present", "Strategic_Recommendations sheet present",
                      "Stock_Trends has GC=F row with correct stats",
                      "Stock_Trends has AMZN row with correct stats",
                      "Product_Margins has >=5 categories",
                      "Correlation_Analysis has 2 entries",
                      "Strategic_Recommendations has rows for each category"]:
            check(label, False, "Excel missing")
        return
    check("Excel file exists", True)

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        check("Excel readable", False, str(e))
        return

    sheets = wb.sheetnames

    def _norm(s): return re.sub(r'[^a-z0-9]', '', s.lower())

    sheets_norm = {_norm(s): s for s in sheets}
    st_name = next((sheets_norm[k] for k in sheets_norm if "stocktrends" in k), None)
    pm_name = next((sheets_norm[k] for k in sheets_norm if "productmargins" in k), None)
    ca_name = next((sheets_norm[k] for k in sheets_norm if "correlationanalysis" in k), None)
    sr_name = next((sheets_norm[k] for k in sheets_norm if "strategicrecommendations" in k), None)

    check("Stock_Trends sheet present", st_name is not None, f"Sheets: {sheets}")
    check("Product_Margins sheet present", pm_name is not None, f"Sheets: {sheets}")
    check("Correlation_Analysis sheet present", ca_name is not None, f"Sheets: {sheets}")
    check("Strategic_Recommendations sheet present", sr_name is not None, f"Sheets: {sheets}")

    yf_expected = _fetch_yf_expected() or YF_FALLBACK

    # Stock_Trends
    if st_name:
        ws = wb[st_name]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            headers = [str(c).lower().strip() if c is not None else "" for c in rows[0]]
            data = [r for r in rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]

            def _get_col(names):
                for nm in names:
                    nm_l = nm.lower()
                    for i, h in enumerate(headers):
                        if h == nm_l or nm_l in h:
                            return i
                return None

            sym_idx = _get_col(["symbol"])
            avg_idx = _get_col(["avg_price", "avg price", "avg"])
            pct_idx = _get_col(["price_change", "change_pct", "pct"])
            vol_idx = _get_col(["volatility", "stddev", "std"])

            for ticker, label in [("GC=F", "GC=F"), ("AMZN", "AMZN")]:
                expected = yf_expected.get(ticker, {})
                row = next((r for r in data
                            if sym_idx is not None and r[sym_idx]
                            and ticker.lower() in str(r[sym_idx]).lower()), None)
                if not row:
                    check(f"Stock_Trends has {label} row with correct stats",
                          False, f"No row for {label}")
                    continue
                ok_avg = avg_idx is not None and num_close(row[avg_idx], expected.get("avg_price"), max(1.0, abs(expected.get("avg_price", 0)) * 0.02))
                ok_vol = vol_idx is not None and num_close(row[vol_idx], expected.get("volatility"), max(1.0, abs(expected.get("volatility", 0)) * 0.05))
                ok_pct = pct_idx is not None and num_close(row[pct_idx], expected.get("pct_change"), 5.0)
                check(f"Stock_Trends has {label} row with correct stats",
                      ok_avg and ok_vol and ok_pct,
                      f"row={row}, expected={expected}, ok_avg={ok_avg}, ok_vol={ok_vol}, ok_pct={ok_pct}")
        else:
            check("Stock_Trends has GC=F row with correct stats", False, "Empty sheet")
            check("Stock_Trends has AMZN row with correct stats", False, "Empty sheet")
    else:
        check("Stock_Trends has GC=F row with correct stats", False, "Sheet missing")
        check("Stock_Trends has AMZN row with correct stats", False, "Sheet missing")

    # Product_Margins
    if pm_name:
        ws = wb[pm_name]
        rows = list(ws.iter_rows(values_only=True))
        data = [r for r in rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]
        check("Product_Margins has >=5 categories", len(data) >= 5,
              f"Found {len(data)} category rows")

        # Check that key categories are present
        cats_lower = " ".join(str(r[0]).lower() for r in data if r and r[0])
        check("Product_Margins includes Audio + Cameras + Electronics",
              "audio" in cats_lower and "camera" in cats_lower and "electronics" in cats_lower,
              f"Categories: {[r[0] for r in data if r]}")

        # All margin_pct must be 40 (since cost is 60% of price)
        # Find margin column
        if rows:
            headers = [str(c).lower().strip() if c is not None else "" for c in rows[0]]
            margin_idx = next((i for i, h in enumerate(headers) if "margin" in h), None)
            if margin_idx is not None:
                bad = [r for r in data if not num_close(r[margin_idx], 40, 0.5)]
                check("All Product_Margins rows have margin_pct == 40",
                      len(bad) == 0,
                      f"Bad rows: {[(r[0], r[margin_idx]) for r in bad[:3]]}")
    else:
        check("Product_Margins has >=5 categories", False, "Sheet missing")
        check("Product_Margins includes Audio + Cameras + Electronics", False, "Sheet missing")

    # Correlation_Analysis
    if ca_name:
        ws = wb[ca_name]
        rows = list(ws.iter_rows(values_only=True))
        data = [r for r in rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]
        check("Correlation_Analysis has 2 entries", len(data) >= 2,
              f"Found {len(data)} rows")
        # Per-row keyword requirements: at least one row mentions gold (GC=F),
        # at least one row mentions amzn/amazon, and consumer should appear somewhere.
        all_text = " ".join(str(c) for r in rows for c in r if c is not None).lower()
        per_row_text = [" ".join(str(c) for c in r if c is not None).lower() for r in data]
        has_gold_row = any(("gold" in t) or ("gc=f" in t) for t in per_row_text)
        has_stock_row = any(("amzn" in t) or ("amazon" in t) for t in per_row_text)
        check("Correlation_Analysis: at least one row references gold/GC=F",
              has_gold_row,
              f"Rows: {per_row_text[:3]}")
        check("Correlation_Analysis: at least one row references AMZN/Amazon",
              has_stock_row,
              f"Rows: {per_row_text[:3]}")
        check("Correlation_Analysis mentions consumer somewhere",
              "consumer" in all_text,
              f"Text: {all_text[:200]}")
    else:
        check("Correlation_Analysis has 2 entries", False, "Sheet missing")
        check("Correlation_Analysis: at least one row references gold/GC=F", False, "Sheet missing")
        check("Correlation_Analysis: at least one row references AMZN/Amazon", False, "Sheet missing")
        check("Correlation_Analysis mentions consumer somewhere", False, "Sheet missing")

    # Strategic_Recommendations
    if sr_name:
        ws = wb[sr_name]
        rows = list(ws.iter_rows(values_only=True))
        data = [r for r in rows[1:] if any(c is not None and str(c).strip() != "" for c in r)]
        check("Strategic_Recommendations has rows for each category",
              len(data) >= 5,
              f"Found {len(data)}")
        if rows:
            headers = [str(c).lower().strip() if c is not None else "" for c in rows[0]]
            cur_idx = next((i for i, h in enumerate(headers) if "current" in h and "margin" in h), None)
            tar_idx = next((i for i, h in enumerate(headers) if "target" in h and "margin" in h), None)
            if cur_idx is not None and tar_idx is not None:
                # All current_margin should be 40
                cur_bad = [r for r in data if not num_close(r[cur_idx], 40, 0.5)]
                check("All current_margin == 40", len(cur_bad) == 0, f"Bad: {[(r[0], r[cur_idx]) for r in cur_bad[:3]]}")
                # Gold rose 138% > 50%, so target should be 35
                tar_bad = [r for r in data if not num_close(r[tar_idx], 35, 0.5)]
                check("Target margin == 35 (gold rose > 50%)",
                      len(tar_bad) == 0,
                      f"Bad: {[(r[0], r[tar_idx]) for r in tar_bad[:3]]}")
    else:
        check("Strategic_Recommendations has rows for each category", False, "Sheet missing")


def check_word(workspace):
    print("\n=== Check 2: Pricing_Strategy_Memo.docx ===")
    path = os.path.join(workspace, "Pricing_Strategy_Memo.docx")
    if not os.path.exists(path):
        check("Word document exists", False, f"Not found at {path}")
        for label in ["Word: title section", "Word: market overview section",
                      "Word: product margin section", "Word: strategic recommendations section",
                      "Word: contains gold pct number", "Word: substantial content"]:
            check(label, False, "Word missing")
        return
    check("Word document exists", True)

    try:
        doc = Document(path)
    except Exception as e:
        check("Word readable", False, str(e))
        return

    full_text = " ".join(p.text for p in doc.paragraphs)
    full_lower = full_text.lower()

    # Title check: require literal phrase 'Commodity Impact and Pricing Strategy'
    # Look in headings/title styles first, fallback to first paragraph
    headings = [p.text.strip() for p in doc.paragraphs
                if p.style and (p.style.name.startswith("Heading") or p.style.name == "Title")]
    headings_lower = [h.lower() for h in headings]
    first_para = (doc.paragraphs[0].text.strip().lower() if doc.paragraphs else "")
    expected_title = "commodity impact and pricing strategy"
    title_match = (
        any(expected_title in h for h in headings_lower)
        or expected_title in first_para
    )
    check("Word: title 'Commodity Impact and Pricing Strategy' present (heading or first paragraph)",
          title_match,
          f"Headings: {headings[:5]}; first para: {first_para[:120]}")
    check("Word: market overview section",
          "market overview" in full_lower or "market trends" in full_lower,
          "")
    check("Word: product margin section",
          "margin" in full_lower and ("category" in full_lower or "categories" in full_lower),
          "")
    check("Word: strategic recommendations section",
          "recommendation" in full_lower or "strategic" in full_lower,
          "")
    # Specific numbers expected
    check("Word: contains gold pct number (~138%)",
          re.search(r'\b13[0-9]\.?\d*%?', full_text) is not None or "138" in full_text,
          f"Text snippet: {full_text[:300]}")
    check("Word: substantial content", len(full_text) > 400, f"Length: {len(full_text)}")


def check_notion():
    print("\n=== Check 3: Notion Market Research Dashboard ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT id, title FROM notion.databases")
        dbs = cur.fetchall()
    except Exception as e:
        check("Notion DB query OK", False, str(e))
        return

    dashboard_db = None
    for db_id, title in dbs:
        title_str = ""
        if isinstance(title, list):
            title_str = " ".join(item.get("text", {}).get("content", "") for item in title if isinstance(item, dict))
        elif isinstance(title, str):
            try:
                parsed = json.loads(title)
                if isinstance(parsed, list):
                    title_str = " ".join(item.get("text", {}).get("content", "") for item in parsed if isinstance(item, dict))
                else:
                    title_str = str(title)
            except Exception:
                title_str = str(title)
        else:
            title_str = str(title) if title else ""
        # Stricter: must contain both "market research" and "dashboard"
        tl = title_str.lower().strip()
        if "market research dashboard" in tl:
            dashboard_db = (db_id, title_str)
            break
    check("'Market Research Dashboard' database exists",
          dashboard_db is not None,
          f"Databases: {[d[1] for d in dbs[:5]]}")

    if dashboard_db:
        try:
            cur.execute("""
                SELECT COUNT(*) FROM notion.pages
                WHERE parent->>'database_id' = %s
            """, (dashboard_db[0],))
            count = cur.fetchone()[0]
            check("Dashboard has exactly 2 entries", count == 2, f"Found {count}")
        except Exception as e:
            check("Dashboard entry count check", False, str(e))

    try:
        cur.close()
        conn.close()
    except Exception:
        pass


def check_script(workspace):
    print("\n=== Check 4: correlation_analysis.py ===")
    path = os.path.join(workspace, "correlation_analysis.py")
    check("correlation_analysis.py exists", os.path.exists(path))


def check_reverse_validation(workspace):
    """Verify things that should NOT exist in the output."""
    print("\n=== Reverse Validation ===")
    path = os.path.join(workspace, "Commodity_Impact_Analysis.xlsx")
    if os.path.exists(path):
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            expected_keywords = {"stock", "trend", "product", "margin", "correlation",
                                 "strategic", "recommend"}
            unexpected = [s for s in wb.sheetnames
                          if not any(kw in s.lower() for kw in expected_keywords)]
            check("No unexpected sheets in Excel", len(unexpected) == 0,
                  f"Unexpected: {unexpected}")
        except Exception:
            pass

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT id, title FROM notion.databases")
        dbs = cur.fetchall()
        market_dbs = []
        for db_id, title in dbs:
            title_str = json.dumps(title).lower() if title else ""
            if "market research dashboard" in title_str:
                market_dbs.append(db_id)
        check("No duplicate Market Research Dashboard databases",
              len(market_dbs) <= 1,
              f"Found {len(market_dbs)} matching")
        cur.close()
        conn.close()
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace)
    check_word(args.agent_workspace)
    check_notion()
    check_script(args.agent_workspace)
    check_reverse_validation(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed")

    result = {"total_passed": PASS_COUNT, "total_checks": total, "failed": FAIL_COUNT}
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
