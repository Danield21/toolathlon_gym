"""Evaluation for sf-sales-discount-gsheet-email."""
import argparse
import os
import sys
import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": int(os.environ.get("PGPORT", "5432")), "dbname": "toolathlon_gym", "user": "eigent", "password": "camel"}


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    # All checks for this task are DB-based (gsheet + email).
    file_errors = []
    db_errors = []

    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except Exception as e:
        db_errors.append(f"Could not connect to PostgreSQL: {e}")
        _print_and_exit(file_errors, db_errors)

    # Compute expected values from Snowflake data
    try:
        cur.execute('''
            SELECT c."SEGMENT", COUNT(o.*) as orders,
                   COUNT(CASE WHEN o."DISCOUNT" > 0 THEN 1 END) as disc_orders,
                   ROUND(100.0 * COUNT(CASE WHEN o."DISCOUNT" > 0 THEN 1 END)/COUNT(*)::numeric, 1) as disc_rate,
                   ROUND(AVG(CASE WHEN o."DISCOUNT" > 0 THEN o."DISCOUNT" END)::numeric * 100, 2) as avg_disc_pct,
                   ROUND(SUM(o."TOTAL_AMOUNT")::numeric, 2) as total_rev,
                   ROUND(SUM(CASE WHEN o."DISCOUNT" > 0 THEN o."TOTAL_AMOUNT" ELSE 0 END)::numeric, 2) as disc_rev,
                   ROUND(100.0 * SUM(CASE WHEN o."DISCOUNT" > 0 THEN o."TOTAL_AMOUNT" ELSE 0 END) / SUM(o."TOTAL_AMOUNT")::numeric, 1) as impact
            FROM sf_data."SALES_DW__PUBLIC__ORDERS" o
            JOIN sf_data."SALES_DW__PUBLIC__CUSTOMERS" c ON o."CUSTOMER_ID" = c."CUSTOMER_ID"
            WHERE o."STATUS" = 'Delivered'
            GROUP BY c."SEGMENT" ORDER BY disc_rate DESC
        ''')
        expected = cur.fetchall()
        expected_map = {r[0]: r for r in expected}
    except Exception as e:
        db_errors.append(f"Could not compute expected values: {e}")
        expected_map = {}

    # Check Google Sheet exists
    print("  Checking Google Sheet...")
    try:
        cur.execute("SELECT id FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%discount%analysis%'")
        sheets = cur.fetchall()
        if not sheets:
            db_errors.append("No Google Sheet with 'discount analysis' in title found")
        else:
            sheet_id = sheets[0][0]
            # Check for Segment Analysis sheet
            cur.execute("SELECT id FROM gsheet.sheets WHERE spreadsheet_id = %s AND LOWER(title) LIKE '%%segment%%'", (sheet_id,))
            seg_sheets = cur.fetchall()
            if not seg_sheets:
                db_errors.append("No 'Segment Analysis' sheet found in spreadsheet")
            else:
                seg_sheet_id = seg_sheets[0][0]
                cur.execute("SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id = %s ORDER BY row_index, col_index", (seg_sheet_id,))
                cells = cur.fetchall()
                # Build grid as {row_idx: {col_idx: value}}
                grid = {}
                for row_idx, col_idx, value in cells:
                    if row_idx not in grid:
                        grid[row_idx] = {}
                    grid[row_idx][col_idx] = value

                # ---- Dynamic header detection (avoid hardcoded col indices) ----
                def _norm(s):
                    return str(s or "").strip().lower().replace(" ", "_").replace("-", "_")

                # Aliases each canonical column may appear as.
                col_aliases = {
                    "segment":            ["segment", "customer_segment", "cust_segment"],
                    "order_count":        ["order_count", "orders", "total_orders", "num_orders"],
                    "discounted_orders":  ["discounted_orders", "disc_orders", "orders_discounted"],
                    "discount_rate_pct":  ["discount_rate_pct", "discount_rate", "disc_rate", "disc_rate_pct"],
                    "avg_discount_pct":   ["avg_discount_pct", "avg_disc_pct", "average_discount_pct"],
                    "total_revenue":      ["total_revenue", "revenue", "total_rev"],
                    "discounted_revenue": ["discounted_revenue", "disc_revenue", "discounted_rev"],
                    "revenue_impact_pct": ["revenue_impact_pct", "revenue_impact", "rev_impact_pct", "impact_pct"],
                }
                header_row_idx = None
                col_map = {}  # canonical_name -> col_idx
                for r in sorted(grid.keys()):
                    row_cells = grid[r]
                    norm_cells = {col: _norm(v) for col, v in row_cells.items()}
                    candidate_map = {}
                    for canon, aliases in col_aliases.items():
                        for col_idx, normed in norm_cells.items():
                            if normed in aliases:
                                candidate_map[canon] = col_idx
                                break
                    # Header row: must contain at least 'segment' and 4+ canonical columns mapped.
                    if "segment" in candidate_map and len(candidate_map) >= 4:
                        header_row_idx = r
                        col_map = candidate_map
                        break

                if header_row_idx is None:
                    db_errors.append(
                        f"Could not find header row mapping (need 'Segment' + 3 metric columns). "
                        f"Top rows: {dict(list(grid.items())[:3])}"
                    )
                else:
                    data_rows = {k: v for k, v in grid.items() if k > header_row_idx}
                    if expected_map and len(data_rows) < len(expected_map):
                        db_errors.append(f"Expected {len(expected_map)} data rows, found {len(data_rows)}")
                    elif expected_map:
                        for seg, exp in expected_map.items():
                            found = False
                            seg_col = col_map.get("segment")
                            for row_idx, row_data in data_rows.items():
                                seg_val = row_data.get(seg_col, "") if seg_col is not None else ""
                                if seg_val and str(seg_val).strip().lower() == seg.lower():
                                    found = True
                                    # Map canonical -> tolerance and exp index
                                    metric_specs = [
                                        ("order_count",        exp[1], 0,    "Order_Count"),
                                        ("discounted_orders",  exp[2], 0,    "Discounted_Orders"),
                                        ("discount_rate_pct",  exp[3], 0.2,  "Discount_Rate_Pct"),
                                        ("avg_discount_pct",   exp[4], 0.05, "Avg_Discount_Pct"),
                                        ("total_revenue",      exp[5], 1.0,  "Total_Revenue"),
                                        ("discounted_revenue", exp[6], 1.0,  "Discounted_Revenue"),
                                        ("revenue_impact_pct", exp[7], 0.2,  "Revenue_Impact_Pct"),
                                    ]
                                    for canon, exp_val, tol, label in metric_specs:
                                        col_idx = col_map.get(canon)
                                        if col_idx is None:
                                            db_errors.append(f"{seg}: column '{label}' not present in header")
                                            continue
                                        actual = row_data.get(col_idx, "")
                                        if not num_close(actual, exp_val, tol):
                                            db_errors.append(
                                                f"{seg}.{label}: {actual} vs {exp_val} (tol={tol})"
                                            )
                                    break
                            if not found:
                                db_errors.append(f"Segment '{seg}' not found in sheet")
    except Exception as e:
        db_errors.append(f"Google Sheet check error: {e}")

    # Check email - exact subject + recipient match
    print("  Checking email...")
    try:
        cur.execute("""SELECT subject, to_addr, COALESCE(body_text, body_html, '')
                       FROM email.messages""")
        email_rows = cur.fetchall()
        target_subj = "segment discount analysis"
        target_to = "finance-team@company.com"
        matched = None
        for subj, to_addr, body in email_rows:
            if target_subj in (subj or "").lower() and target_to in str(to_addr or "").lower():
                matched = (subj, to_addr, body)
                break
        if not matched:
            db_errors.append(f"No email '{target_subj}' to {target_to} found (checked {len(email_rows)})")
        else:
            body_l = (matched[2] or "").lower()
            if "discount" not in body_l or "segment" not in body_l:
                db_errors.append("Email body must mention 'discount' and 'segment'")
    except Exception as e:
        db_errors.append(f"Email check error: {e}")

    cur.close()
    conn.close()

    _print_and_exit(file_errors, db_errors)


def _print_and_exit(file_errors, db_errors):
    print(f"\n=== SUMMARY ===")
    print(f"  File errors: {len(file_errors)}")
    print(f"  DB errors:   {len(db_errors)}")
    if db_errors:
        for e in db_errors[:10]:
            print(f"    [DB] {e}")
    if file_errors:
        for e in file_errors[:10]:
            print(f"    [FILE] {e}")
    total_errors = len(file_errors) + len(db_errors)
    if total_errors == 0:
        print(f"  Overall: PASS")
        sys.exit(0)
    else:
        print(f"  Overall: FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
