"""Evaluation for yf-gold-performance-gsheet-notion task.

Check 1: Google Sheet "Gold vs Stocks Analysis" with "Returns Comparison" sheet
- All 6 symbols pinned to specific rows; per-row return values must match DB
- Rows sorted by Return_1Y_Pct descending
- Beat_Gold columns Yes/No correct per row

Check 2: Notion page titled "Gold vs Stocks Performance"
- Mentions gold's 6M and 1Y returns
- States count of stocks outperforming gold over each period
- Mentions top performer
"""

import argparse
import json
import os
import re
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"),
          port=int(os.environ.get("PGPORT", "5432")),
          dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol=2.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_data():
    """Compute expected returns from DB. Returns (results_sorted_by_1y_desc, gold_6m, gold_1y)."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    symbols = ['AMZN', 'GC=F', 'GOOGL', 'JNJ', 'JPM', 'XOM']

    cur.execute("""SELECT symbol, close FROM yf.stock_prices
        WHERE date = (SELECT MAX(date) FROM yf.stock_prices WHERE symbol='AMZN')
        AND symbol = ANY(%s)""", (symbols,))
    latest = dict(cur.fetchall())

    cur.execute("""
        WITH ranked AS (SELECT symbol, date, close, ROW_NUMBER() OVER(PARTITION BY symbol ORDER BY date DESC) as rn
        FROM yf.stock_prices WHERE symbol = ANY(%s))
        SELECT symbol, close FROM ranked WHERE rn = 127
    """, (symbols,))
    six_m = dict(cur.fetchall())

    cur.execute("""
        WITH ranked AS (SELECT symbol, date, close, ROW_NUMBER() OVER(PARTITION BY symbol ORDER BY date DESC) as rn
        FROM yf.stock_prices WHERE symbol = ANY(%s))
        SELECT symbol, close FROM ranked WHERE rn = 253
    """, (symbols,))
    one_y = dict(cur.fetchall())

    cur.execute("""SELECT symbol, data->>'shortName' FROM yf.stock_info
        WHERE symbol = ANY(%s)""", (symbols,))
    names = dict(cur.fetchall())

    cur.close()
    conn.close()

    if 'GC=F' not in latest or 'GC=F' not in six_m or 'GC=F' not in one_y:
        return None, None, None

    gold_6m_ret = round((float(latest['GC=F']) - float(six_m['GC=F'])) / float(six_m['GC=F']) * 100, 2)
    gold_1y_ret = round((float(latest['GC=F']) - float(one_y['GC=F'])) / float(one_y['GC=F']) * 100, 2)

    results = []
    for s in symbols:
        if s not in latest or s not in six_m or s not in one_y:
            continue
        l = float(latest[s])
        s6 = float(six_m[s])
        s1 = float(one_y[s])
        ret_6m = round((l - s6) / s6 * 100, 2)
        ret_1y = round((l - s1) / s1 * 100, 2)
        beat_6m = "Yes" if ret_6m > gold_6m_ret and s != 'GC=F' else "No"
        beat_1y = "Yes" if ret_1y > gold_1y_ret and s != 'GC=F' else "No"
        results.append({
            "symbol": s, "name": names.get(s, s),
            "latest": l, "price_6m": s6, "ret_6m": ret_6m,
            "price_1y": s1, "ret_1y": ret_1y,
            "beat_6m": beat_6m, "beat_1y": beat_1y,
        })

    results.sort(key=lambda x: x["ret_1y"], reverse=True)
    return results, gold_6m_ret, gold_1y_ret


def parse_num(v):
    """Parse cell value to float; strip %, , and quotes."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().strip('"').strip("'")
    s = s.replace("%", "").replace(",", "")
    m = re.match(r"^-?\d+(?:\.\d+)?$", s)
    if m:
        try:
            return float(s)
        except Exception:
            return None
    # try first number in string
    m2 = re.search(r"-?\d+(?:\.\d+)?", s)
    if m2:
        try:
            return float(m2.group(0))
        except Exception:
            return None
    return None


def check_gsheet():
    """Check the Google Sheet with returns comparison data."""
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except Exception as e:
        check("DB connection", False, str(e))
        return

    try:
        cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%gold%'")
        spreadsheets = cur.fetchall()
    except Exception as e:
        check("Spreadsheet query", False, str(e))
        cur.close(); conn.close()
        return

    check("Spreadsheet with 'Gold' in title exists", len(spreadsheets) > 0,
          "No spreadsheet found with 'Gold' in title")
    if not spreadsheets:
        # cascade fail downstream so they aren't omitted
        for k in ["Sheet 'Returns Comparison' exists", "Sheet has >= 6 data rows",
                  "All 6 symbols present (per row)", "All 1Y returns match (per row)",
                  "All 6M returns match (per row)", "Beat_Gold_6M column correct (per row)",
                  "Beat_Gold_1Y column correct (per row)", "Rows sorted by 1Y return desc"]:
            check(k, False, "skipped (no spreadsheet)")
        cur.close(); conn.close()
        return

    ss_id = spreadsheets[0][0]
    print(f"  Found spreadsheet: '{spreadsheets[0][1]}' (id={ss_id})")

    cur.execute("""
        SELECT id FROM gsheet.sheets
        WHERE spreadsheet_id = %s AND LOWER(title) LIKE '%%return%%'
    """, (ss_id,))
    sheets = cur.fetchall()
    check("Sheet 'Returns Comparison' exists", len(sheets) > 0)
    if not sheets:
        for k in ["Sheet has >= 6 data rows",
                  "All 6 symbols present (per row)", "All 1Y returns match (per row)",
                  "All 6M returns match (per row)", "Beat_Gold_6M column correct (per row)",
                  "Beat_Gold_1Y column correct (per row)", "Rows sorted by 1Y return desc"]:
            check(k, False, "skipped (no sheet)")
        cur.close(); conn.close()
        return

    sheet_id = sheets[0][0]

    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (ss_id, sheet_id))
    cells = cur.fetchall()
    cur.close()
    conn.close()

    grid = {}
    for row_idx, col_idx, value in cells:
        grid[(row_idx, col_idx)] = value

    # Build header row by inspecting row 0
    max_col = max((c for r, c in grid.keys() if r == 0), default=-1)
    header = [str(grid.get((0, ci), "") or "").strip().lower().replace(" ", "_") for ci in range(max_col + 1)]
    # Find columns by name
    def col_idx(*kws):
        for i, h in enumerate(header):
            for k in kws:
                if k in h:
                    return i
        return -1
    ci_sym = col_idx("symbol")
    ci_ret_6m = col_idx("return_6m", "ret_6m", "6m_pct")
    ci_ret_1y = col_idx("return_1y", "ret_1y", "1y_pct")
    ci_beat_6m = col_idx("beat_gold_6m", "beat_6m")
    ci_beat_1y = col_idx("beat_gold_1y", "beat_1y")

    max_row = max((r for r, c in grid.keys()), default=0)
    data_row_indices = [r for r in range(1, max_row + 1)
                        if any(grid.get((r, c)) not in (None, "") for c in range(max_col + 1))]
    check("Sheet has >= 6 data rows", len(data_row_indices) >= 6, f"row count={len(data_row_indices)}")

    expected_data, gold_6m, gold_1y = get_expected_data()
    if expected_data is None:
        check("DB expected data fetched", False, "missing GC=F data")
        return
    expected_by_sym = {e["symbol"]: e for e in expected_data}

    # Build agent rows by symbol
    agent_rows = {}
    if ci_sym >= 0:
        for r in data_row_indices:
            sym_v = grid.get((r, ci_sym))
            if sym_v:
                agent_rows[str(sym_v).strip().upper()] = r

    expected_syms = ["AMZN", "GC=F", "GOOGL", "JNJ", "JPM", "XOM"]
    sym_present_count = sum(1 for s in expected_syms if s in agent_rows)
    check("All 6 symbols present (per row)", sym_present_count == 6,
          f"present: {sorted(agent_rows.keys())}")

    # Per-row 1Y return
    ret_1y_ok = True
    detail_1y = []
    for s in expected_syms:
        if s not in agent_rows or ci_ret_1y < 0:
            ret_1y_ok = False
            detail_1y.append(f"{s} no row/col")
            continue
        v = parse_num(grid.get((agent_rows[s], ci_ret_1y)))
        if v is None or not num_close(v, expected_by_sym[s]["ret_1y"], 1.0):
            ret_1y_ok = False
            detail_1y.append(f"{s}: got {v} expected {expected_by_sym[s]['ret_1y']}")
    check("All 1Y returns match (per row)", ret_1y_ok, "; ".join(detail_1y)[:280])

    ret_6m_ok = True
    detail_6m = []
    for s in expected_syms:
        if s not in agent_rows or ci_ret_6m < 0:
            ret_6m_ok = False
            detail_6m.append(f"{s} no row/col")
            continue
        v = parse_num(grid.get((agent_rows[s], ci_ret_6m)))
        if v is None or not num_close(v, expected_by_sym[s]["ret_6m"], 1.0):
            ret_6m_ok = False
            detail_6m.append(f"{s}: got {v} expected {expected_by_sym[s]['ret_6m']}")
    check("All 6M returns match (per row)", ret_6m_ok, "; ".join(detail_6m)[:280])

    # Beat_Gold checks
    def cell_norm(v):
        return str(v or "").strip().lower()
    beat_6m_ok = True
    beat_6m_detail = []
    for s in expected_syms:
        if s not in agent_rows or ci_beat_6m < 0:
            beat_6m_ok = False
            beat_6m_detail.append(f"{s} no row/col")
            continue
        got = cell_norm(grid.get((agent_rows[s], ci_beat_6m)))
        exp = expected_by_sym[s]["beat_6m"].lower()
        if got != exp:
            beat_6m_ok = False
            beat_6m_detail.append(f"{s}: got {got!r} expected {exp!r}")
    check("Beat_Gold_6M column correct (per row)", beat_6m_ok, "; ".join(beat_6m_detail)[:280])

    beat_1y_ok = True
    beat_1y_detail = []
    for s in expected_syms:
        if s not in agent_rows or ci_beat_1y < 0:
            beat_1y_ok = False
            beat_1y_detail.append(f"{s} no row/col")
            continue
        got = cell_norm(grid.get((agent_rows[s], ci_beat_1y)))
        exp = expected_by_sym[s]["beat_1y"].lower()
        if got != exp:
            beat_1y_ok = False
            beat_1y_detail.append(f"{s}: got {got!r} expected {exp!r}")
    check("Beat_Gold_1Y column correct (per row)", beat_1y_ok, "; ".join(beat_1y_detail)[:280])

    # Sort order by 1Y return descending
    expected_order = [e["symbol"] for e in expected_data]
    actual_order = []
    for r in data_row_indices:
        if ci_sym >= 0:
            sym_v = grid.get((r, ci_sym))
            if sym_v:
                actual_order.append(str(sym_v).strip().upper())
    check("Rows sorted by 1Y return desc", actual_order == expected_order,
          f"got {actual_order}, expected {expected_order}")


def check_notion():
    """Check Notion page 'Gold vs Stocks Performance'."""
    print("\n=== Checking Notion Page ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("SELECT id, properties FROM notion.pages")
        pages = cur.fetchall()
    except Exception as e:
        check("Notion DB query", False, str(e))
        try:
            cur.close(); conn.close()
        except Exception:
            pass
        for k in ["Notion 'Gold vs Stocks Performance' page", "Notion page has content blocks",
                  "Notion mentions gold's 6M and 1Y returns",
                  "Notion mentions count of outperforming stocks",
                  "Notion mentions top 1Y performer"]:
            check(k, False, "skipped")
        return

    found_page = None
    for page_id, props in pages:
        if not props:
            continue
        # extract title text
        title_text = ""
        try:
            for k, v in props.items():
                if isinstance(v, dict) and v.get("type") == "title":
                    items = v.get("title", [])
                    title_text = " ".join(
                        i.get("text", {}).get("content", "") for i in items if isinstance(i, dict)
                    )
                    break
        except Exception:
            pass
        title_l = title_text.lower()
        if "gold" in title_l and ("stock" in title_l or "performance" in title_l):
            found_page = (page_id, props)
            break
    check("Notion 'Gold vs Stocks Performance' page", found_page is not None,
          f"found {len(pages)} pages")

    if not found_page:
        cur.close(); conn.close()
        for k in ["Notion page has content blocks",
                  "Notion mentions gold's 6M and 1Y returns",
                  "Notion mentions count of outperforming stocks",
                  "Notion mentions top 1Y performer"]:
            check(k, False, "skipped (no page)")
        return

    page_id = found_page[0]
    cur.execute("SELECT block_data FROM notion.blocks WHERE parent_id = %s", (page_id,))
    blocks = cur.fetchall()
    cur.close(); conn.close()

    check("Notion page has content blocks", len(blocks) > 0, f"blocks={len(blocks)}")

    all_text = " ".join(json.dumps(b[0]) if isinstance(b[0], dict) else str(b[0])
                        for b in blocks if b[0]).lower()

    expected_data, gold_6m, gold_1y = get_expected_data()
    if expected_data is None:
        check("DB expected data for notion", False, "no GC=F prices")
        return

    # Mentions gold's 6M and 1Y - require both numbers present (within tol)
    def number_in_text(target, text, tol=1.5):
        for m in re.finditer(r"-?\d+(?:\.\d+)?", text):
            try:
                if abs(float(m.group(0)) - target) <= tol:
                    return True
            except Exception:
                pass
        return False
    has_6m = number_in_text(gold_6m, all_text)
    has_1y = number_in_text(gold_1y, all_text)
    check("Notion mentions gold's 6M and 1Y returns",
          has_6m and has_1y, f"gold_6m={gold_6m}, 1y={gold_1y}")

    # Mentions count of outperforming stocks (6M and 1Y counts)
    n_beat_6m = sum(1 for e in expected_data if e["beat_6m"] == "Yes")
    n_beat_1y = sum(1 for e in expected_data if e["beat_1y"] == "Yes")
    # numbers 0..5 — cannot use simple substring; check phrase
    # accept either "<n_beat> of (the )?five" or "<n_beat>" near "outperform"/"beat"
    has_count = False
    pattern_options = []
    for n in [n_beat_6m, n_beat_1y]:
        # word number variants
        word_map = {0: "zero", 1: "one", 2: "two", 3: "three", 4: "four", 5: "five"}
        words = [str(n), word_map.get(n, "")]
        for w in words:
            if not w:
                continue
            # must appear near outperform/beat
            for m in re.finditer(rf"\b{re.escape(w)}\b", all_text):
                start = max(0, m.start() - 60)
                end = min(len(all_text), m.end() + 60)
                window = all_text[start:end]
                if any(kw in window for kw in ["outperform", "beat", "exceed", "of the five", "out of"]):
                    pattern_options.append(n)
                    break
    has_count = len(set(pattern_options)) >= 2 or (n_beat_6m == n_beat_1y and n_beat_6m in pattern_options)
    check("Notion mentions count of outperforming stocks", has_count,
          f"6M-count={n_beat_6m}, 1Y-count={n_beat_1y}")

    # Top 1Y performer mention
    top = expected_data[0]["symbol"]
    top_kw = expected_data[0]["name"] or top
    # For GC=F handle as "gold"
    if top == "GC=F":
        ok_top = "gold" in all_text and ("highest" in all_text or "top" in all_text or "best" in all_text)
    else:
        ok_top = (top.lower() in all_text or top_kw.lower() in all_text)
    check("Notion mentions top 1Y performer", ok_top, f"top={top}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("YF GOLD PERFORMANCE GSHEET NOTION - EVALUATION")
    print("=" * 70)

    check_gsheet()
    check_notion()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
