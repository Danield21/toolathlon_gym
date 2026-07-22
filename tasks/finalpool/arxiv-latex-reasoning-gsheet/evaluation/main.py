"""
Evaluation for arxiv-latex-reasoning-gsheet task.
Checks Google Sheet and Word document.
"""
import argparse
import json
import os
import sys

import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_ONLY_FAIL_COUNT = 0
IN_RUNTIME_BLOCK = False

# The 5 reasoning papers from arxiv_latex.papers (year derived from arxiv id YYMM.nnnn)
EXPECTED_REASONING_PAPERS = [
    {"id": "2201.11903", "title": "Chain-of-Thought Prompting Elicits Reasoning in Large Language Models", "year": 2022},
    {"id": "2203.11171", "title": "Self-Consistency Improves Chain of Thought Reasoning in Language Models", "year": 2022},
    {"id": "2205.11916", "title": "Large Language Models are Zero-Shot Reasoners", "year": 2022},
    {"id": "2210.03493", "title": "Automatic Chain of Thought Prompting in Large Language Models", "year": 2022},
    {"id": "2305.10601", "title": "Tree of Thoughts: Deliberate Problem Solving with Large Language Models", "year": 2023},
]

PAPER_KEYWORDS = [
    ["chain-of-thought prompting", "chain of thought prompting"],
    ["self-consistency"],
    ["zero-shot"],
    ["automatic chain", "auto-cot"],
    ["tree of thoughts"],
]

# Suggested method labels from task.md
SUGGESTED_METHODS = ["chain-of-thought", "chain of thought", "self-consistency", "zero-shot", "automatic cot", "auto-cot", "automatic chain", "tree of thoughts"]

# Word embedding papers that should NOT be included
NOISE_KEYWORDS = ["word2vec", "glove", "word representation", "skip-gram", "word embedding"]


def record(name, passed, detail="", runtime_only=False):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_ONLY_FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        # If this check is in a runtime-only block (e.g. GSheet with no data) or marked runtime_only,
        # don't block overall result but still log.
        if runtime_only or IN_RUNTIME_BLOCK:
            RUNTIME_ONLY_FAIL_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [FAIL-runtime] {name}{msg}")
        else:
            FAIL_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [FAIL] {name}{msg}")


def check_gsheet():
    """Check Google Sheet exists with correct data."""
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Find spreadsheet
        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        # If NO spreadsheets at all, treat as runtime-only (agent didn't create).
        # If some spreadsheets exist but none match, fail hard - agent created wrong thing.
        if not spreadsheets:
            IN_RUNTIME_BLOCK = True
        target_ss = None
        for sid, title in spreadsheets:
            if title and "reasoning" in title.lower():
                target_ss = sid
                break

        record("Google Sheet 'Reasoning Methods Comparison' exists",
               target_ss is not None,
               f"Found spreadsheets: {[t for _, t in spreadsheets]}")

        if target_ss is None:
            conn.close()
            IN_RUNTIME_BLOCK = False
            return

        # Check "Papers" sheet exists
        cur.execute("""
            SELECT id, title FROM gsheet.sheets
            WHERE spreadsheet_id = %s
        """, (target_ss,))
        sheets = cur.fetchall()
        sheet_names = [t for _, t in sheets]

        papers_sheet_id = None
        for sid, sname in sheets:
            if sname and sname.strip().lower() == "papers":
                papers_sheet_id = sid
                break

        record("Sheet 'Papers' exists", papers_sheet_id is not None,
               f"Found sheets: {sheet_names}")

        if papers_sheet_id is None:
            conn.close()
            return

        # Read cells from Papers sheet
        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
            ORDER BY row_index, col_index
        """, (target_ss, papers_sheet_id))
        cells = cur.fetchall()

        # Build grid
        grid = {}
        for row_idx, col_idx, val in cells:
            if row_idx not in grid:
                grid[row_idx] = {}
            grid[row_idx][col_idx] = val

        if not grid:
            record("Papers sheet has data", False, "No cells found")
            conn.close()
            return

        min_row = min(grid.keys())
        header_row = grid.get(min_row, {})
        header_vals = [header_row.get(i, "") for i in range(max(header_row.keys()) + 1)] if header_row else []

        # Find columns
        def find_col(key_substrings):
            for i, h in enumerate(header_vals):
                hn = str(h or "").strip().lower()
                for k in key_substrings:
                    if hn == k.lower() or k.lower() in hn:
                        return i
            return None

        title_col = find_col(["title"])
        year_col = find_col(["year"])
        method_col = find_col(["method"])
        kc_col = find_col(["contribution", "key_contribution", "key contribution"])

        record("Title column exists", title_col is not None, f"Header: {header_vals}")
        record("Year column exists", year_col is not None, f"Header: {header_vals}")
        record("Method column exists", method_col is not None, f"Header: {header_vals}")
        record("Key_Contribution column exists", kc_col is not None, f"Header: {header_vals}")

        # Check data rows
        data_rows = {r: grid[r] for r in grid if r > min_row}
        record("Papers sheet has 5 data rows", len(data_rows) == 5,
               f"Found {len(data_rows)} rows")

        # Check paper titles are present
        if title_col is not None:
            found_titles = []
            for r in sorted(data_rows.keys()):
                val = data_rows[r].get(title_col, "")
                if val:
                    found_titles.append(str(val).lower())

            for paper in EXPECTED_REASONING_PAPERS:
                t_lower = paper["title"].lower()
                found = any(t_lower in t or t in t_lower for t in found_titles)
                record(f"Has paper: {paper['title'][:50]}...", found)

            # Check noise papers (word embedding) NOT in sheet
            all_titles_joined = " ".join(found_titles)
            for nk in NOISE_KEYWORDS:
                record(f"Papers sheet does NOT include noise keyword '{nk}'",
                       nk.lower() not in all_titles_joined,
                       f"Found in titles: {all_titles_joined[:200]}")

        # Check Year/Method/Key_Contribution non-empty per row, Year matches expected
        sorted_rows = sorted(data_rows.keys())
        for idx, r in enumerate(sorted_rows):
            row_data = data_rows[r]
            title_val = str(row_data.get(title_col, "") or "").lower() if title_col is not None else ""
            # Find which expected paper this row corresponds to
            matched = None
            for paper in EXPECTED_REASONING_PAPERS:
                if paper["title"].lower() in title_val or title_val in paper["title"].lower():
                    matched = paper
                    break
            # Year check
            if year_col is not None:
                yv = row_data.get(year_col, "")
                yv_str = str(yv or "").strip()
                if matched:
                    # Should contain the expected year
                    ok = str(matched["year"]) in yv_str
                    record(f"Row {idx+1} Year matches expected {matched['year']}", ok, f"Got: {yv_str}")
                else:
                    record(f"Row {idx+1} Year is non-empty", bool(yv_str), f"Got: {yv_str}")
            # Method check
            if method_col is not None:
                mv = str(row_data.get(method_col, "") or "").strip().lower()
                ok_nonempty = bool(mv)
                record(f"Row {idx+1} Method non-empty", ok_nonempty, f"Got: {mv!r}")
                if ok_nonempty:
                    ok_suggested = any(sm in mv for sm in SUGGESTED_METHODS)
                    record(f"Row {idx+1} Method uses a suggested approach", ok_suggested, f"Got: {mv!r}")
            # Key_Contribution non-empty + reasonable length
            if kc_col is not None:
                kcv = str(row_data.get(kc_col, "") or "").strip()
                record(f"Row {idx+1} Key_Contribution non-empty", bool(kcv) and len(kcv) >= 10, f"Got: {kcv[:80]!r}")

        conn.close()
    except Exception as e:
        record("GSheet connection", False, str(e))


def check_word(agent_workspace):
    """Check Word document."""
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "Reasoning_Methods_Review.docx")

    if not os.path.isfile(doc_path):
        record("Word file exists", False, f"Not found: {doc_path}")
        return

    record("Word file exists", True)

    try:
        doc = Document(doc_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        return

    record("Word file readable", True)

    full_text = "\n".join(p.text for p in doc.paragraphs)
    full_lower = full_text.lower()

    # Check title
    has_title = "chain-of-thought" in full_lower and "reasoning" in full_lower and "comparison" in full_lower
    if not has_title:
        has_title = "chain of thought" in full_lower and "methods" in full_lower
    record("Word doc has correct title", has_title)

    # Check date
    has_date = "2026-03-06" in full_text or "march 6, 2026" in full_lower or "march 2026" in full_lower
    record("Word doc has date", has_date)

    # Check each reasoning paper is mentioned
    for i, paper in enumerate(EXPECTED_REASONING_PAPERS):
        found = any(kw in full_lower for kw in PAPER_KEYWORDS[i])
        record(f"Word mentions: {paper['title'][:50]}...", found,
               f"Keywords: {PAPER_KEYWORDS[i]}")

    # Check noise papers NOT mentioned
    for noise in NOISE_KEYWORDS:
        absent = noise not in full_lower
        record(f"Word does NOT mention: {noise}", absent)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_word(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== Results: {PASS_COUNT}/{total} passed ({RUNTIME_ONLY_FAIL_COUNT} runtime-only inconclusive) ===")
    if FAIL_COUNT > 0:
        print(f"{FAIL_COUNT} checks failed")
        sys.exit(1)
    else:
        print("All checks passed!")
        sys.exit(0)


if __name__ == "__main__":
    main()
