"""
Evaluation for arxiv-latex-reasoning-gsheet task.
Checks Google Sheet and Word document.
"""
import argparse
import json
import os
import sys
import time

import psycopg2
from docx import Document

# DB connection is env-driven (PGHOST/PGPORT/PGDATABASE/PGUSER/PGPASSWORD).
# preprocess/main.py and evaluation/main.py must point at the same database,
# so both read the same environment with the same defaults.
DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
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
    ["automatic chain", "auto-cot", "automatic cot"],
    ["tree of thoughts"],
]

# Suggested method labels from task.md. The trailing generic fallbacks
# ("cot", "prompting") only widen acceptance for abbreviations a correct
# agent might write ("CoT", "... prompting"); they never reject a correct label.
SUGGESTED_METHODS = ["chain-of-thought", "chain of thought", "self-consistency", "zero-shot", "automatic cot", "auto-cot", "automatic chain", "tree of thoughts", "cot", "prompting"]

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


def _evaluate_spreadsheet(cur, ss_id, ss_title):
    """Run every sub-check against one candidate spreadsheet.

    A swarm run may create more than one spreadsheet whose title contains
    "reasoning" (e.g. one per sub-agent). We evaluate each candidate and let
    the caller keep the one that passes the most checks, so a complete and
    correct artifact is never masked by an incomplete duplicate.

    Returns (checks, passed) where checks is a list of
    (check_name, passed_bool, detail) tuples.
    """
    checks = []

    # Check "Papers" sheet exists
    cur.execute("""
        SELECT id, title FROM gsheet.sheets
        WHERE spreadsheet_id = %s
    """, (ss_id,))
    sheets = cur.fetchall()
    sheet_names = [t for _, t in sheets]

    papers_sheet_id = None
    for sid, sname in sheets:
        if sname and sname.strip().lower() == "papers":
            papers_sheet_id = sid
            break

    checks.append(("Sheet 'Papers' exists", papers_sheet_id is not None,
                   f"Found sheets: {sheet_names}"))

    if papers_sheet_id is None:
        return checks, sum(1 for _, p, _ in checks if p)

    # Read cells from Papers sheet
    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (ss_id, papers_sheet_id))
    cells = cur.fetchall()

    # Build grid
    grid = {}
    for row_idx, col_idx, val in cells:
        if row_idx not in grid:
            grid[row_idx] = {}
        grid[row_idx][col_idx] = val

    if not grid:
        checks.append(("Papers sheet has data", False, "No cells found"))
        return checks, sum(1 for _, p, _ in checks if p)

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

    checks.append(("Title column exists", title_col is not None, f"Header: {header_vals}"))
    checks.append(("Year column exists", year_col is not None, f"Header: {header_vals}"))
    checks.append(("Method column exists", method_col is not None, f"Header: {header_vals}"))
    checks.append(("Key_Contribution column exists", kc_col is not None, f"Header: {header_vals}"))

    # Check data rows
    data_rows = {r: grid[r] for r in grid if r > min_row}
    checks.append(("Papers sheet has 5 data rows", len(data_rows) == 5,
                   f"Found {len(data_rows)} rows"))

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
            checks.append((f"Has paper: {paper['title'][:50]}...", found, ""))

        # Check noise papers (word embedding) NOT in sheet
        all_titles_joined = " ".join(found_titles)
        for nk in NOISE_KEYWORDS:
            checks.append((f"Papers sheet does NOT include noise keyword '{nk}'",
                           nk.lower() not in all_titles_joined,
                           f"Found in titles: {all_titles_joined[:200]}"))

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
                checks.append((f"Row {idx+1} Year matches expected {matched['year']}", ok, f"Got: {yv_str}"))
            else:
                checks.append((f"Row {idx+1} Year is non-empty", bool(yv_str), f"Got: {yv_str}"))
        # Method check
        if method_col is not None:
            mv = str(row_data.get(method_col, "") or "").strip().lower()
            ok_nonempty = bool(mv)
            checks.append((f"Row {idx+1} Method non-empty", ok_nonempty, f"Got: {mv!r}"))
            if ok_nonempty:
                ok_suggested = any(sm in mv for sm in SUGGESTED_METHODS)
                checks.append((f"Row {idx+1} Method uses a suggested approach", ok_suggested, f"Got: {mv!r}"))
        # Key_Contribution non-empty + reasonable length
        if kc_col is not None:
            kcv = str(row_data.get(kc_col, "") or "").strip()
            checks.append((f"Row {idx+1} Key_Contribution non-empty", bool(kcv) and len(kcv) >= 10, f"Got: {kcv[:80]!r}"))

    passed = sum(1 for _, p, _ in checks if p)
    return checks, passed


def check_gsheet():
    """Check Google Sheet exists with correct data."""
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Find spreadsheets, most recently updated first, so that among
        # multiple "reasoning" spreadsheets we bias toward the final one.
        cur.execute("""
            SELECT id, title FROM gsheet.spreadsheets
            ORDER BY updated_at DESC, created_at DESC
        """)
        spreadsheets = cur.fetchall()

        # If NO spreadsheets at all, treat as runtime-only (agent didn't create).
        # If some spreadsheets exist but none match, fail hard - agent created wrong thing.
        if not spreadsheets:
            IN_RUNTIME_BLOCK = True
        candidates = [s for s in spreadsheets if s[1] and "reasoning" in s[1].lower()]

        record("Google Sheet 'Reasoning Methods Comparison' exists",
               bool(candidates),
               f"Found spreadsheets: {[t for _, t in spreadsheets]}")

        if not candidates:
            conn.close()
            IN_RUNTIME_BLOCK = False
            return

        # In a T3 swarm run several sub-agents may each create their own
        # "reasoning" spreadsheet. Evaluate every candidate and keep the best
        # one. "Best" means: prefer a candidate that passes *every* check (a
        # fully correct artifact must never be masked by a larger but imperfect
        # one, e.g. a table with an extra noise row that yields more per-row
        # checks to pass but also several FAILs); otherwise the most passed
        # checks, breaking ties toward the fewest failures.
        best_checks = None
        best_score = -1
        best_fail = None
        for sid, title in candidates:
            checks, score = _evaluate_spreadsheet(cur, sid, title)
            fail = len(checks) - score
            if best_checks is None:
                better = True
            elif fail == 0 and best_fail != 0:
                # A fully-correct candidate always wins over an imperfect one.
                better = True
            elif (fail == 0) == (best_fail == 0):
                # Same "has failures?" status: most passed wins, fewest
                # failures breaks ties.
                better = (score > best_score) or (score == best_score and fail < best_fail)
            else:
                # Best already has no failures and this candidate does.
                better = False
            if better:
                best_checks = checks
                best_score = score
                best_fail = fail

        for name, passed, detail in (best_checks or []):
            record(name, passed, detail)

        conn.close()
    except Exception as e:
        record("GSheet connection", False, str(e))


def _read_docx_text(doc_path):
    """Return the plain text of a .docx, trying python-docx first and then a
    raw XML extraction.

    A concurrent writer that rewrites the file non-atomically can leave it in
    a state python-docx rejects (package/metadata XML mangled) while the zip
    container and word/document.xml are still readable. The fallback pulls
    <w:t> text straight out of the archive so such a slightly-damaged file
    does not turn into a false FAIL. A genuinely truncated (non-zip) file
    still raises, which is correct - such an artifact is unreadable.
    """
    try:
        doc = Document(doc_path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        pass

    import re
    import zipfile
    from xml.sax.saxutils import unescape

    with zipfile.ZipFile(doc_path) as zf:
        xml_data = zf.read("word/document.xml").decode("utf-8", errors="replace")
    pieces = re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", xml_data, flags=re.DOTALL)
    return "\n".join(unescape(p) for p in pieces)


def check_word(agent_workspace):
    """Check Word document."""
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "Reasoning_Methods_Review.docx")

    if not os.path.isfile(doc_path):
        record("Word file exists", False, f"Not found: {doc_path}")
        return

    record("Word file exists", True)

    # Retry briefly in case a (concurrent) writer is mid-write when we open
    # the file; a transient lock should not turn into a false FAIL. Each
    # attempt uses the resilient reader in _read_docx_text.
    full_text = None
    last_err = None
    for _ in range(3):
        try:
            full_text = _read_docx_text(doc_path)
            break
        except Exception as e:
            last_err = e
            time.sleep(0.5)
    if full_text is None:
        record("Word file readable", False, str(last_err))
        return

    record("Word file readable", True)

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
