"""
Evaluation script for arxiv-lit-review-gsheet task.

Checks:
1. Google Sheet spreadsheet exists with "prompt engineering" or "literature review" in title
2. "Paper Comparison" sheet exists with at least 5 data rows
3. Paper IDs match the 5 injected target papers
4. Citation counts approximately match expected values
5. "Technique Analysis" sheet exists with at least 3 rows
6. review_summary.txt exists in workspace
7. Memory file has been updated with entities
"""

import argparse
import json
import os
import re
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
}

TARGET_IDS = ["2201.11903", "2203.11171", "2210.03493", "2205.11916", "2305.10601"]
NOISE_IDS = ["1301.03781", "1310.04546", "1405.01512"]

TARGET_TITLES_LOWER = [
    "chain-of-thought prompting",
    "self-consistency",
    "automatic chain of thought",
    "zero-shot reasoners",
    "tree of thoughts",
]

# Alias groups used when checking review_summary.txt: the summary counts as
# mentioning a paper if ANY of that paper's alias phrases appears.
TARGET_TITLE_ALIASES = [
    ["chain-of-thought", "chain of thought"],
    ["self-consistency", "self consistency"],
    ["automatic chain of thought", "auto-cot", "auto cot"],
    ["zero-shot reasoners", "zero-shot reasoning", "zero-shot-cot",
     "zero-shot cot", "zero shot cot", "zero shot"],
    ["tree of thoughts", "tree-of-thoughts", "tree of thought", "tree-of-thought"],
]

EXPECTED_CITATIONS = {
    "2201.11903": 6500,
    "2203.11171": 3200,
    "2210.03493": 1800,
    "2205.11916": 4100,
    "2305.10601": 2400,
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
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def _to_float(x):
    """Robustly coerce a value to float (None-safe).

    Handles int/float, strips currency symbols, thousands separators, spaces
    and trailing percent signs. Returns None when the value cannot be parsed.
    """
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return None
    s = s.lstrip("$¥€£").strip()
    if s.endswith("%"):
        s = s[:-1]
    s = s.replace(",", "").replace(" ", "")
    if not s:
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def num_close(a, b, tol=500):
    """Numeric closeness with a safe fallback to case-insensitive equality.

    Only falls back to string comparison when at least one side cannot be
    parsed as a number (previously the try/except swallowed the parse failure
    and returned False even for identical strings).
    """
    fa, fb = _to_float(a), _to_float(b)
    if fa is None or fb is None:
        return str(a).strip().lower() == str(b).strip().lower()
    return abs(fa - fb) <= tol


REASONING_VALID_TYPES = {"arithmetic", "commonsense", "symbolic", "general", "creative"}
REASONING_SUFFIXES = (" reasoning", " reasoning tasks", " tasks")


def _reasoning_token_ok(tok):
    """True when a Reasoning_Type token names one of the allowed categories.

    Tolerates common natural-language phrasings ("commonsense reasoning",
    "general reasoning tasks") by stripping a trailing descriptor, plus
    hyphen variants. Any other value still fails, so the check keeps its
    discriminating power.
    """
    t = tok.strip().lower().replace("-", " ")
    if t in REASONING_VALID_TYPES:
        return True
    for suffix in REASONING_SUFFIXES:
        core = t[: -len(suffix)].strip()
        if t.endswith(suffix) and core in REASONING_VALID_TYPES:
            return True
    return False


# Header aliases: map each logical field to keyword substrings that may appear
# in the header row (case-insensitive). Column locations are discovered from
# the header row instead of assuming a fixed column order, so an agent that
# follows the prompt but lays out columns differently is still graded fairly.
# Positional indexes are only used as a fallback when the header row is missing
# or unparseable.
PAPER_HEADER_ALIASES = {
    "Paper_ID": ["paper_id", "paper id", "arxiv id", "arxiv_id", "id"],
    "Title": ["title"],
    "Authors": ["authors", "author"],
    "Year": ["year"],
    "Venue": ["venue"],
    "Citation_Count": ["citation", "citations"],
    "Primary_Category": ["category", "primary_category"],
    "Methodology_Summary": ["methodology", "summary"],
}

TECHNIQUE_HEADER_ALIASES = {
    "Technique_Name": ["technique", "name"],
    "Paper_ID": ["paper_id", "paper id", "arxiv id", "arxiv_id", "id"],
    "Key_Innovation": ["innovation", "key_innovation"],
    "Reasoning_Type": ["reasoning"],
    "Requires_Examples": ["requires", "example"],
}


def find_column_indices(cells, aliases):
    """Return {field: col_index} discovered from the header row (row 0)."""
    header_row = cells.get(0, {})
    indices = {}
    for field, keywords in aliases.items():
        found = None
        for col_idx, hval in header_row.items():
            h = str(hval or "").strip().lower()
            for kw in keywords:
                if kw in h:
                    found = col_idx
                    break
            if found is not None:
                break
        if found is not None:
            indices[field] = found
    return indices


def get_gsheet_data():
    """Read Google Sheet data from the database."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Gather every spreadsheet whose title matches the expected theme and keep
    # the most complete one (most cells). In a multi-agent run an extra,
    # half-written duplicate can be created; preferring the fullest one avoids
    # grading a truncated copy.
    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        ORDER BY created_at DESC
    """)
    spreadsheets = cur.fetchall()

    result = {"spreadsheet": None, "sheets": {}, "cells": {}}

    best = None
    best_count = -1
    for ss_id, ss_title in spreadsheets:
        title_lower = (ss_title or "").lower()
        if "prompt" in title_lower or "literature" in title_lower or "engineering" in title_lower:
            cur.execute("SELECT COUNT(*) FROM gsheet.cells WHERE spreadsheet_id = %s", (ss_id,))
            (cnt,) = cur.fetchone()
            if cnt > best_count:
                best_count = cnt
                best = (ss_id, ss_title)
    result["spreadsheet"] = best

    # Do NOT fall back to the first spreadsheet - that's a FP risk
    # If no matching title, keep result["spreadsheet"] as None

    if result["spreadsheet"]:
        ss_id = result["spreadsheet"][0]

        # Get sheets
        cur.execute("""
            SELECT id, title FROM gsheet.sheets
            WHERE spreadsheet_id = %s
            ORDER BY index
        """, (ss_id,))
        for sheet_id, sheet_title in cur.fetchall():
            result["sheets"][sheet_title.lower()] = sheet_id

            # Get cells for this sheet
            cur.execute("""
                SELECT row_index, col_index, value
                FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                ORDER BY row_index, col_index
            """, (ss_id, sheet_id))

            cells = {}
            for row_idx, col_idx, value in cur.fetchall():
                if row_idx not in cells:
                    cells[row_idx] = {}
                cells[row_idx][col_idx] = value

            result["cells"][sheet_title.lower()] = cells

    cur.close()
    conn.close()
    return result


def check_gsheet():
    """Check Google Sheet content."""
    print("\n=== Checking Google Sheet ===")

    data = get_gsheet_data()

    # Check spreadsheet exists. A missing (or wrongly titled) spreadsheet is a
    # hard FAIL: with the DB connection configured correctly, "no spreadsheets"
    # can only mean the agent did not produce the required artifact.
    check("Spreadsheet exists", data["spreadsheet"] is not None,
          "No spreadsheet found")
    if not data["spreadsheet"]:
        return

    ss_id, ss_title = data["spreadsheet"]
    print(f"  Found spreadsheet: {ss_title}")

    # Check for Paper Comparison sheet
    paper_sheet_key = None
    for key in data["sheets"]:
        if "paper" in key or "comparison" in key or "index" in key:
            paper_sheet_key = key
            break

    check("Paper Comparison sheet exists", paper_sheet_key is not None,
          f"Sheets found: {list(data['sheets'].keys())}")

    if paper_sheet_key and paper_sheet_key in data["cells"]:
        cells = data["cells"][paper_sheet_key]
        # Count data rows (excluding header row 0)
        data_rows = {r: cells[r] for r in cells if r > 0}
        check("Paper sheet has at least 5 data rows",
              len(data_rows) >= 5,
              f"Found {len(data_rows)} data rows")

        # Discover columns from the header row (fall back to positional).
        col_idx = find_column_indices(cells, PAPER_HEADER_ALIASES)
        pid_col = col_idx.get("Paper_ID", 0)
        title_col = col_idx.get("Title", 1)
        methodology_col = col_idx.get("Methodology_Summary", 7)

        # Check each target paper independently (5/5 required)
        for i, (pid, title_kw) in enumerate(zip(TARGET_IDS, TARGET_TITLES_LOWER)):
            row_match = False
            for row_data in data_rows.values():
                row_text = " ".join(str(v or "").lower() for v in row_data.values())
                if pid in row_text or title_kw in row_text:
                    row_match = True
                    break
            check(f"Target paper present: {title_kw[:40]}", row_match,
                  f"ID={pid} not found")

        # Noise papers must be ZERO. Only inspect the Paper_ID and Title
        # columns: a legitimate mention of "word2vec"/"distributed
        # representations" inside a Methodology_Summary comparison paragraph
        # must not be mistaken for a noise paper being added to the sheet.
        noise_titles = ["word2vec", "glove", "word representations", "distributed representations"]
        noise_id_set = set(NOISE_IDS)
        noise_found = 0
        noise_detected = []
        for row_data in data_rows.values():
            pid_val = str(row_data.get(pid_col, "") or "").strip().lower()
            title_val = str(row_data.get(title_col, "") or "").strip().lower()
            if pid_val in noise_id_set:
                noise_found += 1
                noise_detected.append(pid_val)
            for nt in noise_titles:
                if nt in title_val:
                    noise_found += 1
                    noise_detected.append(nt)
        check("No noise papers in sheet (word2vec/glove/etc.)",
              noise_found == 0,
              f"Found noise: {noise_detected}")

        # Approximately 5 data rows (task prompt says "around 5"). The 5
        # targets are already required individually above, so a small range
        # tolerates an extra relevant paper without letting noise through.
        check("Paper sheet has roughly 5 data rows",
              5 <= len(data_rows) <= 7,
              f"Found {len(data_rows)} data rows")

        # Check citation counts for any found papers. Candidate numbers are
        # read from the Citation_Count column ONLY. Scanning the whole row
        # would let a venue year (e.g. "ICLR 2023" -> 2023) stand in for the
        # citation count of papers whose expected count sits near a year
        # (2210.03493 -> 1800, 2305.10601 -> 2400, both within 1000 of 2023).
        # Thousands separators (e.g. "6,500") and a moderate absolute delta
        # are still accepted; a blank/garbage citation cell now fails.
        citation_col = col_idx.get("Citation_Count", 5)
        citation_checks = 0
        citation_failures = []
        for row_data in data_rows.values():
            row_text = " ".join(str(v) for v in row_data.values())
            for pid, expected_count in EXPECTED_CITATIONS.items():
                if pid in row_text:
                    cell_val = row_data.get(citation_col)
                    numbers = re.findall(r'[\d,]+', str(cell_val or ""))
                    matched = any(
                        num_close(num_str.replace(",", ""), expected_count, 500)
                        for num_str in numbers
                    )
                    if matched:
                        citation_checks += 1
                    else:
                        citation_failures.append(f"{pid} (cell='{cell_val}')")

        check("All 5 papers have approximately correct citation counts",
              citation_checks >= 5,
              ("; ".join(citation_failures[:3])
               + (" ..." if len(citation_failures) > 3 else ""))
              if citation_failures
              else f"Found {citation_checks}/5 papers with matching citations")

        # Every row must have non-empty Methodology_Summary
        empty_methodology = 0
        for row_idx, row in data_rows.items():
            m = row.get(methodology_col)
            if m is None or not str(m).strip():
                empty_methodology += 1
        check("All rows have non-empty Methodology_Summary",
              empty_methodology == 0,
              f"{empty_methodology} rows with empty methodology")

    # Check for Technique Analysis sheet
    technique_sheet_key = None
    for key in data["sheets"]:
        if "technique" in key or "analysis" in key or "method" in key:
            technique_sheet_key = key
            break

    check("Technique Analysis sheet exists", technique_sheet_key is not None,
          f"Sheets found: {list(data['sheets'].keys())}")

    if technique_sheet_key and technique_sheet_key in data["cells"]:
        cells = data["cells"][technique_sheet_key]
        data_rows = {r: cells[r] for r in cells if r > 0}
        check("Technique sheet has at least 3 data rows",
              len(data_rows) >= 3,
              f"Found {len(data_rows)} data rows")

        col_idx = find_column_indices(cells, TECHNIQUE_HEADER_ALIASES)
        reasoning_col = col_idx.get("Reasoning_Type", 3)
        req_examples_col = col_idx.get("Requires_Examples", 4)

        # Check for technique-related content
        all_values = " ".join(
            str(v).lower() for row in data_rows.values() for v in row.values()
        )
        has_technique_content = any(
            kw in all_values for kw in [
                "chain", "thought", "self-consistency", "zero-shot",
                "tree", "auto", "prompting", "reasoning"
            ]
        )
        check("Technique sheet has prompting-related content",
              has_technique_content,
              "No prompting technique keywords found")

        # Validate Reasoning_Type values. The prompt lists a set of reasoning
        # categories and a single technique may target several, so accept one
        # or more comma-separated tokens, each within the accepted set
        # (including natural-language paraphrases of a category).
        invalid_types = []
        for row_idx, row in data_rows.items():
            rt = row.get(reasoning_col)
            if rt is not None:
                val = str(rt).strip().lower()
                if val:
                    tokens = [t.strip() for t in re.split(r'[,;/]+', val) if t.strip()]
                    for tok in tokens:
                        if not _reasoning_token_ok(tok):
                            invalid_types.append(val)
                            break
        check("All Reasoning_Type values are valid",
              len(invalid_types) == 0,
              f"Invalid values: {invalid_types}")

        # Validate Requires_Examples in {Yes, No}. A leading Yes/No answer is
        # accepted even when the cell also carries a brief justification
        # ("Yes, needs few-shot"), matching the task's "say Yes or No" rule.
        invalid_req = []
        for row_idx, row in data_rows.items():
            re_val = row.get(req_examples_col)
            if re_val is not None:
                v = str(re_val).strip().lower()
                if v:
                    first = v.split()[0].strip(".,;:!?/\\-")
                    if first not in {"yes", "no", "y", "n", "true", "false"}:
                        invalid_req.append(v)
        check("All Requires_Examples values are Yes/No",
              len(invalid_req) == 0,
              f"Invalid: {invalid_req}")


def check_review_summary(agent_workspace):
    """Check review_summary.txt exists and has content."""
    print("\n=== Checking review_summary.txt ===")

    summary_path = os.path.join(agent_workspace, "review_summary.txt")
    check("review_summary.txt exists", os.path.isfile(summary_path),
          f"Not found at {summary_path}")

    if os.path.isfile(summary_path):
        with open(summary_path, "r") as f:
            content = f.read()

        check("review_summary.txt has at least 200 characters",
              len(content.strip()) >= 200,
              f"File has {len(content.strip())} characters")

        content_lower = content.lower()
        # Check it mentions key papers. Alias groups allow phrasing variants
        # (e.g. "zero-shot reasoning", "Auto-CoT", "tree-of-thoughts").
        papers_mentioned = sum(
            1 for aliases in TARGET_TITLE_ALIASES
            if any(a in content_lower for a in aliases)
        )
        check("review_summary mentions at least 3 papers",
              papers_mentioned >= 3,
              f"Found {papers_mentioned}/5 paper references")


def load_memory_graph(memory_path):
    """Parse the memory file into {"entities": [...], "relations": [...]}.

    The memory MCP server writes JSONL (one JSON object per line, tagged with
    "type": "entity"/"relation"), while hand-written / ground-truth files use a
    wrapped object {"entities": [...], "relations": [...]}. Accept both, plus a
    bare list or a bare entity/relation object.
    """
    with open(memory_path, "r") as f:
        content = f.read().strip()
    if not content:
        return None

    # 1) Try whole-file JSON first.
    data = None
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        data = None

    if isinstance(data, dict):
        if "entities" in data or "relations" in data:
            return {
                "entities": data.get("entities") or [],
                "relations": data.get("relations") or [],
            }
        # A single entity/relation object written without a wrapper.
        if data.get("type") in ("entity", "relation"):
            if data["type"] == "entity":
                return {"entities": [data], "relations": []}
            return {"entities": [], "relations": [data]}
        return {"entities": [], "relations": []}

    if isinstance(data, list):
        return {"entities": data, "relations": []}

    # 2) JSONL fallback: one JSON object per line.
    entities, relations = [], []
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError:
            continue
        if not isinstance(item, dict):
            continue
        if item.get("type") == "entity":
            entities.append(item)
        elif item.get("type") == "relation":
            relations.append(item)
        elif "observations" in item or "name" in item:
            entities.append(item)
    if entities or relations:
        return {"entities": entities, "relations": relations}
    return None


def check_memory(agent_workspace):
    """Check memory file has been updated with entities."""
    print("\n=== Checking Memory ===")

    memory_path = os.path.join(agent_workspace, "memory", "memory.json")
    check("memory.json exists", os.path.isfile(memory_path),
          f"Not found at {memory_path}")

    if os.path.isfile(memory_path):
        try:
            data = load_memory_graph(memory_path)
        except OSError as e:
            check("memory.json is valid JSON", False, str(e))
            return

        if data is None:
            check("memory.json is valid JSON", False, "Empty or unparseable file")
            return
        check("memory.json is valid JSON", True)

        entities = data.get("entities", [])
        check("memory has at least 1 entity", len(entities) >= 1,
              f"Found {len(entities)} entities")

        # Check that at least one entity has observations
        has_observations = any(
            isinstance(e, dict) and len(e.get("observations") or []) > 0
            for e in entities
        )
        check("At least one entity has observations", has_observations,
              "No entities with observations found")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_review_summary(args.agent_workspace)
    check_memory(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks performed.")
        sys.exit(1)

    pass_rate = PASS_COUNT / total
    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Pass Rate: {pass_rate:.1%}")

    result = {
        "passed": PASS_COUNT,
        "failed": FAIL_COUNT,
        "pass_rate": round(pass_rate, 3),
        "success": FAIL_COUNT == 0,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
