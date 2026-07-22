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
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
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

EXPECTED_CITATIONS = {
    "2201.11903": 6500,
    "2203.11171": 3200,
    "2210.03493": 1800,
    "2205.11916": 4100,
    "2305.10601": 2400,
}

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_INCONCLUSIVE = 0
IN_RUNTIME_BLOCK = False


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_INCONCLUSIVE
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        if IN_RUNTIME_BLOCK:
            RUNTIME_INCONCLUSIVE += 1
            detail_str = f": {detail[:200]}" if detail else ""
            print(f"  [FAIL-runtime] {name}{detail_str}")
        else:
            FAIL_COUNT += 1
            detail_str = f": {detail[:200]}" if detail else ""
            print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol=500):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_gsheet_data():
    """Read Google Sheet data from the database."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Find spreadsheet
    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        ORDER BY created_at DESC
    """)
    spreadsheets = cur.fetchall()

    result = {"spreadsheet": None, "sheets": {}, "cells": {}}

    for ss_id, ss_title in spreadsheets:
        title_lower = (ss_title or "").lower()
        if "prompt" in title_lower or "literature" in title_lower or "engineering" in title_lower:
            result["spreadsheet"] = (ss_id, ss_title)
            break

    # Do NOT fall back to first spreadsheet - that's a FP risk
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
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Google Sheet ===")

    data = get_gsheet_data()

    # If absolutely no spreadsheets exist, flag as runtime-only (GT V1).
    # If some spreadsheets exist but none match expected title, do not downgrade - fail.
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM gsheet.spreadsheets")
    (n_ss,) = cur.fetchone()
    cur.close(); conn.close()
    if n_ss == 0:
        IN_RUNTIME_BLOCK = True

    # Check spreadsheet exists
    check("Spreadsheet exists", data["spreadsheet"] is not None,
          "No spreadsheet found")
    if not data["spreadsheet"]:
        IN_RUNTIME_BLOCK = False
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

        # Check that target paper IDs or titles appear
        all_values = " ".join(
            str(v).lower() for row in data_rows.values() for v in row.values()
        )

        found_papers = 0
        for title_kw in TARGET_TITLES_LOWER:
            if title_kw in all_values:
                found_papers += 1

        # Also check by ID
        for pid in TARGET_IDS:
            if pid in all_values:
                found_papers += 1

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

        # Noise papers must be ZERO
        noise_found = 0
        noise_detected = []
        for nid in NOISE_IDS:
            if nid in all_values:
                noise_found += 1
                noise_detected.append(nid)
        noise_titles = ["word2vec", "glove", "word representations", "distributed representations"]
        for nt in noise_titles:
            if nt in all_values:
                noise_found += 1
                noise_detected.append(nt)
        check("No noise papers in sheet (word2vec/glove/etc.)",
              noise_found == 0,
              f"Found noise: {noise_detected}")

        # Exactly 5 data rows
        check("Paper sheet has exactly 5 data rows",
              len(data_rows) == 5,
              f"Found {len(data_rows)} data rows")

        # Check citation counts for any found papers
        citation_checks = 0
        for row_data in data_rows.values():
            row_text = " ".join(str(v) for v in row_data.values())
            for pid, expected_count in EXPECTED_CITATIONS.items():
                if pid in row_text:
                    # Try to find a number close to expected citation count
                    numbers = re.findall(r'\d+', row_text)
                    for num_str in numbers:
                        if num_close(num_str, expected_count, 1000):
                            citation_checks += 1
                            break

        check("All 5 papers have approximately correct citation counts",
              citation_checks >= 5,
              f"Found {citation_checks}/5 papers with matching citations")

        # Every row must have non-empty Methodology_Summary (last column, col index 7)
        empty_methodology = 0
        for row_idx, row in data_rows.items():
            # Expected 8 columns: 0=Paper_ID ... 7=Methodology_Summary
            m = row.get(7)
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

        # Validate Reasoning_Type values (col 3) in {arithmetic, commonsense, symbolic, general}
        valid_types = {"arithmetic", "commonsense", "symbolic", "general"}
        invalid_types = []
        for row_idx, row in data_rows.items():
            rt = row.get(3)
            if rt is not None:
                val = str(rt).strip().lower()
                # Accept single-valued entries only
                if val and val not in valid_types:
                    invalid_types.append(val)
        check("All Reasoning_Type values are valid",
              len(invalid_types) == 0,
              f"Invalid values: {invalid_types}")

        # Validate Requires_Examples (col 4) in {Yes, No}
        invalid_req = []
        for row_idx, row in data_rows.items():
            re_val = row.get(4)
            if re_val is not None:
                v = str(re_val).strip().lower()
                if v and v not in {"yes", "no", "y", "n", "true", "false"}:
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
        # Check it mentions key papers
        papers_mentioned = sum(1 for kw in TARGET_TITLES_LOWER if kw in content_lower)
        check("review_summary mentions at least 3 papers",
              papers_mentioned >= 3,
              f"Found {papers_mentioned}/5 paper references")


def check_memory(agent_workspace):
    """Check memory file has been updated with entities."""
    print("\n=== Checking Memory ===")

    memory_path = os.path.join(agent_workspace, "memory", "memory.json")
    check("memory.json exists", os.path.isfile(memory_path),
          f"Not found at {memory_path}")

    if os.path.isfile(memory_path):
        with open(memory_path, "r") as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                check("memory.json is valid JSON", False, "JSON parse error")
                return

        check("memory.json is valid JSON", True)

        entities = data.get("entities", [])
        check("memory has at least 1 entity", len(entities) >= 1,
              f"Found {len(entities)} entities")

        # Check that at least one entity has observations
        has_observations = any(
            len(e.get("observations", [])) > 0 for e in entities
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
