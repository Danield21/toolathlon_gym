"""Evaluation script for arxiv-research-pipeline-notion-excel."""
import os
import argparse, json, os, sys
import openpyxl

def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel"
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
        detail_str = str(detail)[:200] if detail else ""
        print(f"  [FAIL] {name}: {detail_str}")

def safe_float(val, default=None):
    try:
        if val is None: return default
        return float(str(val).replace(",", "").replace("%", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return default

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    global PASS_COUNT, FAIL_COUNT
    PASS_COUNT = 0
    FAIL_COUNT = 0

    # Check Research_Knowledge_Base.xlsx
    excel_path = os.path.join(agent_workspace, "Research_Knowledge_Base.xlsx")
    check("Research_Knowledge_Base.xlsx exists", os.path.exists(excel_path))
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
        gt_path = os.path.join(groundtruth_workspace, "Research_Knowledge_Base.xlsx")
        gt_wb = openpyxl.load_workbook(gt_path) if os.path.exists(gt_path) else None

        if gt_wb:
            for sheet_name in gt_wb.sheetnames:
                check(f"{sheet_name} sheet exists", sheet_name in wb.sheetnames)
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    gt_ws = gt_wb[sheet_name]
                    # Check headers
                    gt_headers = [str(c.value).strip().lower() if c.value else "" for c in gt_ws[1]]
                    headers = [str(c.value).strip().lower() if c.value else "" for c in ws[1]]
                    for h in gt_headers:
                        if h:
                            check(f"{sheet_name} has {h} column", h in headers, f"headers: {headers[:10]}")
                    # Check row count
                    gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True)
                               if any(c is not None for c in r)]
                    data_rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                                 if any(c is not None for c in r)]

                    sn_lower = sheet_name.lower()
                    if "research_gap" in sn_lower or "research gap" in sn_lower:
                        # Task: at least 4 identified gaps.
                        check(f"{sheet_name} has >= 4 data rows", len(data_rows) >= 4,
                              f"got {len(data_rows)}")
                        # Validate Priority column values are in expected enum.
                        header_map = {h: i for i, h in enumerate(headers)}
                        prio_idx = header_map.get("priority", -1)
                        if prio_idx >= 0:
                            valid_priorities = {"critical", "important", "nice-to-have"}
                            invalid = []
                            for r in data_rows:
                                if prio_idx < len(r) and r[prio_idx] is not None:
                                    val = str(r[prio_idx]).strip().lower()
                                    if val and val not in valid_priorities:
                                        invalid.append(val)
                            check(f"{sheet_name} Priority values are in {{Critical,Important,Nice-to-have}}",
                                  not invalid, f"invalid: {invalid[:5]}")
                        # Verify all rows have non-empty Gap_Area, Current_State,
                        # Opportunity, Priority.
                        non_empty_ok = True
                        for h in ["gap_area", "current_state", "opportunity", "priority"]:
                            idx = header_map.get(h, -1)
                            if idx < 0:
                                continue
                            for r in data_rows:
                                if idx >= len(r) or r[idx] is None or str(r[idx]).strip() == "":
                                    non_empty_ok = False
                                    break
                            if not non_empty_ok:
                                break
                        check(f"{sheet_name} all rows non-empty in core columns",
                              non_empty_ok, "")
                    elif "paper_catalog" in sn_lower or "paper catalog" in sn_lower:
                        # Task: at least 5 relevant papers. GT has only 3 rows
                        # (illustrative). Validate >= 5 rows here.
                        check(f"{sheet_name} has >= 5 data rows (task: at least 5 relevant papers)",
                              len(data_rows) >= 5, f"got {len(data_rows)}")
                        # Validate columns Paper_ID, Title, Authors, Year are non-empty.
                        header_map = {h: i for i, h in enumerate(headers)}
                        non_empty_ok = True
                        for h in ["paper_id", "title", "authors", "year"]:
                            idx = header_map.get(h, -1)
                            if idx < 0:
                                continue
                            for r in data_rows:
                                if idx >= len(r) or r[idx] is None or str(r[idx]).strip() == "":
                                    non_empty_ok = False
                                    break
                            if not non_empty_ok:
                                break
                        check(f"{sheet_name} all rows non-empty in Paper_ID/Title/Authors/Year",
                              non_empty_ok, "")
                        # Validate sort order: Citation_Count descending.
                        ci_idx = header_map.get("citation_count", -1)
                        if ci_idx >= 0 and len(data_rows) >= 2:
                            sort_ok = True
                            prev = None
                            for r in data_rows:
                                if ci_idx >= len(r):
                                    continue
                                v = safe_float(r[ci_idx])
                                if v is None:
                                    continue
                                if prev is not None and v > prev + 0.001:
                                    sort_ok = False
                                    break
                                prev = v
                            check(f"{sheet_name} sorted by Citation_Count descending",
                                  sort_ok, "")
                        # Anti-fabrication: Paper_IDs must match real injected
                        # arxiv IDs (the 5 relevant ones). Reject the 2 noise
                        # papers (Quantum/Ocean) and any fabricated IDs.
                        pid_idx = header_map.get("paper_id", -1)
                        if pid_idx >= 0:
                            relevant_ids = {
                                "2301.00001",  # LLM Reasoning Survey
                                "2302.00002",  # Prompt Engineering Guide
                                "2303.00003",  # In-Context Learning Theory
                                "2304.00004",  # Chain-of-Thought Reasoning
                                "2305.00005",  # Survey of In-Context Learning
                            }
                            noise_ids = {"2304.99901", "2305.99902"}
                            seen_pids = []
                            for r in data_rows:
                                if pid_idx < len(r) and r[pid_idx] is not None:
                                    seen_pids.append(str(r[pid_idx]).strip())
                            included_relevant = relevant_ids & set(seen_pids)
                            included_noise = noise_ids & set(seen_pids)
                            check(
                                f"{sheet_name} includes the 5 relevant arxiv IDs (no fabricated catalog)",
                                len(included_relevant) == 5,
                                f"included relevant: {included_relevant}, all seen: {seen_pids}",
                            )
                            check(
                                f"{sheet_name} excludes noise papers (Quantum Computing, Ocean Modeling)",
                                len(included_noise) == 0,
                                f"noise leaked: {included_noise}",
                            )
                    elif "method_comparison" in sn_lower or "method comparison" in sn_lower:
                        # GT has 3 illustrative methods - any valid method choices
                        # acceptable. Just verify >=3 rows and required columns
                        # are populated.
                        check(f"{sheet_name} has >= 3 data rows", len(data_rows) >= 3,
                              f"got {len(data_rows)}")
                        header_map = {h: i for i, h in enumerate(headers)}
                        non_empty_ok = True
                        for h in ["method_name", "paper_source", "key_innovation", "applicability"]:
                            idx = header_map.get(h, -1)
                            if idx < 0:
                                continue
                            for r in data_rows:
                                if idx >= len(r) or r[idx] is None or str(r[idx]).strip() == "":
                                    non_empty_ok = False
                                    break
                            if not non_empty_ok:
                                break
                        check(f"{sheet_name} all rows non-empty in Method/Paper/Innovation/Applicability",
                              non_empty_ok, "")
                        # Validate Applicability enum
                        ap_idx = header_map.get("applicability", -1)
                        if ap_idx >= 0:
                            valid_ap = {"high", "medium", "low"}
                            invalid_ap = []
                            for r in data_rows:
                                if ap_idx < len(r) and r[ap_idx] is not None:
                                    val = str(r[ap_idx]).strip().lower()
                                    if val and val not in valid_ap:
                                        invalid_ap.append(val)
                            check(f"{sheet_name} Applicability values in {{High,Medium,Low}}",
                                  not invalid_ap, f"invalid: {invalid_ap[:5]}")
                    else:
                        # Default: >= GT row count.
                        check(f"{sheet_name} has >= {len(gt_rows)} data rows",
                              len(data_rows) >= len(gt_rows), f"got {len(data_rows)}")

    # Check the specifically-named Python script and JSON outputs exist.
    workspace_files = set(os.listdir(agent_workspace))
    check(
        "research_synthesizer.py exists",
        "research_synthesizer.py" in workspace_files,
        f"workspace: {sorted(workspace_files)[:25]}",
    )
    check(
        "papers_metadata.json exists",
        "papers_metadata.json" in workspace_files,
        f"workspace: {sorted(workspace_files)[:25]}",
    )
    check(
        "paper_contents.json exists",
        "paper_contents.json" in workspace_files,
        f"workspace: {sorted(workspace_files)[:25]}",
    )
    check(
        "research_synthesis.json exists",
        "research_synthesis.json" in workspace_files,
        f"workspace: {sorted(workspace_files)[:25]}",
    )

    # Database checks - validate Notion title and heading content.
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT id, properties FROM notion.pages WHERE archived = false")
        page_rows = cur.fetchall()

        page_count = len(page_rows)
        check("At least one non-archived Notion page exists", page_count >= 1,
              f"page count: {page_count}")

        # Find the page titled 'LLM Research Hub' (or similar).
        hub_page_id = None
        titles_seen = []
        for pid, props in page_rows:
            title = ""
            try:
                if isinstance(props, dict):
                    title_obj = props.get("title", {})
                    if isinstance(title_obj, dict):
                        title_list = title_obj.get("title", [])
                        if isinstance(title_list, list):
                            for t in title_list:
                                if isinstance(t, dict):
                                    title += t.get("text", {}).get("content", "")
            except Exception:
                title = ""
            titles_seen.append(title)
            tl = title.lower()
            if ("llm" in tl or "large language" in tl) and (
                "research hub" in tl or "research" in tl
            ):
                hub_page_id = pid
                break

        check(
            "Notion page 'LLM Research Hub' exists",
            hub_page_id is not None,
            f"Titles seen: {titles_seen[:10]}",
        )

        # Look for heading 'Large Language Model Research Dashboard' anywhere
        # in notion.blocks under that page (or any block).
        cur.execute(
            """
            SELECT parent_id, type, block_data FROM notion.blocks
            WHERE archived = false
            """
        )
        blocks = cur.fetchall()
        joined = " ".join(json.dumps(c) if c is not None else "" for _, _, c in blocks).lower()
        check(
            "Notion contains 'Large Language Model Research Dashboard' heading",
            "large language model research dashboard" in joined,
            f"sample: {joined[:200]}",
        )
        # Heuristic: page should mention each of 4 topics from task.md.
        # Use phrase-level matches (not loose substrings) to reduce false-positives.
        for topic, kw in [
            ("research landscape", "landscape"),
            ("key papers", "key paper"),
            ("methodology comparison", "methodology"),
            ("research gaps", "research gap"),
        ]:
            check(
                f"Notion page mentions topic: {topic}",
                kw in joined,
                f"sample: {joined[:200]}",
            )
        conn.close()
    except Exception as e:
        check("DB checks", False, str(e))

    return FAIL_COUNT == 0, f"Passed {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} checks"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    success, message = run_evaluation(
        args.agent_workspace, args.groundtruth_workspace,
        args.launch_time, args.res_log_file
    )
    print(message)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()