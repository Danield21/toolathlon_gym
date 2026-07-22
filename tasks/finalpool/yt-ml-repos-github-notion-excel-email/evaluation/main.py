"""
Evaluation for yt-ml-repos-github-notion-excel-email task.

Checks:
1. ML_Research_Tracker.xlsx exists
2. "Videos" sheet has >= 5 data rows, has Video_ID and Title columns
3. "Papers" sheet has >= 3 data rows, has ArXiv_ID and Title columns
4. "Summary" sheet has >= 3 rows
5. Notion page exists with "ML Tech" or "Research" in title
6. Email sent to research@lab.edu
"""
import os
import sys
import json
from argparse import ArgumentParser

import psycopg2
import openpyxl

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def check_excel(agent_workspace, groundtruth_workspace="."):
    print("\n=== Check 1-4: ML_Research_Tracker.xlsx ===")

    xlsx_path = os.path.join(agent_workspace, "ML_Research_Tracker.xlsx")
    if not os.path.exists(xlsx_path):
        record("ML_Research_Tracker.xlsx exists", False, f"Not found at {xlsx_path}")
        for msg in ["Videos sheet has >= 5 rows with Video_ID and Title",
                    "Papers sheet has >= 3 rows with ArXiv_ID and Title",
                    "Summary sheet has >= 3 rows"]:
            record(msg, False, "File missing")
        return
    record("ML_Research_Tracker.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path)
    except Exception as e:
        record("Excel readable", False, str(e))
        return

    sheet_names_lower = {s.lower(): s for s in wb.sheetnames}

    # Videos sheet — exactly 7 rows; precise headers
    videos_key = sheet_names_lower.get("videos") or next((sheet_names_lower[k] for k in sheet_names_lower if "video" in k), None)
    if not videos_key:
        record("Videos sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        ws = wb[videos_key]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)] if rows else []
        headers = [str(c).strip().lower().replace(" ", "_") if c else "" for c in rows[0]] if rows else []
        expected_h = ["video_id", "title", "topic", "github_url", "published_date", "view_count"]
        for i, eh in enumerate(expected_h):
            ah = headers[i] if i < len(headers) else "MISSING"
            record(f"Videos header[{i}] = '{eh}'", ah == eh, f"Got '{ah}'")
        record("Videos sheet has exactly 7 data rows",
               len(data_rows) == 7,
               f"Rows: {len(data_rows)}")

    # Papers sheet — at least 5 rows (1+ paper per topic possible, GT has 5)
    papers_key = sheet_names_lower.get("papers") or next((sheet_names_lower[k] for k in sheet_names_lower if "paper" in k), None)
    if not papers_key:
        record("Papers sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        ws2 = wb[papers_key]
        rows2 = list(ws2.iter_rows(values_only=True))
        data_rows2 = [r for r in rows2[1:] if any(c for c in r)] if rows2 else []
        headers2 = [str(c).strip().lower().replace(" ", "_") if c else "" for c in rows2[0]] if rows2 else []
        # Strict header check (note: 'arxiv_id' explicit; 'id' alone is too permissive)
        expected_h2 = ["arxiv_id", "title", "authors", "published", "topic", "related_video_id"]
        for i, eh in enumerate(expected_h2):
            ah = headers2[i] if i < len(headers2) else "MISSING"
            record(f"Papers header[{i}] = '{eh}'", ah == eh, f"Got '{ah}'")
        record("Papers sheet has >= 5 data rows",
               len(data_rows2) >= 5,
               f"Rows: {len(data_rows2)}")

    # Summary sheet — exactly 7 (one per topic)
    summary_key = sheet_names_lower.get("summary") or next((sheet_names_lower[k] for k in sheet_names_lower if "summar" in k), None)
    if not summary_key:
        record("Summary sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        ws3 = wb[summary_key]
        rows3 = list(ws3.iter_rows(values_only=True))
        data_rows3 = [r for r in rows3[1:] if any(c for c in r)] if rows3 else []
        headers3 = [str(c).strip().lower().replace(" ", "_") if c else "" for c in rows3[0]] if rows3 else []
        expected_h3 = ["topic", "video_count", "paper_count", "github_repos_count"]
        for i, eh in enumerate(expected_h3):
            ah = headers3[i] if i < len(headers3) else "MISSING"
            record(f"Summary header[{i}] = '{eh}'", ah == eh, f"Got '{ah}'")
        record("Summary sheet has exactly 7 data rows",
               len(data_rows3) == 7,
               f"Rows: {len(data_rows3)}")

    # --- Groundtruth XLSX value comparison ---
    gt_path = os.path.join(groundtruth_workspace, "ML_Research_Tracker.xlsx")
    if os.path.isfile(gt_path):
        gt_wb = openpyxl.load_workbook(gt_path, data_only=True)
        for gt_sname in gt_wb.sheetnames:
            gt_ws = gt_wb[gt_sname]
            a_ws = None
            for asn in wb.sheetnames:
                if asn.strip().lower() == gt_sname.strip().lower():
                    a_ws = wb[asn]
                    break
            if a_ws is None:
                record(f"GT sheet '{gt_sname}' exists in agent xlsx", False, f"Available: {wb.sheetnames}")
                continue
            gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
            a_rows = [r for r in a_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
            record(f"GT '{gt_sname}' row count", len(a_rows) == len(gt_rows),
                   f"Expected {len(gt_rows)}, got {len(a_rows)}")
            for ri in range(len(gt_rows)):
                if ri >= len(a_rows):
                    break
                ok = True
                for ci in range(min(len(gt_rows[ri]), len(a_rows[ri]))):
                    gv, av = gt_rows[ri][ci], a_rows[ri][ci]
                    if gv is None:
                        continue
                    if isinstance(gv, (int, float)) and not isinstance(gv, bool):
                        # Tighten tolerance: 1% relative or abs 1.0 floor
                        if not num_close(av, gv, max(abs(gv) * 0.01, 1.0)):
                            ok = False
                            break
                    else:
                        if not str_match(av, gv):
                            ok = False
                            break
                record(f"GT '{gt_sname}' row {ri+1} values", ok,
                       f"gt={gt_rows[ri][:4]}, agent={a_rows[ri][:4] if ri < len(a_rows) else 'missing'}")
        gt_wb.close()


def check_notion():
    print("\n=== Check 5: Notion 'ML Tech Research Hub' page + 'Research Items' database ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT id, properties FROM notion.pages WHERE archived = false")
        pages = cur.fetchall()

        target_title = "ml tech research hub"
        page_id = None
        for pid, props in pages:
            if not isinstance(props, dict):
                continue
            for v in props.values():
                if isinstance(v, dict) and v.get("type") == "title":
                    arr = v.get("title", [])
                    name = "".join(t.get("plain_text", "") or t.get("text", {}).get("content", "") for t in arr).strip().lower()
                    if target_title in name:
                        page_id = pid
                        break
            if page_id:
                break
        record("Notion page 'ML Tech Research Hub' exists", page_id is not None,
               "Not found")

        # Database 'Research Items'
        cur.execute("SELECT id, title, properties FROM notion.databases WHERE archived = false")
        dbs = cur.fetchall()
        target_db = None
        for did, title, dprops in dbs:
            title_text = ""
            if isinstance(title, list):
                for t in title:
                    if isinstance(t, dict):
                        title_text += t.get("plain_text", "") or t.get("text", {}).get("content", "")
            else:
                title_text = str(title or "")
            if title_text.strip().lower() == "research items":
                target_db = (did, dprops)
                break
        record("Notion database 'Research Items' exists",
               target_db is not None,
               "Not found")
        if target_db:
            db_id, dprops = target_db
            # Validate properties contain Type select with options Video/Paper, Status select with New/In Review/Archived
            if isinstance(dprops, dict):
                has_type_select = False
                has_status_select = False
                for k, v in dprops.items():
                    if not isinstance(v, dict):
                        continue
                    kn = (k or "").lower()
                    t = v.get("type")
                    if t == "select" and kn == "type":
                        opts = v.get("select", {}).get("options", [])
                        names = {o.get("name", "").lower() for o in opts if isinstance(o, dict)}
                        if "video" in names and "paper" in names:
                            has_type_select = True
                    if t == "select" and kn == "status":
                        opts = v.get("select", {}).get("options", [])
                        names = {o.get("name", "").lower() for o in opts if isinstance(o, dict)}
                        if "new" in names and "in review" in names and "archived" in names:
                            has_status_select = True
                record("Research Items has 'Type' select with Video/Paper", has_type_select)
                record("Research Items has 'Status' select with New/In Review/Archived", has_status_select)

            # 7+ entries (videos) plus paper entries
            cur.execute("SELECT COUNT(*) FROM notion.pages WHERE archived = false AND parent->>'database_id' = %s", (db_id,))
            n = cur.fetchone()[0]
            record("Research Items database has >= 7 entries", n >= 7, f"Got {n}")
        cur.close()
        conn.close()
    except Exception as e:
        record("Notion check", False, str(e))


def check_email():
    print("\n=== Check 6: Email to research@lab.edu (subject + body) ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT subject, to_addr, COALESCE(body_text, body_html, '') FROM email.messages")
        rows = cur.fetchall()
        cur.close(); conn.close()

        target_to = "research@lab.edu"
        match = None
        for subj, to_addr, body in rows:
            tos = []
            if isinstance(to_addr, list):
                tos = [str(t).lower() for t in to_addr]
            elif isinstance(to_addr, str):
                try:
                    p = json.loads(to_addr)
                    tos = [str(t).lower() for t in p] if isinstance(p, list) else [to_addr.lower()]
                except Exception:
                    tos = [to_addr.lower()]
            if any(target_to in t for t in tos):
                match = (subj, body)
                break

        record("Email to research@lab.edu exists", match is not None, "")
        if match:
            subj_lower = (match[0] or "").strip().lower()
            body_lower = (match[1] or "").lower()
            # Subject must reference ML research catalog: 'ml' AND (catalog or tracker)
            record("Email subject references ML research catalog",
                   "ml" in subj_lower and ("catalog" in subj_lower or "tracker" in subj_lower),
                   f"Subject: {match[0]}")
            # Body must mention number of videos and papers
            record("Email body mentions video count (7)",
                   "7" in body_lower or "seven" in body_lower,
                   "")
            record("Email body mentions ML_Research_Tracker.xlsx file",
                   "ml_research_tracker" in body_lower or "research_tracker" in body_lower or "tracker.xlsx" in body_lower,
                   "")
            record("Email body mentions knowledge base / Notion",
                   "knowledge base" in body_lower or "notion" in body_lower or "research hub" in body_lower,
                   "")
    except Exception as e:
        record("Email check", False, str(e))


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    print(f"Running evaluation for yt-ml-repos-github-notion-excel-email")
    print(f"Agent workspace: {agent_workspace}")

    check_excel(agent_workspace, groundtruth_workspace)
    check_notion()
    check_email()

    all_passed = FAIL_COUNT == 0
    summary = f"Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}"
    print(f"\n{'='*40}")
    print(f"Result: {'PASS' if all_passed else 'FAIL'} - {summary}")

    if res_log_file:
        with open(res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "all_passed": all_passed}, f)

    return all_passed, summary


def main():
    parser = ArgumentParser()
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
