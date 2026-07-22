#!/usr/bin/env python3
"""Evaluation script for research-lab-publication-pipeline.

Aligned with task.md 6 phases:
  - Literature search results (spreadsheet/csv with bibliographic info)
  - Bibliography file (.bib or equivalent)
  - Research roadmap document (Word)
  - Email to lab members
  - Calendar meeting
"""

from argparse import ArgumentParser
import json
import os
import sys
from pathlib import Path
import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
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
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def _read_text_any(path):
    p = Path(path)
    if not p.exists():
        return ""
    try:
        ext = p.suffix.lower()
        if ext == ".docx":
            from docx import Document
            return "\n".join(par.text for par in Document(str(p)).paragraphs)
        if ext == ".xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(str(p), data_only=True)
            txts = []
            for sn in wb.sheetnames:
                ws = wb[sn]
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell is not None:
                            txts.append(str(cell))
            return " ".join(txts)
        if ext in (".txt", ".md", ".csv", ".json", ".tsv", ".bib"):
            return p.read_text(errors="ignore")
    except Exception:
        return ""
    return ""


def check_literature_db(workspace):
    """Verify a consolidated literature database / spreadsheet exists with bibliographic entries."""
    base = Path(workspace)
    candidates = []
    for pat in ["*literature*", "*publications*", "*bibliography*", "*reference*",
                "*research_db*", "*paper*", "*citation*", "*submission_checklist*"]:
        for ext in ["xlsx", "csv", "json", "tsv", "md"]:
            candidates.extend(base.rglob(f"{pat}.{ext}"))
    record("Literature/publications database file exists",
           len(candidates) >= 1, f"found {len(candidates)} candidates")
    if not candidates:
        return
    text = "\n".join(_read_text_any(str(c)) for c in candidates).lower()
    record("Literature DB substantive content (>=300 chars)",
           len(text) >= 300, f"length={len(text)}")
    # Bibliographic field tokens
    bib_tokens = ["author", "year", "title"]
    missing = [t for t in bib_tokens if t not in text]
    record("Literature DB has bibliographic fields (author/year/title)",
           len(missing) <= 1, f"missing: {missing}")
    # Tightened: count unique paper rows in literature/publications/reference files specifically
    # (exclude submission_checklist which is a different artifact).
    paper_db_patterns = ["literature", "publications", "bibliography", "reference",
                          "research_db", "paper", "citation"]
    unique_papers = 0
    for c in candidates:
        # Only count files whose basename matches paper-database-style names
        name_l = c.stem.lower()
        if not any(p in name_l for p in paper_db_patterns):
            continue
        ext = c.suffix.lower()
        try:
            if ext == ".xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(str(c), data_only=True)
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                            if r and any(v not in (None, '') for v in r)]
                    unique_papers = max(unique_papers, len(rows))
                wb.close()
            elif ext == ".csv":
                with open(c, encoding="utf-8", errors="ignore") as fh:
                    lines = [ln for ln in fh.read().splitlines() if ln.strip()]
                unique_papers = max(unique_papers, max(0, len(lines) - 1))
            elif ext == ".json":
                with open(c, encoding="utf-8", errors="ignore") as fh:
                    try:
                        data = json.load(fh)
                        if isinstance(data, list):
                            unique_papers = max(unique_papers, len(data))
                        elif isinstance(data, dict):
                            for v in data.values():
                                if isinstance(v, list):
                                    unique_papers = max(unique_papers, len(v))
                    except Exception:
                        pass
            elif ext == ".tsv":
                with open(c, encoding="utf-8", errors="ignore") as fh:
                    lines = [ln for ln in fh.read().splitlines() if ln.strip()]
                unique_papers = max(unique_papers, max(0, len(lines) - 1))
        except Exception:
            pass
    record("Literature DB has >=10 unique paper rows",
           unique_papers >= 10, f"unique_papers={unique_papers}")


def check_bibliography(workspace):
    """Verify a .bib file or comparable references file exists."""
    base = Path(workspace)
    bib_files = list(base.rglob("*.bib")) + list(base.rglob("*references*"))
    record("Bibliography file (.bib or *references*) exists",
           len(bib_files) >= 1, f"found {len(bib_files)}")
    if bib_files:
        text = "\n".join(_read_text_any(str(b)) for b in bib_files).lower()
        entries = text.count("@article") + text.count("@book") + text.count("@inproceedings") + text.count("@misc") + text.count("@conference") + text.count("@incollection")
        record("Bibliography contains at least 5 entries",
               entries >= 5,
               f"@-entries: {entries}")


def check_roadmap(workspace):
    """Verify a research roadmap / synthesis document exists with required content."""
    base = Path(workspace)
    candidates = []
    for pat in ["*roadmap*", "*synthesis*", "*research_plan*", "*future_directions*", "*priorities*"]:
        for ext in ["docx", "pdf", "md", "txt"]:
            candidates.extend(base.rglob(f"{pat}.{ext}"))
    record("Research roadmap document exists",
           len(candidates) >= 1, f"found {len(candidates)}")
    if not candidates:
        return
    text = "\n".join(_read_text_any(str(c)) for c in candidates).lower()
    record("Roadmap doc substantive content (>=400 chars)",
           len(text) >= 400, f"length={len(text)}")
    # Roadmap must mention research directions/gaps/recommendations
    must = ["direction", "gap", "recommendation"]
    found_any = sum(1 for m in must if m in text)
    record("Roadmap covers directions/gaps/recommendations (>=2 topics)",
           found_any >= 2, f"found: {found_any}/3 topics")
    # Tightened: must include explicit 'Future Directions' or 'Recommendations' heading
    has_future = "future direction" in text or "future research" in text
    has_recs = "recommendation" in text
    record("Roadmap has explicit 'Future Directions' or 'Recommendations' heading",
           has_future or has_recs, f"has_future={has_future}, has_recs={has_recs}")


def check_emails():
    try:
        conn = psycopg2.connect(**DB); cur = conn.cursor()
        cur.execute("SELECT subject, to_addr, COALESCE(body_text, body_html, '') FROM email.messages")
        rows = cur.fetchall(); cur.close(); conn.close()
    except Exception as e:
        record("Email DB reachable", False, str(e)); return
    matched = None
    for subj, to_addr, body in rows:
        text = ((subj or "") + " " + (body or "")).lower()
        # Tightened: also require recipient looks like a lab/group address (contains 'lab', 'group',
        # 'team', 'research' or at least an email '@')
        to_str = str(to_addr or "").lower()
        recipient_ok = "@" in to_str and any(k in to_str for k in ["lab", "group", "team", "research", "members"])
        # If at least the topic matches and the recipient is plausibly a group/lab address, accept
        if (any(k in text for k in ["roadmap", "literature review", "publication pipeline",
                                    "research priorities", "lab member"])
                and recipient_ok):
            matched = (subj, to_addr); break
    record("Email to lab members (group address) about roadmap/literature exists",
           matched is not None, f"checked {len(rows)} emails")


def check_calendar():
    try:
        conn = psycopg2.connect(**DB); cur = conn.cursor()
        cur.execute("SELECT id, summary, COALESCE(description, '') FROM gcal.events")
        rows = cur.fetchall(); cur.close(); conn.close()
    except Exception as e:
        record("Calendar DB reachable", False, str(e)); return
    matched = None
    for _id, summary, desc in rows:
        text = ((summary or "") + " " + (desc or "")).lower()
        if any(k in text for k in ["roadmap", "research priorities", "lab meeting",
                                   "literature review", "publication pipeline"]):
            matched = (summary, desc); break
    record("Calendar meeting scheduled to discuss priorities/literature",
           matched is not None, f"checked {len(rows)} events")


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    global PASS_COUNT, FAIL_COUNT
    PASS_COUNT = 0; FAIL_COUNT = 0
    if not agent_workspace or not Path(agent_workspace).exists():
        return False, f"Agent workspace not found: {agent_workspace}"
    # GT self-test toleration: when agent==groundtruth (V1 GT-only smoke test),
    # the email/gcal checks fail because no email/event was actually sent.
    is_gt_self_test = (
        agent_workspace and groundtruth_workspace
        and os.path.abspath(agent_workspace) == os.path.abspath(groundtruth_workspace)
    )
    check_literature_db(agent_workspace)
    check_bibliography(agent_workspace)
    check_roadmap(agent_workspace)
    if not is_gt_self_test:
        check_emails()
        check_calendar()
    else:
        print("  [SKIP] email and calendar checks (GT self-test mode)")
    return FAIL_COUNT == 0, f"Passed {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} checks"


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    success, message = run_evaluation(
        args.agent_workspace, args.groundtruth_workspace, args.launch_time, args.res_log_file
    )
    print(message)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
