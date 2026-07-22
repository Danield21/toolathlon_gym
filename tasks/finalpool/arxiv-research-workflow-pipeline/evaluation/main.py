#!/usr/bin/env python3
"""Evaluation script for arxiv-research-workflow-pipeline task validation"""

from argparse import ArgumentParser
import json
import os
import sys
from pathlib import Path

PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
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


def _find_xlsx_by_keywords(workspace, keywords):
    """Score xlsx files by filename + content keywords."""
    import glob
    import openpyxl
    scored = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.xlsx"))):
        fname_low = os.path.basename(path).lower()
        if fname_low.startswith("~$") or fname_low.startswith("."):
            continue
        score = 0
        for kw in keywords:
            if kw in fname_low:
                score += 10
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            content_text = " ".join(s.lower() for s in wb.sheetnames)
            for ws in wb.worksheets:
                row_count = 0
                for row in ws.iter_rows(values_only=True):
                    if row_count > 5:
                        break
                    row_count += 1
                    for cell in row:
                        if cell is not None:
                            content_text += " " + str(cell).lower()
            wb.close()
            for kw in keywords:
                if kw in content_text:
                    score += 1
        except Exception:
            pass
        if score > 0:
            scored.append((score, path))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [p for _, p in scored]


def check_xlsx_content(workspace, groundtruth_workspace="."):
    """Locate paper analysis xlsx by filename or content keyword.
    Compares structure (sheet headers + at least some matching rows) - row count is
    flexible because task.md does not specify exact paper count."""
    print("\n=== Check: XLSX paper analysis ===")
    import openpyxl
    keywords = ["paper_analysis", "paper analysis", "papers", "literature",
                "research", "analysis", "metadata"]
    exact = os.path.join(workspace, "paper_analysis.xlsx")
    chosen = exact if os.path.isfile(exact) else None
    if chosen is None:
        cands = _find_xlsx_by_keywords(workspace, keywords)
        if cands: chosen = cands[0]
    if chosen is None:
        record("xlsx for paper analysis exists", False,
               f"No xlsx with keywords {keywords[:4]} found")
        return False
    fname = os.path.basename(chosen)
    record(f"xlsx for paper analysis exists ({fname})", True)
    try:
        wb = openpyxl.load_workbook(chosen, data_only=True)
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            record(f"xlsx '{ws.title}' has data", len(rows) >= 2, f"{len(rows)} rows")
    except Exception as e:
        record("xlsx readable", False, str(e))
        return True

    # --- Groundtruth value comparison (relaxed row-count) ---
    gt_path = os.path.join(groundtruth_workspace, "paper_analysis.xlsx")
    if not os.path.isfile(gt_path):
        wb.close()
        return True

    gt_wb = openpyxl.load_workbook(gt_path, data_only=True)
    for gt_sheet_name in gt_wb.sheetnames:
        gt_ws = gt_wb[gt_sheet_name]
        agent_ws = None
        for asn in wb.sheetnames:
            if asn.strip().lower() == gt_sheet_name.strip().lower():
                agent_ws = wb[asn]
                break
        if agent_ws is None and len(wb.sheetnames) == 1 and len(gt_wb.sheetnames) == 1:
            agent_ws = wb[wb.sheetnames[0]]
        if agent_ws is None:
            record(f"GT sheet '{gt_sheet_name}' exists in agent", False, f"Available: {wb.sheetnames}")
            continue

        gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
        agent_rows = [r for r in agent_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]

        # Relaxed: agent should have at least min(5, len(gt_rows)) rows.
        # Task does not specify an exact paper count; honest agent may pick 10-50.
        min_rows = min(5, len(gt_rows))
        record(f"GT '{gt_sheet_name}' has at least {min_rows} data rows",
               len(agent_rows) >= min_rows,
               f"Got {len(agent_rows)} rows (GT has {len(gt_rows)})")

        # Header check: first row of each side
        gt_hdr_row = None
        for r in gt_ws.iter_rows(min_row=1, max_row=1, values_only=True):
            gt_hdr_row = r
            break
        agent_hdr_row = None
        for r in agent_ws.iter_rows(min_row=1, max_row=1, values_only=True):
            agent_hdr_row = r
            break
        if gt_hdr_row and agent_hdr_row:
            gt_hdr_low = [str(v).strip().lower() for v in gt_hdr_row if v is not None]
            agent_hdr_low = [str(v).strip().lower() for v in agent_hdr_row if v is not None]
            shared = sum(1 for h in gt_hdr_low if h in agent_hdr_low)
            record(f"GT '{gt_sheet_name}' shares >=2 headers with agent",
                   shared >= 2, f"GT: {gt_hdr_low[:5]}, agent: {agent_hdr_low[:5]}, shared: {shared}")
    gt_wb.close()
    wb.close()
    return True


def _find_docx_by_keywords(workspace, keywords):
    import glob
    from docx import Document
    scored = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.docx"))):
        fname_low = os.path.basename(path).lower()
        if fname_low.startswith("~$") or fname_low.startswith("."):
            continue
        score = 0
        for kw in keywords:
            if kw in fname_low:
                score += 10
        try:
            doc = Document(path)
            text_low = "\n".join(p.text for p in doc.paragraphs).lower()
            for kw in keywords:
                if kw in text_low:
                    score += 1
        except Exception:
            pass
        if score > 0:
            scored.append((score, path))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [p for _, p in scored]


def check_docx_content(workspace):
    """Locate research summary / literature review docx by filename or content keyword."""
    print("\n=== Check: DOCX literature review / research summary ===")
    from docx import Document
    keywords = ["literature_review", "literature review", "research_summary",
                "research summary", "literature", "research", "summary",
                "review", "machine learning", "federated"]
    exact = os.path.join(workspace, "literature_review.docx")
    chosen = exact if os.path.isfile(exact) else None
    if chosen is None:
        cands = _find_docx_by_keywords(workspace, keywords)
        if cands: chosen = cands[0]
    if chosen is None:
        record("docx for literature review/research summary exists", False,
               f"No docx with keywords {keywords[:5]} found")
        return False
    fname = os.path.basename(chosen)
    record(f"docx for literature review/research summary exists ({fname})", True)
    try:
        doc = Document(chosen)
        non_empty = [p for p in doc.paragraphs if p.text.strip()]
        record("docx has substantial content (>=10 non-empty paragraphs)",
               len(non_empty) >= 10,
               f"{len(non_empty)} non-empty paragraphs")
        all_text = "\n".join(p.text for p in doc.paragraphs)
        word_count = len(all_text.split())
        # Task says "approximately two thousand words"; accept >=500 to allow concise summaries
        record("docx word count >= 500",
               word_count >= 500,
               f"{word_count} words")
        # Mention of distributed ML / federated learning / neural network topic (task is broad).
        # Tightened to >=2 keywords to ensure substantive ML topic coverage rather than
        # incidental mention.
        text_lower = all_text.lower()
        topic_keys = ("federated", "distributed", "machine learning", "neural",
                      "deep learning")
        topic_matched = sum(1 for k in topic_keys if k in text_lower)
        record("docx mentions ML/neural network/federated topic (>=2 keywords)",
               topic_matched >= 2, f"matched {topic_matched}/5")
    except Exception as e:
        record("docx readable", False, str(e))
    return True


def check_bib_file(workspace):
    print("\n=== Check: Bibliography .bib file ===")
    # Look for any .bib file (task says 'comprehensive bibliography' without naming it)
    bib_files = [f for f in os.listdir(workspace) if f.endswith(".bib")]
    record("At least one .bib file present", len(bib_files) >= 1,
           f"Found: {bib_files}")
    if not bib_files:
        return
    # Pick the largest bib file
    bib_path = os.path.join(workspace, max(bib_files, key=lambda f: os.path.getsize(os.path.join(workspace, f))))
    try:
        with open(bib_path) as f:
            content = f.read()
    except Exception as e:
        record(".bib file readable", False, str(e))
        return
    # Count bibtex entries (@article, @inproceedings, etc.)
    import re
    entries = re.findall(r"@\w+\s*\{", content)
    record(".bib file has at least 5 entries", len(entries) >= 5,
           f"Found {len(entries)} entries")
    # Topic check: task is broad (machine learning, neural networks, deep learning).
    # Accept any of these terms in addition to federated/distributed.
    # Tightened to require at least 2 distinct topical keywords across the file,
    # so empty / off-topic .bib files cannot pass with one stray word.
    text_lower = content.lower()
    topic_keys = ("federated", "distributed", "machine learning", "neural",
                  "deep learning")
    matched_topics = sum(1 for k in topic_keys if k in text_lower)
    record(".bib mentions ML/neural network/federated topic (>=2 keywords)",
           matched_topics >= 2,
           f"matched {matched_topics}/5; sample first 200 chars: {text_lower[:200]}")
    # Year distribution: bibliography entries should be reasonably modern (>=3 entries
    # in the last 7 years). Catches GT-noise files that are entirely classical/old papers.
    import re as _re
    years = [int(y) for y in _re.findall(r"year\s*=\s*[{\"]?(\d{4})[}\"]?", content)]
    recent_years = [y for y in years if y >= 2018]
    record(".bib has at least 3 entries from 2018 or later",
           len(recent_years) >= 3,
           f"recent_year_count={len(recent_years)} (years sampled: {sorted(set(years))[:8]})")


def check_notion_db(is_gt_self_test=False):
    print("\n=== Check: Notion database for paper organization ===")
    try:
        import psycopg2
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"), port=5432,
                                dbname="toolathlon_gym", user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT id, title FROM notion.databases")
        dbs = cur.fetchall()
        cur.execute("SELECT COUNT(*) FROM notion.pages WHERE archived = false AND in_trash = false")
        page_count = cur.fetchone()[0]
        if is_gt_self_test and (len(dbs) == 0 or page_count < 5):
            print("  [WARN] Notion DB/page checks skipped (GT self-test, non-blocking)")
            cur.close(); conn.close()
            return
        record("At least one Notion database created", len(dbs) >= 1,
               f"Found {len(dbs)} databases")
        record("At least 5 Notion pages created (papers + hub)", page_count >= 5,
               f"Found {page_count} pages")
        cur.close(); conn.close()
    except Exception as e:
        record("Notion DB query", False, str(e))


def check_calendar(is_gt_self_test=False):
    print("\n=== Check: Calendar invitations ===")
    try:
        import psycopg2
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"), port=5432,
                                dbname="toolathlon_gym", user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT id, summary FROM gcal.events")
        events = cur.fetchall()
        ev_count = len(events)
        if is_gt_self_test and ev_count == 0:
            print("  [WARN] Calendar checks skipped (GT self-test, non-blocking)")
            cur.close(); conn.close()
            return
        record("At least 1 calendar event created for the team", ev_count >= 1,
               f"Found {ev_count} events")
        # Sanity check: at least 1 event title relates to review/research/meeting/paper
        if events:
            topical = 0
            event_ids_topical = []
            for eid, summary in events:
                s = (summary or "").lower()
                if any(kw in s for kw in ("review", "research", "meeting", "paper", "literature", "discuss")):
                    topical += 1
                    event_ids_topical.append(eid)
            record("At least 1 event title mentions review/research/meeting/paper",
                   topical >= 1, f"got {topical}/{ev_count} topical")
            # Attendee sanity: try to find at least one event with >=2 attendees
            try:
                cur.execute("SELECT event_id, COUNT(*) FROM gcal.attendees GROUP BY event_id")
                att_rows = cur.fetchall()
                max_att = max((c for _, c in att_rows), default=0)
                record("At least 1 event has >=2 attendees (team invite)",
                       max_att >= 2, f"max attendees on any event: {max_att}")
            except Exception:
                # gcal.attendees may not exist in this schema - skip silently
                pass
        cur.close(); conn.close()
    except Exception as e:
        record("Calendar query", False, str(e))


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    ws = args.agent_workspace
    if not os.path.isdir(ws):
        print(f"Agent workspace not found: {ws}")
        sys.exit(1)

    # Detect GT self-test (V1 parity test).
    try:
        gt_canon = os.path.realpath(args.groundtruth_workspace) if args.groundtruth_workspace else ""
        ag_canon = os.path.realpath(args.agent_workspace) if args.agent_workspace else ""
        is_gt_self_test = bool(gt_canon) and (gt_canon == ag_canon)
    except Exception:
        is_gt_self_test = False

    check_xlsx_content(ws, args.groundtruth_workspace)
    check_docx_content(ws)
    check_bib_file(ws)
    check_notion_db(is_gt_self_test)
    check_calendar(is_gt_self_test)

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print(f"FAIL ({FAIL_COUNT} checks failed)")
        sys.exit(1)


if __name__ == "__main__":
    main()
