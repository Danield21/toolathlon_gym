#!/usr/bin/env python3
"""Evaluation script for budget-variance-analysis-quarterly task validation"""

from argparse import ArgumentParser
import json
import os
import sys

PASS_COUNT = 0
FAIL_COUNT = 0
WARN_COUNT = 0
IS_GT_SELF_TEST = False


def record(name, passed, detail="", db_side=False):
    global PASS_COUNT, FAIL_COUNT, WARN_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        if IS_GT_SELF_TEST and db_side:
            WARN_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [WARN] {name} (GT self-test mode, DB-side){msg}")
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


def _find_xlsx_by_keywords(workspace, keywords, exclude=None):
    """Score xlsx files by keyword matches; filename match weighs 10x more than content."""
    import glob
    import openpyxl
    exclude = exclude or set()
    scored = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.xlsx"))):
        if path in exclude:
            continue
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


def _compare_xlsx_to_gt(agent_xlsx_path, gt_xlsx_path, label):
    """Compare agent xlsx against GT xlsx. Records per-row pass/fail."""
    import openpyxl
    fname = os.path.basename(agent_xlsx_path)
    try:
        wb = openpyxl.load_workbook(agent_xlsx_path, data_only=True)
    except Exception as e:
        record(f"xlsx {fname} readable", False, str(e))
        return False
    try:
        gt_wb = openpyxl.load_workbook(gt_xlsx_path, data_only=True)
    except Exception as e:
        record(f"GT {label} readable", False, str(e))
        wb.close()
        return False

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
            record(f"GT {label} sheet '{gt_sheet_name}' present in {fname}", False,
                   f"Available: {wb.sheetnames}")
            continue

        gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
        agent_rows = [r for r in agent_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]

        # Tolerant row count: allow ±20% deviation (min ±2) to accommodate
        # legitimate variance from fewer/more cost centers reported.
        row_tol = max(2, int(len(gt_rows) * 0.20))
        row_diff = abs(len(agent_rows) - len(gt_rows))
        record(f"GT {label} '{gt_sheet_name}' row count (within tolerance)",
               row_diff <= row_tol and len(agent_rows) >= max(2, len(gt_rows) - row_tol),
               f"Expected ~{len(gt_rows)} (±{row_tol}), got {len(agent_rows)}")

        check_indices = list(range(min(3, len(gt_rows))))
        if len(gt_rows) > 3:
            check_indices.append(len(gt_rows) - 1)
        for idx in check_indices:
            gt_row = gt_rows[idx]
            if idx < len(agent_rows):
                a_row = agent_rows[idx]
                row_ok = True
                for col_idx in range(min(len(gt_row), len(a_row) if a_row else 0)):
                    gt_val = gt_row[col_idx]
                    a_val = a_row[col_idx]
                    if gt_val is None:
                        continue
                    if isinstance(gt_val, (int, float)):
                        ok = num_close(a_val, gt_val, max(abs(gt_val) * 0.1, 1.0))
                    else:
                        ok = str_match(a_val, gt_val)
                    if not ok:
                        record(f"GT {label} '{gt_sheet_name}' row {idx+1} col {col_idx+1}",
                               False, f"Expected {gt_val}, got {a_val}")
                        row_ok = False
                        break
                if row_ok:
                    record(f"GT {label} '{gt_sheet_name}' row {idx+1} values match", True)
            else:
                record(f"GT {label} '{gt_sheet_name}' row {idx+1} exists", False, "Row missing in agent")
    gt_wb.close()
    wb.close()
    return True


def check_xlsx_content(workspace, groundtruth_workspace="."):
    """Locate three xlsx deliverables (variance analysis, variance tracking, budget forecast)
    by filename or content keyword and compare each to its GT file."""
    print("\n=== Check: XLSX files ===")
    import openpyxl

    targets = [
        ("variance_analysis.xlsx", "variance analysis xlsx",
         ["variance_analysis", "variance-analysis", "variance analysis",
          "variance", "analysis"]),
        ("variance_tracking.xlsx", "variance tracking xlsx",
         ["variance_tracking", "variance-tracking", "variance tracking",
          "tracking", "monitor", "schedule"]),
        ("budget_forecast.xlsx", "budget forecast xlsx",
         ["budget_forecast", "budget-forecast", "budget forecast",
          "forecast", "projection", "revised", "outlook"]),
    ]

    used_paths = set()
    for gt_fname, label, keywords in targets:
        gt_path = os.path.join(groundtruth_workspace, gt_fname)
        if not os.path.isfile(gt_path):
            continue

        exact = os.path.join(workspace, gt_fname)
        chosen = None
        if os.path.isfile(exact) and exact not in used_paths:
            chosen = exact
        else:
            cands = [p for p in _find_xlsx_by_keywords(workspace, keywords, exclude=used_paths)]
            if cands:
                chosen = cands[0]

        if chosen is None:
            record(f"xlsx for {label} exists", False,
                   f"No xlsx with keywords {keywords[:3]} found in workspace")
            continue
        used_paths.add(chosen)
        record(f"xlsx for {label} exists ({os.path.basename(chosen)})", True)

        try:
            wb = openpyxl.load_workbook(chosen, data_only=True)
            for ws in wb.worksheets:
                rows = list(ws.iter_rows(values_only=True))
                record(f"xlsx {os.path.basename(chosen)} '{ws.title}' has data",
                       len(rows) >= 2, f"{len(rows)} rows")
            wb.close()
        except Exception as e:
            record(f"xlsx {os.path.basename(chosen)} readable", False, str(e))
            continue

        _compare_xlsx_to_gt(chosen, gt_path, label)


def _find_docx_by_keywords(workspace, keywords, exclude=None):
    import glob
    from docx import Document
    exclude = exclude or set()
    candidates = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.docx"))):
        if path in exclude:
            continue
        fname_low = os.path.basename(path).lower()
        if fname_low.startswith("~$") or fname_low.startswith("."):
            continue
        if any(kw in fname_low for kw in keywords):
            candidates.append(path)
            continue
        try:
            doc = Document(path)
            text_low = "\n".join(p.text for p in doc.paragraphs).lower()
            if any(kw in text_low for kw in keywords):
                candidates.append(path)
        except Exception:
            continue
    return candidates


def check_docx_content(workspace):
    """Locate department variance report docx by filename or content keywords."""
    print("\n=== Check: DOCX files ===")
    from docx import Document
    keywords = ["variance", "department", "budget", "report", "dept"]
    exact = os.path.join(workspace, "dept_variance_reports.docx")
    if os.path.isfile(exact):
        chosen = exact
    else:
        cands = _find_docx_by_keywords(workspace, keywords)
        chosen = cands[0] if cands else None
    if chosen is None:
        record("docx for department variance reports exists", False,
               f"No docx with keywords {keywords} found in workspace")
        return False
    record(f"docx for department variance reports exists ({os.path.basename(chosen)})", True)
    try:
        doc = Document(chosen)
        text = "\n".join(p.text for p in doc.paragraphs)
        word_count = len([w for w in text.split() if w.strip()])
        record(f"docx {os.path.basename(chosen)} substantive (>=150 words)",
               word_count >= 150, f"{word_count} words")
        text_low = text.lower()
        topic_keys = ["variance", "budget", "department", "forecast", "spending"]
        matched = sum(1 for k in topic_keys if k in text_low)
        record(f"docx {os.path.basename(chosen)} mentions topical keywords (>=2)",
               matched >= 2, f"matched {matched}/5")
    except Exception as e:
        record(f"docx {os.path.basename(chosen)} readable", False, str(e))
    return True


def _find_pptx_by_keywords(workspace, keywords):
    import glob
    from pptx import Presentation
    candidates = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.pptx"))):
        fname_low = os.path.basename(path).lower()
        if fname_low.startswith("~$") or fname_low.startswith("."):
            continue
        if any(kw in fname_low for kw in keywords):
            candidates.append(path)
            continue
        try:
            prs = Presentation(path)
            text_low = ""
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text_low += " " + (shape.text or "").lower()
            if any(kw in text_low for kw in keywords):
                candidates.append(path)
        except Exception:
            continue
    return candidates


def check_pptx_content(workspace):
    """Locate executive presentation pptx by filename or content keywords."""
    print("\n=== Check: PPTX executive presentation ===")
    from pptx import Presentation
    keywords = ["executive", "presentation", "variance", "budget", "leadership"]
    exact = os.path.join(workspace, "executive_presentation.pptx")
    if os.path.isfile(exact):
        chosen = exact
    else:
        cands = _find_pptx_by_keywords(workspace, keywords)
        chosen = cands[0] if cands else None
    if chosen is None:
        record("pptx for executive presentation exists", False,
               f"No pptx with keywords {keywords} found in workspace")
        return False
    record(f"pptx for executive presentation exists ({os.path.basename(chosen)})", True)
    try:
        prs = Presentation(chosen)
        record(f"pptx {os.path.basename(chosen)} has >=3 slides",
               len(prs.slides) >= 3, f"{len(prs.slides)} slides")
        all_text = ""
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    all_text += " " + shape.text
        record(f"pptx {os.path.basename(chosen)} mentions 'variance'",
               "variance" in all_text.lower(),
               f"sample: {all_text[:200]}")
        record(f"pptx {os.path.basename(chosen)} mentions 'budget' or 'forecast'",
               "budget" in all_text.lower() or "forecast" in all_text.lower(),
               f"sample: {all_text[:200]}")
    except Exception as e:
        record(f"pptx {os.path.basename(chosen)} readable", False, str(e))
    return True


def check_gcal_meetings():
    """Phase 6 requires scheduling budget review meetings."""
    print("\n=== Check: GCal budget review meetings ===")
    try:
        import psycopg2
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"), port=5432,
                                dbname="toolathlon_gym", user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT summary FROM gcal.events")
        events = cur.fetchall()
        cur.close(); conn.close()
        summaries = [str(e[0]).lower() for e in events if e[0]]
        record("gcal has >=1 event scheduled (budget review meeting)",
               len(events) >= 1, f"Total events: {len(events)}", db_side=True)
        record("gcal event mentions 'budget' or 'variance' or 'review'",
               any(k in s for s in summaries for k in ["budget", "variance", "review", "forecast"]),
               f"Summaries: {summaries[:5]}", db_side=True)
    except Exception as e:
        record("gcal check ran", False, str(e), db_side=True)


def check_email_to_management():
    """Phase 6 requires sending detailed reports to senior management."""
    print("\n=== Check: Email to senior management ===")
    try:
        import psycopg2
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"), port=5432,
                                dbname="toolathlon_gym", user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
        messages = cur.fetchall()
        cur.close(); conn.close()
        record("email: >=1 message sent (variance report distribution)",
               len(messages) >= 1, f"Total messages: {len(messages)}", db_side=True)
        # Subject/body topic match
        topical = []
        for subj, to, body in messages:
            text = ((subj or "") + " " + (body or "")).lower()
            if any(k in text for k in ["variance", "budget", "forecast", "quarterly"]):
                topical.append(subj)
        record("email: at least 1 topical message (variance/budget/forecast)",
               len(topical) >= 1, f"Sample subjects: {topical[:3]}", db_side=True)
    except Exception as e:
        record("email check ran", False, str(e), db_side=True)


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

    # Detect GT self-test mode
    global IS_GT_SELF_TEST
    try:
        if args.groundtruth_workspace and os.path.exists(args.groundtruth_workspace):
            IS_GT_SELF_TEST = (
                os.path.realpath(ws) ==
                os.path.realpath(args.groundtruth_workspace)
            )
    except Exception:
        IS_GT_SELF_TEST = False

    check_xlsx_content(ws, args.groundtruth_workspace)
    check_docx_content(ws)
    check_pptx_content(ws)
    check_gcal_meetings()
    check_email_to_management()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    print(f"\nOverall: {PASS_COUNT}/{total} checks passed")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "failed": FAIL_COUNT,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print(f"FAIL ({FAIL_COUNT} failures)")
        sys.exit(1)


if __name__ == "__main__":
    main()
