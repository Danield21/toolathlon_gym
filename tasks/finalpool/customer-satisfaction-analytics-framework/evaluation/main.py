#!/usr/bin/env python3
"""Evaluation script for customer-satisfaction-analytics-framework task validation.

Six-phase narrative customer satisfaction analytics task. Validates structural
deliverables that any reasonable agent would produce based on task.md:

  - A satisfaction-analysis xlsx (NPS / CSAT / metric data)
  - An action-plan xlsx (improvement initiatives)
  - A satisfaction report docx with substantive findings
  - An executive-summary docx
  - Outgoing email to stakeholders + calendar review meeting (Phase 6)

The task description is intentionally narrative (no exact filenames, segments
or NPS targets) so verification uses structural + numeric-range overlap
checks against GT rather than strict row-by-row comparison.
"""

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
        # In GT self-test mode, DB-side checks (email/calendar/gsheet/notion)
        # naturally fail because GT files cannot pre-populate DB state.
        if IS_GT_SELF_TEST and db_side:
            WARN_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [WARN] {name} (GT self-test mode, DB-side){msg}")
        else:
            FAIL_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [FAIL] {name}{msg}")


def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"), port=5432,
        dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
        user="eigent", password="camel",
    )


def check_phase6_distribution():
    """Phase 6: distribute findings via email + calendar review meetings."""
    print("\n=== Check: Phase 6 distribution (email/gcal) ===")
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """SELECT id, subject FROM email.messages
               WHERE subject ILIKE %s OR subject ILIKE %s OR subject ILIKE %s
                  OR body_text ILIKE %s OR body_text ILIKE %s""",
            ('%satisfaction%', '%customer experience%', '%findings%',
             '%satisfaction%', '%customer experience%'))
        emails = cur.fetchall()
        cur.execute(
            """SELECT id, summary FROM gcal.events
               WHERE summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s""",
            ('%satisfaction%', '%review%', '%customer%', '%feedback%'))
        events = cur.fetchall()
        conn.close()
    except Exception as e:
        record("Phase 6 distribution (db query)", False, f"db error: {e}", db_side=True)
        return
    record("Phase 6: at least one satisfaction/findings email sent",
           len(emails) >= 1, f"matching emails: {len(emails)}", db_side=True)
    record("Phase 6: at least one review/satisfaction calendar meeting",
           len(events) >= 1, f"matching events: {len(events)}", db_side=True)


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
        # Skip the input survey template
        if fname_low == "survey_template.xlsx":
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


def _collect_numerics(wb):
    """Collect all numeric cell values from a workbook (across sheets)."""
    nums = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, (int, float)) and not isinstance(cell, bool):
                    nums.append(float(cell))
    return nums


def _collect_text_cells(wb):
    """Collect all text cells (lower-cased) from a workbook for header/keyword checks."""
    texts = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is not None and isinstance(cell, str):
                    texts.append(cell.lower())
    return texts


def _structural_xlsx_check(agent_xlsx_path, gt_xlsx_path, label, expected_keywords):
    """Structural validation of an agent xlsx vs GT.

    Replaces strict row-by-row matching (incompatible with the narrative task
    description) with the following checks:
      - File readable and has at least one sheet with substantive rows
      - Total agent row count is within a tolerant range derived from GT
      - At least one content keyword from `expected_keywords` appears in the file
      - Numeric values overlap GT numeric range (catches empty / nonsense files)
    """
    import openpyxl
    fname = os.path.basename(agent_xlsx_path)

    try:
        a_wb = openpyxl.load_workbook(agent_xlsx_path, data_only=True)
    except Exception as e:
        record(f"xlsx {fname} readable", False, str(e))
        return False
    try:
        gt_wb = openpyxl.load_workbook(gt_xlsx_path, data_only=True)
    except Exception as e:
        record(f"GT {label} readable", False, str(e))
        a_wb.close()
        return False

    # Row count: tolerant range based on GT's total non-empty rows
    def total_nonempty_rows(wb):
        total = 0
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                if any(c is not None and str(c).strip() != "" for c in r):
                    total += 1
        return total

    gt_rows = total_nonempty_rows(gt_wb)
    a_rows = total_nonempty_rows(a_wb)
    # Allow agent to produce between ~half and ~triple the GT row count
    lo = max(3, gt_rows // 2)
    hi = max(gt_rows * 3, gt_rows + 10)
    record(f"xlsx {fname} row count in tolerant range [{lo}, {hi}] (GT={gt_rows})",
           lo <= a_rows <= hi, f"agent_rows={a_rows}")

    # Keyword coverage check
    a_texts = _collect_text_cells(a_wb)
    a_blob = " ".join(a_texts)
    matched_kw = [kw for kw in expected_keywords if kw in a_blob]
    record(f"xlsx {fname} content keyword coverage (>=2 of {len(expected_keywords)})",
           len(matched_kw) >= 2,
           f"matched: {matched_kw}")

    # Numeric overlap (catches empty / completely off files)
    gt_nums = _collect_numerics(gt_wb)
    a_nums = _collect_numerics(a_wb)
    if gt_nums:
        gt_min, gt_max = min(gt_nums), max(gt_nums)
        # Allow some buffer below and above the GT range
        buf = max((gt_max - gt_min) * 0.10, 1.0)
        overlap = sum(1 for v in a_nums if (gt_min - buf) <= v <= (gt_max + buf))
        record(f"xlsx {fname} numeric values overlap GT range (>=2 values)",
               overlap >= 2,
               f"overlap_count={overlap}, gt_range=[{gt_min}, {gt_max}]")
    else:
        record(f"xlsx {fname} GT numeric range available (non-blocking)", True,
               "no GT numerics; skipping overlap check")

    gt_wb.close()
    a_wb.close()
    return True


def check_xlsx_content(workspace, groundtruth_workspace="."):
    """Locate two analysis-type xlsx files (satisfaction/analysis + action_plans) by content keywords,
    then perform structural validation against GT. Allows agent-chosen filenames as long as
    content matches."""
    print("\n=== Check: XLSX files ===")
    import openpyxl

    # Each entry: (gt_filename, label, [accept-keywords for locating], [content-keywords for validation])
    targets = [
        ("satisfaction_analysis.xlsx", "satisfaction analysis xlsx",
         ["satisfaction", "analysis", "nps", "metric", "score"],
         ["satisfaction", "nps", "score", "rating", "segment", "category", "rate"]),
        ("action_plans.xlsx", "action plans xlsx",
         ["action", "plan", "initiative", "improvement", "recommendation"],
         ["action", "initiative", "improvement", "owner", "metric", "timeline", "goal"]),
    ]

    used_paths = set()
    for gt_fname, label, locate_keywords, content_keywords in targets:
        gt_path = os.path.join(groundtruth_workspace, gt_fname)
        if not os.path.isfile(gt_path):
            continue

        # 1) Prefer the exact GT filename if present
        exact = os.path.join(workspace, gt_fname)
        chosen = None
        if os.path.isfile(exact) and exact not in used_paths:
            chosen = exact
        else:
            # 2) Fuzzy locate by filename or content keywords (scored)
            cands = _find_xlsx_by_keywords(workspace, locate_keywords, exclude=used_paths)
            if cands:
                chosen = cands[0]

        if chosen is None:
            record(f"xlsx for {label} exists", False,
                   f"No xlsx with keywords {locate_keywords} found in workspace")
            continue
        used_paths.add(chosen)
        record(f"xlsx for {label} exists ({os.path.basename(chosen)})", True)

        # Sanity: at least one sheet has data
        try:
            wb = openpyxl.load_workbook(chosen, data_only=True)
            any_data_sheet = False
            for ws in wb.worksheets:
                rows = list(ws.iter_rows(values_only=True))
                if len(rows) >= 2:
                    any_data_sheet = True
                    break
            wb.close()
            record(f"xlsx {os.path.basename(chosen)} has at least one populated sheet",
                   any_data_sheet, "no sheet has >=2 rows")
        except Exception as e:
            record(f"xlsx {os.path.basename(chosen)} readable", False, str(e))
            continue

        # Structural comparison vs GT (replaces strict row-by-row matching)
        _structural_xlsx_check(chosen, gt_path, label, content_keywords)


def _find_docx_by_keywords(workspace, keywords, exclude=None):
    """Find docx files in workspace whose filename OR content matches any keyword."""
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
        # Content match
        try:
            doc = Document(path)
            text_low = "\n".join(p.text for p in doc.paragraphs).lower()
            if any(kw in text_low for kw in keywords):
                candidates.append(path)
        except Exception:
            continue
    return candidates


def check_docx_content(workspace):
    """Locate two report-type docx files (satisfaction report + executive summary)
    by filename or content keywords, then verify each is substantive."""
    print("\n=== Check: DOCX files ===")
    from docx import Document

    targets = [
        ("satisfaction_report.docx", "satisfaction report docx",
         ["satisfaction", "report", "finding", "customer", "nps"]),
        ("executive_summary.docx", "executive summary docx",
         ["executive", "summary", "strategic", "stakeholder", "leadership"]),
    ]

    used_paths = set()
    for gt_fname, label, keywords in targets:
        # 1) Prefer exact GT filename
        exact = os.path.join(workspace, gt_fname)
        chosen = None
        if os.path.isfile(exact) and exact not in used_paths:
            chosen = exact
        else:
            cands = _find_docx_by_keywords(workspace, keywords, exclude=used_paths)
            if cands:
                chosen = cands[0]

        if chosen is None:
            record(f"docx for {label} exists", False,
                   f"No docx with keywords {keywords} found in workspace")
            continue
        used_paths.add(chosen)
        record(f"docx for {label} exists ({os.path.basename(chosen)})", True)

        try:
            doc = Document(chosen)
            text = "\n".join(p.text for p in doc.paragraphs)
            text_low = text.lower()
            word_count = len([w for w in text.split() if w.strip()])
            record(f"docx {os.path.basename(chosen)} substantive (>=150 words)",
                   word_count >= 150, f"{word_count} words")
            content_keywords = ['satisfaction', 'nps', 'recommend', 'customer',
                                'survey', 'review', 'analysis', 'finding']
            matches = sum(1 for kw in content_keywords if kw in text_low)
            record(f"docx {os.path.basename(chosen)} mentions satisfaction/nps/recommend (>=2 keywords)",
                   matches >= 2, f"matched {matches}/8")
        except Exception as e:
            record(f"docx {os.path.basename(chosen)} readable", False, str(e))


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

    # Detect GT self-test mode (agent_ws == groundtruth_ws)
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
    check_phase6_distribution()

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
