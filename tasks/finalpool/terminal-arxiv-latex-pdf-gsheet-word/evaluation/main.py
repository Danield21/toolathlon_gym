"""
Evaluation for terminal-arxiv-latex-pdf-gsheet-word task.

Checks:
1. Google Sheet "Paper Review Matrix" with 3 sheets (Review Scores, Methodology Comparison, Rankings)
2. Conference_Review_Summary.docx
3. Intermediate JSON files (methodology_analysis.json, comparison_matrix.json, final_rankings.json)
"""
import argparse
import json
import os
import re
import sys

import psycopg2

# All DB connection params are read from environment (harness-injected) with
# local defaults. Must stay symmetric with preprocess/main.py so both target
# the same worker database.
DB = dict(host=os.environ.get("PGHOST", "localhost"),
          port=int(os.environ.get("PGPORT", "5432")),
          dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
          user=os.environ.get("PGUSER", "eigent"),
          password=os.environ.get("PGPASSWORD", "camel"))

PASS_COUNT = 0
FAIL_COUNT = 0
LOCAL_FAIL_COUNT = 0
CURRENT_CATEGORY = "runtime"
# Required deliverables that are genuinely absent (not just unverifiable).
# A correct model always produces all of them, so a missing one is a hard fail.
HARD_MISSING = 0

TARGET_IDS = {"2401.00101", "2401.00102", "2401.00103", "2401.00104"}
NOISE_IDS = {"2401.00201", "2401.00202"}

# Contexts that indicate a noise paper was intentionally excluded (rather than
# reviewed). A faithful model may legitimately mention the excluded papers by
# id (e.g. in the Overview) to say they are out of scope; only a mention that
# is NOT accompanied by any exclusion context should be treated as including
# the noise paper in the review.
EXCLUDE_CONTEXT = (
    "exclud", "ignor", "omit", "discard", "filter", "noise", "unrelat",
    "unassigned", "unreviewed", "out of scope", "outside", "not review",
    "not evaluat", "not assess", "not on", "not assign", "not part",
    "not include", "not consider", "not in our", "did not", "does not",
    "do not", "skip", "beyond",
)

# Candidate JSON keys for "paper identifier" / "total score" / "recommendation".
# Evaluated case-insensitively so a correct model is not punished for a slightly
# different (but equally reasonable) field name.
ID_KEYS = ("paper_id", "id", "paper", "arxiv_id", "paperid",
           "paper_identifier", "identifier", "pid")
SCORE_KEYS = ("total_score", "total", "score", "totalscore")
REC_KEYS = ("recommendation", "recommend", "decision", "verdict")


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT, LOCAL_FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        if CURRENT_CATEGORY == "local":
            LOCAL_FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def _to_float(v):
    """Robustly convert a cell / JSON value to float.

    Accepts int/float/str. For strings: strip whitespace, a leading '='
    (spreadsheet formula), thousands separators, currency symbols and '%',
    then parse. Returns None when the value is not numeric.
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    if s.startswith("="):
        s = s[1:]
    s = s.replace(",", "").replace("$", "").replace("¥", "").replace("€", "")
    s = s.replace(" ", "").replace("%", "")
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def num_close(a, b, tol=2.0):
    """Return |a-b| <= tol when both sides parse as numbers.

    Falls back to case-insensitive string equality only when one side is
    non-numeric (e.g. an Excel formula string or a human-formatted value).
    """
    fa, fb = _to_float(a), _to_float(b)
    if fa is not None and fb is not None:
        return abs(fa - fb) <= tol
    return str(a).strip().lower() == str(b).strip().lower()


def _sheet_headers(grid, header_row=0, limit=12):
    return [str(grid.get((header_row, c), "")).lower() for c in range(limit)]


def _probe_header_row(grid, limit_rows=4, limit_cols=12):
    """Locate the header row: the first row (0-based) where several cells look
    like column headers (contain 'paper'/'score'/'title'/'rank'/'recommend'/
    'method'/'total'/'novelty'). Defaults to row 0 when no such row exists, so
    a leading title/blank row does not shift the header detection."""
    for r in range(limit_rows):
        hits = 0
        for c in range(limit_cols):
            cell = grid.get((r, c))
            if cell is None:
                continue
            low = str(cell).lower()
            if any(k in low for k in ("paper", "score", "title", "rank",
                                      "recommend", "method", "total", "novelty")):
                hits += 1
        if hits >= 2:
            return r
    return 0


def _find_col(grid, keywords, header_row=0, limit=12, fallback=None):
    """Find the first header cell containing every keyword (case-insensitive)."""
    for c in range(limit):
        cell = grid.get((header_row, c))
        if cell is None:
            continue
        low = str(cell).lower()
        if all(k in low for k in keywords):
            return c
    return fallback


def _paper_id_col(grid, header_row=0, limit=12):
    return _find_col(grid, ("paper", "id"), header_row, limit, fallback=0)


def _data_ids(grid, col, header_row=0):
    """Collect non-empty values from the given column, skipping the header row."""
    ids = set()
    for (r, c), v in grid.items():
        if r > header_row and c == col and v:
            ids.add(str(v).strip())
    return ids


def _norm_word(v):
    """Normalize a single-word label: lowercase, keep only the first token."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    tok = re.split(r"[\s()\-/,:;]+", s)[0]
    return tok.strip(".").lower()


def _load_grid(cur, ss_id, sheet_id):
    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (ss_id, sheet_id))
    grid = {}
    for r, c, v in cur.fetchall():
        grid[(r, c)] = v
    return grid


def _count_review_blocks(text, max_gap=4000):
    """Count distinct per-paper review regions: a strengths discussion followed
    (within max_gap chars) by a weaknesses discussion.

    review_template.md asks for a 'Strengths' paragraph and a 'Weaknesses'
    paragraph per paper, so each paper contributes one block. Used as a layout-
    independent signal so the per-paper check does not depend on where paper ids
    happen to be anchored in the document.
    """
    strengths = [m.start() for m in re.finditer(r"strength", text)]
    weaknesses = [m.start() for m in re.finditer(r"weakness", text)]
    blocks = 0
    covered = -1
    wi = 0
    for i in strengths:
        if i < covered:
            continue
        while wi < len(weaknesses) and weaknesses[wi] < i:
            wi += 1
        if wi < len(weaknesses) and weaknesses[wi] - i <= max_gap:
            blocks += 1
            covered = weaknesses[wi]
    return blocks


def _pid_near_sw(text, pid, before=1500, after=2500):
    """True if ANY occurrence of pid sits inside a window that also mentions
    'strength' or 'weakness'. Anchors on every occurrence, not just the first,
    so a paper whose id is mentioned in the Overview and/or the Recommendations
    list is still matched to its strengths/weaknesses discussion."""
    start = 0
    while True:
        idx = text.find(pid, start)
        if idx < 0:
            return False
        window = text[max(0, idx - before): idx + after]
        if "strength" in window or "weakness" in window:
            return True
        start = idx + len(pid)


def check_gsheet():
    """Check Google Sheet via database.

    The spreadsheet is a required deliverable, so its content checks are binding
    (local category). A DB-connectivity failure is NOT a model failure: it is
    reported as a single runtime FAIL and the exit gate ignores it.
    """
    global HARD_MISSING, CURRENT_CATEGORY
    print("\n=== Checking Google Sheet 'Paper Review Matrix' ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        CURRENT_CATEGORY = "local"

        # Find spreadsheets. The task is a shared-artifact multi-agent task, so
        # more than one "Paper Review Matrix" may exist; evaluate the most
        # data-complete one (most cells) rather than blindly picking rows[0].
        cur.execute("""
            SELECT s.id, s.title,
                   (SELECT COUNT(*) FROM gsheet.cells c WHERE c.spreadsheet_id = s.id) AS n
            FROM gsheet.spreadsheets s
            WHERE LOWER(s.title) LIKE '%paper review%'
            ORDER BY n DESC, s.created_at ASC
        """)
        rows = cur.fetchall()
        check("Spreadsheet 'Paper Review Matrix' exists", len(rows) >= 1,
              f"Found {len(rows)} matching spreadsheets")
        if not rows:
            # Spreadsheet genuinely absent (query succeeded) -> hard requirement.
            HARD_MISSING += 1
            cur.close(); conn.close()
            return

        ss_id = rows[0][0]

        # Check sheets
        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
        sheets = cur.fetchall()
        sheet_names = {s[1].strip().lower(): s[0] for s in sheets}

        # Sheet 1: Review Scores
        review_key = None
        for name, sid in sheet_names.items():
            if "review" in name and "score" in name:
                review_key = sid
                break
        check("Sheet 'Review Scores' exists", review_key is not None,
              f"Sheets: {list(sheet_names.keys())}")
        if review_key is None:
            HARD_MISSING += 1

        if review_key is not None:
            grid = _load_grid(cur, ss_id, review_key)

            # Locate the header row (tolerate a leading title/blank row)
            hdr = _probe_header_row(grid)
            headers = _sheet_headers(grid, hdr)
            check("Review Scores has Paper_ID column",
                  any("paper" in h and "id" in h for h in headers),
                  f"Headers: {headers}")
            check("Review Scores has Total_Score column",
                  any("total" in h for h in headers),
                  f"Headers: {headers}")

            # Check 4 data rows (find paper id column, fall back to col 0)
            pid_col = _paper_id_col(grid, hdr)
            data_rows = _data_ids(grid, pid_col, hdr)
            check("Review Scores has 4 target papers",
                  data_rows.issuperset(TARGET_IDS) or len(data_rows) >= 4,
                  f"Found IDs: {data_rows}")

            # Check no noise papers
            noise_in_review = data_rows.intersection(NOISE_IDS)
            check("Review Scores excludes noise papers",
                  len(noise_in_review) == 0,
                  f"Found noise: {noise_in_review}")

            # Check total scores are reasonable (between 5 and 25)
            total_col = _find_col(grid, ("total",), hdr)
            if total_col is not None:
                for r in range(hdr + 1, hdr + 6):
                    val = grid.get((r, total_col))
                    tval = _to_float(val)
                    if tval is None:
                        # non-numeric cell (e.g. a live =SUM formula whose value
                        # the mock never computes) - cannot range-check, skip
                        continue
                    check(f"Row {r} total score in range",
                          5 <= tval <= 25,
                          f"Got {val}")

        # Sheet 2: Methodology Comparison
        method_key = None
        for name, sid in sheet_names.items():
            if "method" in name:
                method_key = sid
                break
        check("Sheet 'Methodology Comparison' exists", method_key is not None,
              f"Sheets: {list(sheet_names.keys())}")
        if method_key is None:
            HARD_MISSING += 1

        if method_key is not None:
            grid = _load_grid(cur, ss_id, method_key)

            hdr = _probe_header_row(grid)
            pid_col = _paper_id_col(grid, hdr)
            data_rows = _data_ids(grid, pid_col, hdr)
            check("Methodology Comparison has target papers",
                  len(data_rows.intersection(TARGET_IDS)) >= 4,
                  f"Found: {data_rows}")

            # Check methods column has content (find by header, fall back to col 2)
            method_col = _find_col(grid, ("method",), hdr, fallback=2)
            methods_populated = 0
            for r in range(hdr + 1, hdr + 6):
                val = grid.get((r, method_col))
                if val and len(str(val).strip()) > 0:
                    methods_populated += 1
            check("Methodology rows have methods content",
                  methods_populated >= 3,
                  f"Populated: {methods_populated}")

        # Sheet 3: Rankings
        rank_key = None
        for name, sid in sheet_names.items():
            if "rank" in name:
                rank_key = sid
                break
        check("Sheet 'Rankings' exists", rank_key is not None,
              f"Sheets: {list(sheet_names.keys())}")
        if rank_key is None:
            HARD_MISSING += 1

        if rank_key is not None:
            grid = _load_grid(cur, ss_id, rank_key)

            hdr = _probe_header_row(grid)
            headers = _sheet_headers(grid, hdr)
            check("Rankings has Recommendation column",
                  any("recommend" in h for h in headers),
                  f"Headers: {headers}")

            # Check recommendations
            rec_col = _find_col(grid, ("recommend",), hdr)
            if rec_col is not None:
                recs = set()
                for r in range(hdr + 1, hdr + 6):
                    val = grid.get((r, rec_col))
                    if val:
                        recs.add(_norm_word(val))
                check("Rankings has Accept/Revise/Reject values",
                      recs.issubset({"accept", "revise", "reject"}) and len(recs) >= 1,
                      f"Found: {recs}")

                # Recommendations must match the total-score bands stated in the
                # task (Accept >= 20, Revise 15-19, Reject < 15). This check is
                # judgment-free: a conservative reviewer who scores every paper
                # <= 19 and writes all-Revise is internally consistent and must
                # not be penalized for the absence of an Accept.
                score_col = _find_col(grid, ("total",), hdr, fallback=3)
                if score_col is not None:
                    checked = 0
                    consistent = 0
                    for r in range(hdr + 1, hdr + 6):
                        fs = _to_float(grid.get((r, score_col)))
                        rval = grid.get((r, rec_col))
                        if fs is None or not rval:
                            continue
                        rec = _norm_word(rval)
                        ok = (rec == "accept" and fs >= 20) or \
                             (rec == "revise" and 15 <= fs < 20) or \
                             (rec == "reject" and fs < 15)
                        checked += 1
                        consistent += 1 if ok else 0
                    if checked > 0:
                        check("Rankings recommendations match total-score bands",
                              consistent == checked,
                              f"{consistent}/{checked} rows consistent")

        cur.close()
        conn.close()
    except Exception as e:
        # DB unreachable / schema issue: not a model failure. Report as runtime
        # so it never turns into a hard (local) fail by itself.
        CURRENT_CATEGORY = "runtime"
        check("GSheet check", False, str(e))


def check_word(agent_workspace):
    """Check Conference_Review_Summary.docx."""
    global CURRENT_CATEGORY
    CURRENT_CATEGORY = "local"
    print("\n=== Checking Conference_Review_Summary.docx ===")
    docx_path = os.path.join(agent_workspace, "Conference_Review_Summary.docx")
    check("Conference_Review_Summary.docx exists", os.path.isfile(docx_path))
    if not os.path.isfile(docx_path):
        return
    try:
        from docx import Document
        doc = Document(docx_path)
        text = " ".join(p.text for p in doc.paragraphs).lower()
        check("Document has substantial content", len(text) > 500, f"Length: {len(text)}")

        # Check sections
        check("Contains 'overview' section",
              "overview" in text)
        check("Contains 'per-paper review' or individual reviews",
              "per-paper" in text or "per paper" in text or "2401.00101" in text)
        check("Contains 'comparative analysis'",
              "comparative" in text or "comparison" in text)
        check("Contains 'recommendation' section",
              "recommendation" in text)

        # Check paper references
        for pid in TARGET_IDS:
            check(f"Mentions paper {pid}", pid in text, "Not found in document")

        # Check key terms
        check("Mentions strengths/weaknesses",
              "strength" in text or "weakness" in text)
        # Recommendations section lists accept/revise/reject for each paper. A
        # conservative but internally consistent reviewer may legitimately end
        # up with all-Revise (or all-Reject), so any single recommendation word
        # is sufficient.
        check("Contains accept/revise/reject recommendations",
              any(w in text for w in ("accept", "revise", "reject")))

        # Check noise papers. A faithful model may explicitly state that it
        # excluded the noise papers (a legitimate, even desirable, handling),
        # so only flag a mention that is NOT accompanied by an exclusion context.
        for nid in NOISE_IDS:
            idx = text.find(nid)
            if idx < 0:
                check(f"Does not include noise paper {nid}", True)
                continue
            window = text[max(0, idx - 200): idx + 200]
            excluded = any(k in window for k in EXCLUDE_CONTEXT)
            check(f"Does not include noise paper {nid}",
                  excluded,
                  f"Found noise paper {nid} without exclusion context")

        # Verify each target paper has a strengths/weaknesses discussion.
        # review_template.md suggests paper-title subsection headings, so a
        # correct document may list all paper ids in the Overview or only in the
        # Recommendations list while the per-paper strengths/weaknesses sit under
        # a title heading. Anchor on EVERY id occurrence with a generous window;
        # as a layout-independent fallback, also accept when the document has at
        # least one complete review block (strengths followed by weaknesses) per
        # paper.
        review_blocks = _count_review_blocks(text)
        for pid in TARGET_IDS:
            check(f"Paper {pid} has strengths/weaknesses discussion",
                  _pid_near_sw(text, pid) or review_blocks >= 4,
                  f"No strength/weakness discussion found near {pid} "
                  f"(review blocks: {review_blocks})")

    except ImportError:
        check("python-docx available", False)
    except Exception as e:
        check("Word document readable", False, str(e))


def _first_value(entry, keys):
    """Return the first present value of any candidate key (case-insensitive)."""
    if not isinstance(entry, dict):
        return None
    for k, v in entry.items():
        if str(k).lower() in keys:
            return v
    return None


def _extract_ids(entries):
    """Collect paper identifiers from a list of dict entries.

    Tolerant of common field-name variants (case-insensitive); as a last resort
    accepts any string value that looks like an arxiv id (e.g. 2401.00101).
    """
    ids = set()
    for e in entries:
        if not isinstance(e, dict):
            continue
        val = _first_value(e, ID_KEYS)
        if val is not None and str(val).strip():
            ids.add(str(val).strip())
            continue
        # last resort: any value that looks like an arxiv id
        for v in e.values():
            if isinstance(v, str) and re.fullmatch(r"\d{4}\.\d{4,5}", v.strip()):
                ids.add(v.strip())
                break
    return ids


def _json_entry_ids(data):
    if isinstance(data, list):
        return _extract_ids(data)
    if isinstance(data, dict):
        # Unwrap common wrapper containers (e.g. {"papers": [...]}) before
        # falling back to treating the keys themselves as the ids.
        for k, v in data.items():
            if isinstance(v, list) and v:
                ids = _extract_ids(v)
                if ids:
                    return ids
        return {str(k).strip() for k in data.keys()}
    return set()


def _find_recommendation(entry):
    """Return the normalized recommendation of an entry.

    Tries the declared field name first; as a fallback scans every string value
    (so an unknown-but-reasonable field name such as 'final_recommendation' is
    still recognized). _norm_word only keeps the first token, so prose values
    cannot false-positive.
    """
    if isinstance(entry, dict):
        v = _first_value(entry, REC_KEYS)
        if v is not None and str(v).strip():
            return _norm_word(v)
        for val in entry.values():
            if isinstance(val, str):
                nw = _norm_word(val)
                if nw in ("accept", "revise", "reject"):
                    return nw
    return ""


def check_json_files(agent_workspace):
    """Check intermediate JSON files.

    These are binding (local category): the task requires the scripts to be
    written and run, so a correct model always produces well-formed JSON with
    the four target papers. A missing or empty JSON therefore fails the task.
    """
    global CURRENT_CATEGORY
    CURRENT_CATEGORY = "local"
    print("\n=== Checking Intermediate JSON Files ===")

    # methodology_analysis.json
    ma_path = os.path.join(agent_workspace, "methodology_analysis.json")
    check("methodology_analysis.json exists", os.path.isfile(ma_path))
    if os.path.isfile(ma_path):
        try:
            with open(ma_path) as f:
                ma = json.load(f)
            ids = _json_entry_ids(ma)
            check("methodology_analysis has 4 target papers",
                  ids.issuperset(TARGET_IDS) or len(ids) >= 4,
                  f"Found: {ids}")
        except Exception as e:
            check("methodology_analysis readable", False, str(e))

    # comparison_matrix.json
    cm_path = os.path.join(agent_workspace, "comparison_matrix.json")
    check("comparison_matrix.json exists", os.path.isfile(cm_path))
    if os.path.isfile(cm_path):
        try:
            with open(cm_path) as f:
                cm = json.load(f)
            ids = _json_entry_ids(cm)
            check("comparison_matrix has 4 target papers",
                  ids.issuperset(TARGET_IDS) or len(ids) >= 4,
                  f"Found: {ids}")
        except Exception as e:
            check("comparison_matrix readable", False, str(e))

    # final_rankings.json
    fr_path = os.path.join(agent_workspace, "final_rankings.json")
    check("final_rankings.json exists", os.path.isfile(fr_path))
    if os.path.isfile(fr_path):
        try:
            with open(fr_path) as f:
                fr = json.load(f)
            if isinstance(fr, list):
                check("final_rankings has 4 entries", len(fr) >= 4, f"Got {len(fr)}")
                # Check sorted by score descending
                if len(fr) >= 2:
                    scores = []
                    for e in fr:
                        s = _to_float(_first_value(e, SCORE_KEYS))
                        scores.append(s if s is not None else 0.0)
                    check("final_rankings sorted by score descending",
                          all(scores[i] >= scores[i+1] for i in range(len(scores)-1)),
                          f"Scores: {scores}")
                # Check recommendations present
                recs = {_find_recommendation(e) for e in fr}
                check("final_rankings has recommendations",
                      bool(recs.intersection({"accept", "revise", "reject"})),
                      f"Found: {recs}")
        except Exception as e:
            check("final_rankings readable", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_word(args.agent_workspace)
    check_json_files(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    accuracy = PASS_COUNT / total * 100 if total > 0 else 0
    print(f"\nOverall: {PASS_COUNT}/{total} ({accuracy:.1f}%)")
    result = {"total_passed": PASS_COUNT, "total_checks": total, "accuracy": accuracy}
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)
    print(f"Local FAIL_COUNT: {LOCAL_FAIL_COUNT}, Total FAIL_COUNT: {FAIL_COUNT}, "
          f"Missing deliverables: {HARD_MISSING}")
    # Pass requires: every local check clean (Word doc + the three JSON files
    # with their content), no required deliverable missing (spreadsheet with the
    # three sheets + JSON files present), and a reasonable overall accuracy.
    # This keeps the gsheet/JSON sections binding instead of allowing a model
    # that only produces the Word document to pass.
    sys.exit(0 if LOCAL_FAIL_COUNT == 0 and HARD_MISSING == 0 and accuracy >= 70 else 1)


if __name__ == "__main__":
    main()
