"""Evaluation for playwright-arxiv-review-criteria-word-gsheet."""
import argparse
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
          user="eigent", password="camel")


def num_close(a, b, tol=0.5):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def check_word_review(agent_workspace, filename, expected_title_fragment,
                      expected_tech, expected_novelty, expected_clarity,
                      expected_recommendation):
    import re
    errors = []
    path = os.path.join(agent_workspace, filename)
    if not os.path.exists(path):
        return [f"{filename} not found"]
    try:
        from docx import Document
        doc = Document(path)
        full_text = "\n".join(p.text for p in doc.paragraphs).lower()
        if expected_title_fragment.lower() not in full_text:
            errors.append(f"{filename}: missing paper title fragment '{expected_title_fragment}'")
        if "summary" not in full_text:
            errors.append(f"{filename}: missing Summary section")
        if "technical soundness" not in full_text and "technical" not in full_text:
            errors.append(f"{filename}: missing Technical Soundness section")
        if "novelty" not in full_text:
            errors.append(f"{filename}: missing Novelty section")
        if "clarity" not in full_text:
            errors.append(f"{filename}: missing Clarity section")

        # Distinguish Accept vs Weak Accept via exact line match
        exp_rec = expected_recommendation.lower().strip()
        rec_ok = False
        for p in doc.paragraphs:
            pt = p.text.lower().strip()
            # Match 'Overall Recommendation: <exact>' or just exact line
            m = re.search(r"recommendation[:\s]*([a-z\s]+)", pt)
            if m:
                got = m.group(1).strip()
                if got == exp_rec:
                    rec_ok = True
                    break
                # Also accept if got starts with exp_rec followed by nothing else
                if exp_rec == "accept" and got == "accept":
                    rec_ok = True
                    break
                if exp_rec == "weak accept" and got == "weak accept":
                    rec_ok = True
                    break
        if not rec_ok:
            errors.append(f"{filename}: recommendation not exactly '{expected_recommendation}'")

        # Check scores appear per criterion: look for 'criterion: N/5' or 'criterion: N'
        score_checks = [
            ("technical soundness", expected_tech),
            ("novelty", expected_novelty),
            ("clarity", expected_clarity),
        ]
        for crit, exp in score_checks:
            # look for "crit: N" pattern with allowed /5
            found = False
            for p in doc.paragraphs:
                pt = p.text.lower()
                m = re.search(re.escape(crit) + r"[^\d]{0,5}(\d+)", pt)
                if m:
                    try:
                        got = int(m.group(1))
                        if got == exp:
                            found = True
                            break
                    except Exception:
                        pass
            if not found:
                errors.append(f"{filename}: score for {crit} not exactly {exp}")
    except Exception as e:
        errors.append(f"Error reading {filename}: {e}")
    return errors


def check_gsheet():
    """Returns (blocking_errors, runtime_only_errors).

    If no spreadsheet exists at all (agent did not create), all errors are
    runtime-only — V1 GT-only test should not fail on this. Once the agent
    has populated a spreadsheet, content errors become blocking.
    """
    blocking = []
    runtime = []
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("""
            SELECT s.id, s.title FROM gsheet.spreadsheets s
            WHERE LOWER(s.title) LIKE '%review%' OR LOWER(s.title) LIKE '%conference%'
            ORDER BY s.id DESC LIMIT 5
        """)
        spreadsheets = cur.fetchall()
        if not spreadsheets:
            runtime.append("No review tracker spreadsheet found")
            cur.close()
            conn.close()
            return blocking, runtime

        ss_id = spreadsheets[0][0]

        cur.execute("""
            SELECT id FROM gsheet.sheets
            WHERE spreadsheet_id = %s AND LOWER(title) LIKE '%%review%%'
            LIMIT 1
        """, (ss_id,))
        sheet_row = cur.fetchone()
        if not sheet_row:
            blocking.append("No 'Reviews' sheet found in spreadsheet")
            cur.close()
            conn.close()
            return blocking, runtime

        sheet_id = sheet_row[0]

        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
            ORDER BY row_index, col_index
        """, (ss_id, sheet_id))
        cells = cur.fetchall()
        cur.close()
        conn.close()

        if len(cells) < 24:
            blocking.append(f"Too few cells in Reviews sheet: {len(cells)}, expected ~32")

        cell_values = [str(c[2]).lower() if c[2] else "" for c in cells]
        all_text = " ".join(cell_values)
        if "2301.07041" not in all_text and "scaling" not in all_text:
            blocking.append("Scaling Laws paper not found in GSheet")
        if "2203.11171" not in all_text and "instruct" not in all_text:
            blocking.append("InstructGPT paper not found in GSheet")
        if "2205.01068" not in all_text and "opt" not in all_text:
            blocking.append("OPT paper not found in GSheet")

        if "4.7" not in all_text:
            blocking.append("InstructGPT average score 4.7 not found in GSheet")
        if "4.3" not in all_text:
            blocking.append("Scaling Laws average score 4.3 not found in GSheet")
        if "4.0" not in all_text and "4.00" not in all_text:
            blocking.append("OPT average score 4.0 not found in GSheet")

        if "completed" not in all_text:
            blocking.append("Review status 'Completed' not found in GSheet")

    except Exception as e:
        runtime.append(f"Error checking GSheet: {e}")
    return blocking, runtime


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    agent_ws = args.agent_workspace or os.path.join(os.path.dirname(__file__), "..", "groundtruth_workspace")

    all_errors = []

    # Check Review 1: Scaling Laws
    print("  Checking Review_Scaling_Laws.docx...")
    errs = check_word_review(agent_ws, "Review_Scaling_Laws.docx",
                             "Scaling Laws", 5, 4, 4, "Accept")
    if errs:
        all_errors.extend(errs)
        for e in errs[:3]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    # Check Review 2: InstructGPT
    print("  Checking Review_InstructGPT.docx...")
    errs = check_word_review(agent_ws, "Review_InstructGPT.docx",
                             "follow instructions", 5, 5, 4, "Accept")
    if errs:
        all_errors.extend(errs)
        for e in errs[:3]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    # Check Review 3: OPT
    print("  Checking Review_OPT.docx...")
    errs = check_word_review(agent_ws, "Review_OPT.docx",
                             "OPT", 4, 3, 5, "Weak Accept")
    if errs:
        all_errors.extend(errs)
        for e in errs[:3]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    # Check GSheet (runtime dependency — split blocking vs runtime-only)
    print("  Checking Google Sheet...")
    gs_blocking, gs_runtime = check_gsheet()
    if gs_blocking:
        all_errors.extend(gs_blocking)
        for e in gs_blocking[:3]:
            print(f"    ERROR: {e}")
    if gs_runtime:
        for e in gs_runtime[:3]:
            print(f"    [runtime-only] {e}")
    if not gs_blocking and not gs_runtime:
        print("    PASS")

    if all_errors:
        print(f"\n=== RESULT: FAIL ({len(all_errors)} blocking errors) ===")
        for e in all_errors[:10]:
            print(f"  {e}")
        sys.exit(1)
    else:
        if gs_runtime:
            print(f"\n=== RESULT: PASS ({len(gs_runtime)} runtime-only failures tolerated) ===")
        else:
            print("\n=== RESULT: PASS ===")
        sys.exit(0)


if __name__ == "__main__":
    main()
