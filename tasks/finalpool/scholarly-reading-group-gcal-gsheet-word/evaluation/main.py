"""
Evaluation for scholarly-reading-group-gcal-gsheet-word task.
Checks: GSheet reading schedule, GCal events, Word document.
"""
import argparse
import os
import sys

import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0

TRANSFORMER_KEYWORDS = ["attention is all you need", "bert", "language models are few-shot", "gpt"]


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        target_ss = None
        # Prefer exact title match, then near-match with both 'transformer' and 'reading'
        for sid, title in spreadsheets:
            if title and title.strip().lower() == "transformer reading group schedule":
                target_ss = sid
                break
        if target_ss is None:
            for sid, title in spreadsheets:
                tl = (title or "").lower()
                if "transformer" in tl and "reading" in tl and "schedule" in tl:
                    target_ss = sid
                    break

        record("GSheet 'Transformer Reading Group Schedule' exists",
               target_ss is not None,
               f"Found sheets: {[t for _, t in spreadsheets]}")

        if target_ss is None:
            conn.close()
            return

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (target_ss,))
        sheets = cur.fetchall()
        if not sheets:
            record("GSheet has at least one sheet", False)
            conn.close()
            return

        sheet_id = sheets[0][0]
        cur.execute("""
            SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 0
        """, (target_ss, sheet_id))
        data_rows = cur.fetchone()[0]
        # Exactly 3 data rows expected (one per paper / per week)
        record("GSheet has exactly 3 paper rows (one per week)", data_rows == 3,
               f"Found {data_rows} data rows")

        # Header row: Week, Date, Paper_Title, Authors, Discussion_Lead
        cur.execute("""
            SELECT col_index, LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index = 0
            ORDER BY col_index
        """, (target_ss, sheet_id))
        header_cells = cur.fetchall()
        header_text = " ".join(v or "" for _, v in header_cells)
        for col in ["week", "date", "paper", "author", "discussion"]:
            ok = col in header_text
            record(f"GSheet header includes '{col}'", ok,
                   f"Header row: {[v for _, v in header_cells]}")

        # All Discussion_Lead values are TBD
        # Find the Discussion_Lead col_index by header
        disc_col = None
        for col_idx, val in header_cells:
            if val and ("discussion" in val):
                disc_col = col_idx
                break
        if disc_col is not None:
            cur.execute("""
                SELECT LOWER(value) FROM gsheet.cells
                WHERE spreadsheet_id = %s AND sheet_id = %s
                  AND col_index = %s AND row_index > 0
            """, (target_ss, sheet_id, disc_col))
            disc_vals = [r[0] for r in cur.fetchall() if r[0] is not None]
            all_tbd = len(disc_vals) >= 1 and all("tbd" in v for v in disc_vals)
            record("GSheet Discussion_Lead is TBD for all rows", all_tbd,
                   f"Discussion_Lead values: {disc_vals}")
        else:
            record("GSheet has Discussion_Lead column", False,
                   f"Header: {[v for _, v in header_cells]}")

        # Check paper content in cells (keep word-boundary-ish matches but require all 3)
        cur.execute("""
            SELECT LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
        """, (target_ss, sheet_id))
        cell_values = [row[0] for row in cur.fetchall() if row[0]]
        all_text = " ".join(cell_values)

        has_attention = "attention is all you need" in all_text
        has_bert = "bert" in all_text
        has_gpt = ("few-shot learners" in all_text or "few shot learners" in all_text
                   or "language models are few" in all_text or "gpt-3" in all_text)
        record("GSheet contains 'Attention Is All You Need' paper entry", has_attention)
        record("GSheet contains BERT paper entry", has_bert)
        record("GSheet contains 'Language Models are Few-Shot Learners' paper entry", has_gpt)

        conn.close()
    except Exception as e:
        record("GSheet connection", False, str(e))


def check_gcal():
    print("\n=== Checking Google Calendar ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Match exact summary "Reading Group: Transformers Week N" (case-insensitive)
        cur.execute("""
            SELECT summary, start_datetime, end_datetime FROM gcal.events
            WHERE LOWER(summary) LIKE '%reading group%' AND LOWER(summary) LIKE '%transformer%'
            ORDER BY start_datetime
        """)
        events = cur.fetchall()
        record("GCal has Reading Group: Transformers events",
               len(events) >= 3,
               f"Found {len(events)} events: {[(e[0], str(e[1])[:16]) for e in events]}")

        # Required: each week's date + correct title + 90 min duration
        # Reading Group: Transformers Week 1 -> 2026-03-16 ; Week 2 -> 03-23 ; Week 3 -> 03-30
        required = {
            1: "2026-03-16",
            2: "2026-03-23",
            3: "2026-03-30",
        }
        for week, date_iso in required.items():
            week_evs = []
            for summary, sd, ed in events:
                sl = (summary or "").lower()
                if f"week {week}" in sl or f"week{week}" in sl:
                    week_evs.append((summary, sd, ed))
            ok_exists = len(week_evs) >= 1
            record(f"GCal has 'Reading Group: Transformers Week {week}' event",
                   ok_exists,
                   f"Matching events: {[(s, str(sd)[:16]) for s, sd, _ in week_evs]}")
            if not ok_exists:
                continue
            ok_date = any(date_iso in str(sd) for _, sd, _ in week_evs)
            record(f"  Week {week} event dated {date_iso}", ok_date,
                   f"start_datetimes: {[str(sd) for _, sd, _ in week_evs]}")
            # 3pm to 4:30pm = 90 minutes; allow either local 15:00-16:30 or UTC equivalent
            ok_dur = False
            for _, sd, ed in week_evs:
                if sd is None or ed is None:
                    continue
                try:
                    dur_min = (ed - sd).total_seconds() / 60.0
                except Exception:
                    continue
                if abs(dur_min - 90) <= 5:
                    ok_dur = True
                    break
            record(f"  Week {week} event is 90 min (3pm-4:30pm)", ok_dur,
                   f"start/end: {[(str(sd), str(ed)) for _, sd, ed in week_evs]}")
            # 3pm in some timezone: hour-of-day in start should be 15 (any tz). Allow 15 or 19/20 (UTC offsets)
            ok_time = False
            for _, sd, _ in week_evs:
                hours = []
                # try hour-of-string at positions 11:13
                s = str(sd)
                if len(s) >= 13:
                    try:
                        hours.append(int(s[11:13]))
                    except ValueError:
                        pass
                if any(h in (15, 19, 20) for h in hours):
                    ok_time = True
                    break
            record(f"  Week {week} event starts at 3pm-equivalent (local 15:00 or UTC 19/20)", ok_time,
                   f"starts: {[str(sd) for _, sd, _ in week_evs]}")

        conn.close()
    except Exception as e:
        record("GCal connection", False, str(e))


def check_word(agent_workspace):
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "Transformer_Reading_List.docx")
    if not os.path.isfile(doc_path):
        record("Word file Transformer_Reading_List.docx exists", False, f"Not found at: {doc_path}")
        return
    record("Word file Transformer_Reading_List.docx exists", True)

    try:
        doc = Document(doc_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        return
    record("Word file readable", True)

    full_text = "\n".join(p.text for p in doc.paragraphs).lower()

    has_heading = "transformer" in full_text and ("reading group" in full_text or "reading list" in full_text or "architecture" in full_text)
    record("Word has 'Transformer' heading with reading group context", has_heading)

    has_intro = len(full_text) > 300
    record("Word has substantial content", has_intro, f"Text length: {len(full_text)}")

    has_attention = "attention is all you need" in full_text or ("attention" in full_text and "vaswani" in full_text)
    has_bert = "bert" in full_text and ("devlin" in full_text or "bidirectional" in full_text)
    has_fewshot = "few-shot" in full_text or "few shot" in full_text or "gpt-3" in full_text or "language models are" in full_text

    papers_mentioned = sum([has_attention, has_bert, has_fewshot])
    record("Word mentions 'Attention Is All You Need'", has_attention)
    record("Word mentions BERT", has_bert)
    record("Word mentions few-shot learners (GPT-3)", has_fewshot)
    # Require ALL 3 (one paper per week, 3 weeks specified)
    record("Word mentions all 3 transformer papers", papers_mentioned == 3,
           f"Found {papers_mentioned}/3 papers")

    # Check for ALL 3 week assignments
    has_w1 = "week 1" in full_text or "week1" in full_text
    has_w2 = "week 2" in full_text or "week2" in full_text
    has_w3 = "week 3" in full_text or "week3" in full_text
    record("Word assigns Week 1", has_w1)
    record("Word assigns Week 2", has_w2)
    record("Word assigns Week 3", has_w3)

    # Each paper should include a year (2017, 2018, 2020 expected per GT)
    has_y1 = "2017" in full_text
    has_y2 = "2018" in full_text or "2019" in full_text  # BERT publication can be 2018/2019
    has_y3 = "2020" in full_text
    record("Word includes publication year 2017 (Attention)", has_y1)
    record("Word includes publication year 2018/2019 (BERT)", has_y2)
    record("Word includes publication year 2020 (GPT-3)", has_y3)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_gcal()
    check_word(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== Results: {PASS_COUNT}/{total} passed ===")
    if FAIL_COUNT > 0:
        print(f"{FAIL_COUNT} checks failed")
        sys.exit(1)
    else:
        print("All checks passed!")
        sys.exit(0)


if __name__ == "__main__":
    main()
