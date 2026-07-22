"""
Evaluation for arxiv-latex-review-notion-word task.
Checks: Notion page, Word document, GSheet.
"""
import argparse
import os
import sys

import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
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
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def norm(s):
    return str(s or "").strip().lower().replace("_", " ")


# Expected papers (from preprocess)
EXPECTED_PAPERS = [
    {"arxiv_id": "2305.20050", "kw": ["dpo", "direct preference"]},
    {"arxiv_id": "2307.09288", "kw": ["llama"]},
    {"arxiv_id": "2310.06825", "kw": ["mistral"]},
]


def _extract_notion_title(props):
    try:
        title_arr = props.get('title', {}).get('title', [])
        return "".join([t.get('text', {}).get('content', '') for t in title_arr])
    except Exception:
        return ""


def check_notion():
    print("\n=== Checking Notion Page ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
    except Exception as e:
        record("Notion connection", False, str(e))
        return

    try:
        cur.execute("""
            SELECT id, properties FROM notion.pages
            WHERE archived = false AND in_trash = false
        """)
        pages = cur.fetchall()

        target_page = None
        for pid, props in pages:
            t = _extract_notion_title(props or {})
            if norm(t) == norm("LLM Fine-Tuning Knowledge Base"):
                target_page = pid
                break
        record("Notion page 'LLM Fine-Tuning Knowledge Base' exists (exact title)",
               target_page is not None,
               f"Searched {len(pages)} pages")

        if target_page:
            # Check that the page has block content covering each relevant paper
            cur.execute("SELECT block_type, content FROM notion.blocks WHERE page_id=%s", (target_page,))
            blocks = cur.fetchall()
            full_text = ""
            for btype, content in blocks:
                if isinstance(content, dict):
                    rt = content.get("rich_text", []) or content.get("text", [])
                    if isinstance(rt, list):
                        for piece in rt:
                            if isinstance(piece, dict):
                                full_text += " " + piece.get("plain_text", "")
                                inner = piece.get("text", {})
                                if isinstance(inner, dict):
                                    full_text += " " + str(inner.get("content", ""))
                else:
                    full_text += " " + str(content)
            full_text_lower = full_text.lower()
            import re as _re
            paper_hits = 0
            for ep in EXPECTED_PAPERS:
                hit = False
                for kw in ep["kw"]:
                    # word-boundary match against false positives like 'mistralian'/'tdpo'
                    if _re.search(r"(?<![a-z0-9])" + _re.escape(kw) + r"(?![a-z0-9])", full_text_lower):
                        hit = True
                        break
                if hit:
                    paper_hits += 1
            record(f"Notion page covers all {len(EXPECTED_PAPERS)} relevant papers",
                   paper_hits == len(EXPECTED_PAPERS),
                   f"hits={paper_hits}, text[:300]={full_text_lower[:300]}")

        conn.close()
    except Exception as e:
        record("Notion check", False, str(e))


def check_word(agent_workspace):
    print("\n=== Checking Word Document ===")
    doc_path = os.path.join(agent_workspace, "LLM_Paper_Synthesis.docx")
    if not os.path.isfile(doc_path):
        record("Word file LLM_Paper_Synthesis.docx exists", False, f"Not found at: {doc_path}")
        return
    record("Word file LLM_Paper_Synthesis.docx exists", True)

    try:
        doc = Document(doc_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        return
    record("Word file readable", True)

    paragraph_texts = [p.text for p in doc.paragraphs]
    full_text = "\n".join(paragraph_texts)
    full_text_lower = full_text.lower()

    # Heading 'LLM Fine-Tuning and Alignment Survey'
    has_heading = "llm fine-tuning and alignment survey" in full_text_lower
    record("Word has heading 'LLM Fine-Tuning and Alignment Survey'", has_heading,
           f"first 200 chars: {full_text[:200]}")

    # Has substantial content
    record("Word has substantial content (> 600 chars)", len(full_text) > 600,
           f"Text length: {len(full_text)}")

    # All 3 relevant papers covered (word-boundary regex to avoid 'tdpo'/'mistralian'/'llamafication')
    import re as _re_word
    def _wb(kw, text):
        return _re_word.search(r"(?<![a-z0-9])" + _re_word.escape(kw) + r"(?![a-z0-9])", text) is not None
    has_dpo = _wb("dpo", full_text_lower) or "direct preference" in full_text_lower
    has_llama = _wb("llama", full_text_lower) or _wb("llama2", full_text_lower) or _wb("llama-2", full_text_lower)
    has_mistral = _wb("mistral", full_text_lower)
    record("Word mentions DPO paper", has_dpo)
    record("Word mentions Llama 2 paper", has_llama)
    record("Word mentions Mistral 7B paper", has_mistral)

    # Authors present (from preprocess data)
    record("Word mentions Rafael Rafailov (DPO)", "rafailov" in full_text_lower)
    record("Word mentions Hugo Touvron (Llama)", "touvron" in full_text_lower)
    record("Word mentions Albert Jiang (Mistral)", "jiang" in full_text_lower)


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
    except Exception as e:
        record("GSheet connection", False, str(e))
        return

    try:
        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        target_ss = None
        for sid, title in spreadsheets:
            if norm(title) == norm("LLM Paper Registry"):
                target_ss = sid
                break

        record("GSheet 'LLM Paper Registry' exists (exact title)",
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

        # Header check
        cur.execute("""
            SELECT value FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index = 0
            ORDER BY col_index
        """, (target_ss, sheet_id))
        headers = [norm(r[0]) for r in cur.fetchall()]
        for h in ["arxiv id", "title", "authors", "published date", "key contribution", "method category"]:
            record(f"GSheet has '{h}' header", h in headers, f"got: {headers}")

        # Row count
        cur.execute("""
            SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 0
        """, (target_ss, sheet_id))
        data_rows = cur.fetchone()[0]
        record("GSheet has 3 data rows (one per paper)",
               data_rows == 3, f"Found {data_rows} data rows")

        # Each expected arxiv ID present in cells
        cur.execute("""
            SELECT LOWER(value) FROM gsheet.cells
            WHERE spreadsheet_id = %s AND sheet_id = %s
        """, (target_ss, sheet_id))
        cell_values = [row[0] for row in cur.fetchall() if row[0]]
        all_text = " ".join(cell_values)
        for ep in EXPECTED_PAPERS:
            record(f"GSheet contains arxiv_id {ep['arxiv_id']}",
                   ep['arxiv_id'] in all_text,
                   f"sample: {all_text[:300]}")

        conn.close()
    except Exception as e:
        record("GSheet check", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_notion()
    check_word(args.agent_workspace)
    check_gsheet()

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== Results: {PASS_COUNT}/{total} passed ===")
    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
