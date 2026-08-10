"""
Evaluation for arxiv-latex-review-notion-word task.
Checks: Notion page, Word document, GSheet.

CLI contract (must not change):
    python evaluation/main.py --agent_workspace <ws> --groundtruth_workspace <gt>
        [--res_log_file <f>] [--launch_time <iso>]
Exits 0 iff FAIL_COUNT == 0.
"""
import argparse
import json
import os
import re
import sys

import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
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

EXPECTED_GSHEET_HEADERS = [
    "arxiv id", "title", "authors", "published date", "key contribution", "method category",
]


def _rich_text_to_str(obj):
    """Convert a rich-text value (single item, array, or nested wrapper) to its
    plain text, tolerating every shape seen in the Notion MCP payloads:
    {"text": {"content": "..."}} / {"plain_text": "..."} / bare string / etc.
    """
    if isinstance(obj, str):
        return obj
    if isinstance(obj, list):
        return "".join(_rich_text_to_str(x) for x in obj)
    if isinstance(obj, dict):
        pt = obj.get("plain_text")
        if isinstance(pt, str):
            return pt
        text = obj.get("text")
        if isinstance(text, dict):
            c = text.get("content")
            if isinstance(c, str):
                return c
            return _rich_text_to_str(c)
        if isinstance(text, str):
            return text
        # Last resort: recurse into any string-valued field we can find.
        for key in ("content", "rich_text", "title", "text"):
            v = obj.get(key)
            if isinstance(v, str):
                return v
        return ""
    return ""


def _extract_notion_title(props):
    """Robustly extract the title text from a notion page's `properties` jsonb.

    Accepts all documented storage shapes:
      {"title": {"title": [{"text": {"content": "..."}}]}}   (API echo shape)
      {"title": [{"text": {"content": "..."}}], "type": "title"}  (create_page schema)
      {"title": "bare string"} / {"title": {"title": "bare string"}}
    and falls back to any property whose `type` is "title".
    """
    if isinstance(props, str):
        return props
    if not isinstance(props, dict):
        return ""
    try:
        title_obj = props.get("title")
        if title_obj is None:
            # The title property may be keyed by an arbitrary name but typed
            # "title" (e.g. a database-style "Name" property).
            for v in props.values():
                if isinstance(v, dict) and v.get("type") == "title":
                    title_obj = v
                    break
        if isinstance(title_obj, str):
            return title_obj
        if isinstance(title_obj, list):
            return _rich_text_to_str(title_obj)
        if isinstance(title_obj, dict):
            inner = title_obj.get("title")
            if inner is not None:
                return _rich_text_to_str(inner)
    except Exception:
        pass
    return ""


def _extract_block_text(block_data):
    """Recursively pull human-readable text out of a notion block_data jsonb.

    Tolerates every block payload shape the MCP server may store:
      {"paragraph": {"rich_text": [{"text": {"content": "..."}, "plain_text": "..."}]}}
      {"heading_2": {"rich_text": [{"text": "..."}]}}       (bare-string text)
      {"bullet": {"rich_text": "bare string"}}              (bare-string rich_text)
      '"plain string block"'                                (JSON-string block_data)
      {"child_page": {"title": "..."}} / title fields, etc.
    Structural noise (block-type labels, links) is harmless for the keyword
    checks the evaluator performs.
    """
    pieces = []

    def walk(obj):
        if isinstance(obj, str):
            pieces.append(obj)
            return
        if isinstance(obj, list):
            for item in obj:
                walk(item)
            return
        if not isinstance(obj, dict):
            return
        pt = obj.get("plain_text")
        inner = obj.get("text")
        rt = obj.get("rich_text")
        if isinstance(pt, str):
            pieces.append(pt)
        if isinstance(inner, str):
            pieces.append(inner)
        elif isinstance(inner, dict):
            c = inner.get("content")
            if isinstance(c, str):
                pieces.append(c)
            else:
                walk(c)
        if isinstance(rt, str):
            pieces.append(rt)
        else:
            walk(rt)
        for key, val in obj.items():
            if val is pt or val is inner or val is rt:
                continue
            walk(val)

    if isinstance(block_data, str):
        try:
            parsed = json.loads(block_data)
            if isinstance(parsed, str):
                return parsed
            block_data = parsed
        except Exception:
            # A bare text string stored directly
            return block_data
    walk(block_data)
    return " ".join(p for p in pieces if p)


def _word_boundary_hit(kw, text):
    return re.search(r"(?<![a-z0-9])" + re.escape(kw) + r"(?![a-z0-9])", text) is not None


def _paper_keyword_hit(kw, text):
    """Match a paper keyword in lowercase text, tolerating common
    spacing/hyphen concatenation variants (e.g. 'llama' also matches 'llama2'
    / 'Llama-2'; 'mistral' also matches 'Mistral7B')."""
    if _word_boundary_hit(kw, text):
        return True
    kw_norm = re.sub(r"[\s\-.]+", "", kw)
    text_norm = re.sub(r"[\s\-.]+", "", text)
    return bool(kw_norm) and kw_norm in text_norm


def _title_norm(s):
    """Collapse whitespace/underscore/hyphen variants for a fuzzy page-title
    match ('LLM Fine-Tuning Knowledge Base' == 'LLM Fine Tuning Knowledge Base')."""
    return re.sub(r"[\s_\-–—]+", "", norm(s))


def _collect_page_text(cur, root_id):
    """Return the concatenated text of a page's block tree by BFS over block
    parent_id chains.  Covers both direct children (parent_type='page_id') and
    arbitrarily nested descendants (parent_type='block_id' under a heading /
    child_page block), which the MCP server stores the same way an agent may
    append content to a specific block.
    """
    pieces = []
    visited = set()
    frontier = [root_id]
    while frontier:
        cur.execute("""
            SELECT id, block_data FROM notion.blocks
            WHERE parent_id = ANY(%s) AND archived = false AND in_trash = false
        """, (frontier,))
        rows = cur.fetchall()
        nxt = []
        for bid, bd in rows:
            if bid in visited:
                continue
            visited.add(bid)
            pieces.append(_extract_block_text(bd))
            nxt.append(bid)
        frontier = nxt
    return " ".join(p for p in pieces if p)


def check_notion():
    print("\n=== Checking Notion Page ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
    except Exception as e:
        record("Notion connection", False, str(e))
        return

    try:
        # --- find all target pages (multi-agent may create several) ---
        cur.execute("""
            SELECT id, properties FROM notion.pages
            WHERE archived = false AND in_trash = false
        """)
        pages = cur.fetchall()

        targets = []
        titles_seen = []
        expected_title_norm = _title_norm("LLM Fine-Tuning Knowledge Base")
        for pid, props in pages:
            t = _extract_notion_title(props or {})
            titles_seen.append(t)
            if _title_norm(t) == expected_title_norm:
                targets.append(pid)

        # Fallback: some runtimes represent pages as blocks (type child_page/page)
        # whose title lives in block_data; their children then use parent_id = block id.
        if not targets:
            try:
                cur.execute("""
                    SELECT id, block_data FROM notion.blocks
                    WHERE type IN ('child_page', 'page') AND archived = false
                """)
                for bid, bdata in cur.fetchall():
                    t = _extract_block_text(bdata)
                    titles_seen.append(t)
                    if _title_norm(t) == expected_title_norm:
                        targets.append(bid)
            except Exception:
                pass

        record("Notion page 'LLM Fine-Tuning Knowledge Base' exists (exact title)",
               len(targets) > 0,
               f"Searched pages, titles seen: {titles_seen[:10]}")

        # --- check block content covers each relevant paper.
        # Aggregate across every matching page: in homogeneous multi-agent mode
        # the first matching page could be an empty duplicate, so PASS as long
        # as ANY correct-titled page covers all papers.  Block text is collected
        # across the whole block tree (direct children + nested descendants), so
        # content stored under a heading/child_page block is still visible. ---
        best_hits = 0
        best_text = ""
        for target in targets:
            full_text = _collect_page_text(cur, target)
            full_text_lower = full_text.lower()

            paper_hits = 0
            for ep in EXPECTED_PAPERS:
                for kw in ep["kw"]:
                    if _paper_keyword_hit(kw.lower(), full_text_lower):
                        paper_hits += 1
                        break
            if paper_hits > best_hits:
                best_hits = paper_hits
                best_text = full_text_lower

        record(f"Notion page covers all {len(EXPECTED_PAPERS)} relevant papers",
               best_hits == len(EXPECTED_PAPERS),
               f"hits={best_hits}, pages={len(targets)}, text[:300]={best_text[:300]}")

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

    # Heading 'LLM Fine-Tuning and Alignment Survey' (hyphen/space tolerant)
    has_heading = ("llm fine-tuning and alignment survey" in full_text_lower
                   or "llm fine tuning and alignment survey" in full_text_lower)
    record("Word has heading 'LLM Fine-Tuning and Alignment Survey'", has_heading,
           f"first 200 chars: {full_text[:200]}")

    # Has substantial content
    record("Word has substantial content (> 600 chars)", len(full_text) > 600,
           f"Text length: {len(full_text)}")

    # All 3 relevant papers covered (keyword matcher tolerates 'llama2'/'Llama-2'/
    # 'Mistral7B' style variants while still rejecting unrelated words)
    has_dpo = _paper_keyword_hit("dpo", full_text_lower) or "direct preference" in full_text_lower
    has_llama = _paper_keyword_hit("llama", full_text_lower)
    has_mistral = _paper_keyword_hit("mistral", full_text_lower)
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

        matching = [sid for sid, title in spreadsheets if norm(title) == norm("LLM Paper Registry")]

        record("GSheet 'LLM Paper Registry' exists (exact title)",
               len(matching) > 0,
               f"Found sheets: {[t for _, t in spreadsheets]}")

        if not matching:
            conn.close()
            return

        # Aggregate across every matching spreadsheet. In homogeneous multi-agent
        # mode each sub-agent may create its own spreadsheet (or append to one),
        # so checks use the union of headers/cells and the max data-row count.
        union_headers = set()
        max_data_rows = 0
        total_sheets = 0
        all_cell_text = ""
        for sid in matching:
            cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (sid,))
            sheets = cur.fetchall()
            total_sheets += len(sheets)
            for sheet_id, _ in sheets:
                cur.execute("""
                    SELECT value FROM gsheet.cells
                    WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index = 0
                    ORDER BY col_index
                """, (sid, sheet_id))
                headers = [norm(r[0]) for r in cur.fetchall() if r[0] is not None]
                union_headers.update(headers)

                cur.execute("""
                    SELECT COUNT(DISTINCT row_index) FROM gsheet.cells
                    WHERE spreadsheet_id = %s AND sheet_id = %s AND row_index > 0
                """, (sid, sheet_id))
                data_rows = cur.fetchone()[0] or 0
                if data_rows > max_data_rows:
                    max_data_rows = data_rows

                cur.execute("""
                    SELECT LOWER(value) FROM gsheet.cells
                    WHERE spreadsheet_id = %s AND sheet_id = %s
                """, (sid, sheet_id))
                all_cell_text += " " + " ".join(r[0] for r in cur.fetchall() if r[0])

        record("GSheet has at least one sheet",
               total_sheets > 0,
               f"Sheets across {len(matching)} matching spreadsheet(s): {total_sheets}")

        # Header check
        for h in EXPECTED_GSHEET_HEADERS:
            record(f"GSheet has '{h}' header", h in union_headers,
                   f"got: {sorted(union_headers)}")

        # Row count (>= 3, since multi-agent may append duplicate rows)
        record("GSheet has >= 3 data rows (one per paper)",
               max_data_rows >= 3, f"Found max {max_data_rows} data rows")

        # Each expected arxiv ID present in cells
        for ep in EXPECTED_PAPERS:
            record(f"GSheet contains arxiv_id {ep['arxiv_id']}",
                   ep['arxiv_id'] in all_cell_text,
                   f"sample: {all_cell_text[:300]}")

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
