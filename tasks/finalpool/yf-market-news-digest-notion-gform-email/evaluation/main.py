"""Evaluation for yf-market-news-digest-notion-gform-email.

Checks:
1. Notion database "Market News Digest" exists with at least 5 pages
2. GForm "Investor Market Sentiment Survey" exists with 4 questions
3. Email sent to subscribers@newsletter.example.com with subject containing 'Market' and 'Digest'
"""
import argparse
import json
import os
import sys

import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "toolathlon_gym", "user": "eigent", "password": "camel"}

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        d = (detail[:300] + "...") if len(detail) > 300 else detail
        print(f"  [FAIL] {name}: {d}")


def check_notion():
    print("\n=== Check 1: Notion Database ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Look for a database with title exactly 'Market News Digest'
    cur.execute("SELECT id, title FROM notion.databases")
    dbs = cur.fetchall()

    found_db_id = None
    found_db_title = None
    for db_id, title_json in dbs:
        title_str = ""
        if isinstance(title_json, list):
            for item in title_json:
                if isinstance(item, dict):
                    title_str += item.get("plain_text", item.get("text", {}).get("content", ""))
        elif isinstance(title_json, str):
            title_str = title_json
        ts_lower = title_str.lower()
        if "market news digest" in ts_lower or ("market" in ts_lower and "news" in ts_lower and "digest" in ts_lower):
            found_db_id = db_id
            found_db_title = title_str
            break

    check("Notion database titled 'Market News Digest' exists",
          found_db_id is not None,
          f"Found DBs: {[(str(r[0])[:20], str(r[1])[:80]) for r in dbs]}")

    if found_db_id is None:
        cur.close()
        conn.close()
        return

    # Filter pages by parent database id (correctly - not all pages)
    cur.execute("""
        SELECT id, properties FROM notion.pages
        WHERE parent::text LIKE %s
    """, (f"%{found_db_id}%",))
    db_pages = cur.fetchall()
    check("At least 5 news entries in Market News Digest DB",
          len(db_pages) >= 5,
          f"Total pages in DB: {len(db_pages)}")

    # Check pages cover required stock symbols (at least 3 of 5)
    symbols = ["GOOGL", "AMZN", "JPM", "JNJ", "XOM"]
    found_symbols = set()
    summary_violations = 0

    def _extract_text_from_prop(v):
        """Extract plain text from a Notion property value."""
        text = ""
        if isinstance(v, dict):
            # rich_text or title
            for key in ("rich_text", "title"):
                rt = v.get(key)
                if isinstance(rt, list):
                    for piece in rt:
                        if isinstance(piece, dict):
                            text += piece.get("plain_text", "")
                            inner = piece.get("text") or {}
                            if isinstance(inner, dict):
                                text += inner.get("content", "")
            # select
            sel = v.get("select")
            if isinstance(sel, dict):
                text += str(sel.get("name") or "")
            # multi_select
            ms = v.get("multi_select")
            if isinstance(ms, list):
                for item in ms:
                    if isinstance(item, dict):
                        text += " " + str(item.get("name") or "")
            # plain string fallback
        elif isinstance(v, str):
            text = v
        return text

    for pid, props in db_pages:
        # Extract symbol property value precisely
        page_symbol_text = ""
        sum_val = ""
        if isinstance(props, dict):
            for k, v in props.items():
                kl = k.lower()
                if "symbol" in kl:
                    page_symbol_text += " " + _extract_text_from_prop(v)
                if "summary" in kl:
                    sum_val = _extract_text_from_prop(v)
        # Match symbols using word boundary approach: symbol must equal or be a token in symbol property text
        page_sym_upper = page_symbol_text.upper()
        # tokenize
        import re as _re
        tokens = set(_re.findall(r"[A-Z]+", page_sym_upper))
        for sym in symbols:
            if sym in tokens:
                found_symbols.add(sym)
        if sum_val and len(sum_val) > 200:
            summary_violations += 1
    check("Notion pages cover at least 3 of 5 tracked stocks (GOOGL/AMZN/JPM/JNJ/XOM)",
          len(found_symbols) >= 3,
          f"Found symbols: {found_symbols}")
    check("All page Summary properties are <=200 chars",
          summary_violations == 0,
          f"{summary_violations} page(s) exceed 200-char limit")

    # Check DB schema/properties
    cur.execute("SELECT properties FROM notion.databases WHERE id = %s", (found_db_id,))
    db_props_row = cur.fetchone()
    if db_props_row and db_props_row[0]:
        db_props = db_props_row[0]
        ds = str(db_props).lower() if db_props else ""
        for prop_name in ["title", "symbol", "publisher", "published_date", "summary"]:
            check(f"Database has '{prop_name}' property", prop_name.replace("_", "") in ds.replace("_", "") or prop_name in ds,
                  f"DB properties (sample): {ds[:300]}")

    cur.close()
    conn.close()


def check_gform():
    print("\n=== Check 2: Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    check("At least one Google Form exists", len(forms) > 0,
          "No forms found in gform.forms")

    found_form_id = None
    for form_id, title in forms:
        tl = (title or "").lower()
        if "investor" in tl and "market sentiment" in tl and "survey" in tl:
            found_form_id = form_id
            break
    # Stricter: do not fall back to any form
    check("Form titled 'Investor Market Sentiment Survey' found",
          found_form_id is not None,
          f"Forms: {[(str(r[0])[:20], r[1]) for r in forms]}")

    if found_form_id:
        cur.execute("""
            SELECT id, title, question_type, config
            FROM gform.questions WHERE form_id = %s ORDER BY position
        """, (found_form_id,))
        questions = cur.fetchall()
        check("Form has exactly 4 questions", len(questions) == 4,
              f"Got {len(questions)} questions: {[(q[1] or '')[:60] for q in questions]}")

        if len(questions) >= 4:
            titles = [(q[1] or "").lower() for q in questions]
            types = [(q[2] or "").upper() for q in questions]

            def _choices(cfg):
                if cfg is None:
                    return []
                if isinstance(cfg, str):
                    try:
                        cfg = json.loads(cfg)
                    except (TypeError, ValueError):
                        return []
                if isinstance(cfg, dict):
                    return [str(c).strip().lower() for c in (cfg.get("choices") or cfg.get("options") or [])]
                return []

            # Q1: market outlook with 5 sentiment options
            q1_idx = next((i for i, t in enumerate(titles) if "outlook" in t and ("30 day" in t or "next 30" in t)), None)
            if q1_idx is not None:
                ch = _choices(questions[q1_idx][3])
                expected = ["very bullish", "bullish", "neutral", "bearish", "very bearish"]
                ok = sum(1 for e in expected if any(e == c or e in c for c in ch)) >= 4
                check("Q1 market outlook has 5 sentiment options", ok,
                      f"Choices: {ch}")
            else:
                check("Q1 about market outlook in next 30 days exists", False,
                      f"Question titles: {titles}")

            # Q2: sector outperform
            q2_idx = next((i for i, t in enumerate(titles)
                           if "sector" in t and "outperform" in t), None)
            if q2_idx is not None:
                ch = _choices(questions[q2_idx][3])
                expected = ["communication services", "consumer cyclical",
                            "financial services", "healthcare", "energy"]
                ok = sum(1 for e in expected if any(e in c for c in ch)) >= 4
                check("Q2 sector outperform has 5 specific sector options", ok,
                      f"Choices: {ch}")
            else:
                check("Q2 about sector outperformance exists", False,
                      f"Question titles: {titles}")

            # Q3: investment concern
            q3_idx = next((i for i, t in enumerate(titles)
                           if "investment concern" in t or "primary concern" in t), None)
            if q3_idx is not None:
                ch = _choices(questions[q3_idx][3])
                expected = ["inflation", "interest rates", "geopolitical risk",
                            "earnings slowdown", "valuations"]
                ok = sum(1 for e in expected if any(e in c for c in ch)) >= 4
                check("Q3 investment concern has 5 specific options", ok,
                      f"Choices: {ch}")
            else:
                check("Q3 about investment concern exists", False,
                      f"Question titles: {titles}")

            # Q4: equity allocation
            q4_idx = next((i for i, t in enumerate(titles)
                           if "equity allocation" in t or "equity allocations" in t), None)
            if q4_idx is not None:
                ch = _choices(questions[q4_idx][3])
                expected = ["very likely", "likely", "neutral", "unlikely", "very unlikely"]
                ok = sum(1 for e in expected if any(e == c or e in c for c in ch)) >= 4
                check("Q4 equity allocation has 5 likelihood options", ok,
                      f"Choices: {ch}")
            else:
                check("Q4 about equity allocation exists", False,
                      f"Question titles: {titles}")

    cur.close()
    conn.close()


def check_email():
    print("\n=== Check 3: Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Strictly: to subscribers@newsletter.example.com AND subject has Weekly+Market+Digest
    cur.execute("""
        SELECT subject, to_addr, from_addr, body_text FROM email.messages
        WHERE to_addr::text ILIKE '%subscribers@newsletter.example.com%'
    """)
    rows = cur.fetchall()
    check("Email sent to subscribers@newsletter.example.com",
          len(rows) > 0,
          f"No email to subscribers@newsletter.example.com")

    if rows:
        ok_subj = False
        ok_from = False
        ok_body = False
        for subject, to_addr, from_addr, body in rows:
            sl = (subject or "").lower()
            bl = (body or "").lower()
            fl = (from_addr or "").lower()
            if "weekly" in sl and "market" in sl and "digest" in sl:
                ok_subj = True
            if "newsletter@fund.example.com" in fl:
                ok_from = True
            # Body must mention >=3 of GOOGL/AMZN/JPM/JNJ/XOM
            sym_count = sum(1 for s in ["googl", "amzn", "jpm", "jnj", "xom"] if s in bl)
            if sym_count >= 3:
                ok_body = True
        check("Email subject contains 'Weekly' AND 'Market' AND 'Digest'",
              ok_subj,
              f"Subjects: {[r[0] for r in rows]}")
        check("Email from newsletter@fund.example.com",
              ok_from,
              f"From: {[r[2] for r in rows]}")
        check("Email body mentions at least 3 of 5 tracked stocks (GOOGL/AMZN/JPM/JNJ/XOM)",
              ok_body,
              f"Body sample: {[(r[3] or '')[:200] for r in rows][:1]}")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=== Evaluation: yf-market-news-digest-notion-gform-email ===")

    check_notion()
    check_gform()
    check_email()

    print(f"\n=== SUMMARY: {PASS_COUNT} passed, {FAIL_COUNT} failed ===")

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"pass": PASS_COUNT, "fail": FAIL_COUNT}, f)

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
