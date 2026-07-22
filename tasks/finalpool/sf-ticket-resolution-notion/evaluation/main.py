"""
Evaluation for sf-ticket-resolution-notion.
Checks:
1. Notion database "Slow Response Tickets Tracker" with pages per priority
2. Email sent to support.manager@company.example.com with summary
"""
import argparse
import json
import os
import sys

import psycopg2

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


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_data():
    """Query Snowflake mirror for expected slow response ticket stats."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT "PRIORITY",
               COUNT(*) as count,
               ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2) as avg_resp,
               ROUND(AVG("SLA_HOURS")::numeric, 2) as avg_sla,
               ROUND(AVG("RESPONSE_TIME_HOURS" / "SLA_HOURS" * 100)::numeric, 2) as avg_util,
               ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2) as avg_csat
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        WHERE "RESPONSE_TIME_HOURS" / "SLA_HOURS" > 0.5
        GROUP BY "PRIORITY"
        ORDER BY "PRIORITY"
    """)
    priority_stats = cur.fetchall()

    cur.execute("""
        SELECT COUNT(*) FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        WHERE "RESPONSE_TIME_HOURS" / "SLA_HOURS" > 0.5
    """)
    total_slow = cur.fetchone()[0]

    cur.execute('SELECT COUNT(*) FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"')
    total_all = cur.fetchone()[0]

    cur.close()
    conn.close()

    return {
        "total_slow": total_slow,
        "total_all": total_all,
        "slow_pct": round(total_slow / total_all * 100, 2),
        "by_priority": {r[0]: {"count": int(r[1]), "avg_resp": float(r[2]),
                                "avg_sla": float(r[3]), "avg_util": float(r[4]),
                                "avg_csat": float(r[5])} for r in priority_stats},
    }


def check_notion(expected):
    """Check Notion database and pages."""
    print("\n=== Checking Notion ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Check for database with EXACT title match
    cur.execute("""
        SELECT id, title FROM notion.databases
        WHERE archived = false
    """)
    databases = cur.fetchall()

    found_db = None
    for db_id, title in databases:
        title_str = json.dumps(title).lower() if title else ""
        if "slow response tickets tracker" in title_str:
            found_db = (db_id, title)
            break

    record("Notion database 'Slow Response Tickets Tracker' exists",
           found_db is not None,
           f"No database with exact title 'Slow Response Tickets Tracker' found among {len(databases)}")

    if found_db is None:
        cur.close()
        conn.close()
        return False

    db_id = found_db[0]
    # Find pages parented by this database (require parent->database_id match,
    # not generic LIKE on full json — too permissive).
    cur.execute("""
        SELECT id, properties, parent FROM notion.pages
        WHERE archived = false AND in_trash = false
    """)
    all_pages = cur.fetchall()
    db_pages = []
    for page_id, props, parent in all_pages:
        try:
            parent_obj = parent if isinstance(parent, dict) else json.loads(parent or "null")
        except Exception:
            parent_obj = None
        if isinstance(parent_obj, dict):
            pdb = parent_obj.get("database_id") or parent_obj.get("databaseId")
            if pdb and str(pdb) == str(db_id):
                db_pages.append((page_id, props))
    # Strict — no fallback to "all pages" because that would let unrelated
    # pages count as DB rows and contaminate the priority lookup below.

    record("Database has exactly 3 priority pages",
           len(db_pages) == 3, f"Found {len(db_pages)}")

    def _extract_title(props):
        """Extract canonical title text from a Notion page properties dict."""
        if not props:
            return ""
        try:
            obj = props if isinstance(props, dict) else json.loads(props)
        except Exception:
            return ""
        # Look for any property of type "title" (Notion stores titles as a
        # rich-text array under a property whose type=='title').
        title_text = ""
        if isinstance(obj, dict):
            for _, val in obj.items():
                if not isinstance(val, dict):
                    continue
                if val.get("type") == "title" and isinstance(val.get("title"), list):
                    parts = []
                    for rt in val["title"]:
                        if isinstance(rt, dict):
                            pt = rt.get("plain_text") or (rt.get("text", {}) or {}).get("content")
                            if pt:
                                parts.append(str(pt))
                    title_text = " ".join(parts).strip()
                    if title_text:
                        return title_text
            # Fallback: title under properties.title.title
            t = obj.get("title")
            if isinstance(t, dict) and isinstance(t.get("title"), list):
                parts = []
                for rt in t["title"]:
                    if isinstance(rt, dict):
                        pt = rt.get("plain_text") or (rt.get("text", {}) or {}).get("content")
                        if pt:
                            parts.append(str(pt))
                return " ".join(parts).strip()
        return ""

    # Per-priority page validation: title must match the priority name
    priority_pages = {}  # priority -> (page_id, full_text)
    for page_id, props in db_pages:
        title = _extract_title(props).strip().lower()
        for priority in ["high", "medium", "low"]:
            # Accept exact title match or "<priority> priority"
            if title == priority or title == f"{priority} priority":
                if priority not in priority_pages:
                    cur.execute("""
                        SELECT block_data::text FROM notion.blocks
                        WHERE parent_id = %s AND archived = false
                    """, (page_id,))
                    blocks = cur.fetchall()
                    full_text = (
                        json.dumps(props).lower() + " " +
                        " ".join(str(b[0] or "").lower() for b in blocks)
                    )
                    priority_pages[priority] = (page_id, full_text)
                break

    for priority in ["high", "medium", "low"]:
        record(f"Page titled '{priority}' exists in tracker DB",
               priority in priority_pages,
               f"'{priority}' page not found")

    # Per-priority content validation: check the 5 stats are within tolerance.
    # Use regex word-boundaries to avoid the substring false positives where a
    # number like "12.3" matches inside another value like "112.3" or "12.34".
    import re as _re
    def _has_number_token(text, val_str):
        # \D or string boundary on both sides; treat '.' specially so we don't
        # match "12.3" inside "12.34".
        pat = r'(?<!\d)' + _re.escape(val_str) + r'(?!\d|\.\d)'
        return bool(_re.search(pat, text))

    for priority, stats in expected["by_priority"].items():
        plower = priority.lower()
        if plower not in priority_pages:
            continue
        full_text = priority_pages[plower][1]
        # Count: word-boundary-bounded number match
        count = stats["count"]
        count_tokens = [str(count), f"{count:,}"]
        count_found = any(_has_number_token(full_text, t) for t in count_tokens)
        record(f"{priority} page has count ~{count}", count_found,
               f"Expected {count} in content")
        # avg_resp, avg_sla, avg_util, avg_csat - check value with tolerance
        for metric_name, metric_val, tol in [
            ("avg_resp", stats["avg_resp"], 0.2),
            ("avg_sla", stats["avg_sla"], 0.2),
            ("avg_util", stats["avg_util"], 1.0),
            ("avg_csat", stats["avg_csat"], 0.1),
        ]:
            # Generate tolerance candidates within +/- tol, in 0.1 increments,
            # rounded to 1 and 2 decimals.
            cand_strs = set()
            base = float(metric_val)
            steps = max(1, int(tol * 10))
            for delta_steps in range(-steps, steps + 1):
                v = round(base + delta_steps * 0.1, 1)
                cand_strs.add(f"{v:.1f}")
                cand_strs.add(f"{v:.2f}")
                # also non-decimal for whole numbers
                if abs(v - round(v)) < 1e-9:
                    cand_strs.add(str(int(round(v))))
            cand_strs.add(f"{metric_val:.1f}")
            cand_strs.add(f"{metric_val:.2f}")
            found = any(_has_number_token(full_text, c) for c in cand_strs)
            record(f"{priority} page has {metric_name}~{metric_val}", found,
                   f"Expected {metric_val}")

    cur.close()
    conn.close()
    return found_db is not None


def check_email(expected):
    """Check email to support manager."""
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("At least 1 email sent", len(emails) >= 1, f"Found {len(emails)}")

    # Find the matching email: must be sent to support.manager AND have subject "Slow Response Tickets Report"
    matched_email = None
    for subject, from_addr, to_addr, body_text in emails:
        to_str = str(to_addr or "").lower()
        subject_lower = (subject or "").lower()
        if ("support.manager@company.example.com" in to_str
            and "slow response tickets report" in subject_lower):
            matched_email = (subject, from_addr, to_addr, body_text)
            break

    record("Email to support.manager with subject 'Slow Response Tickets Report'",
           matched_email is not None,
           "Did not find email with required to/subject")

    if matched_email:
        subject, from_addr, to_addr, body_text = matched_email
        body_lower = (body_text or "").lower()

        # Check from address
        record("From address is escalations@support.example.com",
               "escalations@support.example.com" in str(from_addr or "").lower(),
               f"From: {from_addr}")

        # Total slow count - check as a standalone token (regex)
        import re
        total_str = str(expected['total_slow'])
        # Match token: surrounded by non-digit chars or string boundary
        pattern = r'(?<!\d)' + re.escape(total_str) + r'(?!\d)'
        record(f"Email body has total slow count token ({total_str})",
               bool(re.search(pattern, body_text or "")),
               f"Expected token {total_str}")

        # Slow pct
        pct_str = str(expected['slow_pct'])
        # Allow decimal-rounded variants
        candidates = {pct_str, str(round(expected['slow_pct'])), str(round(expected['slow_pct'], 1))}
        record(f"Email body has slow percentage (~{pct_str})",
               any(c in (body_text or "") for c in candidates),
               f"Expected {pct_str}")

        # Per-priority counts and util
        for priority, stats in expected["by_priority"].items():
            plower = priority.lower()
            count_pat = r'(?<!\d)' + re.escape(str(stats["count"])) + r'(?!\d)'
            count_found = bool(re.search(count_pat, body_text or ""))
            record(f"Body mentions {priority} count {stats['count']}",
                   count_found, f"Expected {stats['count']}")
            # utilization mention - allow rounded
            util = stats["avg_util"]
            util_cands = {str(util), str(round(util, 1)), str(round(util))}
            util_found = any(c in (body_text or "") for c in util_cands)
            record(f"Body mentions {priority} util ~{util}", util_found,
                   f"Expected {util}")

    cur.close()
    conn.close()
    return matched_email is not None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = get_expected_data()
    print(f"[eval] Total slow response: {expected['total_slow']}/{expected['total_all']} ({expected['slow_pct']}%)")

    notion_ok = check_notion(expected)
    email_ok = check_email(expected)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = PASS_COUNT > 0 and FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
