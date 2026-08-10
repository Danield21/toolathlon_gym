"""Evaluation for sf-hr-turnover-notion."""
import argparse
import os
import re
import sys
import json
import psycopg2


def _normalize_title(s):
    """Lowercase, replace punctuation with spaces, collapse whitespace."""
    return " ".join(re.sub(r"[^a-z0-9]+", " ", str(s).lower()).split())


def _levenshtein(a, b):
    """Small-scale DP edit distance between two strings."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def titles_match(expected, actual):
    """Exact Notion title match (task.md prescribes exact titles). After
    normalization the titles must be identical — no fuzzy/edit-distance
    tolerance, so the model must follow the titles given in task.md."""
    e = _normalize_title(expected)
    a = _normalize_title(actual)
    if not e or not a:
        return False
    return e == a

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

def _derive_expected_depts():
    """Derive per-department metrics from Snowflake at runtime to prevent drift fragility."""
    fallback = {
        "engineering": {"count": 7096, "avg": 58991.61, "min": 15360, "max": 695267, "range": 679907},
        "finance": {"count": 7148, "avg": 57878.19, "min": 15760, "max": 638897, "range": 623137},
        "hr": {"count": 7077, "avg": 58920.45, "min": 18307, "max": 692232, "range": 673925},
        "operations": {"count": 7120, "avg": 57808.74, "min": 17168, "max": 656505, "range": 639337},
        "r&d": {"count": 7083, "avg": 57905.93, "min": 15128, "max": 680490, "range": 665362},
        "sales": {"count": 7232, "avg": 58864.79, "min": 15885, "max": 652806, "range": 636921},
        "support": {"count": 7244, "avg": 58400.48, "min": 15916, "max": 608157, "range": 592241},
    }
    fallback_total = 50000
    fallback_top = ("engineering", 58991.61)
    try:
        c = psycopg2.connect(**DB_CONFIG)
        cur = c.cursor()
        cur.execute('''
            SELECT "DEPARTMENT", COUNT(*), AVG("SALARY"), MIN("SALARY"), MAX("SALARY"),
                   MAX("SALARY") - MIN("SALARY")
            FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"
            GROUP BY "DEPARTMENT"
        ''')
        derived = {}
        for r in cur.fetchall():
            key = str(r[0]).strip().lower()
            derived[key] = {
                "count": int(r[1]),
                "avg":   round(float(r[2]), 2),
                "min":   int(float(r[3])),
                "max":   int(float(r[4])),
                "range": int(float(r[5])),
            }
        cur.execute('SELECT COUNT(*) FROM sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES"')
        total = int(cur.fetchone()[0])
        # Highest avg dept
        top = max(derived.items(), key=lambda x: x[1]["avg"])
        cur.close()
        c.close()
        return derived, total, (top[0], top[1]["avg"])
    except Exception as e:
        print(f"[WARN] Could not derive expected depts from DB: {e}; using fallback")
        return fallback, fallback_total, fallback_top


EXPECTED_DEPTS, TOTAL_EMPLOYEES, (TOP_DEPT_KEY, TOP_DEPT_AVG) = _derive_expected_depts()


def num_close(a, b, tol=10.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    all_errors = []
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # ---- Check Notion Page ----
    print("  Checking Notion page...")
    cur.execute("SELECT id, properties FROM notion.pages")
    pages = cur.fetchall()
    found_page = False
    found_page_id = None
    target_phrase = "hr department workforce analysis"
    for page in pages:
        props = page[1] if isinstance(page[1], dict) else json.loads(page[1]) if page[1] else {}
        title_text = ""
        # Look at any title-typed property (not only one keyed 'title')
        if isinstance(props, dict):
            for k, v in props.items():
                if isinstance(v, dict) and v.get("type") == "title":
                    for item in v.get("title", []):
                        if isinstance(item, dict):
                            if "plain_text" in item:
                                title_text += item.get("plain_text", "")
                            elif isinstance(item.get("text"), dict):
                                title_text += item["text"].get("content", "")
        # Exact title match (task.md prescribes the exact page title)
        if titles_match(target_phrase, title_text):
            found_page = True
            found_page_id = page[0]
            break
    if not found_page:
        all_errors.append(
            f"Notion page titled 'HR Department Workforce Analysis' not found "
            f"(exact title match against '{target_phrase}')"
        )
    else:
        print("    Page found")

    # ---- Check Notion Database ----
    print("  Checking Notion database...")
    cur.execute("SELECT id, title, properties FROM notion.databases")
    databases = cur.fetchall()
    found_db = False
    db_id = None
    for db in databases:
        title_text = ""
        if db[1]:
            title_data = db[1] if isinstance(db[1], list) else json.loads(db[1]) if isinstance(db[1], str) else []
            for item in title_data:
                if isinstance(item, dict):
                    title_text += item.get("plain_text", "") or item.get("text", {}).get("content", "")
        if titles_match("Department Metrics", title_text):
            found_db = True
            db_id = db[0]
            break
    if not found_db:
        all_errors.append("Notion database 'Department Metrics' not found")
    else:
        print("    Database found")
        # Check database entries (pages with parent db)
        cur.execute("SELECT properties FROM notion.pages WHERE parent->>'database_id' = %s", (str(db_id),))
        db_pages = cur.fetchall()
        if len(db_pages) != 7:
            all_errors.append(f"Expected exactly 7 department entries, found {len(db_pages)}")
        else:
            print(f"    {len(db_pages)} entries found")

        # Validate per-department metric values
        def extract_property_value(props, key, vtype):
            """Extract notion property value: vtype = 'title' | 'number'"""
            if not isinstance(props, dict):
                return None
            for k, v in props.items():
                if k.lower() == key.lower() and isinstance(v, dict):
                    if vtype == "title":
                        items = v.get("title") or []
                        return "".join(
                            (it.get("text", {}) or {}).get("content", "")
                            or it.get("plain_text", "")
                            for it in items if isinstance(it, dict)
                        )
                    if vtype == "number":
                        return v.get("number")
            return None

        agent_dept_metrics = {}
        for (props_raw,) in db_pages:
            props = props_raw if isinstance(props_raw, dict) else (
                json.loads(props_raw) if props_raw else {}
            )
            dept_name = extract_property_value(props, "Department", "title") or ""
            dept_key = dept_name.strip().lower()
            if not dept_key:
                continue
            agent_dept_metrics[dept_key] = {
                "count": extract_property_value(props, "Employee_Count", "number"),
                "avg":   extract_property_value(props, "Avg_Salary", "number"),
                "min":   extract_property_value(props, "Min_Salary", "number"),
                "max":   extract_property_value(props, "Max_Salary", "number"),
                "range": extract_property_value(props, "Salary_Range", "number"),
            }

        for dept_key, expected in EXPECTED_DEPTS.items():
            if dept_key not in agent_dept_metrics:
                all_errors.append(f"Notion DB missing department '{dept_key}'")
                continue
            actual = agent_dept_metrics[dept_key]
            for field, exp_val in expected.items():
                act_val = actual.get(field)
                if act_val is None:
                    all_errors.append(f"Notion DB '{dept_key}'.{field}: missing")
                    continue
                if field == "count":
                    tol = 1
                elif field == "avg":
                    # Loosen avg tolerance: agent rounding to 2 decimals can
                    # legitimately differ by up to ~0.01; using 0.5 to handle
                    # possible cents/rounding-mode differences. We also accept
                    # rounding to integer (~$1) which the task spec does not
                    # forbid.
                    tol = 1.0
                elif field in ("min", "max", "range"):
                    # Tolerate $5 rounding for min/max (some DBs may store
                    # to cents while task spec doesn't pin precision).
                    tol = 5
                else:
                    tol = 1
                if not num_close(act_val, exp_val, tol):
                    all_errors.append(f"Notion DB '{dept_key}'.{field}: {act_val} vs {exp_val} (tol={tol})")

    # ---- Check paragraph block on the page (highest/lowest dept) ----
    if found_page_id:
        print("  Checking paragraph block...")
        cur.execute(
            "SELECT block_data FROM notion.blocks "
            "WHERE parent_id = %s AND type IN ('paragraph','rich_text','text')",
            (str(found_page_id),)
        )
        block_rows = cur.fetchall()
        # Concatenate all paragraph plain text on this page
        paragraph_text = ""
        for (bd,) in block_rows:
            try:
                bd_obj = bd if isinstance(bd, dict) else (json.loads(bd) if bd else {})
            except Exception:
                bd_obj = {}
            # Common notion shape: {"rich_text": [{"plain_text": "..."}, ...]}
            rt = (bd_obj.get("rich_text") if isinstance(bd_obj, dict) else None) or []
            if isinstance(rt, list):
                for item in rt:
                    if isinstance(item, dict):
                        paragraph_text += item.get("plain_text", "") or item.get("text", {}).get("content", "")
                paragraph_text += " "
        ptxt = paragraph_text.lower()
        # Find lowest-avg dept (deterministic from derived data)
        if EXPECTED_DEPTS:
            sorted_by_avg = sorted(EXPECTED_DEPTS.items(), key=lambda x: x[1]["avg"])
            lowest_dept = sorted_by_avg[0][0]
            highest_dept = sorted_by_avg[-1][0]
        else:
            lowest_dept = "operations"
            highest_dept = "engineering"
        if highest_dept not in ptxt or lowest_dept not in ptxt:
            all_errors.append(
                f"Paragraph block on page should mention highest avg dept '{highest_dept}' "
                f"AND lowest avg dept '{lowest_dept}' "
                f"(got paragraph excerpt: {ptxt[:200]!r})"
            )
        else:
            print("    PASS")

    # ---- Check Email ----
    print("  Checking email sent...")
    cur.execute("SELECT to_addr, subject, body_text, from_addr FROM email.messages")
    messages = cur.fetchall()
    found_email = False
    total_str = str(TOTAL_EMPLOYEES)
    total_str_c = f"{TOTAL_EMPLOYEES:,}"
    avg_int = int(round(TOP_DEPT_AVG))
    avg_int_str = str(avg_int)
    avg_int_str_c = f"{avg_int:,}"
    avg_full = f"{TOP_DEPT_AVG:.2f}"
    avg_full_c = f"{TOP_DEPT_AVG:,.2f}" if TOP_DEPT_AVG >= 1000 else avg_full
    for msg in messages:
        to = str(msg[0]).lower() if msg[0] else ""
        subj = str(msg[1]).lower() if msg[1] else ""
        body = str(msg[2]).lower() if msg[2] else ""
        from_addr = str(msg[3]).lower() if msg[3] else ""
        if "hr-director" in to and "workforce" in subj.lower():
            found_email = True
            # Required body content per task.md
            if total_str not in body and total_str_c not in body:
                all_errors.append(f"Email body should mention total of {TOTAL_EMPLOYEES} employees")
            if TOP_DEPT_KEY not in body:
                all_errors.append(f"Email body should mention {TOP_DEPT_KEY} as highest avg salary dept")
            # Validate avg salary value precisely (no '59,000' truncation accepted)
            avg_candidates = {avg_int_str, avg_int_str_c, avg_full, avg_full_c}
            # Accept the integer-rounded form OR the 2-decimal form, with optional comma separator
            if not any(ac in body for ac in avg_candidates):
                all_errors.append(
                    f"Email body should mention {TOP_DEPT_KEY} avg salary value "
                    f"(any of {avg_candidates})"
                )
            # NOTE: the sender address is fixed by the email environment and
            # cannot be changed by the agent, so we do not validate from_addr.
            break
    if not found_email:
        all_errors.append("Email to hr-director with workforce analysis subject not found")
    else:
        print("    PASS")

    cur.close()
    conn.close()

    if all_errors:
        print(f"\n=== RESULT: FAIL ({len(all_errors)} errors) ===")
        for e in all_errors[:15]:
            print(f"  {e}")
        sys.exit(1)
    else:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)


if __name__ == "__main__":
    main()
