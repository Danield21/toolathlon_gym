"""
Evaluation script for playwright-canvas-curriculum-word-notion task.

Checks:
1. Word document Accreditation_Compliance_Report.docx with course evaluations
2. Notion database "Course Compliance Tracker" with 22 entries
3. Emails to departments with non-compliant courses
"""

import argparse
import json
import os
import sys

import psycopg2

try:
    from docx import Document
except ImportError:
    Document = None

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0

# Non-compliant departments based on standards (min 8 assignments, min 3 quizzes)
NON_COMPLIANT_DEPTS = [
    "Applied Analytics",
    "Biochemistry",
    "Data-Driven",
    "Environmental",
]


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def str_contains(haystack, needle):
    if haystack is None or needle is None:
        return False
    return needle.strip().lower() in str(haystack).strip().lower()


def check_word(agent_workspace):
    """Check Word compliance report."""
    print("\n=== Checking Word Document ===")

    doc_path = os.path.join(agent_workspace, "Accreditation_Compliance_Report.docx")
    if not os.path.isfile(doc_path):
        record("Accreditation_Compliance_Report.docx exists", False)
        return False

    record("Accreditation_Compliance_Report.docx exists", True)

    if Document is None:
        record("python-docx available", False, "Cannot import docx")
        return True

    try:
        doc = Document(doc_path)
        full_text = "\n".join(p.text for p in doc.paragraphs).lower()

        record(
            "Doc mentions accreditation/compliance",
            "accreditation" in full_text or "compliance" in full_text,
        )

        # Check course names mentioned (spot check)
        record(
            "Doc mentions Applied Analytics",
            "applied analytics" in full_text,
        )
        record(
            "Doc mentions Foundations of Finance",
            "foundations of finance" in full_text,
        )

        # Compliance status indicators: must show BOTH compliant and non-compliant
        # OR BOTH pass and fail indicators.
        has_compliant_pair = "compliant" in full_text and "non-compliant" in full_text
        has_pass_fail_pair = ("pass" in full_text or "passes" in full_text) and (
            "fail" in full_text or "fails" in full_text)
        record("Doc has compliance status indicators (both compliant and non-compliant, or pass/fail)",
               has_compliant_pair or has_pass_fail_pair)

        # Check summary statistics
        has_summary = any(
            kw in full_text for kw in ["summary", "total", "rate", "overall"]
        )
        record("Doc has summary section", has_summary)

        # Check mentions all 7 departments (task: "section for each course" of 22 courses)
        course_keywords = [
            "applied analytics", "biochemistry", "creative computing",
            "data-driven", "environmental", "foundations of finance",
            "global governance",
        ]
        mentioned = sum(1 for kw in course_keywords if kw in full_text)
        record(
            "Doc covers all 7 department names",
            mentioned == len(course_keywords),
            f"Found {mentioned}/7 department names",
        )

        return True
    except Exception as e:
        record("Word doc readable", False, str(e))
        return False


def check_notion():
    """Check Notion database for Course Compliance Tracker."""
    print("\n=== Checking Notion Database ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title, properties FROM notion.databases")
        dbs = cur.fetchall()

        found_db = None
        for db_id, title_raw, props in dbs:
            if isinstance(title_raw, list):
                title = " ".join(t.get("plain_text", "") for t in title_raw if isinstance(t, dict))
            elif isinstance(title_raw, str):
                try:
                    parsed = json.loads(title_raw)
                    if isinstance(parsed, list):
                        title = " ".join(t.get("plain_text", "") for t in parsed if isinstance(t, dict))
                    else:
                        title = title_raw
                except (json.JSONDecodeError, TypeError):
                    title = title_raw
            else:
                title = str(title_raw) if title_raw else ""

            title_lower = title.lower()
            if ("compliance" in title_lower or "course" in title_lower) and (
                "tracker" in title_lower or "compliance" in title_lower
            ):
                found_db = db_id
                break

        if not found_db:
            record(
                "Course Compliance Tracker database exists",
                False,
                f"Found {len(dbs)} databases",
            )
            cur.close()
            conn.close()
            return False

        record("Course Compliance Tracker database exists", True)

        # Check entries - exactly 22 (per task.md)
        cur.execute(
            "SELECT id, properties FROM notion.pages WHERE parent::text LIKE %s",
            (f'%{found_db}%',),
        )
        pages = cur.fetchall()

        record(
            "Database has == 22 entries",
            len(pages) == 22,
            f"Found {len(pages)} entries",
        )

        # Validate all 6 required properties (Department, Assignment_Count, Quiz_Count,
        # Assignments_Compliant, Quizzes_Compliant, Overall_Status, Follow_Up_Date) plus title
        cur.execute("SELECT properties FROM notion.databases WHERE id = %s", (found_db,))
        db_props_row = cur.fetchone()
        if db_props_row and db_props_row[0]:
            db_props = db_props_row[0]
            if isinstance(db_props, str):
                db_props = json.loads(db_props)
            prop_names = [k.lower() for k in db_props.keys()] if isinstance(db_props, dict) else []
            for required in ["course_name", "department", "assignment_count", "quiz_count",
                             "assignments_compliant", "quizzes_compliant", "overall_status",
                             "follow_up_date"]:
                # Allow flexible naming via substring of meaningful tokens
                core = required.replace("_", "")
                ok = any(p.replace("_", "").replace(" ", "") == core for p in prop_names)
                if not ok:
                    # Looser fallback: any prop containing core key tokens
                    ok = any(required in p.replace(" ", "_") for p in prop_names)
                record(f"Notion DB has '{required}' property", ok, f"props: {prop_names}")

        # Compliant / Non-Compliant + Follow_Up_Date checks
        compliant_count = 0
        non_compliant_count = 0
        nc_with_followup = 0
        nc_without_followup = 0
        c_with_followup = 0
        for page_id, page_props in pages:
            if isinstance(page_props, str):
                page_props = json.loads(page_props)
            if not isinstance(page_props, dict):
                continue
            status = None
            follow_up = None
            for key, val in page_props.items():
                kl = key.lower()
                if ("status" in kl or "overall" in kl) and isinstance(val, dict):
                    select_val = val.get("select")
                    if isinstance(select_val, dict):
                        status = (select_val.get("name", "") or "").lower()
                if "follow" in kl and "date" in kl:
                    if isinstance(val, dict):
                        d = val.get("date")
                        if isinstance(d, dict):
                            follow_up = d.get("start")
                        elif isinstance(d, str):
                            follow_up = d
            if status:
                if "non" in status:
                    non_compliant_count += 1
                    if follow_up and "2026-04-15" in str(follow_up):
                        nc_with_followup += 1
                    else:
                        nc_without_followup += 1
                elif "compliant" in status:
                    compliant_count += 1
                    if follow_up:
                        c_with_followup += 1

        record(
            "Mix of compliant and non-compliant entries",
            compliant_count > 0 and non_compliant_count > 0,
            f"Compliant: {compliant_count}, Non-Compliant: {non_compliant_count}",
        )
        record(
            "Compliant + Non-Compliant counts sum to 22",
            compliant_count + non_compliant_count == 22,
            f"sum {compliant_count + non_compliant_count}",
        )
        record(
            "All non-compliant entries have Follow_Up_Date == 2026-04-15",
            non_compliant_count > 0 and nc_with_followup == non_compliant_count,
            f"{nc_with_followup}/{non_compliant_count}",
        )
        record(
            "Compliant entries have empty Follow_Up_Date",
            c_with_followup == 0,
            f"{c_with_followup} compliant entries unexpectedly have a follow-up date",
        )

        cur.close()
        conn.close()
        return True

    except Exception as e:
        record("Notion DB accessible", False, str(e))
        return False


def check_emails():
    """Check department notification emails."""
    print("\n=== Checking Emails ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute(
            "SELECT subject, from_addr, to_addr, body_text FROM email.messages"
        )
        emails = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Email DB accessible", False, str(e))
        return False

    # Find accreditation review emails sent from accreditation@university.edu
    # to department-review@university.edu with subject 'Accreditation Review:'
    review_emails = []
    for subject, from_addr, to_addr, body_text in emails:
        subj_lower = (subject or "").lower()
        from_lower = (from_addr or "").lower()
        to_lower = (to_addr or "").lower()
        if ("accreditation review" in subj_lower
                and "accreditation@university.edu" in from_lower
                and "department-review@university.edu" in to_lower):
            review_emails.append((subject, from_addr, to_addr, body_text))

    record(
        "At least 4 accreditation review emails (one per non-compliant dept)",
        len(review_emails) >= 4,
        f"Found {len(review_emails)} review emails",
    )

    # Each unique non-compliant dept must have its own email (subject contains dept)
    depts_with_email = set()
    for subj, _, _, _ in review_emails:
        sl = (subj or "").lower()
        for dept_kw in NON_COMPLIANT_DEPTS:
            if dept_kw.lower() in sl:
                depts_with_email.add(dept_kw)
    record(
        "All 4 non-compliant departments have a dedicated email",
        len(depts_with_email) == len(NON_COMPLIANT_DEPTS),
        f"Found {sorted(depts_with_email)}; expected {NON_COMPLIANT_DEPTS}",
    )

    # Check body mentions non-compliance details (assignment, quiz, etc.)
    all_body = " ".join((e[3] or "").lower() for e in review_emails)
    all_subjects = " ".join((e[0] or "").lower() for e in review_emails)
    combined = all_body + " " + all_subjects

    record("Email body mentions assignments",
           "assignment" in combined)
    record("Email body mentions quizzes",
           "quiz" in combined)

    return len(review_emails) >= 4


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    word_ok = check_word(args.agent_workspace)
    notion_ok = check_notion()
    email_ok = check_emails()

    print(f"\n=== SUMMARY ===")
    print(f"  Word:   {'PASS' if word_ok else 'FAIL'}")
    print(f"  Notion: {'PASS' if notion_ok else 'FAIL'}")
    print(f"  Email:  {'PASS' if email_ok else 'FAIL'}")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}")

    overall = word_ok and notion_ok and email_ok
    print(f"  Overall:  {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
