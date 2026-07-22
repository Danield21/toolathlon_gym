"""Evaluation for canvas-faculty-workload-review."""
import argparse
import os
import sys

import openpyxl
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=2.0):
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower() == sheet_name.strip().lower():
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def get_expected_instructors():
    """Get expected instructor data from DB."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT u.name as instructor,
               COUNT(DISTINCT c.id) as courses_count,
               SUM(sc.cnt) as total_students,
               SUM(ac.cnt) as total_assignments,
               ROUND(SUM(sc.cnt) * 0.5, 1) as est_grading_hours
        FROM canvas.enrollments e
        JOIN canvas.users u ON u.id = e.user_id
        JOIN canvas.courses c ON c.id = e.course_id
        LEFT JOIN (SELECT course_id, COUNT(*) as cnt FROM canvas.enrollments WHERE type='StudentEnrollment' GROUP BY course_id) sc ON sc.course_id = c.id
        LEFT JOIN (SELECT course_id, COUNT(*) as cnt FROM canvas.assignments GROUP BY course_id) ac ON ac.course_id = c.id
        WHERE e.type = 'TeacherEnrollment'
        GROUP BY u.name
        ORDER BY u.name
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return rows


def check_excel(agent_workspace):
    print("\n=== Checking Excel ===")
    xlsx_path = os.path.join(agent_workspace, "Faculty_Workload.xlsx")
    if not os.path.isfile(xlsx_path):
        check("Faculty_Workload.xlsx exists", False, f"Not found: {xlsx_path}")
        return
    check("Faculty_Workload.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        check("Excel readable", False, str(e))
        return
    check("Excel readable", True)

    expected = get_expected_instructors()

    # Instructor Load sheet
    il_rows = load_sheet_rows(wb, "Instructor Load")
    if il_rows is None:
        check("Sheet 'Instructor Load' exists", False, f"Available: {wb.sheetnames}")
    else:
        check("Sheet 'Instructor Load' exists", True)
        data_rows = il_rows[1:] if len(il_rows) > 1 else []
        check(f"Instructor Load has {len(expected)} rows",
              abs(len(data_rows) - len(expected)) <= 1,
              f"Found {len(data_rows)}")

        header = il_rows[0] if il_rows else []
        header_lower = [str(h).lower().replace(" ", "_") if h else "" for h in header]
        for col in ["instructor", "courses_count", "total_students", "overloaded_yn"]:
            check(f"Column '{col}' present", any(col in h for h in header_lower),
                  f"Header: {header}")

        # Build column-index map
        def col_idx(target):
            for i, h in enumerate(header_lower):
                if target in h:
                    return i
            return None

        i_instr = col_idx("instructor")
        i_courses = col_idx("courses_count") if col_idx("courses_count") is not None else col_idx("course")
        i_students = col_idx("total_students") if col_idx("total_students") is not None else col_idx("student")
        i_overload = col_idx("overloaded") if col_idx("overloaded") is not None else (len(header_lower) - 1)

        # Build agent map: instructor_name -> row data
        agent_map = {}
        if i_instr is not None:
            for row in data_rows:
                if row and i_instr < len(row) and row[i_instr]:
                    agent_map[str(row[i_instr]).strip().lower()] = row

        # Validate EVERY expected instructor (not just 2 spot checks)
        for inst_row in expected:
            instr_name, courses_count, total_students, total_assignments, est_hrs = inst_row
            key = str(instr_name).strip().lower()
            actual_row = agent_map.get(key)
            if actual_row is None:
                check(f"Instructor '{instr_name}' present", False,
                      f"agent rows: {list(agent_map.keys())[:10]}")
                continue
            check(f"Instructor '{instr_name}' present", True)
            # courses_count
            if i_courses is not None and i_courses < len(actual_row):
                check(f"'{instr_name}' courses_count == {courses_count}",
                      num_close(actual_row[i_courses], courses_count, 0),
                      f"Got {actual_row[i_courses]}")
            # total_students
            if i_students is not None and i_students < len(actual_row):
                check(f"'{instr_name}' total_students == {total_students}",
                      num_close(actual_row[i_students], total_students, 5),
                      f"Got {actual_row[i_students]}")
            # overloaded computed
            weekly_hours = float(total_students) * 0.5 / 16
            expected_overloaded = (courses_count > 4) or (weekly_hours > 40)
            if i_overload is not None and i_overload < len(actual_row):
                actual_y = str(actual_row[i_overload]).strip().lower() in ("yes", "y", "true", "1")
                check(f"'{instr_name}' overloaded_yn == {'Yes' if expected_overloaded else 'No'}",
                      actual_y == expected_overloaded,
                      f"Got '{actual_row[i_overload]}'; weekly_hours={weekly_hours:.1f}")

    # Department Summary sheet
    ds_rows = load_sheet_rows(wb, "Department Summary")
    if ds_rows is None:
        check("Sheet 'Department Summary' exists", False, f"Available: {wb.sheetnames}")
    else:
        check("Sheet 'Department Summary' exists", True)
        data_rows = ds_rows[1:] if len(ds_rows) > 1 else []
        # Department aggregation depends on subject-area derivation; allow 5-9 range
        check("Department Summary has 5-9 departments (subject area)",
              5 <= len(data_rows) <= 9,
              f"Found {len(data_rows)}")
        # Verify required columns
        if ds_rows:
            ds_header = [str(h).lower().replace(" ", "_") if h else "" for h in ds_rows[0]]
            for col in ["department", "instructor_count", "course_count", "total_students"]:
                check(f"Department Summary has '{col}'",
                      any(col in h for h in ds_header), f"Got: {ds_rows[0]}")


def check_pptx(agent_workspace):
    print("\n=== Checking PowerPoint ===")
    pptx_path = os.path.join(agent_workspace, "Workload_Review.pptx")
    if not os.path.isfile(pptx_path):
        check("Workload_Review.pptx exists", False, f"Not found: {pptx_path}")
        return
    check("Workload_Review.pptx exists", True)

    try:
        from pptx import Presentation
        prs = Presentation(pptx_path)
        check("PPTX has at least 3 slides", len(prs.slides) >= 3,
              f"Found {len(prs.slides)} slides")
        all_text = " ".join(
            shape.text for slide in prs.slides for shape in slide.shapes if shape.has_text_frame
        ).lower()
        check("PPTX mentions workload or overloaded",
              "workload" in all_text or "overload" in all_text,
              f"Sample: {all_text[:200]}")
    except ImportError:
        check("python-pptx available", False)


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("""
            SELECT id, title FROM gsheet.spreadsheets
            WHERE title ILIKE '%%workload%%' OR title ILIKE '%%faculty%%'
        """)
        sheets = cur.fetchall()
        check("Google Sheet created for workload data", len(sheets) >= 1,
              "No matching spreadsheet found")
        if sheets:
            sid = sheets[0][0]
            cur.execute("SELECT COUNT(*) FROM gsheet.cells WHERE spreadsheet_id = %s", (sid,))
            cell_count = cur.fetchone()[0]
            check("Google Sheet has data (cells)", cell_count > 10,
                  f"Found {cell_count} cells")
        cur.close()
        conn.close()
    except Exception as e:
        check("Google Sheet check", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace)
    check_pptx(args.agent_workspace)
    check_gsheet()

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
