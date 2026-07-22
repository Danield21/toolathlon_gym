"""
Check that the agent created the Google Sheet with correct grade summary data.
Queries the gsheet schema in PostgreSQL.
"""

import os
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

# Expected data: Course_Code, Class_Average, Letter_Grade, Distinction, Probation
EXPECTED_ROWS = [
    ("BBB-2014B", "77.35", "C", "No", "No"),
    ("CCC-2014B", "62.77", "D", "No", "No"),
    ("DDD-2014B", "66.47", "D", "No", "No"),
    ("EEE-2014B", "78.80", "C", "No", "No"),
    ("FFF-2014B", "74.67", "C", "No", "No"),
    ("GGG-2014B", "77.38", "C", "No", "No"),
]

EXPECTED_HEADERS = ["Course_Code", "Class_Average", "Letter_Grade", "Distinction", "Probation"]


def num_close(a, b, tol=0.5):
    """Compare two numeric values with tolerance."""
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    """Case-insensitive comparison."""
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return str(a).strip().lower() == str(b).strip().lower()


def check_gsheet():
    """
    Verify the Google Sheet was created with correct data.
    Returns (passed_count, failed_count, error_details).
    """
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    passed = 0
    failed = 0
    errors = []

    # Check that a spreadsheet with the expected title exists
    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%spring 2014%grade%'
           OR LOWER(title) LIKE '%grade%spring 2014%'
    """)
    spreadsheets = cur.fetchall()

    if not spreadsheets:
        cur.close()
        conn.close()
        return 0, 1, ["No Google Sheet found with 'Spring 2014' and 'Grade' in title"]

    ss_id = spreadsheets[0][0]
    ss_title = spreadsheets[0][1]
    passed += 1
    print(f"  [check_gsheet] Found spreadsheet: '{ss_title}' (id={ss_id})")

    # Check that a sheet named "Grades" exists
    cur.execute("""
        SELECT id, title FROM gsheet.sheets
        WHERE spreadsheet_id = %s AND LOWER(title) = 'grades'
    """, (ss_id,))
    sheets = cur.fetchall()

    if not sheets:
        failed += 1
        errors.append("No sheet named 'Grades' found in the spreadsheet")
        cur.close()
        conn.close()
        return passed, failed, errors

    sheet_id = sheets[0][0]
    passed += 1

    # Fetch all cells for this sheet
    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (ss_id, sheet_id))

    cells = {}
    for row_idx, col_idx, value in cur.fetchall():
        cells[(row_idx, col_idx)] = value

    cur.close()
    conn.close()

    if not cells:
        failed += 1
        errors.append("No cells found in the Grades sheet")
        return passed, failed, errors

    max_row = max(r for r, c in cells.keys())
    max_col = max(c for r, c in cells.keys())

    header_vals = {}
    for col_idx in range(0, max_col + 1):
        header_vals[col_idx] = str(cells.get((0, col_idx), "") or "").strip().lower()

    # NEW: Validate header order and exact values (case-insensitive, _/space interchangeable)
    expected_lower = [h.lower().replace("_", " ") for h in EXPECTED_HEADERS]
    actual_headers_ordered = [header_vals.get(i, "").replace("_", " ").strip() for i in range(len(EXPECTED_HEADERS))]
    header_order_ok = actual_headers_ordered == expected_lower
    if not header_order_ok:
        failed += 1
        errors.append(
            f"GSheet header order/exact mismatch: expected {EXPECTED_HEADERS}, got {actual_headers_ordered}"
        )
        return passed, failed, errors
    passed += 1

    col_indices = {}
    for exp_header in EXPECTED_HEADERS:
        key = exp_header.lower().replace("_", " ")
        for col_idx, hv in header_vals.items():
            hv_norm = hv.replace("_", " ")
            if hv_norm == key or key == hv_norm.strip():
                col_indices[exp_header] = col_idx
                break

    missing_headers = [h for h in EXPECTED_HEADERS if h not in col_indices]
    if missing_headers:
        failed += 1
        errors.append(f"GSheet missing headers: {missing_headers}. Found: {list(header_vals.values())}")
        return passed, failed, errors
    passed += 1

    data_rows = {}
    for row_idx in range(1, max_row + 1):
        code_val = cells.get((row_idx, col_indices["Course_Code"]), "")
        if code_val and "2014B" in str(code_val):
            data_rows[str(code_val).strip()] = {h: cells.get((row_idx, col_indices[h]), "") for h in EXPECTED_HEADERS}

    for exp_code, exp_avg, exp_grade, exp_dist, exp_prob in EXPECTED_ROWS:
        if exp_code not in data_rows:
            failed += 1
            errors.append(f"GSheet: Course {exp_code} not found")
            continue

        row_data = data_rows[exp_code]
        row_ok = True
        if not num_close(row_data["Class_Average"], exp_avg, 0.5):
            row_ok = False
            errors.append(f"GSheet {exp_code}: class average expected {exp_avg}, got {row_data['Class_Average']}")
        if not str_match(row_data["Letter_Grade"], exp_grade):
            row_ok = False
            errors.append(f"GSheet {exp_code}: letter grade expected {exp_grade}, got {row_data['Letter_Grade']}")
        if not str_match(row_data["Distinction"], exp_dist):
            row_ok = False
            errors.append(f"GSheet {exp_code}: distinction expected {exp_dist}, got {row_data['Distinction']}")
        if not str_match(row_data["Probation"], exp_prob):
            row_ok = False
            errors.append(f"GSheet {exp_code}: probation expected {exp_prob}, got {row_data['Probation']}")

        if row_ok:
            passed += 1
        else:
            failed += 1

    return passed, failed, errors
