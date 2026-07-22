"""Evaluation for sf-sales-region-pdf-gsheet."""
import argparse
import os
import sys

import psycopg2

DB = {"host": os.environ.get("PGHOST", "localhost"), "port": 5432, "dbname": "toolathlon_gym", "user": "eigent", "password": "camel"}


def num_close(a, b, tol=1.0):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def extract_pdf_text(path):
    """Extract text from PDF using available libraries."""
    for mod_name, klass in [
        ("pypdf", "PdfReader"),
        ("PyPDF2", "PdfReader"),
    ]:
        try:
            mod = __import__(mod_name)
            cls = getattr(mod, klass)
            reader = cls(path)
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""
            return text
        except (ImportError, AttributeError):
            continue
    try:
        import pdfplumber
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    except ImportError:
        pass
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(path)
        text = ""
        for page in doc:
            text += page.get_text() or ""
        doc.close()
        return text
    except ImportError:
        pass
    with open(path, "rb") as f:
        return f.read().decode("latin-1", errors="ignore")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    all_errors = []

    # --- Check PDF ---
    agent_pdf = os.path.join(args.agent_workspace, "Regional_Sales_Report.pdf")
    if not os.path.exists(agent_pdf):
        print(f"FAIL: Agent output not found: {agent_pdf}")
        sys.exit(1)

    print("  Checking Regional_Sales_Report.pdf...")
    text = extract_pdf_text(agent_pdf).lower()

    # Check title
    if "regional sales report" not in text:
        all_errors.append("PDF missing title 'Regional Sales Report'")

    # Check regions present
    expected_regions = ["asia pacific", "europe", "latin america", "middle east", "north america"]
    for region in expected_regions:
        if region not in text:
            all_errors.append(f"PDF missing region: {region}")

    # Check summary section
    if "summary" not in text:
        all_errors.append("PDF missing 'Summary' section")

    # Validate data against DB
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT c."REGION",
          COUNT(DISTINCT o."ORDER_ID") as order_count,
          ROUND(SUM(o."TOTAL_AMOUNT")::numeric, 2) as total_revenue
        FROM sf_data."SALES_DW__PUBLIC__ORDERS" o
        JOIN sf_data."SALES_DW__PUBLIC__CUSTOMERS" c ON o."CUSTOMER_ID" = c."CUSTOMER_ID"
        WHERE o."STATUS" = 'Delivered'
        GROUP BY c."REGION"
        ORDER BY c."REGION"
    """)
    db_regions = cur.fetchall()
    conn.close()

    grand_total = sum(float(r[2]) for r in db_regions)
    top_region = max(db_regions, key=lambda x: float(x[2]))

    # Format-tolerant numeric matching: PDF text may have commas/$ signs
    text_normalized = text.replace(",", "").replace("$", "")

    # Check that order counts appear in the PDF
    for r in db_regions:
        region_name = r[0]
        order_count = str(r[1])
        if order_count not in text_normalized:
            all_errors.append(f"PDF missing order count {order_count} for {region_name}")

    # Check total orders
    total_orders = str(sum(r[1] for r in db_regions))
    if total_orders not in text_normalized:
        all_errors.append(f"PDF missing total orders: {total_orders}")

    # Check revenue per region appears (format-tolerant)
    for r in db_regions:
        region_name = r[0]
        revenue_int = str(int(round(float(r[2]))))
        # Look for either exact 2-decimal or integer match
        revenue_dec = f"{float(r[2]):.2f}".replace(",", "")
        if revenue_dec not in text_normalized and revenue_int not in text_normalized:
            all_errors.append(f"PDF missing revenue ~{revenue_dec} for {region_name}")

    # Check total revenue appears
    grand_total_int = str(int(round(grand_total)))
    grand_total_dec = f"{grand_total:.2f}".replace(",", "")
    if grand_total_dec not in text_normalized and grand_total_int not in text_normalized:
        all_errors.append(f"PDF missing grand total revenue ~{grand_total_dec}")

    # Check top region mentioned
    if top_region[0].lower() not in text:
        all_errors.append(f"PDF missing top region: {top_region[0]}")

    if not all_errors:
        print("    PASS")

    # --- BLOCKING: Check Google Sheet in DB ---
    print("  Checking Google Sheet...")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("""
            SELECT id, title FROM gsheet.spreadsheets
            WHERE LOWER(title) = LOWER('Regional Sales Dashboard')
               OR LOWER(title) LIKE '%regional sales dashboard%'
        """)
        sheets = cur.fetchall()
        if not sheets:
            all_errors.append("Google Sheet 'Regional Sales Dashboard' not found")
        else:
            sid = sheets[0][0]
            cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s AND LOWER(title) = 'overview'", (sid,))
            overview = cur.fetchall()
            if not overview:
                all_errors.append("Google Sheet 'Overview' tab not found")
            else:
                tab_id = overview[0][0]
                cur.execute("SELECT COUNT(*) FROM gsheet.cells WHERE sheet_id = %s", (tab_id,))
                cell_count = cur.fetchone()[0]
                # 5 regions * 6 cols = 30 + header 6 = 36 minimum cells expected
                if cell_count < 30:
                    all_errors.append(f"Google Sheet 'Overview' has only {cell_count} cells (expected >=30)")
        conn.close()
    except Exception as e:
        all_errors.append(f"GSheet DB check error: {e}")

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
