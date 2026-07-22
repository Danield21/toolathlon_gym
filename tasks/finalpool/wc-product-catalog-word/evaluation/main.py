"""Evaluation for wc-product-catalog-word."""
import argparse
import os
import re
import sys
from docx import Document


# Fallback static lists if PDF parsing or DB query fails. These mirror the
# current catalog and store snapshot.
FALLBACK_MATCHING = [
    "Canon M50 Mark II",
    "JBL Flip 4 Portable Wireless Speaker",
    "Boult Audio Powerbuds",
    "AmazonBasics Expanding File Folder",
    "CraftDev A500MB Portable Paper Trimmer",
    "AGARO Adjustable Camera Tripod Stand",
    "Ambrane Mobile Holding Tabletop Stand",
]

FALLBACK_MISSING = [
    "Belkin Ultra HD High Speed HDMI Cable",
    "Sony WH-1000XM5 Wireless Headphones",
    "Logitech MX Master 3S Mouse",
    "Samsung T7 Shield Portable SSD 1TB",
    "Anker PowerCore 26800mAh Battery Pack",
]

TOTAL_CATALOG_FALLBACK = 12


def extract_catalog_from_pdf(pdf_path):
    """Parse the supplier catalog PDF and return a list of product names.
    Returns None if parsing fails."""
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            r = PyPDF2.PdfReader(f)
            text = "\n".join(p.extract_text() or "" for p in r.pages)
        # Parse lines like "<num>\n<product name>\n$<price>"
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        products = []
        i = 0
        while i < len(lines):
            # Look for line that is only a number (catalog index)
            if re.fullmatch(r"\d+", lines[i]) and i + 2 < len(lines):
                next_line = lines[i + 1]
                price_line = lines[i + 2]
                # Verify price_line begins with $
                if re.match(r"\$\s*[0-9]", price_line):
                    products.append(next_line)
                    i += 3
                    continue
            i += 1
        return products if products else None
    except Exception:
        return None


def compute_matching_missing(catalog, store_products):
    """Given catalog names and list of store product names, partition the catalog
    into matching (substring of any store name, case-insensitive) and missing."""
    store_lower = [p.lower() for p in store_products]
    matching = []
    missing = []
    for item in catalog:
        needle = item.lower()
        if any(needle in sname for sname in store_lower):
            matching.append(item)
        else:
            missing.append(item)
    return matching, missing


def get_store_products():
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=os.environ.get("PGHOST", "localhost"),
            port=int(os.environ.get("PGPORT", "5432")),
            dbname="toolathlon_gym",
            user="eigent",
            password="camel",
        )
        cur = conn.cursor()
        cur.execute("SELECT name FROM wc.products WHERE status = 'publish' OR status IS NULL")
        rows = [r[0] for r in cur.fetchall() if r[0]]
        cur.close()
        conn.close()
        return rows or None
    except Exception:
        return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    all_errors = []

    # Dynamically compute expected matching/missing sets from the supplier PDF
    # and the current store. Fall back to static snapshot if parsing fails.
    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    pdf_candidates = [
        os.path.join(args.agent_workspace or ".", "Supplier_Catalog.pdf"),
        os.path.join(task_root, "initial_workspace", "Supplier_Catalog.pdf"),
    ]
    catalog = None
    for p in pdf_candidates:
        if os.path.exists(p):
            catalog = extract_catalog_from_pdf(p)
            if catalog:
                break
    store = get_store_products()
    if catalog and store:
        matching_items, missing_items = compute_matching_missing(catalog, store)
        total_catalog = len(catalog)
    else:
        matching_items = FALLBACK_MATCHING
        missing_items = FALLBACK_MISSING
        total_catalog = TOTAL_CATALOG_FALLBACK
    print(f"  Catalog size = {total_catalog}, matching = {len(matching_items)}, missing = {len(missing_items)}")

    agent_doc = os.path.join(args.agent_workspace, "Product_Comparison.docx")
    if not os.path.exists(agent_doc):
        print("FAIL: Agent output Product_Comparison.docx not found")
        sys.exit(1)

    doc = Document(agent_doc)

    # Extract all text
    all_text = ""
    headings = []
    paragraphs_by_section = {}
    current_section = ""
    for para in doc.paragraphs:
        text = para.text.strip()
        all_text += text.lower() + " "
        if para.style.name.startswith("Heading"):
            current_section = text.lower()
            headings.append(text.lower())
            paragraphs_by_section[current_section] = []
        elif current_section and text:
            paragraphs_by_section.setdefault(current_section, []).append(text)

    # Check document structure
    print("  Checking document structure...")
    # headings already filtered by Heading style in the extraction loop
    has_matching = any("matching" in h and "product" in h for h in headings)
    has_missing = any("missing" in h and "product" in h for h in headings)
    has_summary = any("summary" in h for h in headings)

    if not has_matching:
        all_errors.append("Missing 'Matching Products' heading")
    if not has_missing:
        all_errors.append("Missing 'Missing Products' heading")
    if not has_summary:
        all_errors.append("Missing 'Summary' heading")

    # Check matching products section
    print("  Checking matching products...")
    matching_text = ""
    for key, paras in paragraphs_by_section.items():
        if "matching" in key:
            matching_text = " ".join(paras).lower()
            break

    found_matching = 0
    def _norm(s):
        return " ".join(str(s).lower().split())
    norm_matching_text = _norm(matching_text)
    norm_all_text = _norm(all_text)
    for product in matching_items:
        p = _norm(product)
        if p in norm_matching_text:
            found_matching += 1
        else:
            if p in norm_all_text:
                found_matching += 1
            else:
                all_errors.append(f"Matching product not found: {product}")

    # Require at least N-1 of matching products to be listed (allow small OCR slip)
    min_matching = max(1, len(matching_items) - 1)
    if found_matching < min_matching:
        all_errors.append(f"Only {found_matching}/{len(matching_items)} matching products found (need >= {min_matching})")
    else:
        print(f"    {found_matching}/{len(matching_items)} matching products found")

    # Check missing products section
    print("  Checking missing products...")
    missing_text = ""
    for key, paras in paragraphs_by_section.items():
        if "missing" in key:
            missing_text = " ".join(paras).lower()
            break

    found_missing = 0
    for product in missing_items:
        if product.lower() in missing_text or product.lower() in all_text:
            found_missing += 1
        else:
            all_errors.append(f"Missing product not listed: {product}")

    min_missing = max(1, len(missing_items) - 1)
    if found_missing < min_missing:
        all_errors.append(f"Only {found_missing}/{len(missing_items)} missing products found (need >= {min_missing})")
    else:
        print(f"    {found_missing}/{len(missing_items)} missing products found")

    # Ensure Matching section does NOT list any of the missing products (false positives)
    for product in missing_items:
        if product.lower() in norm_matching_text:
            all_errors.append(f"Missing product '{product}' incorrectly listed in Matching Products section")

    # Check summary section
    print("  Checking summary...")
    summary_text = ""
    for key, paras in paragraphs_by_section.items():
        if "summary" in key:
            summary_text = " ".join(paras).lower()
            break

    if not summary_text:
        summary_text = all_text  # fallback

    # Summary must mention total catalog count, matching count, and missing count
    for tok, label in [(str(total_catalog), "total catalog"),
                        (str(len(matching_items)), "matching"),
                        (str(len(missing_items)), "missing")]:
        if not re.search(r"\b" + re.escape(tok) + r"\b", summary_text):
            all_errors.append(f"Summary should mention {tok} ({label})")

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
