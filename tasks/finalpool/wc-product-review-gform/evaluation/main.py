"""Evaluation for wc-product-review-gform."""
import argparse
import json
import os
import sys

import psycopg2

def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


DB = dict(host=os.environ.get("PGHOST", "localhost"), port=int(os.environ.get("PGPORT", "5432")), dbname="toolathlon_gym", user="eigent", password="camel")
PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1; print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1; print(f"  [FAIL] {name}: {str(detail)[:300]}")


def get_expected():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""SELECT id, name, ROUND(average_rating::numeric,2), total_sales
        FROM wc.products WHERE average_rating IS NOT NULL AND average_rating::numeric > 0
        ORDER BY average_rating ASC, total_sales DESC LIMIT 5""")
    products = [{"id": r[0], "name": r[1], "rating": float(r[2]), "sales": r[3]} for r in cur.fetchall()]
    product_ids = [p["id"] for p in products]

    placeholders = ",".join(["%s"] * len(product_ids))
    cur.execute(f"""SELECT DISTINCT o.customer_id, c.email, c.first_name, c.last_name
        FROM wc.orders o JOIN wc.customers c ON c.id=o.customer_id,
        LATERAL jsonb_array_elements(o.line_items) AS item
        WHERE (item->>'product_id')::int IN ({placeholders}) AND o.customer_id > 0
        ORDER BY o.customer_id""", product_ids)
    customers = [{"id": r[0], "email": r[1], "first_name": r[2], "last_name": r[3]} for r in cur.fetchall()]
    conn.close()
    return {"products": products, "customers": customers}


def check_gform(expected):
    print("\n=== Checking Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title, description FROM gform.forms")
    forms = cur.fetchall()
    record("At least 1 form created", len(forms) >= 1, f"Found {len(forms)}")

    target_form = None
    for fid, title, desc in forms:
        if "product" in (title or "").lower() or "quality" in (title or "").lower() or "feedback" in (title or "").lower():
            target_form = fid
            break
    if target_form is None and forms:
        target_form = forms[0][0]

    if target_form is None:
        conn.close()
        return

    record("Form titled with product/quality/feedback",
           any("product" in (f[1] or "").lower() or "quality" in (f[1] or "").lower() or "feedback" in (f[1] or "").lower() for f in forms))

    cur.execute("SELECT title, question_type, required FROM gform.questions WHERE form_id=%s ORDER BY position", (target_form,))
    questions = cur.fetchall()
    record("Exactly 5 questions", len(questions) == 5, f"Found {len(questions)}")

    # Each question must be required
    record("All 5 questions are required",
           all(q[2] is True for q in questions),
           f"Required flags: {[q[2] for q in questions]}")

    # All 5 product names must appear (one per question)
    matched = []
    for prod in expected["products"]:
        prod_name_low = (prod["name"] or "").lower()
        # Try longest-substring match
        for q_title, q_type, q_req in questions:
            ql = (q_title or "").lower()
            # Require reasonably unique name substring of >=15 chars (or full if shorter)
            key = prod_name_low[:25] if len(prod_name_low) > 25 else prod_name_low
            if key and key in ql:
                matched.append(prod["name"])
                break
    record(f"All 5 lowest-rated product names in questions",
           len(matched) == 5,
           f"Matched: {matched}")

    cur.close()
    conn.close()


def check_emails(expected):
    print("\n=== Checking Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("Emails sent", len(emails) >= len(expected["customers"]),
           f"Found {len(emails)}, expected >= {len(expected['customers'])}")

    # Check that emails were sent to affected customers
    all_to = []
    for subj, to, body in emails:
        if isinstance(to, list):
            all_to.extend([str(r).lower() for r in to])
        elif isinstance(to, str):
            try:
                parsed = json.loads(to)
                if isinstance(parsed, list):
                    all_to.extend([str(r).lower() for r in parsed])
                else:
                    all_to.append(to.lower())
            except:
                all_to.append(to.lower())

    matched_customers = 0
    for cust in expected["customers"]:
        if cust["email"].lower() in " ".join(all_to):
            matched_customers += 1

    record("Emails to 100% of affected customers",
           matched_customers == len(expected["customers"]),
           f"Matched {matched_customers}/{len(expected['customers'])}")

    # Every email must have exact subject + body-with-first-name
    target_subject = "we value your feedback - product quality survey"
    subject_matches = 0
    first_name_matches = 0

    # Build customer email -> first_name map
    cust_map = {c["email"].lower(): c["first_name"] for c in expected["customers"]}

    for subj, to, body in emails:
        # Parse to_addr
        to_list = []
        if isinstance(to, list):
            to_list = [str(r).lower() for r in to]
        elif isinstance(to, str):
            try:
                p = json.loads(to)
                if isinstance(p, list):
                    to_list = [str(r).lower() for r in p]
                else:
                    to_list = [to.lower()]
            except Exception:
                to_list = [to.lower()]

        if (subj or "").strip().lower() == target_subject:
            subject_matches += 1

        # If any recipient matches an expected customer, check first_name in body
        body_low = (body or "").lower()
        for recip in to_list:
            fn = cust_map.get(recip)
            if fn and fn.lower() in body_low:
                first_name_matches += 1
                break

    record("All emails have exact subject 'We Value Your Feedback - Product Quality Survey'",
           subject_matches == len(emails) and subject_matches > 0,
           f"Subject match: {subject_matches}/{len(emails)}")

    record(f"All emails to affected customers include first_name in body",
           first_name_matches >= len(expected["customers"]),
           f"{first_name_matches}/{len(expected['customers'])} had first_name in body")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", default=".")
    parser.add_argument("--groundtruth_workspace", default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    expected = get_expected()
    check_gform(expected)
    check_emails(expected)
    print(f"\n=== SUMMARY: {PASS_COUNT} passed, {FAIL_COUNT} failed ===")
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "success": FAIL_COUNT == 0}, f)
    sys.exit(0 if FAIL_COUNT == 0 else 1)

if __name__ == "__main__":
    main()
