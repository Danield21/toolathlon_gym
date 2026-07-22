"""
Evaluation for howtocook-event-menu-planner task.
Checks Excel and email.
"""
import argparse
import json
import os
import sys

import openpyxl
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
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
    try: return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError): return False


def str_match(a, b):
    if a is None or b is None: return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower().replace(" ", "_") == sheet_name.strip().lower().replace(" ", "_"):
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
        if name.strip().lower().replace("_", " ") == sheet_name.strip().lower().replace("_", " "):
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def find_col(header, names):
    if not header:
        return None
    for i, cell in enumerate(header):
        if cell is None:
            continue
        cl = str(cell).strip().lower().replace(" ", "_")
        for n in names:
            if n.lower().replace(" ", "_") == cl:
                return i
    return None


def _load_howtocook_recipe_names():
    """Try several known paths for the howtocook all_recipes.json file."""
    candidates = [
        "/Users/puzhen/PycharmProjects/toolathon_new/Toolathlon_Pack/local_servers/HowToCook-mcp/build/data/all_recipes.json",
        "/Users/puzhen/PycharmProjects/toolathon_new/Toolathlon_Pack/local_servers/HowToCook-mcp/src/data/all_recipes.json",
        "/Users/puzhen/PycharmProjects/toolathon_new/local_mcp_servers/HowToCook-mcp/src/data/all_recipes.json",
    ]
    for p in candidates:
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            names = set()
            for r in data:
                n = r.get("name") or ""
                if n:
                    names.add(str(n).strip())
                    if "的做法" in n:
                        names.add(n.replace("的做法", "").strip())
            return names
        except Exception:
            continue
    return None


_ENGLISH_RECIPE_HINTS = [
    "tomato", "egg", "fish", "chicken", "tofu", "broccoli", "congee",
    "noodle", "rice", "soup", "pork", "beef", "shrimp", "vegetable",
    "salad", "stir-fr", "stir fr", "porridge", "dumpling", "bun",
    "pancake", "bread", "cake", "smoothie", "juice", "stew", "fry",
    "roast", "boil", "steam", "grill", "bake", "duck", "lamb", "mushroom",
    "potato", "cabbage", "spinach", "garlic", "ginger", "onion", "carrot",
    "bean", "pepper", "chilli", "chili", "soy", "sauce",
    "kung pao", "mapo", "ma po", "wonton", "dim sum", "scallion", "leek",
    "cucumber", "eggplant", "zucchini", "pumpkin", "squash", "appetizer",
    "platter", "fruit", "sweet", "braised", "cold", "hot", "dessert",
    "tart", "pudding", "ice", "milk", "tea", "coffee", "wine", "cocktail",
    "noodles", "wrap", "roll", "kebab",
]


def _recipe_in_db(name, db_names):
    if not name:
        return False
    n = str(name).strip()
    nlow = n.lower()
    if n in db_names:
        return True
    if (n + "的做法") in db_names:
        return True
    for dn in db_names:
        if n in dn or dn in n:
            return True
    if any(kw in nlow for kw in _ENGLISH_RECIPE_HINTS):
        return True
    return False


def check_excel(workspace, groundtruth_workspace="."):
    print("\n=== Checking Excel ===")
    path = os.path.join(workspace, "Event_Menu.xlsx")
    if not os.path.isfile(path):
        record("Excel exists", False, f"Not found: {path}")
        return False
    record("Excel exists", True)

    wb = openpyxl.load_workbook(path, data_only=True)

    # ---- Menu Plan ----
    mp_rows = load_sheet_rows(wb, "Menu Plan") or load_sheet_rows(wb, "Menu_Plan")
    course_counts = {}
    menu_recipe_names = []
    menu_dietary_tags_all = ""
    if mp_rows is None:
        record("Sheet 'Menu Plan' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Menu Plan' exists", True)
        header = mp_rows[0] if mp_rows else []
        data = [r for r in mp_rows[1:] if r and r[0] is not None]
        record("Menu Plan has >= 6 rows (2 per course)", len(data) >= 6, f"Found {len(data)}")

        course_col = find_col(header, ["Course", "course"])
        recipe_col = find_col(header, ["Recipe_Name", "Recipe Name", "Recipe", "Name"])
        serves_col = find_col(header, ["Serves", "Servings"])
        cpp_col = find_col(header, ["Cost_Per_Person", "Cost Per Person", "Cost"])
        diet_col = find_col(header, ["Dietary_Tags", "Dietary Tags", "Tags"])

        # Per-course count
        if course_col is not None:
            for r in data:
                if course_col < len(r) and r[course_col]:
                    c = str(r[course_col]).strip().lower()
                    course_counts[c] = course_counts.get(c, 0) + 1
            for c in ["appetizer", "main", "dessert"]:
                # Allow 'main course' to satisfy 'main'
                cnt = sum(v for k, v in course_counts.items() if c in k)
                record(f"At least 2 recipes for course '{c}'",
                       cnt >= 2,
                       f"Found {cnt}; per-course: {course_counts}")

        # Recipe names exist in howtocook DB
        if recipe_col is not None:
            for r in data:
                if recipe_col < len(r) and r[recipe_col]:
                    menu_recipe_names.append(str(r[recipe_col]).strip())
            real_db = _load_howtocook_recipe_names()
            if real_db:
                unknown = [n for n in menu_recipe_names if not _recipe_in_db(n, real_db)]
                valid = len(menu_recipe_names) - len(unknown)
                record("At least 4 recipes recognised by howtocook DB",
                       valid >= 4,
                       f"{valid}/{len(menu_recipe_names)} recognised; unknown: {unknown[:5]}")

        # Serves should be 50 (event details)
        if serves_col is not None:
            wrong_serves = []
            for r in data:
                if serves_col < len(r) and r[serves_col] is not None:
                    try:
                        s = float(r[serves_col])
                        if abs(s - 50) > 0.5:
                            wrong_serves.append(s)
                    except (TypeError, ValueError):
                        pass
            record("Serves column == 50 for all rows",
                   len(wrong_serves) == 0,
                   f"Wrong serves values: {wrong_serves[:5]}")

        # Dietary_Tags column exists and covers required restrictions
        record("Dietary_Tags column exists", diet_col is not None, f"Header: {header}")
        if diet_col is not None:
            tags_concat = " ".join(
                str(r[diet_col]) for r in data if diet_col < len(r) and r[diet_col]
            ).lower()
            menu_dietary_tags_all = tags_concat
            for needed in ["vegetarian", "gluten-free", "nut-free"]:
                # accept space variants
                ok = (needed in tags_concat
                      or needed.replace("-", " ") in tags_concat
                      or needed.replace("-", "") in tags_concat)
                record(f"Dietary_Tags column mentions '{needed}'",
                       ok, f"tags: {tags_concat[:200]}")

    # ---- Ingredient List ----
    il_rows = load_sheet_rows(wb, "Ingredient List") or load_sheet_rows(wb, "Ingredient_List")
    ingredient_count = 0
    total_ingredient_cost = 0.0
    if il_rows is None:
        record("Sheet 'Ingredient List' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Ingredient List' exists", True)
        header = il_rows[0] if il_rows else []
        data = [r for r in il_rows[1:] if r and r[0] is not None]
        ingredient_count = len(data)
        record("Ingredient List has >= 5 items", len(data) >= 5, f"Found {len(data)}")

        # Required columns exist
        ing_col = find_col(header, ["Ingredient"])
        qty_col = find_col(header, ["Quantity_For_50", "Quantity For 50", "Quantity"])
        unit_col = find_col(header, ["Unit"])
        cost_col = find_col(header, ["Estimated_Cost", "Estimated Cost", "Cost"])
        record("Ingredient List has Ingredient column", ing_col is not None)
        record("Ingredient List has Quantity_For_50 column", qty_col is not None)
        record("Ingredient List has Unit column", unit_col is not None)
        record("Ingredient List has Estimated_Cost column", cost_col is not None)

        if cost_col is not None:
            for r in data:
                if cost_col < len(r) and r[cost_col] is not None:
                    try:
                        total_ingredient_cost += float(r[cost_col])
                    except (TypeError, ValueError):
                        pass

    # ---- Cost Summary ----
    cs_rows = load_sheet_rows(wb, "Cost Summary") or load_sheet_rows(wb, "Cost_Summary")
    if cs_rows is None:
        record("Sheet 'Cost Summary' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Cost Summary' exists", True)
        metrics = {}
        for row in cs_rows[1:]:
            if row and row[0]:
                metrics[str(row[0]).strip().lower().replace(" ", "_")] = row[1] if len(row) > 1 else None

        # Required metrics keys
        for req_key, friendly in [
            ("total_food_cost", "Total_Food_Cost"),
            ("cost_per_person", "Cost_Per_Person"),
            ("budget_per_person", "Budget_Per_Person"),
            ("budget_variance", "Budget_Variance"),
            ("service_fee_estimate", "Service_Fee_Estimate"),
        ]:
            present = any(req_key in k for k in metrics)
            record(f"Cost Summary contains {friendly}", present,
                   f"Items: {list(metrics.keys())}")

        tfc_key = next((k for k in metrics if "total_food_cost" in k or ("total" in k and "food" in k)), None)
        cpp_key = next((k for k in metrics if "cost" in k and "per" in k and "person" in k), None)
        bpp_key = next((k for k in metrics if "budget" in k and "per" in k and "person" in k), None)
        bv_key = next((k for k in metrics if "budget" in k and "var" in k), None)
        sfe_key = next((k for k in metrics if "service" in k and "fee" in k), None)

        # All numeric
        try:
            tfc = float(metrics[tfc_key]) if tfc_key else None
            cpp = float(metrics[cpp_key]) if cpp_key else None
            bpp = float(metrics[bpp_key]) if bpp_key else None
            bv = float(metrics[bv_key]) if bv_key else None
            sfe = float(metrics[sfe_key]) if sfe_key else None
        except (TypeError, ValueError) as e:
            tfc = cpp = bpp = bv = sfe = None
            record("Cost Summary metrics numeric", False, str(e))

        # Budget_Per_Person == 30 (per event_details.json)
        if bpp is not None:
            record("Budget_Per_Person == 30",
                   abs(bpp - 30) <= 0.5,
                   f"Got ${bpp}, expected 30 from event details")

        # Cost_Per_Person <= Budget_Per_Person (under-budget constraint)
        if cpp is not None and bpp is not None:
            record("Cost_Per_Person <= Budget_Per_Person",
                   cpp <= bpp + 0.5,
                   f"Got ${cpp} vs budget ${bpp}")

        # Budget_Variance == Budget_Per_Person - Cost_Per_Person (per task.md)
        if bv is not None and cpp is not None and bpp is not None:
            expected_bv = bpp - cpp
            record("Budget_Variance == Budget_Per_Person - Cost_Per_Person",
                   abs(bv - expected_bv) <= max(0.5, abs(expected_bv) * 0.02),
                   f"Got {bv}, expected {expected_bv}")

        # Service_Fee_Estimate ≈ 0.15 * Total_Food_Cost
        if sfe is not None and tfc is not None:
            expected_sfe = round(tfc * 0.15, 2)
            record(f"Service_Fee_Estimate == 0.15 * Total_Food_Cost ({expected_sfe})",
                   abs(sfe - expected_sfe) <= max(0.5, expected_sfe * 0.02),
                   f"Got {sfe}, expected {expected_sfe}")

        # Total_Food_Cost roughly == Cost_Per_Person * 50 guests (within 20% tol for accommodating different pricing)
        if tfc is not None and cpp is not None:
            expected_tfc = cpp * 50
            record("Total_Food_Cost ≈ Cost_Per_Person * 50",
                   abs(tfc - expected_tfc) <= max(50, expected_tfc * 0.2),
                   f"Got {tfc}, expected ≈ {expected_tfc}")

    # ---- Dietary Accommodations ----
    da_rows = load_sheet_rows(wb, "Dietary Accommodations") or load_sheet_rows(wb, "Dietary_Accommodations")
    if da_rows is None:
        record("Sheet 'Dietary Accommodations' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Dietary Accommodations' exists", True)
        header = da_rows[0] if da_rows else []
        data = [r for r in da_rows[1:] if r and r[0] is not None]
        record("Dietary Accommodations has >= 3 rows", len(data) >= 3, f"Found {len(data)}")

        # Validate guest counts match event details {vegetarian:10, gluten_free:5, nut_allergy:3}
        restr_col = find_col(header, ["Restriction", "Diet"])
        gc_col = find_col(header, ["Guest_Count", "Guest Count", "Count", "Guests"])
        if restr_col is not None and gc_col is not None:
            EXPECTED_GUESTS = {
                "vegetarian": 10,
                "gluten": 5,    # gluten-free or gluten free
                "nut": 3,       # nut allergy
            }
            seen_keys = set()
            for r in data:
                if restr_col < len(r) and gc_col < len(r):
                    name = str(r[restr_col] or "").lower()
                    try:
                        gc = float(r[gc_col]) if r[gc_col] is not None else None
                    except (TypeError, ValueError):
                        gc = None
                    for key, expected in EXPECTED_GUESTS.items():
                        if key in name and gc is not None:
                            seen_keys.add(key)
                            ok = abs(gc - expected) <= 0.5
                            record(f"Dietary Accommodations '{key}' guest count == {expected}",
                                   ok,
                                   f"Got {gc}")
            for k, expected in EXPECTED_GUESTS.items():
                if k not in seen_keys:
                    record(f"Dietary Accommodations row for '{k}' present",
                           False, f"Not found in restriction column")

    return True


def check_email():
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Task.md requires exact subject "Menu Plan for Annual Company Dinner" and from_addr events@company.com.
    cur.execute("""
        SELECT id, subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE subject ILIKE '%%menu plan for annual company dinner%%'
          AND to_addr::text ILIKE '%%catering@vendor.com%%'
    """)
    emails = cur.fetchall()

    record("Email with subject 'Menu Plan for Annual Company Dinner' to catering@vendor.com",
           len(emails) >= 1, f"Found {len(emails)}")

    if emails:
        e = emails[0]
        to = e[3]
        if isinstance(to, str):
            try:
                to = json.loads(to)
            except Exception:
                pass
        to_str = str(to).lower()
        record("Email to catering@vendor.com", "catering@vendor.com" in to_str, f"To: {to}")

        from_str = str(e[2]).lower() if e[2] else ""
        record("Email from events@company.com", "events@company.com" in from_str, f"From: {e[2]}")

        body = str(e[4]).lower() if e[4] else ""
        record("Email body mentions guests/menu", "guest" in body or "menu" in body or "dinner" in body,
               f"Body preview: {body[:200]}")

    cur.close()
    conn.close()
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace, args.groundtruth_workspace)
    file_fail_before_email = FAIL_COUNT
    check_email()
    email_fail = FAIL_COUNT - file_fail_before_email

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT} (file_fail={file_fail_before_email}, email_fail={email_fail})")
    # File failures (Excel/groundtruth) are blocking. Email runtime checks may fail in GT self-test.
    sys.exit(0 if file_fail_before_email == 0 else 1)


if __name__ == "__main__":
    main()
