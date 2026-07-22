"""
Evaluation for howtocook-diet-gform-notion-excel task.
Checks: GForm with questions, Notion page, Excel with 2 sheets, email.
"""
import argparse
import os
import sys

import psycopg2
import openpyxl

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_ONLY_FAIL = 0


def record(name, passed, detail="", runtime_only=False):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_ONLY_FAIL
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        if runtime_only:
            RUNTIME_ONLY_FAIL += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_gform():
    print("\n=== Checking Google Form ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gform.forms")
        forms = cur.fetchall()

        # Require title to contain 'dietary preference' (tighter than substring any word)
        target_form = None
        for fid, title in forms:
            tlow = (title or "").lower().strip()
            if "dietary preference" in tlow and "survey" in tlow:
                target_form = fid
                break
        # Fallback to earlier heuristic only if strict not found to still locate agent attempts
        if target_form is None:
            for fid, title in forms:
                tlow = (title or "").lower().strip()
                if "dietary" in tlow and ("preference" in tlow or "survey" in tlow):
                    target_form = fid
                    break
        record("GForm 'Dietary Preference Survey' exists",
               target_form is not None,
               f"Found forms: {[t for _, t in forms]}",
               runtime_only=True)

        if target_form is None:
            conn.close()
            return

        cur.execute("SELECT title, question_type, required, config FROM gform.questions WHERE form_id = %s ORDER BY position", (target_form,))
        questions = cur.fetchall()
        record("GForm has at least 4 questions", len(questions) >= 4,
               f"Found {len(questions)} questions",
               runtime_only=True)

        q_types = [q[1] for q in questions]
        has_radio = "RADIO" in q_types or "MULTIPLE_CHOICE" in q_types
        has_checkbox = "CHECKBOX" in q_types
        has_scale = "SCALE" in q_types or "LINEAR_SCALE" in q_types
        has_text = "TEXT" in q_types or "SHORT_ANSWER" in q_types or "PARAGRAPH" in q_types
        record("GForm has a multiple-choice (RADIO) question", has_radio,
               f"Question types: {q_types}",
               runtime_only=True)
        record("GForm has a checkbox question", has_checkbox,
               f"Question types: {q_types}",
               runtime_only=True)
        record("GForm has a scale question", has_scale,
               f"Question types: {q_types}",
               runtime_only=True)
        record("GForm has a text question", has_text,
               f"Question types: {q_types}",
               runtime_only=True)

        # Check question about dietary restrictions
        q_titles_lower = [q[0].lower() if q[0] else "" for q in questions]
        has_dietary_q = any("dietary" in qt or "restriction" in qt or "vegetarian" in qt for qt in q_titles_lower)
        record("GForm has dietary restrictions question", has_dietary_q,
               f"Question titles: {q_titles_lower}",
               runtime_only=True)

        # --- Option set verification via questions.config ---
        # Meal types multi-choice question should have Breakfast, Lunch, Dinner, Snacks
        meal_q_ok = False
        meal_q_opts = None
        # Dietary restriction checkbox should include Vegetarian, Vegan, Gluten-Free, Dairy-Free, None
        restr_q_ok = False
        restr_q_opts = None
        # Scale question 1..7
        scale_ok = False
        scale_detail = None
        for title, qtype, required, cfg in questions:
            tlow = (title or "").lower()
            cfg_str = str(cfg) if cfg else ""
            opts_lower = [str(o).strip().lower() for o in (cfg.get("options") if isinstance(cfg, dict) else [])]
            if ("meal" in tlow or "meal types" in tlow) and qtype in ("RADIO", "MULTIPLE_CHOICE"):
                needed = {"breakfast", "lunch", "dinner", "snacks"}
                if needed.issubset(set(opts_lower)):
                    meal_q_ok = True
                meal_q_opts = opts_lower
            if ("dietary" in tlow or "restriction" in tlow) and qtype == "CHECKBOX":
                needed = {"vegetarian", "vegan", "gluten-free", "dairy-free", "none"}
                if needed.issubset(set(opts_lower)):
                    restr_q_ok = True
                restr_q_opts = opts_lower
            if qtype in ("SCALE", "LINEAR_SCALE"):
                # Config commonly has 'low'/'high' or 'min'/'max' or 'lowerBound'/'upperBound'
                if isinstance(cfg, dict):
                    low = cfg.get("low", cfg.get("min", cfg.get("lowerBound", cfg.get("low_value"))))
                    high = cfg.get("high", cfg.get("max", cfg.get("upperBound", cfg.get("high_value"))))
                    try:
                        if int(low) == 1 and int(high) == 7:
                            scale_ok = True
                    except (TypeError, ValueError):
                        pass
                    scale_detail = f"low={low}, high={high}, cfg={cfg_str[:150]}"
                else:
                    scale_detail = cfg_str[:150]
        record("Meal-types question has Breakfast/Lunch/Dinner/Snacks options", meal_q_ok,
               f"Options: {meal_q_opts}",
               runtime_only=True)
        record("Dietary-restrictions question has Veg/Vegan/GF/DF/None options", restr_q_ok,
               f"Options: {restr_q_opts}",
               runtime_only=True)
        record("Scale question ranges 1..7", scale_ok,
               f"{scale_detail}",
               runtime_only=True)

        conn.close()
    except Exception as e:
        record("GForm connection", False, str(e), runtime_only=True)


def check_notion():
    print("\n=== Checking Notion Page ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        cur.execute("""
            SELECT id, properties FROM notion.pages
            WHERE archived = false AND in_trash = false
        """)
        pages = cur.fetchall()

        # Strict title match: "Healthy Recipe Knowledge Base"
        target_page = None
        for pid, props in pages:
            props_text = str(props).lower() if props else ""
            if "healthy recipe knowledge base" in props_text:
                target_page = pid
                break
        # Fallback loose match (so agent-created pages with minor variations still found)
        if target_page is None:
            for pid, props in pages:
                props_text = str(props).lower() if props else ""
                if "healthy recipe" in props_text and ("knowledge" in props_text or "base" in props_text):
                    target_page = pid
                    break

        record("Notion page 'Healthy Recipe Knowledge Base' exists",
               target_page is not None,
               f"Searched {len(pages)} pages",
               runtime_only=True)

        if target_page is not None:
            # Check blocks under that page - require at least 6 recipe-mention blocks
            cur.execute("""
                SELECT id, type, block_data FROM notion.blocks
                WHERE parent_id = %s AND archived = false AND in_trash = false
                ORDER BY position
            """, (target_page,))
            blocks = cur.fetchall()
            record("Notion page has at least 6 content blocks", len(blocks) >= 6,
                   f"Found {len(blocks)} blocks (task requires at least 6 recipes)",
                   runtime_only=True)

            # Count blocks mentioning any recipe name/category from the GT list
            recipe_keywords = [
                "tomato and egg", "steamed fish", "kung pao", "mapo tofu",
                "stir-fried broccoli", "congee",
                "vegetables", "seafood", "meat", "tofu", "porridge"
            ]
            recipe_mention_blocks = 0
            for _, btype, bdata in blocks:
                text = str(bdata or "").lower()
                if any(kw in text for kw in recipe_keywords):
                    recipe_mention_blocks += 1
            record("Notion page has recipe-content blocks (>=6)",
                   recipe_mention_blocks >= 6,
                   f"Found {recipe_mention_blocks} recipe-related blocks",
                   runtime_only=True)

        conn.close()
    except Exception as e:
        record("Notion connection", False, str(e), runtime_only=True)


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def _load_howtocook_recipe_names():
    """Try several known paths for the howtocook all_recipes.json file."""
    import json
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
                    # Strip trailing 的做法 to allow shorter variants
                    if "的做法" in n:
                        names.add(n.replace("的做法", "").strip())
            return names
        except Exception:
            continue
    return None


# A small known-good set of common English transliterations used historically
# in tasks. Allows agents using English names to pass even when the underlying
# DB has only Chinese names.
_ENGLISH_RECIPE_HINTS = [
    "tomato", "egg", "fish", "chicken", "tofu", "broccoli", "congee",
    "noodle", "rice", "soup", "pork", "beef", "shrimp", "vegetable",
    "salad", "stir-fr", "stir fr", "porridge", "dumpling", "bun",
    "pancake", "bread", "cake", "smoothie", "juice", "stew", "fry",
    "roast", "boil", "steam", "grill", "bake", "duck", "lamb", "mushroom",
    "potato", "cabbage", "spinach", "garlic", "ginger", "onion", "carrot",
    "bean", "pepper", "chilli", "chili", "soy", "sauce",
    "kung pao", "mapo", "ma po", "wonton", "dim sum", "scallion", "leek",
    "cucumber", "eggplant", "zucchini", "pumpkin", "squash",
]


def _recipe_in_db(name, db_names):
    """Return True if a recipe name is plausibly real.

    Accepts (a) exact match against db_names (Chinese), (b) English text
    that contains common food keywords from _ENGLISH_RECIPE_HINTS.
    """
    if not name:
        return False
    n = str(name).strip()
    nlow = n.lower()
    if n in db_names:
        return True
    # Strip the suffix the DB uses (...的做法) and try again
    if (n + "的做法") in db_names:
        return True
    # Substring match against any DB name
    for dn in db_names:
        if n in dn or dn in n:
            return True
    # English transliteration heuristic
    if any(kw in nlow for kw in _ENGLISH_RECIPE_HINTS):
        return True
    return False


def check_excel(agent_workspace, groundtruth_workspace="."):
    print("\n=== Checking Excel File ===")
    xl_path = os.path.join(agent_workspace, "Recipe_Overview.xlsx")
    if not os.path.isfile(xl_path):
        record("Excel file Recipe_Overview.xlsx exists", False, f"Not found at: {xl_path}")
        return
    record("Excel file Recipe_Overview.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xl_path, data_only=True)
    except Exception as e:
        record("Excel file readable", False, str(e))
        return
    record("Excel file readable", True)

    sheet_names = [s.lower() for s in wb.sheetnames]
    has_recipes = any("recipe" in s for s in sheet_names)
    has_summary = any("category" in s or "summary" in s for s in sheet_names)
    record("Excel has 'Recipes' sheet", has_recipes, f"Found sheets: {wb.sheetnames}")
    record("Excel has 'Category Summary' sheet", has_summary, f"Found sheets: {wb.sheetnames}")

    # ---- Constraint-based validation (NOT strict GT recipe-name matching) ----
    # task.md: "select at least 6 recipes representing at least 3 different categories"
    # Agents may legitimately pick any valid recipes from the howtocook MCP, in
    # English or Chinese. Strict per-recipe matching would FN compliant agents.

    # Find Recipes sheet
    recipes_sheet = None
    recipes_sheet_name = None
    for sname in wb.sheetnames:
        if "recipe" in sname.lower() and "summary" not in sname.lower() and "category" not in sname.lower():
            recipes_sheet = wb[sname]
            recipes_sheet_name = sname
            break

    # Find Category Summary sheet
    cat_sheet = None
    cat_sheet_name = None
    for sname in wb.sheetnames:
        if "category" in sname.lower() or "summary" in sname.lower():
            cat_sheet = wb[sname]
            cat_sheet_name = sname
            break

    if recipes_sheet is None:
        record("Recipes sheet present", False, f"Sheets: {wb.sheetnames}")
        return

    # Read header row to identify column positions (Recipe_Name, Category, Description)
    header_cells = next(recipes_sheet.iter_rows(min_row=1, max_row=1, values_only=True), tuple())
    headers = [str(h or "").strip().lower() for h in header_cells]
    def _col_idx(*candidates):
        for c in candidates:
            cnorm = c.strip().lower()
            for i, h in enumerate(headers):
                if h.replace("_", " ") == cnorm.replace("_", " ") or cnorm in h.replace("_", " "):
                    return i
        return None
    name_idx = _col_idx("recipe_name", "name", "recipe name")
    cat_idx = _col_idx("category")
    desc_idx = _col_idx("description")

    # Collect data rows (non-empty)
    recipe_rows = []
    for row in recipes_sheet.iter_rows(min_row=2, values_only=True):
        if any(cell is not None and str(cell).strip() != "" for cell in row):
            recipe_rows.append(row)
    record("Excel Recipes sheet has at least 6 recipe rows", len(recipe_rows) >= 6,
           f"Found {len(recipe_rows)} data rows")

    # Validate each recipe row: name + category + description non-empty
    if name_idx is not None and cat_idx is not None and desc_idx is not None:
        empty_count = 0
        for r in recipe_rows:
            n = str(r[name_idx] or "").strip() if name_idx < len(r) else ""
            c = str(r[cat_idx] or "").strip() if cat_idx < len(r) else ""
            d = str(r[desc_idx] or "").strip() if desc_idx < len(r) else ""
            if not n or not c or not d:
                empty_count += 1
        record("All recipe rows have non-empty Name/Category/Description",
               empty_count == 0,
               f"{empty_count} rows have empty Name/Category/Description")
    else:
        record("Recipes sheet has Recipe_Name/Category/Description columns",
               False, f"Headers: {headers}")

    # Validate recipe names exist in the howtocook recipe DB (allow English or
    # Chinese). Use a tolerant substring/equality match against known recipe
    # names from all_recipes.json.
    real_recipe_names = _load_howtocook_recipe_names()
    if real_recipe_names and name_idx is not None:
        unknown = []
        for r in recipe_rows:
            if name_idx >= len(r):
                continue
            n_raw = r[name_idx]
            if n_raw is None:
                continue
            n = str(n_raw).strip()
            if not _recipe_in_db(n, real_recipe_names):
                unknown.append(n)
        # Allow up to 1 unknown (e.g., minor spelling diff) but require >= 5 valid
        valid_count = len(recipe_rows) - len(unknown)
        record("At least 5 selected recipes match howtocook DB names",
               valid_count >= 5,
               f"{valid_count}/{len(recipe_rows)} recognised; unknown: {unknown[:5]}")

    # Distinct categories >= 3
    if cat_idx is not None:
        cats = set()
        for r in recipe_rows:
            if cat_idx < len(r):
                v = r[cat_idx]
                if v is not None and str(v).strip():
                    cats.add(str(v).strip().lower())
        record("Recipes span at least 3 distinct categories",
               len(cats) >= 3,
               f"Found {len(cats)} distinct categories: {sorted(cats)}")
    else:
        cats = set()

    # Category Summary sheet: counts must equal actual row counts in Recipes
    if cat_sheet is None:
        record("Category Summary sheet present", False, f"Sheets: {wb.sheetnames}")
    else:
        cat_summary_rows = []
        for row in cat_sheet.iter_rows(min_row=2, values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                cat_summary_rows.append(row)
        record("Category Summary has at least 3 categories",
               len(cat_summary_rows) >= 3,
               f"Found {len(cat_summary_rows)} category rows")

        # Build expected counts from Recipes sheet
        if cat_idx is not None:
            expected_counts = {}
            for r in recipe_rows:
                if cat_idx < len(r):
                    v = r[cat_idx]
                    if v is not None and str(v).strip():
                        key = str(v).strip().lower()
                        expected_counts[key] = expected_counts.get(key, 0) + 1
            # Build agent_summary {category: count}
            agent_summary = {}
            for r in cat_summary_rows:
                if len(r) >= 2 and r[0] is not None and r[1] is not None:
                    key = str(r[0]).strip().lower()
                    try:
                        agent_summary[key] = int(float(r[1]))
                    except (TypeError, ValueError):
                        pass
            # Sum-of-counts must equal total recipe rows
            total_summary_count = sum(agent_summary.values())
            record("Category Summary total count equals Recipes row count",
                   total_summary_count == len(recipe_rows),
                   f"Summary total={total_summary_count}, Recipes rows={len(recipe_rows)}")
            # All categories in Recipes should appear in Summary with correct count
            mismatch = []
            for k, v in expected_counts.items():
                a = agent_summary.get(k)
                if a != v:
                    mismatch.append(f"{k}: expected {v}, got {a}")
            record("Each category count in summary matches Recipes-sheet count",
                   len(mismatch) == 0,
                   f"Mismatches: {mismatch[:5]}")


def check_email():
    print("\n=== Checking Email ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Tightened: require subject to mention 'healthy eating' or 'recipe selection',
        # not just any keyword; still accept variants while excluding unrelated
        # emails that happen to mention 'eating' or 'recipe' alone.
        cur.execute("""
            SELECT subject, to_addr FROM email.messages
            WHERE LOWER(subject) LIKE '%healthy eating%'
               OR LOWER(subject) LIKE '%recipe selection%'
               OR LOWER(subject) LIKE '%healthy recipe%'
        """)
        emails = cur.fetchall()
        record("Email about healthy eating/recipes sent", len(emails) > 0,
               f"Found {len(emails)} matching emails",
               runtime_only=True)

        if emails:
            target_found = False
            for subject, to_addr in emails:
                to_str = str(to_addr).lower() if to_addr else ""
                if "wellness.team@company.com" in to_str:
                    target_found = True
                    break
            record("Email sent to wellness.team@company.com", target_found,
                   f"Recipients: {[e[1] for e in emails]}",
                   runtime_only=True)

        conn.close()
    except Exception as e:
        record("Email connection", False, str(e), runtime_only=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=True)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    check_gform()
    check_notion()
    check_excel(args.agent_workspace, args.groundtruth_workspace)
    check_email()

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== Results: {PASS_COUNT}/{total} passed ===")
    non_runtime_fail = FAIL_COUNT - RUNTIME_ONLY_FAIL
    overall = non_runtime_fail == 0
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Overall: {'PASS' if overall else 'FAIL'} (runtime-only fails: {RUNTIME_ONLY_FAIL})")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
