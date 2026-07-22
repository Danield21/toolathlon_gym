"""
Evaluation for howtocook-terminal-menu-generator task.

Checks:
1. Weekly_Menu.docx exists and is readable
2. Has at least 5 sections (one per weekday Mon-Fri)
3. Contains dish names from at least 4 different categories
4. Each weekday section has ingredients listed
5. Memory file has been updated with entities
"""
import json
import os
import re
import sys
from argparse import ArgumentParser
from datetime import datetime

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {detail}")


def main():
    global PASS_COUNT, FAIL_COUNT

    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    agent_ws = args.agent_workspace

    # ---- Check 1: Weekly_Menu.docx exists and is readable ----
    docx_path = os.path.join(agent_ws, "Weekly_Menu.docx")
    if not os.path.exists(docx_path):
        check("Weekly_Menu.docx exists", False, f"Not found at {docx_path}")
        # Cannot continue without the file
        print(f"\nResults: {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} passed, {FAIL_COUNT} failed")
        sys.exit(1)

    try:
        from docx import Document
        doc = Document(docx_path)
        all_paragraphs = [p.text for p in doc.paragraphs]
        full_text = "\n".join(all_paragraphs)
        check("Weekly_Menu.docx exists and is readable", True)
    except Exception as e:
        check("Weekly_Menu.docx exists and is readable", False, str(e))
        print(f"\nResults: {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} passed, {FAIL_COUNT} failed")
        sys.exit(1)

    normalized = full_text.lower()

    # ---- Check 2: Has at least 5 weekday sections ----
    weekdays = ["monday", "tuesday", "wednesday", "thursday", "friday"]
    found_days = [d for d in weekdays if d in normalized]
    check("Has at least 5 weekday sections (Mon-Fri)",
          len(found_days) >= 5,
          f"Found {len(found_days)} weekday(s): {found_days}")

    # ---- Check 3: Parse per-day Category: lines and validate ----
    # Structured approach: find "Category: X" lines under each weekday section
    weekdays_order = ["monday", "tuesday", "wednesday", "thursday", "friday"]
    per_day_category = {}
    per_day_difficulty = {}
    per_day_name = {}
    current_day = None
    for p in all_paragraphs:
        t = p.strip()
        t_lower = t.lower()
        # Detect weekday heading
        for d in weekdays_order:
            if t_lower == d or t_lower.startswith(d + " ") or t_lower == d.capitalize().lower():
                current_day = d
                break
        else:
            # Not a weekday heading
            pass
        # Detect Category line
        cat_m = re.match(r"^\s*category\s*:\s*(\S+)", t_lower)
        if cat_m and current_day:
            per_day_category[current_day] = cat_m.group(1).strip().rstrip(",.;:")
        # Detect Difficulty line
        diff_m = re.match(r"^\s*difficulty\s*(?:level)?\s*:\s*(\d+)", t_lower)
        if diff_m and current_day:
            try:
                per_day_difficulty[current_day] = int(diff_m.group(1))
            except Exception:
                pass
        # Detect Dish: line
        dish_m = re.match(r"^\s*dish\s*:\s*(.+)$", t_lower)
        if dish_m and current_day:
            per_day_name[current_day] = dish_m.group(1).strip()

    # Distinct categories among selected days
    sel_cats = [per_day_category[d] for d in weekdays_order if d in per_day_category]
    distinct_cats = set(sel_cats)
    check("All 5 weekdays have a Category line", len(sel_cats) == 5,
          f"per_day_category={per_day_category}")
    check("At least 4 distinct categories across Mon-Fri", len(distinct_cats) >= 4,
          f"categories: {sel_cats}")

    # No consecutive-day category repeats
    consec_ok = True
    violating = []
    for i in range(1, len(sel_cats)):
        if sel_cats[i] == sel_cats[i-1]:
            consec_ok = False
            violating.append((weekdays_order[i-1], weekdays_order[i], sel_cats[i]))
    check("No consecutive-day category repeats", consec_ok,
          f"violations: {violating}")

    # All selected difficulties <= 3
    diff_values = [per_day_difficulty[d] for d in weekdays_order if d in per_day_difficulty]
    check("All 5 weekdays have a Difficulty line", len(diff_values) == 5,
          f"per_day_difficulty={per_day_difficulty}")
    check("All selected dishes have difficulty <= 3",
          all(d <= 3 for d in diff_values) if diff_values else False,
          f"difficulties: {diff_values}")

    # ---- Check 4: Each weekday section has ingredients ----
    # Check that ingredient-related words appear multiple times
    ingredient_indicators = ["ingredient", "ingredients",
                             "\u539f\u6599", "\u6750\u6599", "\u98df\u6750",
                             "salt", "oil", "sugar", "sauce", "water",
                             "garlic", "ginger", "onion", "pepper",
                             "\u76d0", "\u6cb9", "\u7cd6", "\u916b", "\u6c34",
                             "\u849c", "\u59dc", "\u8471", "\u6912"]
    ingredient_count = 0
    for indicator in ingredient_indicators:
        ingredient_count += normalized.count(indicator)

    check("Weekday sections contain ingredient information",
          ingredient_count >= 5,
          f"Found only {ingredient_count} ingredient-related terms")

    # ---- Check 5: Document has sufficient content ----
    check("Document has sufficient content (>= 500 chars)",
          len(full_text.strip()) >= 500,
          f"Only {len(full_text.strip())} characters")

    # ---- Check 6: Memory file has been updated ----
    memory_path = os.path.join(agent_ws, "memory", "memory.json")
    if os.path.exists(memory_path):
        try:
            with open(memory_path) as f:
                mem_data = json.load(f)
            entities = mem_data.get("entities", [])
            check("Memory file has been updated with entities",
                  len(entities) > 0,
                  "No entities found in memory.json")
            # Must contain dietary preferences with >=4 observations
            diet_entity = next(
                (e for e in entities
                 if "dietary" in str(e.get("name", "")).lower()
                 or "preference" in str(e.get("entityType", "")).lower()),
                None)
            if diet_entity:
                obs = diet_entity.get("observations", [])
                check("Memory has dietary_preferences entity with >=4 observations",
                      len(obs) >= 4, f"Got {len(obs)} observations")
            else:
                check("Memory has dietary_preferences entity with >=4 observations",
                      False, f"Entities: {[e.get('name') for e in entities]}")
        except Exception as e:
            check("Memory file has been updated with entities", False, str(e))
    else:
        check("Memory file has been updated with entities", False,
              f"memory.json not found at {memory_path}")

    # ---- Check 7: Title contains 'Weekly Lunch Menu' (or 'Weekly Menu') ----
    title_ok = (
        "weekly lunch menu" in normalized
        or "weekly menu" in normalized
    )
    check("Title contains 'Weekly Lunch Menu' or 'Weekly Menu'",
          title_ok, f"Sample: {full_text[:150]}")

    # ---- Check 8: Intermediate JSON pipeline artifacts REQUIRED ----
    # Task requires saving all_recipes.json and filtered_recipes.json.
    all_recipes_path = None
    filtered_path = None
    for candidate in ("all_recipes.json", "recipes.json", "all_dishes.json"):
        p = os.path.join(agent_ws, candidate)
        if os.path.isfile(p):
            all_recipes_path = p
            break
    for candidate in ("filtered_recipes.json", "filtered.json", "dishes_filtered.json"):
        p = os.path.join(agent_ws, candidate)
        if os.path.isfile(p):
            filtered_path = p
            break
    check("all_recipes.json present (REQUIRED)", all_recipes_path is not None,
          f"Expected at {os.path.join(agent_ws, 'all_recipes.json')}")
    check("filtered_recipes.json present (REQUIRED)", filtered_path is not None,
          f"Expected at {os.path.join(agent_ws, 'filtered_recipes.json')}")
    if filtered_path:
        try:
            with open(filtered_path) as f:
                data = json.load(f)
            # Look for records with difficulty<=3
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                items = data.get("recipes") or data.get("dishes") or list(data.values())
            else:
                items = []
            hard_ones = [r for r in items if isinstance(r, dict)
                         and isinstance(r.get("difficulty"), (int, float))
                         and r["difficulty"] > 3]
            check("Filtered recipes JSON: no difficulty>3 entries",
                  len(hard_ones) == 0,
                  f"Found {len(hard_ones)} difficulty>3")
        except Exception as e:
            check("Filtered recipes JSON readable", False, str(e))

    # ---- Summary ----
    total = PASS_COUNT + FAIL_COUNT
    print(f"\nResults: {PASS_COUNT}/{total} passed, {FAIL_COUNT} failed")

    if args.res_log_file:
        result = {
            "total_passed": PASS_COUNT,
            "total_checks": total,
            "accuracy": (PASS_COUNT / total * 100) if total > 0 else 0,
            "timestamp": datetime.now().isoformat(),
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
