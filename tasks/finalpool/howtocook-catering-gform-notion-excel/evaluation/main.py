"""Evaluation for howtocook-catering-gform-notion-excel."""
import argparse
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_ONLY_FAIL = 0


def check(name, condition, detail="", runtime_only=False):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_ONLY_FAIL
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        if runtime_only:
            RUNTIME_ONLY_FAIL += 1
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def check_excel(agent_workspace):
    """Check Excel budget file."""
    print("\n=== Checking Excel Budget File ===")
    try:
        import openpyxl
    except ImportError:
        check("openpyxl available", False, "openpyxl not installed")
        return

    agent_file = os.path.join(agent_workspace, "Catering_Budget.xlsx")
    check("Catering_Budget.xlsx exists", os.path.isfile(agent_file), f"Expected {agent_file}")
    if not os.path.isfile(agent_file):
        return

    try:
        wb = openpyxl.load_workbook(agent_file, data_only=True)
    except Exception as e:
        check("Excel file readable", False, str(e))
        return

    def get_sheet(wb, name):
        for s in wb.sheetnames:
            if s.strip().lower() == name.strip().lower():
                return wb[s]
        return None

    # Check Menu sheet
    print("\n--- Menu Sheet ---")
    menu_ws = get_sheet(wb, "Menu")
    check("Sheet 'Menu' exists", menu_ws is not None, f"Found: {wb.sheetnames}")

    menu_cost_sum = 0.0
    if menu_ws:
        headers = [c.value for c in list(menu_ws.rows)[0]] if menu_ws.max_row > 0 else []
        check("Menu has Recipe_Name column",
              any("recipe" in str(h).lower() for h in headers if h),
              f"Headers: {headers}")
        check("Menu has Category column",
              any("category" in str(h).lower() for h in headers if h),
              f"Headers: {headers}")
        check("Menu has Servings column",
              any("serving" in str(h).lower() for h in headers if h),
              f"Headers: {headers}")
        check("Menu has Estimated_Cost_Per_Person column",
              any("cost" in str(h).lower() for h in headers if h),
              f"Headers: {headers}")

        data_rows = [row for row in menu_ws.iter_rows(min_row=2, values_only=True)
                     if any(v is not None for v in row)]
        # Per task: exactly 6 to 8 recipes
        check("Menu has 6 to 8 recipe rows", 6 <= len(data_rows) <= 8,
              f"Found {len(data_rows)} rows")

        # Sum Estimated_Cost_Per_Person for later verification
        cost_idx = None
        for i, h in enumerate(headers):
            if h and "cost" in str(h).lower():
                cost_idx = i
                break
        if cost_idx is not None:
            for r in data_rows:
                if cost_idx < len(r) and r[cost_idx] is not None:
                    try:
                        menu_cost_sum += float(r[cost_idx])
                    except (ValueError, TypeError):
                        pass

    # Check Summary sheet
    print("\n--- Summary Sheet ---")
    sum_ws = get_sheet(wb, "Summary")
    check("Sheet 'Summary' exists", sum_ws is not None, f"Found: {wb.sheetnames}")

    if sum_ws:
        summary_data = {}
        for row in sum_ws.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                summary_data[str(row[0]).strip().lower().replace(" ", "_")] = row[1]

        check("Summary has Total_Dishes",
              "total_dishes" in summary_data,
              f"Keys: {list(summary_data.keys())}")
        check("Summary has Total_Budget with value 500",
              "total_budget" in summary_data and num_close(summary_data["total_budget"], 500, 5),
              f"Data: {summary_data}")
        check("Summary has Budget_Per_Person with value 25",
              "budget_per_person" in summary_data and num_close(summary_data["budget_per_person"], 25, 2),
              f"Data: {summary_data}")

        # Validate Total_Dishes == menu row count
        if "total_dishes" in summary_data and menu_ws:
            try:
                td = int(float(summary_data["total_dishes"]))
                data_rows = [row for row in menu_ws.iter_rows(min_row=2, values_only=True)
                             if any(v is not None for v in row)]
                check("Total_Dishes matches Menu row count",
                      td == len(data_rows),
                      f"Summary Total_Dishes={td}, Menu rows={len(data_rows)}")
            except (ValueError, TypeError):
                pass

        # Validate sum(Estimated_Cost_Per_Person) does not exceed 25 (per-person budget)
        # Task: reasonable estimated cost per person should fit within $25 budget
        if menu_cost_sum > 0:
            check("Sum(Estimated_Cost_Per_Person) within budget 25",
                  menu_cost_sum <= 26.0,
                  f"Sum={menu_cost_sum:.2f} exceeds $25 per-person budget")


def check_gform():
    """Check Google Form creation."""
    print("\n=== Checking Google Form ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    check("At least one form created", len(forms) >= 1, f"Found {len(forms)} forms",
          runtime_only=True)

    # Target form: "Team Lunch Dietary Preferences"
    form_id = None
    for fid, title in forms:
        t = (title or "").lower()
        if "dietary" in t and "lunch" in t:
            form_id = fid
            break
    if form_id is None:
        # fallback: forms with team-lunch / dietary preferences
        for fid, title in forms:
            t = (title or "").lower()
            if "dietary" in t or ("team" in t and "lunch" in t):
                form_id = fid
                break

    check("Form titled 'Team Lunch Dietary Preferences' found",
          form_id is not None,
          f"Found titles: {[f[1] for f in forms]}",
          runtime_only=True)

    if form_id:
        cur.execute("SELECT title, question_type FROM gform.questions WHERE form_id = %s", (form_id,))
        questions = cur.fetchall()
        check("Form has at least 3 questions", len(questions) >= 3,
              f"Found {len(questions)} questions",
              runtime_only=True)
        # Per task: questions must cover dietary, allergy, comments
        q_titles = " ".join([str(q[0] or "").lower() for q in questions])
        check("Form mentions dietary/preference",
              "dietary" in q_titles or "vegetarian" in q_titles or "vegan" in q_titles
              or "preference" in q_titles,
              f"Q titles: {q_titles[:150]}",
              runtime_only=True)
        check("Form mentions allergy",
              "allerg" in q_titles,
              f"Q titles: {q_titles[:150]}",
              runtime_only=True)
        check("Form mentions comments/additional",
              "comment" in q_titles or "additional" in q_titles or "request" in q_titles,
              f"Q titles: {q_titles[:150]}",
              runtime_only=True)
        # At least one multiple-choice (radio) or checkbox question per triage.
        q_types = [str(q[1] or "").lower() for q in questions]
        check("Form has at least one multiple-choice/checkbox/radio question",
              any(any(kw in t for kw in ("multiple", "radio", "checkbox", "choice", "choose")) for t in q_types),
              f"Question types: {q_types}",
              runtime_only=True)

    conn.close()


def check_notion():
    """Check Notion page creation."""
    print("\n=== Checking Notion Page ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    # Target page: "Team Lunch Menu"
    cur.execute("SELECT id, properties FROM notion.pages")
    pages = cur.fetchall()
    check("At least one Notion page created", len(pages) >= 1,
          f"Found {len(pages)} pages",
          runtime_only=True)

    def _extract_title(props):
        if not props:
            return ""
        try:
            # properties is a jsonb; look for any key with title rich_text array
            for _, prop in (props.items() if isinstance(props, dict) else []):
                if isinstance(prop, dict):
                    if prop.get("type") == "title":
                        title_list = prop.get("title", [])
                        return "".join(t.get("plain_text", "") for t in title_list if isinstance(t, dict))
        except Exception:
            pass
        return ""

    # Look for a page titled "Team Lunch Menu" (case-insensitive)
    target_page = None
    titles = []
    for pid, props in pages:
        t = _extract_title(props)
        titles.append(t)
        if t and "team lunch menu" in t.lower():
            target_page = pid
            break
    check("Notion page titled 'Team Lunch Menu' exists",
          target_page is not None,
          f"Page titles: {titles[:5]}",
          runtime_only=True)

    if target_page:
        # Check blocks (content) on that page mention 20 people, $25, budget
        cur.execute("SELECT block_data FROM notion.blocks WHERE parent_id = %s", (target_page,))
        blocks = cur.fetchall()
        import json as _json
        texts = []
        for (bd,) in blocks:
            if bd is None:
                continue
            try:
                texts.append(_json.dumps(bd))
            except Exception:
                pass
        all_block_text = " ".join(texts).lower()
        # Stricter regex: look for '20 people' / '20 persons' / 'for 20'
        import re as _re
        twenty_people_match = bool(_re.search(r"\b20\s*(people|persons|guests|attendees|team members)\b", all_block_text)) \
                              or bool(_re.search(r"\bfor\s*20\b", all_block_text))
        check("Notion page mentions '20 people' (exact phrase)",
              twenty_people_match,
              f"Block text sample: {all_block_text[:200]}",
              runtime_only=True)
        check("Notion page mentions 25 (per person) or budget",
              "25" in all_block_text or "budget" in all_block_text,
              f"Block text len: {len(all_block_text)}",
              runtime_only=True)

    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("HOWTOCOOK CATERING GFORM NOTION EXCEL - EVALUATION")
    print("=" * 70)

    check_excel(args.agent_workspace)
    check_gform()
    check_notion()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    non_runtime_fail = FAIL_COUNT - RUNTIME_ONLY_FAIL
    overall = non_runtime_fail == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'} (runtime-only fails: {RUNTIME_ONLY_FAIL})")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
