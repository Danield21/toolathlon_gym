"""
Evaluation script for howtocook-event-registration task.

Checks via psycopg2:
1. Google Form exists with "cooking" or "lunch" in title, has at least 4 questions
2. Calendar event exists with "cooking" or "lunch" in summary, dated 2026-03-20
3. Notion page exists with "cooking" or "event" or "planning" in properties
4. At least 1 email sent with "lunch" or "cooking" in subject, mentioning March 20
"""

import os
import argparse
import json
import sys

import psycopg2

def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
RUNTIME_INCONCLUSIVE = 0
IN_RUNTIME_BLOCK = False


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT, RUNTIME_INCONCLUSIVE
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        if IN_RUNTIME_BLOCK:
            RUNTIME_INCONCLUSIVE += 1
            detail_str = f": {detail[:300]}" if detail else ""
            print(f"  [FAIL-runtime] {name}{detail_str}")
        else:
            FAIL_COUNT += 1
            detail_str = f": {detail[:300]}" if detail else ""
            print(f"  [FAIL] {name}{detail_str}")


def check_google_form():
    """Check that a Google Form was created with the expected structure."""
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Google Form ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Check total forms - if zero, all gform checks are runtime-inconclusive
    cur.execute("SELECT COUNT(*) FROM gform.forms")
    (n_forms,) = cur.fetchone()
    if n_forms == 0:
        IN_RUNTIME_BLOCK = True

    # Require exact title: "Lunch and Learn: Chinese Cooking Registration"
    cur.execute("""
        SELECT id, title
        FROM gform.forms
        WHERE LOWER(title) LIKE '%%lunch and learn%%chinese cooking%%registration%%'
           OR LOWER(title) LIKE '%%chinese cooking registration%%'
    """)
    forms = cur.fetchall()

    check("Google Form titled 'Lunch and Learn: Chinese Cooking Registration' exists",
          len(forms) > 0,
          "No matching forms found - exact title required")

    if not forms:
        cur.close()
        conn.close()
        IN_RUNTIME_BLOCK = False
        return set()

    form_id = forms[0][0]
    form_title = forms[0][1]
    print(f"  Found form: '{form_title}' (id={form_id})")

    # Check number of questions
    cur.execute("""
        SELECT id, title, question_type, required
        FROM gform.questions
        WHERE form_id = %s
        ORDER BY position
    """, (form_id,))
    questions = cur.fetchall()

    check("Form has exactly 5 questions",
          len(questions) == 5,
          f"Found {len(questions)} questions, expected exactly 5")

    # Check question content by title keywords
    q_titles_lower = [q[1].lower() for q in questions]

    check("Has a 'name' question",
          any("name" in t for t in q_titles_lower),
          f"Question titles: {[q[1] for q in questions]}")

    check("Has an 'email' question",
          any("email" in t for t in q_titles_lower),
          f"Question titles: {[q[1] for q in questions]}")

    check("Has a 'department' question",
          any("department" in t for t in q_titles_lower),
          f"Question titles: {[q[1] for q in questions]}")

    check("Has a 'dietary' or 'restriction' question",
          any("diet" in t or "restriction" in t for t in q_titles_lower),
          f"Question titles: {[q[1] for q in questions]}")

    check("Has a 'dish' or 'excited' question",
          any("dish" in t or "excited" in t for t in q_titles_lower),
          f"Question titles: {[q[1] for q in questions]}")

    # Check required fields - expected 4 required (Full Name, Email, Department, Dish)
    required_questions = [q for q in questions if q[3] is True]
    check("Exactly 4 required questions (Name, Email, Department, Dish)",
          len(required_questions) == 4,
          f"Found {len(required_questions)} required questions")

    # Dietary Restrictions question should be OPTIONAL
    diet_q = next((q for q in questions if "diet" in q[1].lower() or "restriction" in q[1].lower()), None)
    if diet_q is not None:
        check("Dietary Restrictions is OPTIONAL",
              diet_q[3] is False,
              f"Got required={diet_q[3]}")

    # Validate option lists
    cur.execute("""
        SELECT title, question_type, required, config
        FROM gform.questions WHERE form_id = %s ORDER BY position
    """, (form_id,))
    full_questions = cur.fetchall()

    expected_dept_opts = {"Engineering", "Marketing", "Finance", "HR", "Operations"}
    expected_diet_opts = {"Vegetarian", "Vegan", "Gluten-free", "Nut-free", "None"}

    dept_q = next((q for q in full_questions if "department" in q[0].lower()), None)
    if dept_q:
        try:
            cfg = dept_q[3] if isinstance(dept_q[3], dict) else json.loads(dept_q[3]) if dept_q[3] else {}
            opts = set(str(o).strip() for o in (cfg.get("options") or []))
            check("Department question has all 5 expected options",
                  expected_dept_opts.issubset(opts),
                  f"Got options: {opts}")
        except Exception as e:
            check("Department options", False, str(e))

    diet_q_full = next((q for q in full_questions if "diet" in q[0].lower() or "restriction" in q[0].lower()), None)
    if diet_q_full:
        try:
            cfg = diet_q_full[3] if isinstance(diet_q_full[3], dict) else json.loads(diet_q_full[3]) if diet_q_full[3] else {}
            opts = set(str(o).strip() for o in (cfg.get("options") or []))
            check("Dietary Restrictions has all 5 expected options",
                  expected_diet_opts.issubset(opts),
                  f"Got options: {opts}")
        except Exception as e:
            check("Dietary options", False, str(e))

    # Dish question should have exactly 2 options (the 2 selected dish names)
    dish_q = next((q for q in full_questions if "dish" in q[0].lower() or "excited" in q[0].lower()), None)
    selected_dishes = set()
    if dish_q:
        try:
            cfg = dish_q[3] if isinstance(dish_q[3], dict) else json.loads(dish_q[3]) if dish_q[3] else {}
            opts = [str(o).strip() for o in (cfg.get("options") or [])]
            check("Dish question has exactly 2 options (the 2 selected dishes)",
                  len(opts) == 2,
                  f"Got {len(opts)} options: {opts}")
            selected_dishes = set(o.lower() for o in opts)
        except Exception as e:
            check("Dish options", False, str(e))

    cur.close()
    conn.close()
    IN_RUNTIME_BLOCK = False
    return selected_dishes


def check_calendar(selected_dishes=None):
    """Check Google Calendar event."""
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Google Calendar ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM gcal.events")
    (n_ev,) = cur.fetchone()
    if n_ev == 0:
        IN_RUNTIME_BLOCK = True

    # Require exact summary 'Lunch and Learn: Chinese Cooking'
    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime, location, start_timezone
        FROM gcal.events
        WHERE LOWER(summary) LIKE '%%lunch and learn%%chinese cooking%%'
    """)
    events = cur.fetchall()

    check("Calendar event titled 'Lunch and Learn: Chinese Cooking' exists",
          len(events) > 0,
          "No matching calendar events found")

    if not events:
        cur.close()
        conn.close()
        IN_RUNTIME_BLOCK = False
        return

    event = events[0]
    summary, description, start_dt, end_dt, location, start_tz = event
    print(f"  Found event: '{summary}'")

    # Check date is 2026-03-20 (with TZ normalization: if tzinfo is UTC/other, convert to local intent)
    if start_dt:
        # If start_dt is timezone-aware and start_timezone is set, convert
        try:
            if start_tz and start_dt.tzinfo is not None:
                import zoneinfo
                local_dt = start_dt.astimezone(zoneinfo.ZoneInfo(start_tz))
            else:
                local_dt = start_dt
        except Exception:
            local_dt = start_dt
        date_str = local_dt.strftime("%Y-%m-%d")
        check("Event date is 2026-03-20",
              date_str == "2026-03-20",
              f"Got {date_str}")
        # Exact 12:00 start
        check("Event starts at 12:00 (local)",
              local_dt.hour == 12 and local_dt.minute == 0,
              f"Got {local_dt.hour:02d}:{local_dt.minute:02d}")
    else:
        check("Event has a start datetime", False, "start_datetime is null")

    # Check duration ~90 minutes (with 2-min tolerance)
    if start_dt and end_dt:
        duration_minutes = (end_dt - start_dt).total_seconds() / 60
        check("Event duration is exactly 90 minutes",
              abs(duration_minutes - 90) <= 2,
              f"Got {duration_minutes:.0f} minutes")

    # Location must equal "Company Kitchen"
    if location:
        check("Event location equals 'Company Kitchen'",
              location.strip().lower() == "company kitchen",
              f"Location: '{location}'")
    else:
        check("Event has location set to Company Kitchen", False, "location is null")

    # Description mentions both selected dishes
    if description and selected_dishes:
        desc_lower = description.lower()
        missing = [d for d in selected_dishes if d not in desc_lower]
        check("Event description mentions both selected dishes",
              len(missing) == 0,
              f"Missing dish names in description: {missing}")

    cur.close()
    conn.close()


def check_notion(selected_dishes=None):
    """Check Notion page for cooking event planning."""
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Notion ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM notion.pages")
    (n_pages,) = cur.fetchone()
    if n_pages == 0:
        IN_RUNTIME_BLOCK = True

    # Find page titled exactly "Cooking Event Planning"
    cur.execute("SELECT id, properties FROM notion.pages")
    target_pid = None
    target_props = None
    for pid, props in cur.fetchall():
        if props is None:
            continue
        try:
            p = props if isinstance(props, dict) else json.loads(props)
            title_val = p.get("title")
            if isinstance(title_val, dict):
                items = title_val.get("title", [])
                title_text = " ".join(
                    (it.get("plain_text") or it.get("text", {}).get("content", ""))
                    for it in items if isinstance(it, dict)
                ).strip().lower()
                if "cooking event planning" in title_text:
                    target_pid = pid
                    target_props = p
                    break
        except Exception:
            continue

    check("Notion page titled 'Cooking Event Planning' exists",
          target_pid is not None,
          "No matching page with exact title")

    if target_pid is None:
        cur.close()
        conn.close()
        IN_RUNTIME_BLOCK = False
        return

    # Gather block text for this specific page
    cur.execute("SELECT block_data FROM notion.blocks WHERE parent_id = %s", (target_pid,))
    blocks = cur.fetchall()

    all_text = ""
    for b in blocks:
        if b[0]:
            all_text += (json.dumps(b[0]) if isinstance(b[0], (dict, list)) else str(b[0])).lower() + " "
    if target_props:
        all_text += json.dumps(target_props).lower()

    has_date = "march 20" in all_text or "2026-03-20" in all_text or "3/20" in all_text
    check("Notion page mentions event date March 20, 2026",
          has_date,
          "Date not found in page content")

    has_time = "12:00" in all_text or "12 pm" in all_text or "1:30" in all_text
    check("Notion page mentions event time 12:00 PM to 1:30 PM",
          has_time,
          "Time not found in page content")

    has_location = "kitchen" in all_text
    check("Notion page mentions Company Kitchen location",
          has_location,
          "Location not found in page content")

    has_recipe_content = ("ingredient" in all_text or "step" in all_text)
    check("Notion page mentions ingredients and preparation steps",
          has_recipe_content,
          "Recipe content not found")

    # Dishes must be present
    if selected_dishes:
        missing = [d for d in selected_dishes if d not in all_text]
        check("Notion page mentions both selected dishes",
              len(missing) == 0,
              f"Missing: {missing}")

    cur.close()
    conn.close()


def check_emails(selected_dishes=None):
    """Check that announcement email was sent."""
    global IN_RUNTIME_BLOCK
    print("\n=== Checking Emails ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM email.messages")
    (n_msgs,) = cur.fetchone()
    if n_msgs == 0:
        IN_RUNTIME_BLOCK = True

    # Require exact subject: "You're Invited: Lunch and Learn Chinese Cooking - March 20"
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_mails = cur.fetchall()

    def norm_subj(s):
        return (s or "").lower().replace("’", "'").replace("‘", "'")

    emails = [m for m in all_mails
              if "you're invited" in norm_subj(m[0]) and "lunch and learn" in norm_subj(m[0])
                 and "chinese cooking" in norm_subj(m[0]) and "march 20" in norm_subj(m[0])]

    check("Email with exact subject 'You're Invited: Lunch and Learn Chinese Cooking - March 20'",
          len(emails) >= 1,
          f"Found {len(emails)} matching. Subjects: {[m[0] for m in all_mails][:5]}")

    if not emails:
        cur.close()
        conn.close()
        IN_RUNTIME_BLOCK = False
        return

    email_row = emails[0]
    subject, from_addr, to_addr, body_text = email_row
    print(f"  Found email with subject: '{subject}'")

    subject_lower = norm_subj(subject)

    # Check body mentions March 20
    body_lower = (body_text or "").lower()
    has_date = ("march 20" in body_lower or "2026-03-20" in body_lower or
                "3/20" in body_lower or "march 20, 2026" in body_lower)
    check("Email body mentions March 20",
          has_date,
          "Date not found in email body")

    # Check body mentions time
    has_time = ("12:00" in body_lower or "12 pm" in body_lower or
                "1:30" in body_lower or "1:30 pm" in body_lower or
                "13:30" in body_lower)
    check("Email body mentions event time",
          has_time,
          "Time not found in email body")

    # Check body mentions location
    has_location = "kitchen" in body_lower or "company kitchen" in body_lower
    check("Email body mentions Company Kitchen",
          has_location,
          "Kitchen/location not found in email body")

    # Body must mention both selected dishes
    if selected_dishes:
        missing = [d for d in selected_dishes if d not in body_lower]
        check("Email body mentions both selected dishes",
              len(missing) == 0,
              f"Missing: {missing}")

    # Check from address - must equal 'hr@company.com' (tight)
    if from_addr:
        from_str = str(from_addr).lower()
        # Handle JSON array or plain string
        import re as _re
        from_addrs = _re.findall(r'[\w.+-]+@[\w.-]+', from_str)
        # Check canonical form present
        ok = "hr@company.com" in from_addrs or from_str.strip() == "hr@company.com"
        check("Email sent from exactly hr@company.com",
              ok,
              f"From: '{from_addr}' (extracted: {from_addrs})")

    # Check to address - must equal 'all-staff@company.com' (tight)
    if to_addr:
        to_str = str(to_addr).lower()
        import re as _re
        to_addrs = _re.findall(r'[\w.+-]+@[\w.-]+', to_str)
        ok = ("all-staff@company.com" in to_addrs or
              "all_staff@company.com" in to_addrs or
              to_str.strip() == "all-staff@company.com")
        check("Email sent to exactly all-staff@company.com",
              ok,
              f"To: '{to_addr}' (extracted: {to_addrs})")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    selected_dishes = check_google_form() or set()
    check_calendar(selected_dishes)
    check_notion(selected_dishes)
    check_emails(selected_dishes)

    total = PASS_COUNT + FAIL_COUNT
    pass_rate = PASS_COUNT / total if total > 0 else 0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Runtime-inconclusive: {RUNTIME_INCONCLUSIVE}")
    print(f"  Pass Rate: {pass_rate:.1%}")
    print(f"  Overall: {'PASS' if FAIL_COUNT == 0 else 'FAIL'}")

    if args.res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "pass_rate": round(pass_rate, 3),
            "success": FAIL_COUNT == 0,
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
