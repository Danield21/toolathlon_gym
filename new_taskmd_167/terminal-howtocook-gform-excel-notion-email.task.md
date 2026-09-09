You are a corporate wellness coordinator responsible for launching an employee lunch program. Your goal is to create a dietary preference survey, select recipes that match common employee preferences, plan a weekly menu, build a recipe knowledge base, and communicate the program to staff.

Start by reading the Cafeteria_Budget.pdf in your workspace, which describes the budget per meal and expected headcount. Also review the dietary_requirements.json file that lists common dietary restrictions and allergen considerations for the employee population.

Create a dietary preference survey using the forms platform. The survey should be titled "Employee Lunch Program Preferences" and include the following five questions: a multiple choice question asking about preferred cuisine type with options Chinese Meat Dishes, Chinese Vegetable Dishes, Chinese Staple Foods, Soups, and Seafood; a multiple choice question about spice tolerance with options Mild, Medium, and Spicy; a multiple choice question about dietary restrictions with options Vegetarian, No Pork, No Seafood, Gluten Free, and No Restrictions; a short text question asking how important variety is in their lunch options (rate on a scale of 1 to 5); and a short text question for any additional food preferences or allergies.

Query the recipe database to find recipes across different categories. Search for at least two meat dishes, two vegetable dishes, one staple food, one soup, and one seafood dish. For each recipe retrieved, note the recipe name, category, difficulty level, and ingredients list.

Write a Python script called menu_planner.py in your workspace and execute it using command-line tools. The script should select five recipes for a Monday through Friday weekly menu, choosing recipes that balance different categories (no two consecutive days with the same category) and keeping difficulty at 4 or below for cafeteria feasibility. The script should estimate a cost per serving of 8 dollars for meat and seafood dishes, 5 dollars for vegetable dishes, 6 dollars for staple foods, and 4 dollars for soups, then compute the total weekly cost assuming 50 servings per day.

Create an Excel workbook called Meal_Program_Plan.xlsx in your workspace with four sheets.

The first sheet should be named Survey_Questions and contain columns for question_num, text, and type. Include one row for each of the five survey questions.

The second sheet should be named Recipe_Selection and contain columns for recipe_name, category, and difficulty. Include one row for each recipe you retrieved from the recipe database, with at least seven recipes total. Record the recipe name and category exactly as returned by the recipe database.

The third sheet should be named Weekly_Menu and contain columns for day, lunch_recipe, and estimated_cost. Include five rows, one for each weekday Monday through Friday, with the assigned recipe and estimated cost per serving. Enter all cost figures as literal numbers (for example, 8), not as spreadsheet formulas.

The fourth sheet should be named Program_Summary and contain columns for metric and value. Include rows for total_weekly_cost (sum of daily cost times 50 servings), avg_daily_cost_per_person, survey_question_count (5), recipes_evaluated (total recipes retrieved), and menu_days (5).

Set up a database in the team wiki system called "Recipe Knowledge Base" with properties for Recipe Name (title), Category (select with options for the main recipe categories), Difficulty (number), and Suitable For (multi-select with options Lunch, Dinner, Quick Meal). Add one entry for each recipe you included in the weekly menu.

Send an email to all_staff@company.com with the subject "New Employee Lunch Program - Take Our Survey" describing the new lunch program, mentioning that the weekly menu has been planned with a variety of dishes, and encouraging employees to fill out the dietary preference survey to help improve future menu selections.

To solve this task efficiently, the main agent completes this small five-day menu task directly; do not dispatch sub-agents. Dynamically resolve the workspace, read Cafeteria_Budget.pdf and dietary_requirements.json, and query exact HowToCook records for at least two meat, two vegetable, one staple, one soup and one seafood dish.

Create Employee Lunch Program Preferences exactly once, retain formId, add exactly the five stated questions sequentially in exact order/types/options, and get_form. Never retry an ambiguous create. In parallel with the independent form branch, write/run menu_planner.py over at least seven verified recipe records. Preserve exact database name/category/difficulty/ingredients; select Monday-Friday with difficulty<=4 and no same-category consecutive days; use per-serving costs meat/seafood=8, vegetable=5, staple=6, soup=4 and 50 servings/day. Persist/read one canonical menu result with five rows and arithmetic controls.

After validating form question order, category minimums, feasibility and weekly total, the main agent directly performs the three independent outputs:

- sole-write Meal_Program_Plan.xlsx with exact Survey_Questions, Recipe_Selection, Weekly_Menu and Program_Summary, literal costs/counts and full read-back;
- fully paginate exact-title search for Notion database Recipe Knowledge Base, create/reuse at most once, block duplicates, enforce Recipe Name/Category/Difficulty/Suitable For, upsert exactly the five menu recipes and read them back;
- paginate Sent search/read candidates for all_staff@company.com and exact subject New Employee Lunch Program - Take Our Survey, send at most once with verified form link and menu-variety statement, then read back.

Final acceptance compares external workbook/Notion/email/form read-backs to the canonical menu. No source, writer, auditor or repair agents are allowed.
