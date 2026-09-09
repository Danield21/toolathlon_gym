I need to prepare a department budget presentation for our upcoming board meeting. Please pull the current department headcount and budget data from the company HR data warehouse and create a PowerPoint presentation called HR_Department_Overview.pptx. The presentation should have at least seven slides: a title slide called HR Department Budget Overview, one slide per department showing the department name, headcount, budget, and location, and a final summary slide showing total headcount and total budget across all departments. Use the department names exactly as they appear in the data source (for example `HR` and `R&D`) — do not expand or rephrase them.

After creating the presentation, also create an Excel file called Department_Budget_Analysis.xlsx with a sheet called Department Summary that has columns Department, Headcount, Budget rounded to 2 decimals, Location, and Budget_Per_Employee which is the budget divided by headcount rounded to 2 decimals, sorted alphabetically by department name. Add a second sheet called Summary with Metric and Value columns showing Total_Headcount, Total_Budget, Avg_Budget_Per_Employee (computed as Total_Budget divided by Total_Headcount) rounded to 2 decimals, and Highest_Budget_Dept. Use these exact metric labels (with underscores) as the Metric values.

Then create a Google Sheet named 'HR Department Tracker' (create it if it does not already exist) and populate it with the same department data, and send an email to hr.directors@company.com with subject HR Department Budget Summary Q1 2026 containing a brief overview of the findings.

To solve this task efficiently, the main agent should first query the complete department table once, preserve exact department names such as HR and R&D, compute headcount and budget controls, and freeze alphabetically sorted rows. Then create the three independent deliverables in one wave.

1. Wave 1 — dispatch exactly 3 coder sub-agents in parallel:

   1. One coder sub-agent owns `HR_Department_Overview.pptx` under one presentation ID. Use the literal source labels as slide text—especially the exact three-character string `R&D`; never expand it to Research and Development. Create the title slide, one slide per returned department with exact name/headcount/budget/location, and the final numeric summary. Use one bounded populate operation per slide, save once, extract the full deck text once, and return the path, slide count/titles, literal department-name coverage, and total checks.

   2. One coder sub-agent owns `Department_Budget_Analysis.xlsx` and writes the exact Department Summary and Summary sheets with literal values, alphabetical rows, Budget_Per_Employee, and the exact underscore metric labels. Use one bulk write per sheet and one read-back.

   3. One coder sub-agent owns exactly one `HR Department Tracker` cloud spreadsheet. Search the exact title, create only if absent, write the same literal source labels and rows once, and read it back once.

After Wave 1, the main agent checks that all three outputs contain every exact department label including `R&D` and that the totals agree, then sends exactly one email to `hr.directors@company.com` with subject `HR Department Budget Summary Q1 2026`. Do not run row-by-row revalidation or dispatch another wave.
