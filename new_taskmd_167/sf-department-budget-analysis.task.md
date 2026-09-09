I need to review how each department is performing against their allocated budget. There is a PDF file called Budget_Targets.pdf in the workspace that shows the approved annual budget for each department. Please read that first.

Then pull the actual employee data from our HR analytics data warehouse. I need to see for each department the planned headcount from the department table, the actual employee count, average salary, and total salary cost. Also calculate a Budget_Utilization_Pct showing total salary cost as a percentage of the department budget, rounded to 1 decimal.

Create an Excel file called Department_Budget_Report.xlsx with two sheets. The first sheet "Budget Analysis" should have columns Department, Budget, Planned_Headcount, Actual_Headcount, Avg_Salary rounded to 2 decimals, Total_Salary_Cost rounded to 2 decimals, and Budget_Utilization_Pct rounded to 1 decimal. Sort by Department alphabetically.

The "Summary" sheet should have Metric and Value columns with Total_Budget, Total_Salary_Cost, Avg_Budget_Utilization rounded to 1 decimal (computed as the average of the per-department Budget_Utilization_Pct values), and Over_Budget_Depts counting departments where utilization exceeds 100 percent.

To solve this task efficiently, complete it directly in the main agent; do not dispatch sub-agents. Read `Budget_Targets.pdf` and issue the independent set-based department aggregate in the same native-parallel response. Join only by exact department label, preserve raw values until final rounding, and create `Department_Budget_Report.xlsx` with the exact two sheets, headers, alphabetical department order, literal values, and requested summary metrics. Use one bulk write per sheet and one final workbook read-back; do not add a separate verification phase.
