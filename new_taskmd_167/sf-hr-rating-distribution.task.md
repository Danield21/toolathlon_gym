I need a breakdown of employee performance ratings across all departments. Pull the data from our HR data warehouse and show me how many employees fall into each rating level within each department.

Create an Excel file called HR_Rating_Distribution.xlsx with two sheets. The "Rating Distribution" sheet should have columns Dept_Rating (formatted as "Department - Rating", e.g. "Engineering - 5"), and Count. Sort by Dept_Rating alphabetically.

The "Summary" sheet with Metric and Value should have Total_Employees, Rating_5_Count, Rating_4_Count, Rating_1_Count, and High_Performers_Pct showing the percentage of employees with rating 4 or 5 out of total, rounded to 1 decimal.

After creating the report, send an email to hr-analytics@company.com with subject "Performance Rating Distribution Report" briefly summarizing the overall distribution.

To solve this task efficiently, complete it directly in the main agent; do not dispatch sub-agents. Freeze the task-visible agent-workspace absolute path from the Visible Boundary/current working directory. Run one set-based query grouping the employee source by exact department and rating and returning all group counts plus company-wide counts for total employees, ratings 5, 4, 1, and combined 4-or-5. Build literal `Dept_Rating = Department - Rating`, sort alphabetically, and compute `High_Performers_Pct` from raw counts once.

From that frozen result, create `<agent_workspace>/HR_Rating_Distribution.xlsx` with only the two exact sheets/headers and five summary metrics in the task, literal numeric cells, and the requested one-decimal percentage; in the same native-parallel response send exactly one email to `hr-analytics@company.com` with subject `Performance Rating Distribution Report`. Read both populated ranges back once, reconcile totals, and finish. Do not search Sent folders, retry folder aliases, send a second message, read the sent message back, or add a verification wave.
