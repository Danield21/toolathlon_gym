The HR leadership team is requesting a compensation analysis across all departments in the company. Please connect to the data warehouse and pull employee records from the human resources database, then calculate key salary statistics for each department.

Create an Excel file called Compensation_Report.xlsx in the workspace with two sheets. The first sheet should be named "Department Compensation" and contain columns Department, Headcount, Avg_Salary (average salary rounded to two decimal places), Min_Salary (rounded to two decimal places), Max_Salary (rounded to two decimal places), and Total_Payroll (sum of all salaries rounded to two decimal places). Sort the rows alphabetically by department name.

The second sheet should be named "Summary" with columns Metric and Value. Include the following entries: Total_Employees (total headcount across all departments), Total_Payroll (total payroll across all departments rounded to two decimal places), Company_Avg_Salary (overall average salary rounded to two decimal places), Highest_Avg_Department (department with the highest average salary), and Lowest_Avg_Department (department with the lowest average salary).

Also create a Google Sheet spreadsheet titled "Workforce Compensation Dashboard" with a sheet called "Overview" containing the same Department Compensation data so the HR team can collaborate on it online.

To solve this task efficiently, use one true dependency wave. The department aggregate is compact, so the main agent performs it directly instead of inserting a one-agent source wave.

Inspect the HR schema, confirm `HR_ANALYTICS.PUBLIC.EMPLOYEES` and its `DEPARTMENT` and `SALARY` fields, then run one server-side department aggregate over the full table plus one company control query. Freeze alphabetically ordered Department, Headcount, raw Avg_Salary, Min_Salary, Max_Salary, Total_Payroll; company headcount/payroll and `Company_Avg_Salary=company payroll/company headcount`; and deterministic highest/lowest raw-average departments with alphabetical ties. Require department headcounts and payrolls to reconcile exactly to company controls and return no employee rows.

Wave 1 — dispatch exactly two coder sub-agents in parallel in the same response:
1. The Excel coder is the sole writer of Compensation_Report.xlsx in the dynamically discovered task workspace. It creates exactly Department Compensation and Summary with the requested headers/order, literal rounded numbers, reads both populated ranges back, and returns row/count/value controls.
2. The Google-Sheet coder is the sole owner of exact title Workforce Compensation Dashboard. It searches/reuses one exact-title spreadsheet or creates once, blocks on duplicates, creates or reuses a single Overview tab, batch-writes the exact Department Compensation rows from the canonical packet, and reads the complete range back.

Each writer must correct its own artifact before returning. The main agent only compares the two receipts to the frozen packet; it must not launch sequential repair agents or reread either artifact.
