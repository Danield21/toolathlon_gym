You are an HR compensation analyst. Your task is to analyze salary distribution across departments and prepare a compensation review package.

Query the HR employee database to calculate salary statistics for each department. You need headcount, minimum salary, maximum salary, average salary, and median salary per department.

Create an Excel file named Salary_Analysis.xlsx in the workspace with two sheets:

Sheet named Department_Stats with columns: Department, Headcount, Min_Salary, Max_Salary, Avg_Salary, Median_Salary. Include one row per department (7 departments total). Sort rows by Avg_Salary descending.

Sheet named Summary with two columns: Metric and Value. Include these rows: Total_Employees (total headcount across all departments), Company_Avg_Salary (the weighted average salary across all employees rounded to 2 decimal places), Highest_Paid_Dept (name of department with highest avg salary), Lowest_Paid_Dept (name of department with lowest avg salary).

Note: all numeric values in Salary_Analysis.xlsx must be written as literal numeric values (numbers or numeric text), not as Excel formulas.

Create a Google Forms survey titled Compensation Satisfaction Survey with exactly 5 questions. The first question asks how satisfied employees are with their current compensation and offers multiple choice answers ranging from Very Satisfied to Very Dissatisfied. The second asks whether their pay is competitive compared to industry standards with Yes, No, or Not Sure options. The third asks what benefits matter most to them as a short answer text question. The fourth asks whether they would consider leaving for better pay with Yes, No, or Maybe options. The fifth invites any additional comments about compensation as an open-ended text question.

Send an email from compensation@hr.example.com to hr-leadership@company.example.com with the subject Compensation Analysis Report - Action Required. The body must mention the company average salary and identify the highest and lowest paid departments.

The Compensation_Policy.pdf file in your workspace contains compensation philosophy guidelines for reference.

To solve this task efficiently, complete it directly in the main agent and dispatch no sub-agent. The source is one compact employee table, the form has five fixed questions, and there is only one workbook writer.

Resolve the runtime workspace dynamically and read `Compensation_Policy.pdf` completely. Discover and verify the employee relation in `HR_ANALYTICS.PUBLIC` and its `DEPARTMENT`/`SALARY` columns, then run one set-based department query returning headcount, minimum, maximum, average, and median salary for every department plus one independent company query returning total headcount and total salary. Require exactly seven nonblank departments and reconcile their headcounts and salary sums to the company controls. Compute `Company_Avg_Salary = total_salary / total_headcount`; rank by unrounded department average descending with department-name ties, and round only displayed salary values to two decimals.

The main agent is the sole owner of `Compensation Satisfaction Survey`. The clean Forms state exposes no list/search operation: call `create_form` exactly once; if it fails, report the blocker and do not issue a second create. Reuse the returned `formId` and add questions sequentially in this exact order: `add_multiple_choice_question` for current-compensation satisfaction with options `Very Satisfied, Satisfied, Neutral, Dissatisfied, Very Dissatisfied`; `add_multiple_choice_question` for pay competitiveness with `Yes, No, Not Sure`; `add_text_question` for benefits that matter most; `add_multiple_choice_question` for leaving for better pay with `Yes, No, Maybe`; and `add_text_question` for additional compensation comments. Preserve each question's task wording and required flag, never parallelize additions to one form, then call `get_form(formId)` and verify title, five-question order, three choice/two text types, and every option.

Create `Salary_Analysis.xlsx` once with exact `Department_Stats` headers `Department,Headcount,Min_Salary,Max_Salary,Avg_Salary,Median_Salary`, seven rows sorted by raw `Avg_Salary` descending, and `Summary(Metric,Value)` containing exactly `Total_Employees,Company_Avg_Salary,Highest_Paid_Dept,Lowest_Paid_Dept`. Write literal values only and read both full ranges back.

For email idempotency, call `search_emails(query="Compensation Analysis Report - Action Required", folder="Sent", page=n, page_size=50)` through the returned final page; because search rows do not prove recipients, call `read_email` on every candidate and match exact From `compensation@hr.example.com`, To `hr-leadership@company.example.com`, and subject. If and only if absent, call `send_email` exactly once with those fields and a body containing the verified company average and highest/lowest departments. Repeat search/read-back and require exactly one exact message. Reconcile workbook, form, and email once, then finish without a Wave.
