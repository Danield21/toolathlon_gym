I need a job satisfaction analysis by department from our HR data warehouse. Include both job satisfaction and work-life balance scores along with performance ratings.

Create an Excel file called HR_Satisfaction_Report.xlsx with two sheets. The "Satisfaction Analysis" sheet should have Department, Avg_Satisfaction, Avg_Work_Life_Balance, Avg_Rating (all rounded to 2 decimals), and Employees count. Sort by raw (un-rounded) Avg_Satisfaction descending; if two departments are equal at the displayed (2-decimal) rounding, the row order should still reflect the raw-value descending sort.

The "Summary" should have Overall_Satisfaction (overall mean of job satisfaction across all employees, rounded to 2 decimals), Overall_WLB (overall mean of work-life balance, rounded to 2 decimals), Happiest_Dept (department with the highest raw Avg_Satisfaction), and Least_Happy_Dept (department with the lowest raw Avg_Satisfaction).

Also create a Word document called Satisfaction_Summary.docx (at least 200 characters) with a narrative covering employee satisfaction, work-life balance, and the highest- and lowest-ranked departments.

To solve this task efficiently, complete it directly in the main agent. Dispatching sub-agents is not necessary. Run the compact department satisfaction aggregate, freeze the verified rows and summary values, then create HR_Satisfaction_Report.xlsx and Satisfaction_Summary.docx with native parallel tool calls in the same response and reconcile both outputs once.
