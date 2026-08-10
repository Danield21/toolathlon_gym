I need help with a sf efficiency analysis. There is external benchmark data available that I need you to visit http://localhost:30334 and extract the relevant metrics.

Then pull our internal data from the company data warehouse for comparison.

Use the terminal to create and run a Python script called sf_efficiency_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs sf_efficiency_results.json.

Create an Excel file called Hr_Dept_Efficiency_Report.xlsx with three sheets. The first sheet Data_Analysis should contain the main comparison data with relevant columns. The second sheet Metrics should summarize key metrics. The third sheet Recommendations should list actionable items.

The Data_Analysis sheet should include columns for the primary dimension, our internal metric values, the external benchmark values, and the gap or difference between them. Sort the data alphabetically by the primary dimension. The Metrics sheet should have two columns Metric and Value summarizing total counts, averages, and key statistics. The Recommendations sheet should list priority actions based on the gap analysis. Schedule a review meeting titled "Analysis Review" on March 14, 2026 from 2:00 PM to 3:00 PM UTC.

Use the department dimension for the comparison (one row per department), and name that first column exactly "Department". The Recommendations sheet must have two columns: "Priority" (the rank or priority level of each action) and "Action" (the recommended action text). Write all values into the Excel cells as literal numbers or text rather than spreadsheet formulas.
