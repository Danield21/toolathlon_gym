I need help with a portfolio health analysis. There is external benchmark data available that I need you to fetch from http://localhost:30314/api/data.json and extract the relevant metrics.

Then pull current market data for the stocks in our portfolio.

Write and run a Python script called yf_portfolio_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs yf_portfolio_results.json.

Create an Excel file called Portfolio_Health_Report.xlsx with three sheets. The first sheet Data_Analysis should have columns Symbol, Current_Price, Target_Price, and Upside, with one row per stock symbol covered by the analysis, sorted by Symbol ascending. The second sheet Metrics should have two columns Metric and Value with at least these metrics: Total_Stocks, Avg_Upside, and Best_Opportunity. The third sheet Recommendations should have columns Priority and Action and list at least two actionable items.

Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings. Add an event to the shared calendar titled "Analysis Review" on March 14, 2026 from 2:00 PM to 3:00 PM UTC.
