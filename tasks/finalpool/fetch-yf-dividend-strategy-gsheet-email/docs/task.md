I need help with a portfolio dividend analysis. There is external benchmark data available that I need you to fetch from http://localhost:30316/api/data.json and extract the relevant metrics.

Then pull current market data for the stocks in our portfolio.

Write and run a Python script called yf_dividend_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs yf_dividend_results.json.

Create an Excel file called Dividend_Strategy_Report.xlsx with three sheets. The first sheet Data_Analysis should have columns Symbol, Current_Price, Target_Price, and Upside, with one row per stock symbol covered by the analysis, sorted by Symbol ascending. The second sheet Metrics should have two columns Metric and Value with at least these metrics: Total_Stocks, Avg_Upside, and Best_Opportunity. The third sheet Recommendations should have columns Priority and Action and list at least two actionable items.

Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings. Create a cloud spreadsheet titled "Portfolio Dividend Tracker" with the key data points.
