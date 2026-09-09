I need help with a portfolio dividend analysis. There is external benchmark data available that I need you to fetch from http://localhost:30316/api/data.json and extract the relevant metrics.

Then pull current market data for the stocks in our portfolio.

Write and run a Python script called yf_dividend_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs yf_dividend_results.json.

Create an Excel file called Dividend_Strategy_Report.xlsx with three sheets. The first sheet Data_Analysis should have columns Symbol, Current_Price, Target_Price, and Upside, with one row per stock symbol covered by the analysis, sorted by Symbol ascending. The second sheet Metrics should have two columns Metric and Value with at least these metrics: Total_Stocks, Avg_Upside, and Best_Opportunity. The third sheet Recommendations should have columns Priority and Action and list at least two actionable items.

Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings. Create a cloud spreadsheet titled "Portfolio Dividend Tracker" with the key data points.

To solve this task efficiently, complete it directly in the main agent; do not dispatch sub-agents. Fetch `http://localhost:30316/api/data.json` exactly once, retain its five GOOGL, AMZN, JPM, JNJ, and XOM target-price rows, and issue the five independent Yahoo Finance current-price lookups as native-parallel calls. Write the two source JSON files, then write and run `yf_dividend_processor.py` once. The script must produce `yf_dividend_results.json` with one row per symbol and compute `Upside` in percentage points as `(Target_Price - Current_Price) / Current_Price * 100`, not as a 0-to-1 fraction. From that one verified result, create `Dividend_Strategy_Report.xlsx` and the cloud spreadsheet `Portfolio Dividend Tracker` as independent output branches, using the exact sheet names, columns, metrics, literal values, and Symbol-ascending order from the task. Verify each output once, then send exactly one email to `team-lead@company.com` with subject `Analysis Report Complete`.
