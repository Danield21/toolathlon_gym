I need to create a comprehensive sector analysis report combining financial data with academic research insights. Pull financial data for stocks across different sectors, focusing on their sector classifications, market performance, and key financial metrics.

Also search for scholarly papers related to "sector rotation", "industry analysis", and "market cycles" to understand academic perspectives on sector performance patterns.

Use the terminal to create and run a Python script called sector_analyst.py in the workspace that reads financial_data.json and research_findings.json (create both first), analyzes sector performance metrics, maps academic findings to real market data, and outputs sector_analysis.json.

Create an Excel file called Sector_Analysis_Report.xlsx with three sheets. The first sheet Sector_Performance should have columns Sector, Stock_Count, Avg_Price (round to 2 decimals), Total_Market_Value (round to 2 decimals), and Volatility_Score (round to 2 decimals), sorted by Sector. The second sheet Research_Mapping should have columns Paper_Title, Key_Finding, Applicable_Sector, Validation_Status ("Confirmed", "Partial", or "Inconclusive"), sorted by Paper_Title. The third sheet Investment_Thesis should have Sector, Outlook ("Bullish", "Neutral", or "Bearish"), Supporting_Evidence, and Risk_Factor columns.

Create a Word document called Sector_Research_Brief.docx with heading "Cross-Disciplinary Sector Analysis", sections for "Financial Performance Review", "Academic Research Insights", "Theory vs Practice Comparison", and "Investment Implications" with specific data points and paper references.

For the Excel file, write every numeric cell as a literal value (e.g., 190.25 or 2.5), not as an Excel formula. Compute Avg_Price as the average of each stock's latest available closing price within that sector (use the most recent close for every stock, not a historical window average), Total_Market_Value as the combined market capitalization of the sector's stocks, and Volatility_Score as a volatility measure (e.g., the standard deviation of daily returns) of the sector's stocks; use one consistent time window for all metrics.

These deliverables are independent files (sector_analyst.py, Sector_Analysis_Report.xlsx, Sector_Research_Brief.docx, and the two input JSON files financial_data.json and research_findings.json). Have different workers handle different files, and never have two workers write to the same file at the same time.
