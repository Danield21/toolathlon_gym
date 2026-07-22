I need help with a market sentiment analysis. There is external benchmark data available that I need you to fetch from http://localhost:30337/api/data.json and extract the relevant metrics. The benchmark feed lists each stock symbol along with its analyst target price and rating.

For every symbol that appears in the benchmark feed, pull the current market price from the financial market data service.

Write and run a Python script called yf_sentiment_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs yf_sentiment_results.json.

Create an Excel file called Market_Sentiment_Report.xlsx with three sheets. The first sheet Data_Analysis should have columns Symbol, Current_Price, Target_Price, and Upside, with one row per stock symbol covered by the benchmark feed, sorted by Symbol ascending. Define Upside = Target_Price - Current_Price, rounded to 2 decimal places (an absolute dollar amount, not a percentage). The second sheet Metrics should have two columns Metric and Value with at least these metrics: Total_Stocks (the count of symbols in Data_Analysis), Avg_Upside (the mean of the Upside column rounded to 2 decimal places), and Best_Opportunity (the Symbol with the largest Upside value). The third sheet Recommendations should have columns Priority and Action, listing at least two actionable items where each Action is one of Buy, Hold, or Sell.

Create a page in the team knowledge base titled "Market Sentiment Dashboard" with a summary of the analysis.
