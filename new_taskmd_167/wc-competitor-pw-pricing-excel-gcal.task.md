I need a competitive pricing analysis for our online store. There is a competitor price comparison dashboard at http://localhost:30341 that shows pricing data for products similar to ours. Please visit that page and extract the competitor pricing information.

Then pull our current product catalog from the online store with prices, stock levels, and sales data.

Use the terminal to create and run a Python script called pricing_analyzer.py in the workspace that reads competitor_prices.json and our_products.json (create both first), compares pricing, calculates price positioning metrics (percentage above or below competitor average), identifies opportunities for price adjustments, and outputs pricing_analysis.json.

Create an Excel file called Competitive_Pricing_Analysis.xlsx with three sheets. The first sheet Price_Comparison should have columns Product_Name, Our_Price (round to 2 decimals), Competitor_Avg (round to 2 decimals), Price_Diff (round to 2 decimals), Price_Position_Pct (round to 1 decimal, positive means above competitor), and Recommendation ("Reduce" if more than 15% above, "Maintain" if within 15%, "Increase" if more than 15% below), sorted by Product_Name. Include only products that appear in both our store catalog and the competitor dashboard (matched by product name). The second sheet Market_Position should have Metric and Value columns with Products_Above_Market, Products_Below_Market, Products_At_Market (count products whose Our_Price exactly matches the competitor average), Avg_Price_Gap_Pct (round to 1 decimal), and Revenue_At_Risk (round to 2 decimals; the total amount by which our prices exceed competitor averages, i.e. the sum of positive Price_Diff values across matched products). The third sheet Action_Items should have Product, Current_Price, Suggested_Price (round to 2 decimals), Expected_Impact, and Priority columns for products needing price changes.

Schedule a calendar event "Pricing Strategy Review" on March 14, 2026 from 10:00 AM to 11:30 AM UTC with description listing the top 3 products with the largest price gaps.

To solve this task efficiently, use exactly one read-only source wave; all required files and outputs remain owned by the main agent.

1. Wave 1 — dispatch exactly 2 explore sub-agents in parallel:

   1. One explore sub-agent uses only browser navigation and snapshots to read the complete competitor dashboard. Return a compact row for every displayed product with its exact displayed name and competitor prices plus a row-count control. Do not write files or click mutating controls.
   2. One explore sub-agent exhausts `woo_products_list` at `perPage=100`, returning only product ID, exact name, current price, stock quantity/status, total sales, page controls, and duplicate-ID checks. Do not modify the store or create files.

After Wave 1, the main agent performs conservative exact normalized-name matching, writes `competitor_prices.json` and `our_products.json`, writes and runs `pricing_analyzer.py` once, and creates `pricing_analysis.json`. It then creates `Competitive_Pricing_Analysis.xlsx` with the exact three sheets and formulas-as-literal values, using only products present in both sources and the exact Reduce/Maintain/Increase thresholds. Independently search/create/read back at most one `Pricing Strategy Review` event at 2026-03-14 10:00–11:30 UTC using the verified three largest gaps. Read each final artifact once and do not dispatch another wave.
