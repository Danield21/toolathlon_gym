You are a portfolio manager at a mid-sized investment firm. Your firm tracks five core holdings: GOOGL, AMZN, JPM, JNJ, and XOM. A research service has published its latest macroeconomic indicator forecasts, which you need to incorporate into your portfolio analysis. Read the Investment_Strategy.pdf file in your workspace first to understand the firm's approach to macro-sensitive portfolio management.

The research service API is available at http://localhost:30211/api/macro_forecast.json and provides GDP growth, inflation rate, and interest rate forecasts by quarter for 2026. Fetch this data to understand the macro outlook.

After retrieving the macro forecast, query the financial data source for current stock information on all five holdings, including their current price, sector, market capitalization, P/E ratio, and beta values.

Create an Excel workbook called Macro_Portfolio_Analysis.xlsx in your workspace with three sheets.

The first sheet should be called "Macro Forecast" and contain columns: Quarter, GDP_Growth_Pct, Inflation_Rate_Pct, Interest_Rate_Pct. Populate it with the four quarters of data from the research API (Q1 2026 through Q4 2026).

The second sheet should be called "Stock Holdings" and contain columns: Symbol, Sector, Price, Market_Cap_B (market cap in billions rounded to 1 decimal), PE_Ratio (rounded to 1 decimal), Beta. List all five stocks sorted alphabetically by Symbol.

The third sheet should be called "Sector Sensitivity" and contain columns: Sector, Avg_GDP_Growth, Rate_Sensitivity, Recommended_Action. For each unique sector among the five holdings, compute the average annual GDP growth from the macro forecast (average of the 4 quarters, rounded to 1 decimal), assign a Rate_Sensitivity label based on the sector's beta values (High if average beta for that sector is above 1.0, Low otherwise), and set Recommended_Action to "Overweight" if Rate_Sensitivity is High and average GDP growth is above 2.0 percent, "Underweight" if Rate_Sensitivity is Low and average GDP growth is below 2.0 percent, and "Hold" for all other cases. Sort by Sector alphabetically.

Schedule a Google Calendar event called "Q2 Portfolio Rebalancing Review" on March 30, 2026 from 14:00 to 15:30. In the description, list the five holdings and note the average GDP growth forecast and the overall recommended portfolio stance based on the majority of sector recommendations.

Send an email to investment_committee@firm.com with the subject "Macro Outlook & Portfolio Impact - Q2 2026". The body should summarize the macro forecast highlights (average GDP growth, inflation trend, interest rate direction) and list any sectors recommended for Overweight or Underweight adjustment.

When you are finished, call claim_done.

To solve this task efficiently, the main agent must complete this compact fixed-five-symbol workflow directly; do not dispatch sub-agents. Resolve the accessible task workspace dynamically.

Read Investment_Strategy.pdf first. Then fetch http://localhost:30211/api/macro_forecast.json and issue get_stock_info for GOOGL, AMZN, JPM, JNJ, and XOM together in one native-parallel response. Require exactly four ordered 2026 quarter rows and exactly one market record per literal ticker with numeric price, market capitalization, P/E ratio and beta plus source-returned sector. Freeze one canonical payload containing Macro Forecast rows, alphabetically sorted Stock Holdings rows, annual average GDP growth from all four unrounded quarters, one Sector Sensitivity row per returned sector, and the deterministic majority portfolio stance. Apply the task's strict beta/GDP rules and rounding exactly once.

The main agent is the sole writer of Macro_Portfolio_Analysis.xlsx in the resolved workspace. Create exactly Macro Forecast, Stock Holdings, and Sector Sensitivity with exact headers, literal values and sort order, then read every populated range back.

Once the canonical payload is frozen, start the workbook branch, exact calendar search, and exact Sent-mail search natively in parallel. The calendar contract is summary Q2 Portfolio Rebalancing Review, start 2026-03-30T14:00:00 and end 2026-03-30T15:30:00. The task supplies no timezone; omit the timeZone field and allow the calendar service's UTC default. Never substitute America/New_York or another timezone. Create only if the exact event is absent, include all five holdings, average GDP growth and majority stance in the description, and read it back. For To investment_committee@firm.com / subject Macro Outlook & Portfolio Impact - Q2 2026, read Sent candidates, send once only if absent with the canonical GDP/inflation/rate highlights and Overweight/Underweight sectors, then read it back.

Finish only after workbook, event and email agree with the same payload and call claim_done. Do not retry an ambiguous external write.
