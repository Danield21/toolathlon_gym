You are an investment assistant helping an investor track upcoming dividend payments for their stock portfolio.

The investor holds positions in several stocks and wants to identify which ones pay dividends, then organize dividend information for easy tracking.

First, look up stock information for the following tickers: GOOGL, AMZN, JPM, JNJ, XOM. For each stock, check whether it pays a dividend by looking at the dividend rate. Only include stocks that have a dividend rate greater than zero.

For each dividend-paying stock, collect the following information: ticker symbol, company name (short name), dividend rate, dividend yield (as a percentage), ex-dividend date, and payout ratio.

Create an Excel file called "Dividend_Tracker.xlsx" in your workspace with two sheets:

The first sheet should be named "Dividend Stocks" and contain columns: Ticker, Company, Dividend_Rate, Dividend_Yield, Ex_Dividend_Date, Payout_Ratio. Include one row per dividend-paying stock, sorted by ticker symbol alphabetically.

The second sheet should be named "Summary" with the following rows of information: Total_Dividend_Stocks (count of stocks with dividends), Avg_Yield (average dividend yield across dividend stocks), Highest_Yield_Ticker (the ticker with the highest dividend yield).

Next, create a calendar event for each dividend-paying stock to remind the investor of the ex-dividend date. Each event should have the summary "Ex-Dividend: [TICKER]" (replacing [TICKER] with the actual ticker symbol), and the description should mention the dividend rate. Schedule each event as an all-day event on the ex-dividend date.

Finally, create a cloud spreadsheet titled "Dividend Watch List" with a sheet named "Overview" that contains the same columns and data as the "Dividend Stocks" sheet in the Excel file.

To solve this task efficiently, complete it directly in the main agent and dispatch no sub-agent. Five fixed stock-info lookups and a small dividend table do not justify a delegation wave.

Resolve the workspace dynamically. In one native-parallel response call `get_stock_info` exactly once for each of `GOOGL`, `AMZN`, `JPM`, `JNJ`, and `XOM`. Preserve returned short name, dividend rate, dividend yield, ex-dividend date, and payout ratio with their source units. Include only rows whose numeric dividend rate is greater than zero; normalize dividend yield to percentage units exactly once, and do not multiply a value already returned as a percent. Sort by ticker and freeze one canonical table plus Total_Dividend_Stocks, numeric Avg_Yield, and Highest_Yield_Ticker with deterministic tie handling.

The main agent is the sole writer of both `Dividend_Tracker.xlsx` and the cloud spreadsheet `Dividend Watch List`. Call `search_spreadsheets(query="Dividend Watch List", max_results=100)` and compare complete returned names: reuse one exact-title match, create exactly once only on zero matches, and block cloud writes on multiple exact matches or an ambiguous create failure. Create the local workbook once, create or reuse exactly one cloud `Overview` sheet, write the exact sheets/headers from the canonical frozen rows, and read their bounded ranges back; both data tables must match cell-for-cell in ticker and values, and retain the verified spreadsheet ID. For each returned ex-dividend civil date, issue `list_events` natively in parallel with ISO midnight `timeMin`, next-civil-midnight `timeMax`, sufficient `maxResults`, and `orderBy="startTime"`; inspect exact title/date locally. Create only missing reminders natively in parallel with `create_event` using `start.dateTime="[date]T00:00:00"`, next-day `end.dateTime`, no timezone-driven date shift, exact `Ex-Dividend: [TICKER]` summary, and verified rate description; this midnight span is the all-day representation supported by the actual schema. Retrieve each ID with `get_event`. Read back both spreadsheets and the final event inventory, complete any original requirement, and call `claim_done` if available.
