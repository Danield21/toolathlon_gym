Your finance team must conduct a quarterly budget variance analysis to monitor spending against approved budgets and ensure financial accountability. The current fiscal year is FY2024 and the first quarter (Q1) covers January–March 2024. Work from the two source files provided in your workspace:

- `approved_budget.xlsx` — approved budgets for all departments and cost centers for FY2024 (sheet `Annual Budget`). Use the `Q1 Budget` column as the budget baseline for this analysis.
- `q1_actual_expenditures.csv` — every expenditure transaction completed in Q1 FY2024 (January–March 2024), with columns `transaction_id, transaction_date, cost_center, department, account_code, account_description, spending_category, amount`.

Work through the six phases below. All dollar figures are USD and must be written into your outputs as literal numeric values (do not write Excel formulas such as `=SUM(...)`).

**Phase 1 — Data extraction.** Extract the approved Q1 budget for all departments and cost centers from `approved_budget.xlsx`. Extract actual spending for all transactions in Q1 from `q1_actual_expenditures.csv`, aggregated by department, cost center, account code, and spending category. Sum the transaction amounts per (cost center, spending category) to get Q1 actual spending.

**Phase 2 — Variance analysis.** Calculate variance amounts and percentages for each department and major cost center. Use the following conventions consistently:
- `Variance $ = Actual − Budget`. A negative variance means spending was under budget (favorable); a positive variance means spending exceeded budget (unfavorable); a variance near zero is "on budget".
- `Variance % = Variance / Budget × 100`.
- Mark each category `Favorable`, `Unfavorable`, or `On Budget` accordingly.
- Flag variances that exceed a materiality threshold of 5% of budget or $10,000.

Segment variance analysis by department, by spending category, and by cost center. At least one favorable and one unfavorable category must be identified.

**Phase 3 — Root cause investigation.** For each significant variance, document a plausible explanation (timing differences, unexpected price increases, higher-than-anticipated volume, unbudgeted projects, etc.). Assess whether each variance is a permanent change requiring budget adjustment or a temporary fluctuation expected to reverse.

**Phase 4 — Reporting.** Prepare detailed schedules showing budget versus actual for all departments with variance calculations, narratives explaining each significant variance, and visualizations showing variance trends by department and category.

**Phase 5 — Forecast update.** Using the Q1 results, revise the forecasts for the remaining three quarters (Q2–Q4 FY2024). Create a revised annual budget forecast (incorporating Q1 actuals) under at least two scenarios (for example, a base case following current trends and a conservative case with cost controls), each with per-department projected full-year spending and a total.

**Phase 6 — Communication.** Prepare a summary report per department showing its specific variances and forecast. Send a detailed variance report to senior management by email, and schedule a budget review meeting with department leaders on the calendar.

## Required deliverables

Produce all five files below with these exact filenames, sheet names, and column layouts. Extra sheets, columns, or rows are allowed; the required columns must be present.

1. **`variance_analysis.xlsx`** — sheet `Variance Analysis`, columns:
   `Cost Center | Department | Category | Q1 Budget | Q1 Actual | Variance $ | Variance % | Status`
   One row per (cost center, spending category), covering all departments and categories from the source files. You may add a summary section (e.g., total budget, total actual, total variance) below the detail rows.

2. **`variance_tracking.xlsx`** — sheet `Variance Tracking`, columns:
   `Department | January | February | March | Q1 Variance | Trend`
   One row per department showing its monthly Q1 variance (actual minus budget) and the quarter total, plus a trend/assessment column.

3. **`budget_forecast.xlsx`** — sheet `Budget Forecast`, with a scenario table whose header row contains `Scenario` and per-department columns (e.g., Operations, Sales, Marketing, IT) plus a `Total`, and at least two scenario rows with numeric full-year projections.

4. **`dept_variance_reports.docx`** — a Word report (at least 100 words) with per-department sections covering variance details, root causes, and mitigation strategies.

5. **`executive_presentation.pptx`** — a PowerPoint deck (at least 3 slides) summarizing Q1 results, departmental performance, and strategic recommendations.

Additionally:
- **Email**: send a summary of the variance analysis to senior management (email subject and body should reference the quarterly budget variance analysis).
- **Calendar**: schedule a budget review meeting with department leaders.
