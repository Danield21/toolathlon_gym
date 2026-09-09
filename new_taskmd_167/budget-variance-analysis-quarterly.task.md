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

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

Before Wave 1, the main agent must inspect the complete task-visible workspace manifest and read all relevant task-visible mapping, contact, and configuration references in read-only mode without assuming their filenames. It must verify that both source files named by the task—approved_budget.xlsx and q1_actual_expenditures.csv—are actually present and readable with the stated sheet/columns. If either named source is absent or unreadable at runtime, the main agent must stop before delegation, report the exact missing-source blocker, and must not substitute any unrelated workspace file, warehouse table, or other unrequested source or fabricate actual spending. The following waves apply only if both named sources are available.

1. Wave 1 — dispatch exactly 1 sub-agent for the following numbered sub-task:

   1. Use one coder sub-agent to own complete source extraction and one runtime-unique normalized analysis packet. Read all Annual Budget rows and the complete Q1 transaction CSV; validate January–March dates and schemas; aggregate actuals by department, cost center, account code, and spending category; reconcile every budget/category key and transaction amount; compute monthly/quarterly actuals, Variance $ = Actual - Budget, Variance %, status, materiality flags, department/category/cost-center summaries, plausible evidence-bounded cause classifications, and at least two forecast scenarios. Write/read back the normalized packet and return only its path, source hashes, row/key/amount controls, department/category counts, favorable/unfavorable controls, and validation. Do not create any final deliverable or external side effect.

After Wave 1, the main agent must verify source-to-packet totals, sign conventions, literal forecast values, materiality logic, and complete runtime department coverage, then freeze one source-backed packet for all deliverable owners.

2. Wave 2 — dispatch exactly 5 sub-agents in parallel, one for each numbered sub-task:

   1. Use one coder sub-agent to own variance_analysis.xlsx exclusively. Create it exactly once with the Variance Analysis sheet, exact required headers, one complete row per runtime cost-center/category key, literal numeric budget/actual/variance values, and consistent status. Read it back and return its path, header/row/key controls, source reconciliation totals, and validation. Do not write any other deliverable or external state.

   2. Use one coder sub-agent to own variance_tracking.xlsx exclusively. Create it exactly once with the Variance Tracking sheet, exact required headers, one row per runtime department, literal January/February/March and Q1 variance values, and a source-backed trend assessment. Read it back and return its path, header/row/department controls, monthly-to-quarter reconciliation totals, and validation. Do not write any other deliverable or external state.

   3. Use one coder sub-agent to own budget_forecast.xlsx exclusively. Create it exactly once with the Budget Forecast sheet, Scenario plus one column for every runtime department and Total, and at least two department-complete scenario rows with literal numeric full-year projections incorporating Q1 actuals. Read it back and return its path, header/scenario/department controls, row-total reconciliation, and validation. Do not write any other deliverable or external state.

   4. Use one coder sub-agent to own dept_variance_reports.docx exclusively. Create it exactly once from the frozen packet with at least 100 words, one substantive section per runtime department, significant-variance causes, temporary/permanent assessment, mitigations, and forecast implications. Read it back and return path, word/section/departments controls, and validation. Do not write any other deliverable.

   5. Use one coder sub-agent to own executive_presentation.pptx exclusively. Create it exactly once with at least three slides covering Q1 results, departmental performance, and strategic recommendations, using only frozen literal totals and source-backed narratives. Read it back and return path, slide titles/count, metric coverage, and validation. Do not write spreadsheets, Word, email, or calendar state.

After Wave 2, the main agent must reconcile all five files against the frozen totals. It may send the quarterly budget variance analysis message only if task-visible contact data unambiguously identifies the intended senior-management recipient or recipients; otherwise it must report the missing-recipient blocker and not guess an address. Likewise, it may create the budget review meeting only if task-visible evidence supplies a definite date, start/end time, and any required timezone interpretation; otherwise it must report the missing-calendar-parameter blocker and not invent or choose a slot. For each safely resolvable side effect, search for an exact existing target, perform it at most once, read it back, and add attachments or attendees only when the task-visible contract explicitly requires and identifies them. It must perform final file, message, and event checks before completion.
