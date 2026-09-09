I need help with a email scheduler analysis. There is external benchmark data available that I need you to fetch the data from http://localhost:30331/api/data.json and extract the relevant metrics. The internal Email Scheduling Efficiency Index values per department have been sent to your inbox by the Operations team (see the email with subject "Q1 Scheduler Internal Metrics"); fetch those from your email.



Use the terminal to create and run a Python script called email_scheduler_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs email_scheduler_results.json.

Create an Excel file called Scheduler_Report.xlsx with three sheets. The first sheet Data_Analysis should contain the main comparison data with relevant columns. The second sheet Metrics should summarize key metrics. The third sheet Recommendations should list actionable items.

The Data_Analysis sheet should include columns for the primary dimension (such as department, product, region, or topic), our internal metric values, the external benchmark values, and the gap or difference between them. The columns must be named exactly: `Item`, `Internal_Value`, `External_Benchmark`, `Gap`. Sort the data alphabetically by the primary dimension. The Metrics sheet should have two columns Metric and Value summarizing total counts, averages, and key statistics. The Recommendations sheet should list priority actions based on the gap analysis, with columns named exactly: `Priority`, `Action`. Also create a Word document called Scheduler_Analysis.docx with an executive summary, key findings, and recommendations sections. Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings. Schedule a review meeting titled "Analysis Review" on March 14, 2026 from 2:00 PM to 3:00 PM UTC.

To solve this task efficiently, preserve the real source-to-processor-to-report dependency and parallelize only after the processor result is frozen. Resolve the accessible task workspace dynamically; never hard-code another run's workspace path.

Before Wave 1, the main agent natively in parallel fetches exactly http://localhost:30331/api/data.json and searches the Inbox for subject Q1 Scheduler Internal Metrics. Paginate the exact-subject search, read every candidate, require one unambiguous Operations message, and extract every parseable item/value record. Validate source identities, schemas and counts, normalize only key spelling, preserve conflicts/gaps explicitly, and freeze the two literal input payloads. Missing or ambiguous data is a blocker.

1. Wave 1 — dispatch exactly 1 coder sub-agent:

   1. The Coder solely owns the processor branch in the resolved workspace. It receives both literal frozen payloads, writes runtime-unique external and internal JSON inputs, creates email_scheduler_processor.py, and runs it to produce email_scheduler_results.json. The script must align items explicitly, define Gap=Internal_Value-External_Benchmark, sort rows alphabetically, and emit literal summary metrics and evidence-bounded recommendations. Prefer one compact Python execution to write/run the processor, then read every file back. Return only exact paths, SHA-256 values, execution status, source/item totals, unmatched items, gap/sort checks and result schema controls; do not create reports or side effects.

After Wave 1, the main agent verifies the two input hashes, complete item accounting, every numeric gap and aggregate, then freezes the processor result path/hash as the sole downstream source.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of Scheduler_Report.xlsx. Verify the result hash, create exactly Data_Analysis(Item,Internal_Value,External_Benchmark,Gap), Metrics(Metric,Value), and Recommendations(Priority,Action), and read all exact populated ranges back.
   2. One Coder is the sole writer of Scheduler_Analysis.docx. Verify the same result hash, create substantive executive summary, key findings and recommendations sections, and read the document text/headings back.

In the same response that launches Wave 2, the main agent natively in parallel searches Sent mail for complete To team-lead@company.com / subject Analysis Report Complete and the shared calendar for Analysis Review at 2026-03-14T14:00:00Z--15:00:00Z. Create only missing exact side effects from the frozen result, leave attendees empty, and read retained IDs/content back. After Wave 2, reconcile both reports, message and UTC event against the same result hash. Do not dispatch a verification wave.
