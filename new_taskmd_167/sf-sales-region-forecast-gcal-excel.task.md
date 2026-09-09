The VP of Sales wants to track how actual regional revenue compares to the annual targets set at the beginning of the year. There is a PDF file called Sales_Targets.pdf in the workspace that contains the 2026 annual revenue targets approved for each sales region. Please read that document first to get the target figures.

Then connect to the data warehouse and query the actual total revenue and order count grouped by customer region. These actuals must be based on ALL orders regardless of order status — include Delivered, Shipped, Processing, and Cancelled orders alike.

Create an Excel file called Regional_Forecast.xlsx in the workspace with two sheets.

The first sheet should be called "Region Performance" with five columns: Region (sorted alphabetically), Target (from the PDF), Actual (from the database), Variance (which is Actual minus Target rounded to 2 decimal places), and Variance_Pct (which is Variance divided by Target multiplied by 100, rounded to 1 decimal place).

The second sheet should be called "Summary" with two columns: Metric and Value. Include rows for Total_Target (sum of all regional targets), Total_Actual (sum of all actual revenues), Total_Variance (Total_Actual minus Total_Target), Met_Target_Count (number of regions where Actual is greater than or equal to Target), and Missed_Target_Count (number of regions where Actual is below Target).

After creating the Excel file, schedule a Google Calendar event called "Quarterly Forecast Review" on March 31, 2026 from 14:00 to 15:30. Schedule it as a single event (do not set a time zone or recurrence rule — the calendar stores times in UTC, so pass the start and end times without a time zone parameter).

Finally, send an email to vp_sales@company.com with a subject referencing the regional forecast report and a body that summarizes which regions are meeting their targets, the total variance figure, and confirms that a quarterly review meeting has been scheduled.

To solve this task efficiently, complete it directly in the main agent; do not dispatch sub-agents. Read `Sales_Targets.pdf` and run one all-status, set-based regional warehouse aggregate as native-parallel calls. Preserve exact region labels, reconcile the two small result sets, and compute literal Target, Actual, Variance, and Variance_Pct values plus the five Summary metrics. Create `Regional_Forecast.xlsx` with two bulk sheet writes and one read-back. Independently search/create/read back the single non-recurring `Quarterly Forecast Review` event at the literal offset-free times 2026-03-31 14:00–15:30, omitting timezone and recurrence fields. Then send exactly one email to `vp_sales@company.com` that names the regions meeting target, the verified total variance, and the scheduled meeting.
