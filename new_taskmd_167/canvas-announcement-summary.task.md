I need a summary of course announcements from our learning management system. Pull all announcements and aggregate by course, showing count and date range.

Create an Excel file called Course_Announcements.xlsx with two sheets. The first sheet, named "Announcement Stats", should have columns Course_Code, Announcements, Earliest_Date, and Latest_Date with one row per course, sorted by Announcements descending. Use UTC dates (YYYY-MM-DD) for Earliest_Date and Latest_Date — do not apply any timezone conversion to the timestamps returned by the announcements API.

The second sheet, named "Summary", should be a two-column table with the header "Metric, Value" and exactly three rows: Total_Announcements (total announcements across all courses), Courses_With_Announcements (number of distinct courses that have at least one announcement), and Most_Active_Course (the course code with the highest announcement count).

Send an email to academic-affairs@openuniversity.ac.uk with subject "Course Announcement Activity Report" briefly summarizing the activity levels.

To solve this task efficiently,

Use the main agent for the compact read-only aggregation and one coder only for the final workbook. Do not create one reader agent per course.

Resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. Call `canvas_list_courses(include_ended=true)` once, deduplicate stable course IDs, then issue one independent complete announcement-list call per returned course in one native-parallel response. If the handler exposes pagination, continue only the courses whose response says another page exists; otherwise do not invent cursors. Canvas is read-only.

For every course compute announcement count and the earliest/latest UTC calendar dates directly from returned timestamps without timezone conversion. Include one row per course, leaving dates empty when count is zero. Sort by count descending then course code ascending. Compute exactly `Total_Announcements`, `Courses_With_Announcements`, and `Most_Active_Course`, breaking a count tie by course code ascending. Freeze one literal canonical payload containing all course rows, the three metrics, and the email summary; no output branch may recompute it.

1. Wave 1 — dispatch exactly 1 coder sub-agent:

   1. Create `<workspace_root>/Course_Announcements.xlsx` exactly once with only `Announcement Stats` and `Summary`. Write exact headers `Course_Code, Announcements, Earliest_Date, Latest_Date`, every canonical course row, UTC `YYYY-MM-DD` dates, and exactly the three vertical Summary metrics. Use literal numeric counts. Read back both complete populated ranges and return path, sheet order, row counts, and cell-equality controls. Do not send email.

In the same response that launches this individual Agent call, the main agent paginates Sent search for exact To `academic-affairs@openuniversity.ac.uk` and Subject `Course Announcement Activity Report`, reading every candidate. If no exact match exists, send exactly one concise message containing the three canonical metrics and referencing `Course_Announcements.xlsx`; never retry an ambiguous send. Read the retained message and workbook back, and accept only when both match the same canonical payload.
