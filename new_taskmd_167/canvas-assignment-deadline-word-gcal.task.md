I need to help students in the Foundations of Finance course from Fall 2013 stay on top of their assignment deadlines by creating a comprehensive deadline management package.

Please look up all assignments for that course and create an Excel file called Assignment_Deadlines_FFF2013J.xlsx with two sheets. The first sheet called All Assignments should have columns Assignment_Name, Points_Possible, and Due_Date formatted as YYYY-MM-DD, sorted by Due_Date ascending with assignments that have no due date placed at the end. The second sheet called Summary should have two columns Metric and Value with rows for Total_Assignments showing the total count of all assignments, Total_Points_Possible showing the sum of all assignment points, and Avg_Points_Per_Assignment showing the average points rounded to 2 decimal places.

Also create a Word document called Assignment_Schedule_FFF2013J.docx with a heading Assignment Schedule for Foundations of Finance Fall 2013, followed by a table listing each assignment with its name, due date formatted as YYYY-MM-DD, and points possible. If an assignment has no due date write TBD in the due date column.

For each assignment that has a due date, schedule a calendar reminder event that starts 7 days before the due date at 8:00 AM Eastern Time (ET) and lasts 30 minutes. The event title should be Assignment Due: followed by the assignment name. Use the America/New_York timezone.

Send an email to fff2013j.students@university.edu with the subject Foundations of Finance Assignment Deadline Reminder. The email body should list all assignments with their due dates and points, and remind students to check their calendars for the scheduled reminders.

To solve this task efficiently,

This compact single-course task uses the main agent for source collection and side effects, with one parallel writer wave for the two independent files. Canvas is read-only.

First resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. Call `canvas_get_course(course_id=16)` and `canvas_list_assignments(course_id=16)` in native parallel; verify course code `FFF-2013J`, deduplicate stable assignment IDs, and use every returned assignment. No per-assignment sub-agent is needed because the list call already returns the required name, points, and due date.

Sort by non-null `due_at` ascending, then name, with null dates last. Compute literal `Total_Assignments`, `Total_Points_Possible`, and two-decimal `Avg_Points_Per_Assignment`. For each non-null due date, convert the instant to America/New_York, subtract exactly seven calendar days, and set a 30-minute reminder beginning at 08:00 local time. Freeze one literal canonical payload containing the ordered rows, three summary metrics, exact event title/start/end/timezone payloads, and exact email body lines. Both file writers and every side effect consume this packet without recomputation.

1. Wave 1 — dispatch exactly 2 coder sub-agents in parallel as individual Agent calls:

   1. Excel writer: create `<workspace_root>/Assignment_Deadlines_FFF2013J.xlsx` exactly once with only `All Assignments` and `Summary`; use exact headers, all canonical rows, `YYYY-MM-DD` or blank due dates, and three vertical metrics with literal numbers. Read both populated ranges back and return path, sheet order, row counts, and equality controls.
   2. Word writer: create `<workspace_root>/Assignment_Schedule_FFF2013J.docx` exactly once, preferably atomically with `python-docx`, with heading `Assignment Schedule for Foundations of Finance Fall 2013` and the complete canonical name/date/points table, writing `TBD` for null due dates. Read it back and return path, row count, headings, and equality controls.

In the same response that launches these two individual Agent calls, the main agent natively parallelizes exact searches for every canonical calendar event and the exact Sent-email recipient/subject. Create each missing event at most once with title `Assignment Due: <name>`, the canonical America/New_York start/end, and no invented attendees; retain and read back IDs. Paginate Sent search with page size 50, read every candidate, match complete To and Subject, and send at most one email to `fff2013j.students@university.edu` with subject `Foundations of Finance Assignment Deadline Reminder`; its body lists every canonical assignment/date/points line and reminds students to check the calendar. Never retry an ambiguous create/send.

Accept only after both files, every expected reminder event, and exactly one exact email read back against the same canonical payload.
