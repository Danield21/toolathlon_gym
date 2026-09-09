I want to get organized for the Data-Driven Design course from Spring 2014 (its course code is DDD-2014B). Please look up this course and pull all of its assignments that have due dates.

Create an Excel file called "Assignment_Tracker.xlsx" in the workspace. The first sheet should be called "Assignments" with these columns: Assignment_Name, Due_Date (the date of the assignment's due_at field, formatted as YYYY-MM-DD), Points_Possible, and Assignment_Group (the name of the assignment group that the assignment belongs to). Sort the rows by Due_Date from earliest to latest.

Add a second sheet called "Summary" with two columns, Metric and Value. Include the following rows: Total_Assignments (the number of assignments with due dates), Avg_Points_Possible (the average points possible rounded to 1 decimal place), Earliest_Due_Date, and Latest_Due_Date (both formatted as YYYY-MM-DD).

Finally, create a calendar event for each assignment that has a due date. Title each event "Due: [Assignment Name]" where you use the actual assignment name. Set the event to start and end on the due date using the time from the due_at field. In the description, mention the points possible for that assignment.

To solve this task efficiently,

This is a small single-course task. The main agent completes it directly; do not dispatch any sub-agent. Canvas is read-only.

First resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. Call `canvas_list_courses(include_ended=true)` once and select the unique returned course whose code is exactly `DDD-2014B`; do not hardcode a course or assignment ID. If the code is absent or duplicated, stop.

For that returned course ID, call `canvas_list_assignments(course_id=<returned id>)` and `canvas_list_assignment_groups(course_id=<returned id>)` in native parallel. Use every distinct returned assignment with non-null `due_at`; if the list response is paginated, continue only until its returned pagination says complete. Join group IDs to returned group names, never to a guessed label. Sort by the UTC date component ascending, then assignment name. Compute literal `Total_Assignments`, one-decimal `Avg_Points_Possible`, `Earliest_Due_Date`, and `Latest_Due_Date`.

Freeze one literal canonical payload containing the sorted workbook rows, four Summary metrics, and one event payload per assignment. Each event title is `Due: <actual assignment name>`; start and end are the exact returned `due_at` instant; description is `Points possible: <literal points_possible>`. The workbook and calendar branch consume this same payload without recomputing dates or names.

Start the workbook operations and all exact-event searches in one native-parallel response. Create `<workspace_root>/Assignment_Tracker.xlsx` exactly once with only `Assignments` and `Summary`; use exact columns `Assignment_Name, Due_Date, Points_Possible, Assignment_Group`, UTC `YYYY-MM-DD` dates, literal numbers, and the four vertical Summary rows. Read both complete populated ranges back.

For each canonical event, search the relevant time/title before creation. Create only a missing exact event, at most once; never retry an ambiguous create. Read all retained event IDs back in one native-parallel batch. Accept only after the workbook and exactly one matching event per canonical assignment agree with the frozen payload, with no extra matching event.
