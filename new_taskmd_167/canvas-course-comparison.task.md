I need a side-by-side comparison of all our Canvas courses. Pull data for each course including enrollment numbers, assignment counts, quiz counts, and discussion topic counts. Count every quiz that the learning management system returns for the course, including quizzes that are not yet published, so that the quiz count matches exactly what the system reports.

Create an Excel file called Canvas_Course_Comparison.xlsx with two sheets. The "Course Comparison" sheet should have Course_Code, Course_Name, Students, Assignments, Quizzes, and Discussions. Sort by Course_Code.

The "Summary" should have Total_Courses, Avg_Students_Per_Course rounded to nearest integer, and Most_Popular_Course code.

Lay out the "Summary" sheet vertically (two columns): put the metric name in column A and its value in column B, one metric per row. For example, row 1 is the header (Metric, Value); row 2 is Total_Courses with its value in column B; row 3 is Avg_Students_Per_Course with its value in column B; row 4 is Most_Popular_Course with its value in column B. Do not lay the summary out horizontally (metrics across columns in a single row) — use the vertical Metric/Value layout.

Also create a Google Sheet titled "Course Comparison Dashboard" with a sheet "Data" containing the same comparison data.

To solve this task efficiently,

Canvas is read-only. Use one dynamic course manifest and one parallel count wave; do not hardcode course or object IDs.

Before delegation, resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. The main agent calls `canvas_list_courses(include_ended=true)` once, deduplicates stable IDs, requires the 22 unique task-visible courses, and freezes the exact returned `{course_id,course_code,course_name,total_students}` manifest sorted by course code. If count or identity is missing/duplicated, stop rather than guessing.

1. Wave 1 — dispatch exactly 88 Explore sub-agents concurrently as individual Agent calls: four per each of the 22 manifest courses. For each course, render these four complete objects with its literal returned ID/code:

   1. Course verifier: call `canvas_get_course(course_id=<id>)` once; return `{course_id,course_code,course_name,total_students}`.
   2. Assignment counter: call `canvas_list_assignments(course_id=<id>)` once; if pagination is exposed follow it to completion, deduplicate stable assignment IDs, and return only `{course_id,assignment_count,duplicate_ids,page_controls}`.
   3. Quiz counter: call `canvas_list_quizzes(course_id=<id>)` once; include unpublished records, follow only returned pagination, deduplicate stable quiz IDs, and return only `{course_id,quiz_count,duplicate_ids,page_controls}`.
   4. Discussion counter: call `canvas_list_discussion_topics(course_id=<id>)` once; follow only returned pagination, deduplicate stable topic IDs, and return only `{course_id,discussion_count,duplicate_ids,page_controls}`.

Each Explore agent handles exactly one returned course/object type, makes no writes, and does not return full item arrays. After Wave 1, require four valid handoffs per manifest course, join only by course ID, and freeze one literal canonical comparison payload with exact columns `Course_Code, Course_Name, Students, Assignments, Quizzes, Discussions`, code-sorted rows, and vertical Summary metrics `Total_Courses`, nearest-integer `Avg_Students_Per_Course`, and `Most_Popular_Course` (students descending, code ascending tie-break). Both writers consume this exact payload without recomputing counts.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. Excel owner: create `<workspace_root>/Canvas_Course_Comparison.xlsx` exactly once with only `Course Comparison` and `Summary`; write the exact six headers, every canonical row, and the two-column `Metric, Value` Summary. Use literal numerics, read both complete populated ranges back, and return path, sheet order, row counts, and equality controls.
   2. Google Sheets owner: search exact spreadsheet title `Course Comparison Dashboard`; reuse one exact match, create once on zero matches, and block on duplicates. Ensure exactly one `Data` tab, write only the literal 23-by-6 matrix (header plus 22 canonical course rows) to `A1:F23`, read `A1:F23` back, and return spreadsheet ID, tab identity, row count, and equality controls. Never create or rewrite after an ambiguous failure.

Accept only after both outputs read back cell-for-cell against the same canonical payload.
