I need a summary of overall student grades across all courses. For each course, query the learning management system for that course's enrollments with grades, and use each enrolled student's **current overall grade** — the `current_score` field inside each enrollment's grade record. Ignore enrollments that have no grade or a null `current_score`.

Create an Excel file called Canvas_Grade_Summary.xlsx with two sheets.

The first sheet, named "Grade Summary", should have exactly one row per course, with these columns in this order: **Course** (the course name), **Students_Graded** (the number of students in the course whose `current_score` is present/non-null), **Avg_Score**, **Max_Score**, and **Min_Score**. Avg_Score, Max_Score, and Min_Score are the average, maximum, and minimum of those students' `current_score` values, each rounded to 2 decimals. Sort the rows by Avg_Score descending.

The second sheet, named "Summary", should contain a header row with columns `Metric` and `Value`, followed by three rows: **Total_Courses** (the total number of courses), **Highest_Avg_Course** (the name of the course with the highest Avg_Score), and **Overall_Avg_Score** (the unweighted average of the per-course Avg_Score values across all courses, rounded to 2 decimals).

Also record these three summary metrics (Total_Courses, Highest_Avg_Course, Overall_Avg_Score) in a Google Sheet titled "Grade Summary Report" for the academic team to review.

To solve this task efficiently, read every required score row but cap each sub-agent's raw Canvas input at four pages. Return only partial aggregates and never hard-code runtime course or page counts.

The main agent calls `canvas_list_courses(include_ended=true)`, validates unique course IDs/codes/full names, and issues one native-parallel discovery batch with `canvas_get_course_grades(course_id=<literal ID>, page=1, per_page=1, type=["StudentEnrollment"])` for every returned course. Take `total_count` from pagination, compute `pages_at_100=ceil(total_count/100)`, and freeze one literal `{course_id,course_code,course_name,total_count,pages_at_100}` row per course. For each course, divide its consecutive page interval `1..pages_at_100` into non-overlapping consecutive work items of at most four pages; freeze every work item as `{course_id,course_code,course_name,total_count,page_start,page_end}` and verify their union is the complete page manifest. The current runtime manifest must fit one Wave of at most 128 items; if it does not, use the minimum number of additional consecutive waves required by the 128-agent cap without changing any item.

1. Wave 1 — dispatch exactly one explore sub-agent for every frozen page work item, all in parallel and never more than 128:

   Each Explore agent receives one complete literal work item. In its first tool-call response it issues the one-to-four independent `canvas_get_course_grades(course_id=<literal ID>, page=<each integer page_start..page_end>, per_page=100, type=["StudentEnrollment"])` calls natively in parallel. It verifies the returned pages equal that exact interval and every page reports the frozen `pagination.total_count`. From `enrollment.grades.current_score`, ignore nulls and return only `{course_id,course_code,course_name,total_count,page_start,page_end,pages_observed,returned_rows,graded_count,score_sum,score_min,score_max,duplicate_pages,missing_pages}`. Do not return enrollment IDs, score lists, or raw rows, and do not modify Canvas.

After the page-work Wave(s), the main agent requires exactly one compact record per frozen work item, rejects missing/duplicate pages, verifies every course's summed `returned_rows` equals its discovery `total_count`, and combines unrounded partials into a canonical course table `{course_name,students_graded,avg_score,max_score,min_score}`. Round course values only after combining unrounded sums, sort by Avg_Score descending with course name ascending for ties, and compute Total_Courses, Highest_Avg_Course, and the unweighted average of per-course rounded Avg_Score once.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of Canvas_Grade_Summary.xlsx in the dynamically resolved task workspace. Create exactly Grade Summary and Summary with the exact headers, canonical rows and literal values; read both populated ranges back.
   2. One Coder is the sole owner of the Google spreadsheet titled Grade Summary Report. Create the spreadsheet once, write only the header plus the three canonical summary rows to the exact range A1:B4, and read A1:B4 back. Do not add diagnostic cells or any extra data range.

The main agent finally verifies both outputs against the same canonical metrics. No raw grade page, extra verification agent, or hidden expected value may be introduced.
