Hi there, the Fall 2014 semester just wrapped up and I need to prepare a summary for the department. Could you pull together an overview of all the courses from that term?

Please look up all courses that ran in Fall 2014 and gather the following for each: the course name, course code, total number of students, enrollment count, number of assignments, and the average points possible across all assignments (rounded to 1 decimal). If a course has no assignments, record Avg Points Possible as 0.0. If an individual assignment has a null or zero points_possible, include it in the count but treat its points as 0 when computing the average.

Create an Excel file called Semester_Summary.xlsx in the workspace with two sheets.

The first sheet should be called "Course Overview" with columns: Course Name, Course Code, Total Students, Enrollments, Assignments, Avg Points Possible. One row per course, sorted alphabetically by course name.

The second sheet should be called "Summary" with columns Metric and Value, containing these rows: Total Courses, Total Students (sum of all course total_students), Total Enrollments (sum of all enrollment counts), Total Assignments (sum of all assignment counts), Average Assignments per Course (rounded to 1 decimal).

Also create a PowerPoint presentation called Semester_Summary.pptx in the workspace. It should have:

A title slide with "Fall 2014 Semester Summary" as the title.

An overview slide showing the summary statistics (total courses, total students, total enrollments, total assignments, average assignments per course).

One slide for each course showing the course name as title and its details (course code, total students, enrollments, assignments, average points possible).

A final "Key Takeaways" slide identifying the largest course by student count, the smallest course by student count, and the course with the most assignments. For each, include both the course name and the specific numeric value (e.g. "Largest: CCC-2014J with 2498 students").

To solve this task efficiently, combine the three small reads for each concrete Fall 2014 course and freeze deterministic extrema before either writer starts.

The main agent calls canvas_list_courses(include_ended=true), retains exactly the returned courses whose full name contains Fall 2014 or whose code ends in 2014J, and freezes every literal ID/code/name and reported total_students. It then instantiates one Explore prompt per retained course.

1. Wave 1 — dispatch exactly 7 explore sub-agents, one for each literal Fall 2014 course, all in parallel:

   Each Explore agent receives one literal course ID/code/name and issues canvas_list_assignments(course_id=<ID>, include_submissions=false) and canvas_get_course_grades(course_id=<ID>, page=1, per_page=1) natively in parallel. Return {course_id,course_code,course_name,total_students,assignment_ids,assignment_count,points_possible_values,total_enrollments:pagination.total_count,pagination_controls}. Include null/zero assignment points in the returned values, do not fetch enrollment rows, and do not modify Canvas.

After Wave 1, the main agent validates one result per frozen course and unique assignment IDs. Compute Avg Points Possible with the task's null/zero rule and requested rounding. Freeze one canonical seven-row Course Overview table and five Summary rows. Compute largest-by-students, smallest-by-students, and most-assignments by running deterministic argmax/argmin over the complete canonical table, using course_code ascending only as a tie-break. Assert each selected row belongs to that table and its metric equals the recomputed extremum. Freeze the three literal takeaway strings; writers may not select or recompute extrema.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of Semester_Summary.xlsx in the dynamically resolved task workspace. Create exactly Course Overview and Summary with the exact headers, canonical rows, literal values, sorting and rounding; read both populated ranges back.
   2. One Coder is the sole writer of Semester_Summary.pptx in that workspace. Consume only the frozen canonical course rows, summary rows and three literal takeaway strings. Create exactly ten slides: title, overview, seven course slides in course-name order, and Key Takeaways. Read back all slide titles/text and require every course and each frozen takeaway to occur once.

After Wave 2, the main agent reconciles both artifacts against the same payload, including the deterministic extrema assertions. Do not hard-code expected course names/counts, write to /workspace directly, or dispatch a verification wave.
