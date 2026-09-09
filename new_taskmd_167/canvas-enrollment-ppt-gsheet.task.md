Hi there, I am an academic administrator and I need to put together an enrollment overview for all courses in our learning management system. This will be shared with the provost next week so it needs to be comprehensive.

Please pull data for every course available in the system. For each course, I need the course name, course code, total number of enrollments, and a breakdown of enrollments by type (how many are students, how many are teachers, how many are teaching assistants). Also break down by enrollment status showing how many are active and how many are completed.

Create an Excel file called Enrollment_Overview.xlsx in the workspace with two sheets.

The first sheet should be called "Enrollment Details" and have columns Course_Name, Course_Code, Total_Enrollments, Students, Teachers, TAs, Active, and Completed. Sort the rows alphabetically by Course_Name.

The second sheet should be called "Summary" and have a two-column layout with Metric in column A and Value in column B. Include the following rows: Total_Courses (the number of courses), Total_Enrollments (sum across all courses), Total_Students, Total_Teachers, Total_TAs, Avg_Enrollment_Per_Course (average total enrollments per course rounded to 2 decimal places), Largest_Course (name of the course with the most enrollments), Largest_Course_Enrollments (its enrollment count), Smallest_Course (name of the course with the fewest enrollments), and Smallest_Course_Enrollments (its enrollment count).

Also create a PowerPoint presentation called Enrollment_Overview.pptx in the workspace. The first slide should have the title "Course Enrollment Overview" and a subtitle "Academic Dashboard Report - <today's date in YYYY-MM-DD format>". The second slide should be titled "Enrollment Summary" and show the total courses, total enrollments, total students, average enrollment per course, and which course is largest and smallest. The third slide should be titled "Top 5 Courses by Enrollment" and list the five courses with the highest enrollment counts, showing course name, code, and enrollment number. The fourth slide should be titled "Enrollment Distribution" and show the breakdown by enrollment type (students, teachers, TAs) and by status (active, completed) with percentages rounded to 1 decimal place.

Additionally, create a shared spreadsheet titled "Enrollment Dashboard" with a sheet called "Course Data" containing the same columns as the "Enrollment Details" sheet described above.

To solve this task efficiently, first isolate each course's six compact counts, then fan out one canonical table to three independent artifact writers. Canvas remains read-only.

The main agent calls canvas_list_courses(include_ended=true), validates unique IDs/codes/names, and freezes the exact returned course inventory. It then dispatches one explicitly instantiated Explore prompt per literal course; every prompt must contain that course's exact ID, code, and full name rather than an ordinal, range, or grouped label.

1. Wave 1 — dispatch exactly 22 explore sub-agents, one for each literal course in the validated 22-course inventory, all in parallel:

   Each Explore agent owns one literal course and issues these six independent canvas_get_course_grades calls natively in parallel, all with page=1 and per_page=1: unfiltered; type=["StudentEnrollment"]; type=["TeacherEnrollment"]; type=["TaEnrollment"]; enrollment_state=["active"]; enrollment_state=["completed"]. Return only {course_id, course_code, course_name, total_enrollments, students, teachers, tas, active, completed, pagination_controls}, taking every count from the matching response's pagination.total_count. Do not fetch later pages, return enrollment rows, or modify Canvas.

After Wave 1, the main agent validates exactly one result per course and freezes a canonical table sorted by course_name with the eight required Enrollment Details columns. Compute Summary values, type/status percentages, and deterministic top-five/largest/smallest selections once; use course_code ascending as the tie-break. Freeze the runtime current date required by the presentation. No writer may recompute these facts.

2. Wave 2 — dispatch exactly 3 coder sub-agents in parallel:

   1. One Coder is the sole writer of Enrollment_Overview.xlsx in the dynamically resolved task workspace. Create exactly Enrollment Details and Summary with the required headers, rows, literal numeric values and rounding, then read both populated ranges back.
   2. One Coder is the sole writer of Enrollment_Overview.pptx in that workspace. Create exactly the four requested slides from the canonical summary/top-five/distribution payload and runtime date, then read back slide titles and all displayed values.
   3. One Coder is the sole owner of the shared spreadsheet titled Enrollment Dashboard. Create or reuse one exact-title spreadsheet and write the canonical table only to sheet Course Data in the exact bounded range A1:H<row count>; do not write diagnostic cells. Read that range back.

After Wave 2, the main agent verifies all three artifacts against the canonical table, including total-enrollment arithmetic, exact type/state counts, percentages, top-five order, extrema and current-date subtitle. Do not dispatch a verification wave.
