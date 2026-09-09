I need to put together an enrollment overview for all our Spring 2014 courses. Please look through the course catalog and find every course that has "Spring 2014" in its name (equivalently, course codes ending in "2014B"). For each of those courses, count how many students are enrolled (only counting StudentEnrollment type) and how many teachers.

Put everything into an Excel file called Spring2014_Enrollment.xlsx. The first sheet should be called "Enrollment" with columns Course_Name, Course_Code, Student_Count, Teacher_Count, and Student_Teacher_Ratio (students divided by teachers, rounded to 1 decimal, use 0 if there are no teachers). Sort this sheet by Student_Count from lowest to highest.

Add a second sheet called "Summary" with two columns Metric and Value. Include these rows: Total_Courses (how many Spring 2014 courses there are), Total_Students (sum of all students across courses), Avg_Students_Per_Course (rounded to the nearest whole number), Smallest_Course (the name of the course with the fewest students), and Largest_Course (the name of the course with the most students).

Also create a calendar event titled "Spring 2014 Enrollment Review" scheduled for March 14, 2026 from 2:00 PM to 3:00 PM UTC.

Finally, send an email to registrar@university.edu with the subject "Spring 2014 Enrollment Summary" that includes the total number of courses, total students, and which course has the highest and lowest enrollment.

To solve this task efficiently, the main agent must complete this six-course workflow directly; do not dispatch sub-agents. Resolve the accessible task workspace dynamically and place Spring2014_Enrollment.xlsx there.

Call canvas_list_courses(include_ended=true), retain every and only course whose full name contains Spring 2014 or whose code ends in 2014B, and freeze its literal ID/code/name. For each retained course, issue StudentEnrollment and TeacherEnrollment canvas_get_course_grades calls together in one native-parallel response with page=1 and per_page=1. Use pagination.total_count for the two counts; do not fetch enrollment rows. Reject identity mismatches or duplicate/missing Spring courses.

Compute a single canonical table with {course_name, course_code, student_count, teacher_count, student_teacher_ratio}, sorted by student_count ascending and course_code ascending for ties. Compute all Summary metrics once from that table, including deterministic smallest/largest course names, and freeze the exact email facts.

The main agent is the sole writer of Spring2014_Enrollment.xlsx. Create exactly Enrollment and Summary with the task's exact headers, literal values and rounding, then read both populated ranges back from the resolved path. Once the canonical payload is frozen, run the workbook branch, the exact calendar lookup for Spring 2014 Enrollment Review at 2026-03-14T14:00:00Z--15:00:00Z, and the exact Sent-mail lookup for registrar@university.edu / Spring 2014 Enrollment Summary natively in parallel where possible. Create only a missing exact event, leaving attendees empty and using the task's UTC times. Send only a missing exact-recipient/subject message and include the canonical course count, total students, largest course and smallest course. Read the event and message back; do not retry ambiguous creates.
