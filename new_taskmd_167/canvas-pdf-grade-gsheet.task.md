I am a course administrator at Open University and I need to generate end-of-semester grade reports for all Spring 2014 courses. The Spring 2014 courses are identified by course codes containing "2014B".

There is a grading_policy.pdf file in the workspace that contains our official grading scale and special designation criteria. You must read this PDF first and apply its rules. The PDF specifies the following canonical thresholds that your output must follow exactly: letter grade A is awarded when the class average is 90.00 or above, B is 80.00 through 89.99, C is 70.00 through 79.99, D is 60.00 through 69.99, and F is below 60.00. The "With Distinction" designation is awarded when the class average is 85.00 or above. The "Academic Probation" flag is triggered when the class average falls below 50.00. All numeric values in the report (individual averages, summary averages) must be rounded to exactly 2 decimal places using standard mathematical rounding.

For each Spring 2014 course, look up the course in our learning management system and find the lead instructor. The lead instructor is the teacher enrolled in the course who comes first alphabetically by name. If a course has no teacher enrolled, use "N/A" for the instructor name and email, and skip sending an email for that course. Also find the number of students who have recorded scores (a non-null current_score in their enrollment grades) and calculate the class average score from those student scores.

Assign each course a letter grade based on its class average using the scale in the PDF. Also determine if the course qualifies for "With Distinction" honors or triggers an "Academic Probation" flag per the PDF criteria.

Create a file called semester_grade_report.xlsx in the workspace with three sheets.

The first sheet should be named "Course Grades" with the following columns: Course_Code, Course_Name, Lead_Instructor, Instructor_Email, Students_Scored, Class_Average (rounded to 2 decimal places), Letter_Grade, Distinction (write "Yes" or "No"), Probation (write "Yes" or "No"). Sort rows by Course_Code alphabetically. Course_Name should include the semester suffix in parentheses, e.g. "Biochemistry & Bioinformatics (Spring 2014)".

The second sheet should be named "Grade Distribution" with columns: Letter_Grade (list A, B, C, D, F in that order), Course_Count (how many courses received that letter grade), Courses (comma-separated course codes that received that grade, sorted alphabetically within each grade; leave empty string if no courses got that grade).

The third sheet should be named "Summary" with two columns (Metric, Value) containing the following rows: Total_Courses (the number of Spring 2014 courses), Avg_Class_Average (the average of all class averages rounded to 2 decimal places), Highest_Average_Course (the course code with the highest class average), Lowest_Average_Course (the course code with the lowest class average), Distinction_Count (how many courses earned the With Distinction designation), Probation_Count (how many courses are flagged for Academic Probation).

Also create a Google Sheet titled "Spring 2014 Grade Summary" with one sheet named "Grades" containing columns: Course_Code, Class_Average, Letter_Grade, Distinction, Probation for all six courses sorted by Course_Code.

Finally, send an email from registrar@openuniversity.ac.uk to each course's lead instructor (for courses that have a teacher) with the subject "End-of-Semester Grade Report: [Course Code]" where [Course Code] is replaced with the actual course code. The body should state the course name, the class average, the letter grade, and whether the course earned distinction or is on academic probation. Note: the registrar inbox and the cloud spreadsheet workspace may contain pre-existing unrelated messages and spreadsheets (such as IT maintenance notices, library notices, or historical registration reports). Do not forward, reference, or otherwise act on those unrelated items; only create the Spring 2014 grade report artifacts described above.

To solve this task efficiently, dynamically collect every Spring 2014 score page, compute one canonical six-row payload, and make the workbook, Google Sheet, and instructor emails consume that payload verbatim. Do not embed expected course averages or instructor identities in the prompt.

The main agent resolves the accessible task workspace, reads grading_policy.pdf, calls canvas_list_courses(include_ended=true), and freezes the literal IDs/codes/full names of exactly the 2014B courses. For each retained course, call canvas_get_course_grades(page=1, per_page=1, type=["StudentEnrollment"]) in a native-parallel discovery batch; take pagination.total_count and derive the exact page set at per_page=100.

1. Wave 1 — dispatch all concrete page and teacher work items in parallel, never more than 128:

   - For every literal {course_id,course_code,course_name,page}, one Explore agent calls canvas_get_course_grades(course_id=<ID>, page=<page>, per_page=100, type=["StudentEnrollment"]) exactly once and returns only {course_id,course_code,page,pagination_page,pagination_total_count,returned_rows,scored_count,score_sum,score_min,score_max}, using enrollment.grades.current_score and ignoring nulls. It must not return enrollment rows.
   - For every literal Spring course, one Explore agent first calls canvas_get_course_grades(course_id=<ID>, page=1, per_page=100, type=["TeacherEnrollment"]) and canvas_list_course_users(course_id=<ID>). It joins teacher enrollment user_ids to users, sorts matched teachers by complete name alphabetically, and returns {course_id,course_code,teachers:[{user_id,name,email}],lead_instructor,lead_email,teacher_count,unmatched_teacher_ids}. If no teacher exists, return N/A for name/email. Do not modify Canvas.

After Wave 1, the main agent verifies every discovered student page exactly once, summed returned_rows against total_count, and exactly one teacher result per course. Combine unrounded page sums, apply the policy and standard mathematical rounding once, and freeze one canonical six-row table:
{course_code,course_name,lead_instructor,instructor_email,students_scored,class_average_2dp,class_average_text_2dp,letter_grade,distinction,probation}.
Derive Grade Distribution and Summary from this table only. Assert every downstream class-average string is exactly class_average_text_2dp; no branch may re-query or recompute grades.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of semester_grade_report.xlsx in the dynamically resolved workspace. Create exactly Course Grades, Grade Distribution, and Summary with the exact headers/order and literal canonical values, then read all populated ranges back.
   2. One Coder is the sole owner of the Google spreadsheet titled Spring 2014 Grade Summary. Create it once, create/use only sheet Grades, write precisely the header and six canonical rows to A1:E7, and read A1:E7 back. Do not add diagnostic cells.

In the same response that launches Wave 2, the main agent searches Sent mail in native parallel for every canonical row whose instructor email is not N/A, matching complete From registrar@openuniversity.ac.uk, To, and subject End-of-Semester Grade Report: <Course Code>. Send only missing messages. Each body must copy the literal canonical course_name, class_average_text_2dp, letter_grade, distinction, and probation fields without recalculation. Read every retained email back.

After Wave 2, reconcile workbook, sheet and emails against the identical canonical rows and exact two-decimal strings. Do not dispatch a verification wave.
