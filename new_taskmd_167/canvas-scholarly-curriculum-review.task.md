You are an accreditation coordinator preparing a comprehensive curriculum review for an upcoming university accreditation visit. Your job is to audit the institution's courses against accreditation requirements, find supporting scholarly literature, and document everything in a structured compliance matrix and an accreditation report.

Start by fetching the accreditation requirements from the accreditation portal at http://localhost:30238/api/accreditation_requirements.json. This document specifies minimum thresholds for course counts, faculty ratios, assessment types, and scholarly integration that the institution must meet.

Next, connect to the learning management system and retrieve a list of all courses offered by the institution. For each course, gather the course name, the number of enrolled students, the number of assignments, the number of quizzes the course has (use 0 if it has none), and the number of discussion topics. Also determine how many faculty members (teachers and teaching assistants combined) are associated with each course. Use the course's enrollment count (total students) directly for the number of enrolled students, and derive the faculty count from each course's enrollment records, where faculty are enrollments of type TeacherEnrollment or TaEnrollment.

Using this data, evaluate each course against the accreditation requirements. A course is considered compliant if the student to faculty ratio does not exceed the maximum specified in the requirements and the course has at least one form of assessment (quiz or assignment). Note which courses are compliant and which are not.

Then search the scholarly literature for research supporting the pedagogical approaches used in the curriculum. Search for papers on the following topics: active learning in higher education, assessment best practices in education, online learning effectiveness, curriculum design frameworks, and student engagement strategies. For each topic, find at least one relevant paper and note its title, authors, and relevance to the curriculum.

Create an Excel file called Curriculum_Review.xlsx in your workspace with three sheets. The first sheet should be called "Course Compliance" with columns for Course_Name, Assignment_Count, Quiz_Count, Discussion_Count, Student_Count, Faculty_Count, Student_Faculty_Ratio, and Compliant (Yes or No based on whether the course meets the accreditation criteria). Include all courses.

The second sheet should be called "Literature Support" with columns for Search_Topic, Paper_Title, Authors, and Relevance_Note. Include at least five rows covering the five pedagogical topics listed above.

The third sheet should be called "Summary" with columns for Metric and Value. Include rows for total courses, compliant courses, non-compliant courses, compliance rate (as a percentage), total faculty across all courses, average student to faculty ratio, and the number of supporting papers found.

Finally, create a page in the team's knowledge base (Notion) titled "Accreditation Review Report" (the page title must contain this exact phrase) that summarizes the curriculum audit findings. The page should mention the overall compliance rate, highlight any courses that fail to meet requirements, reference the scholarly literature that supports the institution's pedagogical approach, and note any recommendations for improvement before the accreditation visit.

When you have completed all tasks, call claim_done.

To solve this task efficiently, group the four small Canvas reads by concrete course and keep the five independent literature searches parallel. Do not launch separate agents for each endpoint.

Before Wave 1, the main agent resolves the task workspace and natively in parallel fetches http://localhost:30238/api/accreditation_requirements.json, reads accreditation_timeline.md and self_study_template.json, and calls canvas_list_courses(include_ended=true). Validate the requirement schema and unique course IDs/codes/names; retain each course's reported total_students. Freeze the exact five task-stated literature topics. Instantiate one course prompt per literal returned course and one literature prompt per literal topic.

1. Wave 1 — dispatch exactly 27 explore sub-agents in parallel: 22, one per literal course in the validated inventory, plus five, one per named literature topic:

   - Each course Explore prompt contains one exact course_id, course_code and full course_name. In one native-parallel response call canvas_list_assignments(course_id=<ID>, include_submissions=false), canvas_list_quizzes(course_id=<ID>), canvas_list_discussion_topics(course_id=<ID>), and canvas_get_course_grades(course_id=<ID>, page=1, per_page=1, type=["TeacherEnrollment","TaEnrollment"]). Return only {course_id,course_code,course_name,assignment_ids,assignment_count,quiz_ids,quiz_count,discussion_ids,discussion_count,faculty_count:pagination.total_count,source_counts,duplicate_ids}. Use zero for empty lists, do not fetch faculty rows beyond page 1, and do not modify Canvas.
   - Each literature Explore prompt contains exactly one of: active learning in higher education; assessment best practices in education; online learning effectiveness; curriculum design frameworks; student engagement strategies. Search Google Scholar with that literal query, retain at least one directly relevant result, and return {search_topic,papers:[{title,authors,year,venue,url,relevance_note}],returned_count}. It must remain read-only.

After Wave 1, the main agent requires one result per literal course and topic, joins course results to total_students, and applies only the fetched accreditation thresholds. Compute Faculty_Count from the filtered pagination total, Student_Faculty_Ratio from unrounded values, and Compliant=Yes only when the ratio is within the maximum and assignment_count>0 or quiz_count>0. Freeze one canonical 22-course compliance table, one five-topic literature table, Summary metrics, the complete non-compliance reasons, and a canonical report paragraph.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of Curriculum_Review.xlsx in the dynamically resolved workspace. Create exactly Course Compliance, Literature Support, and Summary with the task's exact headers, complete course/topic coverage and literal values; read every populated range back.
   2. One Coder solely owns the Notion page Accreditation Review Report. Search/retrieve exact-title candidates under the task-visible Accreditation Documents parent, stop on duplicate parent/report pages, and create the report page exactly once only if absent. Append the canonical paragraph exactly once only if its complete text is absent. The paragraph must state the literal compliance rate, every non-compliant full course name and reason, at least one exact paper from each named topic, and source-supported recommendations/timeline facts. Retrieve the page and all child blocks after writing; do not modify unrelated content.

After Wave 2, the main agent reconciles all 22 course rows, five literature topics, workbook summary and Notion paragraph against the same canonical payload, then calls claim_done. Do not dispatch a verification wave.
