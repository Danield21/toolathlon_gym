You are an academic advisor responsible for identifying students at risk of failing and coordinating early intervention efforts. Your goal is to perform a retention risk analysis across two specific courses, generate reports, build a tracking system, and notify the advising team.

Begin by reading the Retention_Policy.pdf in your workspace, which describes the intervention thresholds and escalation procedures your institution uses. Also review the scoring_model.json file, which contains the risk classification cutoffs and weighting parameters for the analysis.

Retrieve enrollment and grade data from the learning management system for two courses: course 16 (Foundations of Finance Fall 2013) and course 17 (Foundations of Finance Fall 2014). For each course, gather the course name, the total enrollment count, and every individual student submission score. You will compute course-level statistics from this data later (the exact definitions are given in the Course_Overview sheet instructions below).

Next, write a Python script called risk_scorer.py in your workspace and execute it using command-line tools. The script should classify each student based on their average score: students with an average below 60 are classified as High risk, students with an average between 60 and 74.99 are Medium risk, and students with an average of 75 or above are Low risk. The script should output the counts and percentages for each risk level per course and overall.

Create an Excel workbook called Student_Risk_Analysis.xlsx in your workspace with four sheets. Write all numeric values as literal numbers (for example, 77.14), not as Excel formulas.

The first sheet should be named Course_Overview and contain columns for course_id, course_name, enrollment_count, avg_score, and pass_rate. Include one row for each of the two courses.

Use these exact definitions:
- enrollment_count: the total number of enrollment records for the course, i.e. count every row in the course's enrollment data (regardless of role or state).
- avg_score: the pooled mean of every individual submission score in the course — the sum of all submission scores divided by the total number of submissions. Each submission counts equally; do NOT average the per-student averages.
- pass_rate: the percentage of students whose average score is 60 or above, expressed on a 0-100 scale. The denominator is the number of students who have at least one graded submission (not the total enrollment count): pass_rate = (number of students with average score >= 60) / (number of students with at least one graded submission) x 100.

The second sheet should be named Risk_Distribution and contain columns for risk_level, student_count, and pct. Include one row for each risk level (High, Medium, Low) with the combined totals across both courses. The pct column holds values on the 0-100 scale (e.g., write 9.8 to mean 9.8%), not 0-1 fractions.

The third sheet should be named At_Risk_Students and contain columns for course_name, high_risk_count, medium_risk_count, and low_risk_count. Include one row for each course showing how many students fall into each risk category.

The fourth sheet should be named Intervention_Plan and contain columns for risk_level, action, timeline, and responsible. Include one row for each risk level describing the recommended intervention: High risk students should receive immediate one-on-one advising within 1 week by an academic advisor, Medium risk students should receive group tutoring sessions within 2 weeks by a course tutor, and Low risk students should receive a self-paced study resources email within 1 month by the student success office. (Use the numeric forms "1 week", "2 weeks", and "1 month" for the timeline values, matching the scoring_model.json file.)

Create a Word document called Intervention_Plan.docx in your workspace. The document should contain a title "Student Retention Intervention Plan", a brief executive summary describing the analysis scope and key findings, a section on risk distribution summarizing the percentages across both courses, and a section detailing the recommended interventions for each risk level.

Set up a database in the team wiki system called "Student Risk Tracker" with properties for Student Course (title), Risk Level (select with options High, Medium, Low), Student Count (number), Average Score (number), and Pass Rate (number). Add one entry for each course with the aggregated data.

Finally, send an email to academic_advisors@university.edu with the subject "Student Retention Risk Analysis - Action Required" summarizing the key findings, including the total number of high-risk students across both courses and the recommended next steps.
