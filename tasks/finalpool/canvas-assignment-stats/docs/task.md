I need to understand the assignment distribution across our courses in the learning management system. Please pull assignment data for all courses — every course in the LMS across all semesters, with no filtering by term or date — and analyze the counts and types. List every course regardless of its workflow state (include available, completed, and ended courses alike) and regardless of whether you are enrolled in it; do not restrict the list to your own enrollments or to currently active terms.

Create an Excel file called Canvas_Assignment_Stats.xlsx with two sheets. The "Assignment Stats" sheet should have one row per course and columns Course_Code (the Canvas course code of the course, e.g. "CCC-2014J"), Total_Assignments (all assignments in the course), Avg_Points (the average of the assignments' points possible, rounded to 1 decimal), TMA_Count (count of Tutor Marked Assessments, i.e. assignments whose name starts with "TMA"), and CMA_Count (count of Computer Marked Assessments, i.e. assignments whose name starts with "CMA"). Note that some assignments (such as "Final Exam ...") start with neither "TMA" nor "CMA"; they still count toward Total_Assignments and Avg_Points but not toward TMA_Count or CMA_Count, so TMA_Count + CMA_Count may be less than Total_Assignments. Sort by Course_Code alphabetically.

The "Summary" sheet must use a two-column layout with headers "Metric" and "Value" and must contain exactly these four rows — one metric per row, all four rows required, with these exact metric names:
- Total_Assignments: total number of assignments across all courses.
- Total_TMAs: total number of TMA assignments across all courses.
- Total_CMAs: total number of CMA assignments across all courses.
- Course_Most_Assignments: the course code of the course with the most assignments.

Also create a Word document called Assignment_Overview.docx briefly summarizing the findings.
