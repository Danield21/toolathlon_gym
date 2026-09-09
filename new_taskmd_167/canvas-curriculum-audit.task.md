The provost's office has requested a comprehensive curriculum compliance audit ahead of the upcoming accreditation review. The accreditation standards are published on a portal at http://localhost:30217 which lists the minimum structural requirements that each course must meet. Visit this portal to retrieve the current accreditation standards, and also follow the link to the content quality standards page for additional context on assessment difficulty targets.

Once the standards are retrieved, audit each course in the learning management system by counting the number of assignments, quizzes, and modules, and checking whether a syllabus has been published. Compare these counts against the accreditation thresholds from the portal to determine compliance status. A course is compliant only if it meets every structural requirement: at least 8 assignments, at least 3 quizzes, at least 4 modules, and a published syllabus. The content-quality / assessment-difficulty targets on the linked Content Quality Standards page are informational context only and must not affect the Compliant_YN determination.

Review the accreditation checklist file in the workspace for the scoring methodology and the course catalog file for reference information about the university's course offerings. Note that the learning management system is the authoritative source of courses and contains exactly 22 courses; the course_catalog.txt file is a partial, informational summary and does not list every course that must be audited.

Create an Excel file called Curriculum_Audit.xlsx with two sheets. The first sheet named "Compliance Matrix" should contain columns for Course, Assignments_Count, Quizzes_Count, Modules_Count, Has_Syllabus (Yes or No), Compliant_YN (Yes if all standards are met, No otherwise), and Issues (a semicolon-separated list of specific deficiencies, or "None" if compliant). Include all 22 courses sorted alphabetically by course name. The second sheet named "Summary" should contain rows for Total_Courses, Compliant_Courses, Non_Compliant_Courses, and Compliance_Rate (as a percentage).

Also create a Word document called Audit_Report.docx containing a narrative compliance report. The document should include an executive summary describing the overall compliance rate and key findings, a section detailing the most common areas of non-compliance, and recommendations for bringing non-compliant courses into alignment before the accreditation review.

To solve this task efficiently,

Follow this exact orchestration plan. Canvas is read-only. Before Wave 1, resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. After Wave 1, freeze one literal canonical 22-row compliance payload plus its four summary metrics; both writers consume exactly this payload without independently rescoring courses.

1. **Wave 1 — dispatch exactly 93 explore sub-agents concurrently:** render all 93 subtasks as parallel individual Agent calls in one orchestration response; do not use AgentSwarm because the wave mixes nine instruction templates.

- **Explore assignment rows 1-22:** call `canvas_list_assignments(course_id={course_id})` once; deduplicate IDs; return `{course_id,course_code,metric:"Assignments_Count",value,source_ids,duplicate_ids}`.
- **Explore quiz rows 23-44:** call `canvas_list_quizzes(course_id={course_id})` once, including unpublished; deduplicate IDs; return `{course_id,course_code,metric:"Quizzes_Count",value,source_ids,duplicate_ids}`.
- **Explore module rows 45-66:** call `canvas_list_modules(course_id={course_id})` once; deduplicate IDs; return `{course_id,course_code,metric:"Modules_Count",value,source_ids,duplicate_ids}`.
- **Explore syllabus rows 67-88:** call `canvas_get_syllabus(course_id={course_id})` once; set value `Yes` only for non-empty body; return `{course_id,course_code,metric:"Has_Syllabus",value}`.
- **Agent 89 — Explore portal structural standards:** navigate/snapshot `http://localhost:30217` once and return `{portal_url,min_assignments,min_quizzes,min_modules,syllabus_required,content_quality_link}`.
- **Agent 90 — Explore content-quality context:** navigate/snapshot `http://localhost:30217/standards.html` once and return `{url,assessment_difficulty_targets,informational_only:true}`.
- **Agent 91 — Explore checklist reader:** read `<workspace_root>/accreditation_checklist.md` completely and return `{path,binary_scoring_rule,all_requirements_rule,read_errors}`.
- **Agent 92 — Explore catalog reader:** read `<workspace_root>/course_catalog.txt` completely and return `{path,catalog_is_partial:true,reference_subjects,read_errors}`.
- **Agent 93 — Explore identity verifier:** call `canvas_list_courses(include_ended=true)` once; verify exactly the 22 table ID/code pairs, retain names, and return `{course_count,courses:[{course_id,course_code,course_name}],missing,unexpected,identity_mismatches}`.

|agent_no|template|course_id|course_code|
|---|---|---|---|
|1|assignment|1|AAA-2013J|
|2|assignment|2|AAA-2014J|
|3|assignment|3|BBB-2013J|
|4|assignment|4|BBB-2014J|
|5|assignment|5|BBB-2013B|
|6|assignment|6|BBB-2014B|
|7|assignment|7|CCC-2014J|
|8|assignment|8|CCC-2014B|
|9|assignment|9|DDD-2013J|
|10|assignment|10|DDD-2014J|
|11|assignment|11|DDD-2013B|
|12|assignment|12|DDD-2014B|
|13|assignment|13|EEE-2013J|
|14|assignment|14|EEE-2014J|
|15|assignment|15|EEE-2014B|
|16|assignment|16|FFF-2013J|
|17|assignment|17|FFF-2014J|
|18|assignment|18|FFF-2013B|
|19|assignment|19|FFF-2014B|
|20|assignment|20|GGG-2013J|
|21|assignment|21|GGG-2014J|
|22|assignment|22|GGG-2014B|
|23|quiz|1|AAA-2013J|
|24|quiz|2|AAA-2014J|
|25|quiz|3|BBB-2013J|
|26|quiz|4|BBB-2014J|
|27|quiz|5|BBB-2013B|
|28|quiz|6|BBB-2014B|
|29|quiz|7|CCC-2014J|
|30|quiz|8|CCC-2014B|
|31|quiz|9|DDD-2013J|
|32|quiz|10|DDD-2014J|
|33|quiz|11|DDD-2013B|
|34|quiz|12|DDD-2014B|
|35|quiz|13|EEE-2013J|
|36|quiz|14|EEE-2014J|
|37|quiz|15|EEE-2014B|
|38|quiz|16|FFF-2013J|
|39|quiz|17|FFF-2014J|
|40|quiz|18|FFF-2013B|
|41|quiz|19|FFF-2014B|
|42|quiz|20|GGG-2013J|
|43|quiz|21|GGG-2014J|
|44|quiz|22|GGG-2014B|
|45|module|1|AAA-2013J|
|46|module|2|AAA-2014J|
|47|module|3|BBB-2013J|
|48|module|4|BBB-2014J|
|49|module|5|BBB-2013B|
|50|module|6|BBB-2014B|
|51|module|7|CCC-2014J|
|52|module|8|CCC-2014B|
|53|module|9|DDD-2013J|
|54|module|10|DDD-2014J|
|55|module|11|DDD-2013B|
|56|module|12|DDD-2014B|
|57|module|13|EEE-2013J|
|58|module|14|EEE-2014J|
|59|module|15|EEE-2014B|
|60|module|16|FFF-2013J|
|61|module|17|FFF-2014J|
|62|module|18|FFF-2013B|
|63|module|19|FFF-2014B|
|64|module|20|GGG-2013J|
|65|module|21|GGG-2014J|
|66|module|22|GGG-2014B|
|67|syllabus|1|AAA-2013J|
|68|syllabus|2|AAA-2014J|
|69|syllabus|3|BBB-2013J|
|70|syllabus|4|BBB-2014J|
|71|syllabus|5|BBB-2013B|
|72|syllabus|6|BBB-2014B|
|73|syllabus|7|CCC-2014J|
|74|syllabus|8|CCC-2014B|
|75|syllabus|9|DDD-2013J|
|76|syllabus|10|DDD-2014J|
|77|syllabus|11|DDD-2013B|
|78|syllabus|12|DDD-2014B|
|79|syllabus|13|EEE-2013J|
|80|syllabus|14|EEE-2014J|
|81|syllabus|15|EEE-2014B|
|82|syllabus|16|FFF-2013J|
|83|syllabus|17|FFF-2014J|
|84|syllabus|18|FFF-2013B|
|85|syllabus|19|FFF-2014B|
|86|syllabus|20|GGG-2013J|
|87|syllabus|21|GGG-2014J|
|88|syllabus|22|GGG-2014B|

After Wave 1 reject mismatches/duplicates. Require thresholds 8 assignments, 3 quizzes, 4 modules, and syllabus. Mark Yes only if all four pass; content-quality values never affect status. List deficiencies in task-column order, or `None`; compute counts and percentage; sort by course name.

2. **Wave 2 — dispatch exactly 2 coder sub-agents in parallel:**

   1. **Coder Excel writer:** create `<workspace_root>/Curriculum_Audit.xlsx` with only `Compliance Matrix` and `Summary`; write exact seven headers, all 22 sorted canonical rows, and four vertical metrics. Read back and return `{path,sheet_names,row_counts,readback_ok}`.
   2. **Coder Word writer:** create `<workspace_root>/Audit_Report.docx` from the canonical payload with executive summary/rate, most common deficiency counts, and concrete recommendations; read text/info back and return `{path,sections,readback_ok}`.

Accept only after both read-backs.
