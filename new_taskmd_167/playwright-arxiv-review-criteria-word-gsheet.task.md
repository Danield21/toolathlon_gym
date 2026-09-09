You are a professor serving as a peer reviewer for a machine learning conference. You have been assigned three papers to review. The conference review portal is available at http://localhost:30214/review_criteria.html and contains the official scoring rubric, including criteria for technical soundness, novelty, clarity, and overall recommendation.

Read the Review_Guidelines.pdf in your workspace for additional context about how to structure your reviews and what the conference expects from reviewers.

First, browse the conference review portal to confirm the rubric format and scoring scale that the review documents must follow.

Then retrieve information about the three assigned papers from the research paper database. The papers are titled "Scaling Laws for Neural Language Models", "Training language models to follow instructions with human feedback", and "OPT: Open Pre-trained Transformer Language Models". Read each paper's full content and abstract. Additionally, examine the LaTeX source of each paper to analyze the methodology sections in more detail, looking at equations, experimental setup, and section structure.

Create three separate Word documents in your workspace, one for each paper review.

The first document should be called "Review_Scaling_Laws.docx". It should contain a heading with the paper title and authors. Then include sections for Summary (2-3 sentences describing the paper's contribution), Technical Soundness (score 1-5 with justification based on the rubric), Novelty (score 1-5 with justification), Clarity (score 1-5 with justification), and Overall Recommendation (Accept, Weak Accept, Borderline, Weak Reject, or Reject with a brief rationale). The lead reviewer has already finalized the assessment for this paper as Technical Soundness 5, Novelty 4, Clarity 4, with an overall Accept recommendation; the review document must reflect these exact values.

The second document should be called "Review_InstructGPT.docx" for the RLHF paper. Follow the same structure. The reviewer's finalized assessment for this paper records Technical Soundness 5, Novelty 5, Clarity 4, leading to an Accept recommendation; record those values verbatim.

The third document should be called "Review_OPT.docx" for the OPT paper. Follow the same structure. The reviewer's finalized assessment for this paper records Technical Soundness 4, Novelty 3, Clarity 5, leading to a Weak Accept recommendation; record those values verbatim.

Create a cloud spreadsheet titled "Conference Review Tracker" with a single sheet called "Reviews". The sheet should have columns: Paper_ID, Paper_Title, Technical_Soundness, Novelty, Clarity, Average_Score, Recommendation, Review_Status. Include one row for each of the three papers. Paper_ID should be the arxiv ID. Average_Score is the mean of the three scores rounded to 1 decimal. Review_Status should be "Completed" for all three. Sort rows by Average_Score descending.

When you are finished, call claim_done.

To solve this task efficiently, the main agent must first read the review portal and Review_Guidelines.pdf, freeze the rubric, and preserve the task's finalized scores and recommendations. Then review the three independent papers in one wave.

1. Wave 1 — dispatch exactly 3 coder sub-agents in parallel:

   1. One coder sub-agent retrieves and reads Scaling Laws for Neural Language Models, including abstract, full content, and LaTeX methodology, and owns Review_Scaling_Laws.docx. It must use Technical Soundness 5, Novelty 4, Clarity 4, and Accept exactly, then read the document back and return its path, arXiv ID, authors, and section/score checks.

   2. One coder sub-agent retrieves and reads Training language models to follow instructions with human feedback, including abstract, full content, and LaTeX methodology, and owns Review_InstructGPT.docx. It must use Technical Soundness 5, Novelty 5, Clarity 4, and Accept exactly, then read the document back and return its path, arXiv ID, authors, and section/score checks.

   3. One coder sub-agent retrieves and reads OPT: Open Pre-trained Transformer Language Models, including abstract, full content, and LaTeX methodology, and owns Review_OPT.docx. It must use Technical Soundness 4, Novelty 3, Clarity 5, and Weak Accept exactly, then read the document back and return its path, arXiv ID, authors, and section/score checks.

After Wave 1, the main agent must verify all three reviews and directly create the Conference Review Tracker with one Reviews sheet, the exact columns, three rows, one-decimal averages, descending Average_Score order, and Completed status. Read it back and call claim_done only after all four deliverables pass; do not dispatch another wave.
