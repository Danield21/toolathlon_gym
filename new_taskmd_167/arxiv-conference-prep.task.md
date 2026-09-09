You are helping a researcher prepare for an upcoming academic conference on Reinforcement Learning from Human Feedback (RLHF). The conference is called "RLHF Summit 2026" and will take place on April 10, 2026 from 9:00 AM to 5:00 PM at the San Francisco Convention Center.

First, search the academic paper databases for papers related to "reinforcement learning from human feedback". You should find at least 5 relevant papers. Exclude any papers that are not directly about RLHF or human feedback in the context of language model training (for example, papers about self-play in board games should be excluded).

For each of the relevant RLHF papers you find, retrieve the full manuscript source to examine their methodology sections in detail. Pay attention to the key techniques each paper introduces.

Next, create a PowerPoint presentation called RLHF_Conference_Report.pptx in your workspace. The presentation should include the following slides:

First, a title slide with the title "RLHF Summit 2026 Conference Preparation Report" and subtitle "Key Papers in Reinforcement Learning from Human Feedback".

First, an overview slide that introduces the field of RLHF and why it matters for language model alignment.

First, one slide for each of the relevant papers (at least 5 slides). Each paper slide should include the paper title, the authors, and a brief summary of the key contribution and methodology. Make sure to include papers about instruction following with human feedback, learning to summarize from human feedback, constitutional AI, direct preference optimization, and proximal policy optimization.

First, a summary slide that synthesizes the key themes across all papers.

After creating the presentation, create a calendar event for the RLHF Summit 2026 conference. The event should be on April 10, 2026 from 9:00 AM to 5:00 PM, with the location set to "San Francisco Convention Center" and a description that mentions this is a conference about reinforcement learning from human feedback.

Finally, send an email to your collaborators at collaborators@rlhf-lab.org with the subject "RLHF Summit 2026 Preparation Materials". The email body should briefly describe that you have prepared a conference report covering key RLHF papers and created a calendar event for the conference. Mention the date (April 10, 2026) and location in the email.

If you are working in a multi-agent setup, designate exactly one agent to create and write RLHF_Conference_Report.pptx, and have the other agents coordinate with it — do not have multiple agents write to the same file, as concurrent writers can overwrite each other.

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

Before Wave 1, the main agent must freeze exactly five required paper slots—InstructGPT, learning to summarize from human feedback, Constitutional AI, DPO, and PPO—and the conference, presentation, calendar, and email contract. It must assign one stable paper identity per slot and exclude unrelated human-feedback or self-play work. Resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. Because the conference event is fixed and independent of paper research, the main agent must search for the exact event before dispatch and, in the same orchestration response that launches Wave 1, create at most the one missing event for 2026-04-10 from 09:00 to 17:00 at `San Francisco Convention Center`; preserve the calendar tool's default timezone, use no attendees, and retain the event ID for later read-back. Do not retry creation after an ambiguous failure.

1. Wave 1 — dispatch exactly 5 sub-agents in parallel, one for each numbered sub-task:

   1. Use one explore sub-agent to research instruction following with human feedback.
   2. Use one explore sub-agent to research learning to summarize from human feedback.
   3. Use one explore sub-agent to research Constitutional AI.
   4. Use one explore sub-agent to research Direct Preference Optimization.
   5. Use one explore sub-agent to research Proximal Policy Optimization.

Each Wave 1 sub-agent must search the granted academic sources read-only for only its assigned slot, validate title/abstract relevance, retrieve the full available manuscript/source, and return a compact source-backed record with authors, contribution, methodology, stable ID/title, source/section coverage, and missing-field status. Explore is read-only, so do not write files or create slides, calendar events, or email.

After Wave 1, the main agent must reconcile the five stable IDs and distinguish PPO's foundational algorithm role from direct language-model alignment evidence; it must freeze one literal canonical five-paper packet and reject invented claims. The deck and final email consume exactly this packet without re-researching or rephrasing factual fields.

2. Wave 2 — dispatch exactly 1 sub-agent for the following numbered sub-task:

   1. Use one coder sub-agent to own `<workspace_root>/RLHF_Conference_Report.pptx` exclusively. To avoid a long sequence of slide-by-slide calls, use one `python_execute` with the installed `python-pptx` package to create the complete deck atomically from the canonical packet: required title and overview slides, exactly one substantive slide for each of the five frozen papers, and a final synthesis slide. Then use the presentation read tools once to verify slide count, titles, five-paper coverage, and source-backed authors/methods. Return path and validation; do not create calendar or email state.

After Wave 2, the main agent must verify the deck and read back the retained calendar event, requiring exactly one matching RLHF Summit 2026 event on 2026-04-10 from 09:00 to 17:00 at the specified location. It must then search sent mail and send at most one missing email to collaborators@rlhf-lab.org with the exact subject and verified report/event details. Perform final deck, event, and message checks before completion.
