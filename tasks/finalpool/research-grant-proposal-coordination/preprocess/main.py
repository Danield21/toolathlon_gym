#!/usr/bin/env python3
"""Preprocess script for research-grant-proposal-coordination.

Seeds researcher CV submission emails so the agent can extract them, clears
writable schemas, and stages literature_sources from initial_workspace.
"""

import os
import shutil
from argparse import ArgumentParser
from datetime import datetime, timedelta
from pathlib import Path

import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
}

TASK_ROOT = Path(__file__).resolve().parent.parent


def get_conn():
    return psycopg2.connect(**DB)


def clear_writable_schemas():
    conn = get_conn(); cur = conn.cursor()
    try:
        cur.execute("DELETE FROM email.messages WHERE message_id LIKE 'rgpc-%'")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM gcal.events WHERE id LIKE 'rgpc-%'")
    except Exception:
        pass
    try:
        cur.execute("DELETE FROM notion.blocks WHERE id LIKE 'rgpc-%'")
        cur.execute("DELETE FROM notion.pages WHERE id LIKE 'rgpc-%'")
    except Exception:
        pass
    conn.commit(); cur.close(); conn.close()


def seed_researcher_submissions(launch_time):
    """Seed inbox with researcher CV/proposal-section emails."""
    conn = get_conn(); cur = conn.cursor()
    submissions = [
        ("rgpc-cv-lee@inst.edu", "Grant Proposal CV Submission - S. Lee", "lee@inst.edu",
         "PI Lee CV. Focus: Transformer Optimization. arxiv 2401.12345. Available 30% time. Significant publication: Transformer Optimization (2024)."),
        ("rgpc-cv-wang@inst.edu", "Grant Proposal CV Submission - T. Wang", "wang@inst.edu",
         "Co-I Wang CV. Focus: Few-Shot Learning. arxiv 2312.54321. Available 25% time."),
        ("rgpc-cv-kumar@inst.edu", "Grant Proposal CV Submission - R. Kumar", "kumar@inst.edu",
         "Co-I Kumar CV. Focus: Vision Transformers. arxiv 2311.98765. Available 20% time."),
        ("rgpc-cv-zhang@inst.edu", "Grant Proposal CV Submission - Q. Zhang", "zhang@inst.edu",
         "Co-I Zhang CV. Focus: Prompt Engineering. arxiv 2310.11111. Available 15% time."),
    ]
    for mid, subj, frm, body in submissions:
        cur.execute("SELECT id FROM email.messages WHERE message_id = %s", (mid,))
        if cur.fetchone():
            continue
        cur.execute("""
            INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date,
                body_text, is_read, is_important, is_flagged, size, created_at, headers)
            VALUES (3118, %s, %s, %s, '["pi@inst.edu"]'::jsonb, NOW(), %s,
                    false, false, false, %s, NOW(), '{}'::jsonb)
        """, (mid, subj, frm, body, len(body)))
    conn.commit(); cur.close(); conn.close()


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    seed_researcher_submissions(args.launch_time)

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        # Stage initial_workspace files (already includes literature_sources.csv + grant_template.docx)
        src = TASK_ROOT / "initial_workspace"
        for item in src.iterdir():
            dst = agent_ws / item.name
            if item.is_file() and not dst.exists():
                shutil.copy(item, dst)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
