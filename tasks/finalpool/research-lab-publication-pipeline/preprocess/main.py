#!/usr/bin/env python3
"""Preprocess script for research-lab-publication-pipeline."""

import os
import shutil
import json
from argparse import ArgumentParser
from pathlib import Path

import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
}

TASK_ROOT = Path(__file__).resolve().parent.parent


def clear_writable_schemas():
    conn = psycopg2.connect(**DB); cur = conn.cursor()
    # Full cleanup to prevent cross-task contamination — partial WHERE LIKE
    # 'rlpp-%' clauses would not match agent-generated record IDs.
    for tbl in [
        "email.messages",
        "email.attachments",
        "email.sent_log",
        "email.drafts",
        "gcal.events",
    ]:
        try:
            cur.execute(f"DELETE FROM {tbl}")
        except Exception:
            conn.rollback()
            continue
        conn.commit()
    cur.close(); conn.close()


def seed_publications():
    papers = json.loads(
        (TASK_ROOT / "files" / "research_papers.json").read_text(encoding="utf-8")
    )
    if len(papers) < 10:
        raise ValueError("research publication seed must contain at least 10 papers")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM arxiv_latex.papers")
        cur.execute("DELETE FROM arxiv.papers")
        for paper in papers:
            pdf_url = f"https://arxiv.org/pdf/{paper['id']}"
            cur.execute(
                """
                INSERT INTO arxiv.papers
                    (id, title, authors, summary, categories, primary_category,
                     published, updated, pdf_url, links, markdown_content,
                     is_downloaded)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, FALSE)
                """,
                (
                    paper["id"],
                    paper["title"],
                    json.dumps(paper["authors"]),
                    paper["abstract"],
                    json.dumps(paper["categories"]),
                    paper["categories"][0],
                    paper["published"],
                    paper["published"],
                    pdf_url,
                    json.dumps([{"href": pdf_url, "type": "application/pdf"}]),
                    paper["full_prompt"],
                ),
            )
            cur.execute(
                """
                INSERT INTO arxiv_latex.papers
                    (id, title, abstract, full_prompt, sections)
                VALUES (%s, %s, %s, %s, %s)
                """,
                (
                    paper["id"],
                    paper["title"],
                    paper["abstract"],
                    paper["full_prompt"],
                    json.dumps(paper["sections"]),
                ),
            )
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        cur.close()
        conn.close()


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clear_writable_schemas()
    seed_publications()

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)
        src = TASK_ROOT / "initial_workspace"
        for item in src.iterdir():
            dst = agent_ws / item.name
            if item.is_file() and not dst.exists():
                shutil.copy(item, dst)

    print("Preprocess completed successfully with seeded arXiv publications")


if __name__ == "__main__":
    main()
