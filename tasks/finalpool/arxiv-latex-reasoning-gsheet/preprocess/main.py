"""
Preprocess for arxiv-latex-reasoning-gsheet task.
- Clears gsheet data so the agent starts fresh.
- Clears and injects reasoning papers into arxiv_latex.papers.
"""
import os
import argparse
import json

import psycopg2

# Env-driven DB config; must point at the same database as evaluation/main.py.
DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": os.environ.get("PGUSER", "eigent"),
    "password": os.environ.get("PGPASSWORD", "camel"),
}

REASONING_PAPERS = [
    {
        "id": "2201.11903",
        "title": "Chain-of-Thought Prompting Elicits Reasoning in Large Language Models",
        "abstract": "We explore how generating a chain of thought — a series of intermediate reasoning steps — significantly improves the ability of large language models to perform complex reasoning. We show that chain-of-thought prompting improves performance on arithmetic, commonsense, and symbolic reasoning benchmarks.",
    },
    {
        "id": "2203.11171",
        "title": "Self-Consistency Improves Chain of Thought Reasoning in Language Models",
        "abstract": "We introduce a new decoding strategy, self-consistency, to replace the naive greedy decoding used in chain-of-thought prompting. We first sample a diverse set of reasoning paths instead of only taking the greedy one, and then select the most consistent answer by marginalizing out the sampled reasoning paths.",
    },
    {
        "id": "2205.11916",
        "title": "Large Language Models are Zero-Shot Reasoners",
        "abstract": "Pretrained large language models are widely used in many sub-fields of natural language processing. We show that simply adding 'Let's think step by step' before each answer significantly enables zero-shot chain-of-thought reasoning in large language models.",
    },
    {
        "id": "2210.03493",
        "title": "Automatic Chain of Thought Prompting in Large Language Models",
        "abstract": "Large language models can perform chain-of-thought reasoning if provided with demonstrations of the reasoning process. We propose Auto-CoT that automatically constructs demonstrations with questions and reasoning chains using clustering and zero-shot CoT generation.",
    },
    {
        "id": "2305.10601",
        "title": "Tree of Thoughts: Deliberate Problem Solving with Large Language Models",
        "abstract": "Language models are increasingly deployed for general problem solving across a wide range of tasks, but are still confined to token-level, left-to-right decision-making. We introduce a new framework for language model inference, Tree of Thoughts, which generalizes over chain-of-thought prompting and enables exploration over coherent units of text.",
    },
]

# Noise papers about word embeddings - must NOT be included by agent
NOISE_PAPERS = [
    {
        "id": "1301.3781",
        "title": "Efficient Estimation of Word Representations in Vector Space",
        "abstract": "We propose two novel model architectures for computing continuous vector representations of words. We use the skip-gram and CBOW architectures for word2vec, a distributed representation learning system. These word representations capture semantic similarities between words.",
    },
    {
        "id": "1310.4546",
        "title": "Distributed Representations of Words and Phrases and their Compositionality",
        "abstract": "The skip-gram model is an efficient method for learning high-quality distributed word representations. We extend the skip-gram approach with several modifications that improve the quality of word embeddings and accelerate training.",
    },
    {
        "id": "1402.3722",
        "title": "GloVe: Global Vectors for Word Representation",
        "abstract": "Recent methods for learning word vector representations have succeeded in capturing fine-grained semantic and syntactic regularities. We propose a new global log-bilinear regression model that combines the advantages of global matrix factorization and local context window methods for word representation.",
    },
]


def clear_gsheet(conn):
    """Clear all Google Sheets data."""
    with conn.cursor() as cur:
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.spreadsheets")
    conn.commit()
    print("[preprocess] Cleared gsheet data")


def inject_arxiv_latex_papers(conn):
    """Clear and inject reasoning + noise (word embedding) papers into arxiv_latex.papers."""
    with conn.cursor() as cur:
        cur.execute("DELETE FROM arxiv_latex.papers")
        for p in REASONING_PAPERS + NOISE_PAPERS:
            # full_prompt is populated so the arxiv-latex MCP's get_paper_prompt
            # returns each paper's title and abstract (the title appears as a
            # markdown '#' heading; get_paper_abstract alone returns no title).
            full_prompt = f"# {p['title']}\n\n{p['abstract']}"
            cur.execute("""
                INSERT INTO arxiv_latex.papers (id, title, abstract, full_prompt)
                VALUES (%s, %s, %s, %s)
            """, (p["id"], p["title"], p["abstract"], full_prompt))
    conn.commit()
    print(f"[preprocess] Injected {len(REASONING_PAPERS)} reasoning + {len(NOISE_PAPERS)} noise papers into arxiv_latex.papers")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_gsheet(conn)
        inject_arxiv_latex_papers(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
