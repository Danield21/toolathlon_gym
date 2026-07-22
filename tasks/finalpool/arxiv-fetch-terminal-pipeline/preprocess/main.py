"""
Preprocess for arxiv-fetch-terminal-pipeline task.

Clears scholarly and arxiv tables and injects 4 federated learning papers.
Starts mock HTTP server on port 30151 for supplementary data fetch.

Prerequisites:
  - PostgreSQL toolathlon_gym database running on localhost:5432
"""
import argparse
import asyncio
import json
import os
import shutil
import tarfile

import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

# Federated learning papers (matches task.md topic)
TARGET_PAPERS = [
    {
        "arxiv_id": "1602.05629",
        "title": "Communication-Efficient Learning of Deep Networks from Decentralized Data",
        "authors": [{"name": "H. Brendan McMahan"}, {"name": "Eider Moore"}, {"name": "Daniel Ramage"}],
        "categories": ["cs.LG"],
        "primary_category": "cs.LG",
        "abstract": (
            "Modern mobile devices have access to a wealth of data suitable for learning models, "
            "which in turn can greatly improve the user experience. We advocate an alternative that "
            "leaves the training data distributed on the mobile devices, and learns a shared model by "
            "aggregating locally-computed updates. We term this decentralized approach Federated "
            "Learning. We present a practical method for the federated learning of deep networks based "
            "on iterative model averaging, and conduct an extensive empirical evaluation."
        ),
        "published": "2016-02-17",
        "pub_year": 2017,
        "venue": "AISTATS",
        "citation_count": 6500,
    },
    {
        "arxiv_id": "1806.00582",
        "title": "Federated Learning with Non-IID Data",
        "authors": [{"name": "Yue Zhao"}, {"name": "Meng Li"}, {"name": "Liangzhen Lai"}],
        "categories": ["cs.LG", "stat.ML"],
        "primary_category": "cs.LG",
        "abstract": (
            "Federated learning enables resource-constrained edge compute devices, such as mobile phones "
            "and IoT devices, to learn a shared model for prediction, while keeping the training data "
            "local. This decentralized approach to train models provides privacy, security, regulatory "
            "and economic benefits. In this work, we focus on the statistical challenge of federated "
            "learning when local data is non-IID."
        ),
        "published": "2018-06-02",
        "pub_year": 2018,
        "venue": "arXiv",
        "citation_count": 1700,
    },
    {
        "arxiv_id": "2002.05516",
        "title": "Personalized Federated Learning: A Meta-Learning Approach",
        "authors": [{"name": "Alireza Fallah"}, {"name": "Aryan Mokhtari"}, {"name": "Asuman Ozdaglar"}],
        "categories": ["cs.LG"],
        "primary_category": "cs.LG",
        "abstract": (
            "We study a personalized variant of the federated learning, in which our goal is to find a "
            "shared initial model that participating users can quickly adapt to their local datasets. "
            "We use a model-agnostic meta-learning approach to formulate this problem and analyze its "
            "convergence properties for both convex and non-convex settings."
        ),
        "published": "2020-02-13",
        "pub_year": 2020,
        "venue": "NeurIPS",
        "citation_count": 800,
    },
    {
        "arxiv_id": "1812.06127",
        "title": "Federated Optimization in Heterogeneous Networks",
        "authors": [{"name": "Tian Li"}, {"name": "Anit Kumar Sahu"}, {"name": "Manzil Zaheer"}],
        "categories": ["cs.LG"],
        "primary_category": "cs.LG",
        "abstract": (
            "Federated Learning is a distributed learning paradigm with two key challenges that "
            "differentiate it from traditional distributed optimization: (1) significant variability in "
            "terms of the systems characteristics on each device in the network (systems heterogeneity), "
            "and (2) non-identically distributed data across the network (statistical heterogeneity). "
            "We introduce a framework, FedProx, to tackle heterogeneity."
        ),
        "published": "2018-12-14",
        "pub_year": 2020,
        "venue": "MLSys",
        "citation_count": 3000,
    },
]

# Noise paper to test the agent's filtering ability
NOISE_PAPERS = [
    {
        "arxiv_id": "1906.10611",
        "title": "A Survey of Machine Learning for Big Code and Naturalness",
        "authors": [{"name": "Miltiadis Allamanis"}, {"name": "Earl T. Barr"}],
        "categories": ["cs.SE", "cs.LG"],
        "primary_category": "cs.SE",
        "abstract": (
            "Research at the intersection of machine learning and software engineering has recently "
            "seen a surge in interest. This survey is a comprehensive review of the state of the art "
            "in this area, covering probabilistic models of code, neural models for code analysis, "
            "code completion, bug detection, and program repair."
        ),
        "published": "2019-06-25",
        "pub_year": 2019,
        "venue": "ACM Computing Surveys",
        "citation_count": 900,
    },
]

MOCK_PORT = 30151


def clear_tables(conn):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM scholarly.arxiv_papers")
        cur.execute("DELETE FROM scholarly.scholar_papers")
        cur.execute("DELETE FROM arxiv.papers")
    conn.commit()
    print("Cleared scholarly, arxiv tables")


def inject_arxiv_papers(conn, papers):
    with conn.cursor() as cur:
        for p in papers:
            cur.execute("""
                INSERT INTO arxiv.papers
                (id, title, authors, summary, categories, primary_category,
                 published, updated, doi, journal_ref, comment, pdf_url,
                 links, markdown_content, is_downloaded)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title,
                    authors = EXCLUDED.authors,
                    summary = EXCLUDED.summary,
                    categories = EXCLUDED.categories,
                    primary_category = EXCLUDED.primary_category,
                    published = EXCLUDED.published,
                    markdown_content = EXCLUDED.markdown_content,
                    is_downloaded = EXCLUDED.is_downloaded
            """, (
                p["arxiv_id"],
                p["title"],
                json.dumps(p["authors"]),
                p["abstract"],
                json.dumps(p["categories"]),
                p["primary_category"],
                p["published"],
                p["published"],
                None,
                p.get("venue"),
                None,
                f"http://arxiv.org/pdf/{p['arxiv_id']}",
                json.dumps([]),
                "",
                False,
            ))
    conn.commit()
    print(f"Injected {len(papers)} papers into arxiv.papers")


def inject_scholarly_arxiv(conn, papers):
    with conn.cursor() as cur:
        for p in papers:
            cur.execute("""
                INSERT INTO scholarly.arxiv_papers
                (id, title, authors, abstract, categories, primary_category,
                 published, updated, doi, journal_ref, pdf_url, html_url, comment)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    title = EXCLUDED.title,
                    authors = EXCLUDED.authors,
                    abstract = EXCLUDED.abstract,
                    categories = EXCLUDED.categories,
                    primary_category = EXCLUDED.primary_category,
                    published = EXCLUDED.published
            """, (
                p["arxiv_id"],
                p["title"],
                json.dumps(p["authors"]),
                p["abstract"],
                json.dumps(p["categories"]),
                p["primary_category"],
                p["published"],
                p["published"],
                None,
                p.get("venue"),
                f"http://arxiv.org/pdf/{p['arxiv_id']}",
                f"http://arxiv.org/abs/{p['arxiv_id']}",
                None,
            ))
    conn.commit()
    print(f"Injected {len(papers)} papers into scholarly.arxiv_papers")


def inject_scholarly_scholar(conn, papers):
    with conn.cursor() as cur:
        for p in papers:
            cur.execute("""
                INSERT INTO scholarly.scholar_papers
                (title, authors, abstract, pub_year, venue, citation_count,
                 url, eprint_url, pub_url, bib)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                p["title"],
                json.dumps(p["authors"]),
                p["abstract"],
                p["pub_year"],
                p.get("venue"),
                p.get("citation_count", 0),
                f"http://arxiv.org/abs/{p['arxiv_id']}",
                f"http://arxiv.org/pdf/{p['arxiv_id']}",
                f"http://arxiv.org/abs/{p['arxiv_id']}",
                json.dumps({"title": p["title"], "year": p["pub_year"]}),
            ))
    conn.commit()
    print(f"Injected {len(papers)} papers into scholarly.scholar_papers")


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30151."""
    print("[preprocess] Setting up mock server...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.isfile(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)
        print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")
    else:
        print(f"[preprocess] Skipping mock server: tarball not found at {tar_path}")
        return

    mock_dir = os.path.join(tmp_dir, "mock_pages")

    # Kill any existing process on the port
    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{MOCK_PORT}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    # Start HTTP server
    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {MOCK_PORT} --directory {mock_dir} "
        f"> {mock_dir}/server.log 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{MOCK_PORT}")


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        clear_tables(conn)
        all_papers = TARGET_PAPERS + NOISE_PAPERS
        inject_arxiv_papers(conn, all_papers)
        inject_scholarly_arxiv(conn, all_papers)
        inject_scholarly_scholar(conn, all_papers)
    finally:
        conn.close()

    await setup_mock_server()
    print("\nPreprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
