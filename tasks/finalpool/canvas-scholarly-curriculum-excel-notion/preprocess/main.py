"""Preprocess script for canvas-scholarly-curriculum-excel-notion."""
import os
import argparse, json, os, sys, shutil, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM notion.comments")
    cur.execute("DELETE FROM notion.blocks")
    cur.execute("DELETE FROM notion.pages")
    cur.execute("DELETE FROM notion.databases")
    cur.execute("DELETE FROM scholarly.scholar_papers")
    cur.execute("DELETE FROM scholarly.arxiv_papers")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    # Inject scholarly papers matching task topics: machine learning,
    # data analytics, computational thinking.
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2401.00001', 'Recent Advances in Machine Learning', '[{"name": "Author A"}]'::jsonb, 'A survey of recent advances in machine learning techniques and their applications.',
        '["cs.LG", "cs.AI"]'::jsonb, 'cs.LG', 'https://arxiv.org/pdf/2401.00001', '2024-01-10')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2402.00002', 'Modern Data Analytics Techniques', '[{"name": "Author B"}]'::jsonb, 'A comprehensive review of modern data analytics methods including statistical approaches and machine learning.',
        '["cs.DB", "stat.ML"]'::jsonb, 'cs.DB', 'https://arxiv.org/pdf/2402.00002', '2024-02-15')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2403.00003', 'Computational Thinking in Education', '[{"name": "Author C"}]'::jsonb, 'A framework for teaching computational thinking skills in undergraduate computer science programs.',
        '["cs.CY"]'::jsonb, 'cs.CY', 'https://arxiv.org/pdf/2403.00003', '2024-03-05')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2404.00004', 'Deep Learning for Pattern Recognition', '[{"name": "Author F"}]'::jsonb, 'Recent progress in deep learning approaches for pattern recognition tasks.',
        '["cs.LG", "cs.CV"]'::jsonb, 'cs.LG', 'https://arxiv.org/pdf/2404.00004', '2024-04-01')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2405.00005', 'Big Data Analytics Methods', '[{"name": "Author G"}]'::jsonb, 'Methods for analyzing large-scale datasets in business and scientific contexts.',
        '["cs.DB"]'::jsonb, 'cs.DB', 'https://arxiv.org/pdf/2405.00005', '2024-05-12')""")
    # Noise papers - off-topic
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2304.99901', 'Quantum Computing Basics', '[{"name": "Author D"}]'::jsonb, 'Introduction to quantum computing.',
        '["quant-ph"]'::jsonb, 'quant-ph', 'https://arxiv.org/pdf/2304.99901', '2023-04-01')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2305.99902', 'Ocean Modeling Techniques', '[{"name": "Author E"}]'::jsonb, 'Advanced ocean modeling.',
        '["physics.ao-ph"]'::jsonb, 'physics.ao-ph', 'https://arxiv.org/pdf/2305.99902', '2023-05-15')""")
    # Noise notion data
    cur.execute("""INSERT INTO notion.pages (id, parent, properties, archived)
        VALUES ('noise-page-001',
        '{"type": "workspace", "workspace": true}'::jsonb,
        '{"title": {"id": "title", "type": "title", "title": [{"type": "text", "text": {"content": "Meeting Notes Archive"}}]}}'::jsonb,
        false)""")
    conn.commit()
    cur.close()
    conn.close()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)

if __name__ == "__main__":
    main()