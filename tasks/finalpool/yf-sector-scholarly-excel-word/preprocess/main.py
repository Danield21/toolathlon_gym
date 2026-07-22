"""Preprocess script for yf-sector-scholarly-excel-word."""
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
    cur.execute("DELETE FROM scholarly.scholar_papers")
    cur.execute("DELETE FROM scholarly.arxiv_papers")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    # Relevant papers — sector rotation / industry analysis / market cycles
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2301.10001', 'Sector Rotation Patterns', '[{"name": "Carter J."}]'::jsonb,
        'An empirical study of sector rotation patterns across business cycles, finding cyclical rotation every 3-5 years across all sectors and lead-lag relationships between leading and defensive sectors.',
        '["q-fin.PM", "econ.GN"]'::jsonb, 'q-fin.PM', 'https://arxiv.org/pdf/2301.10001', '2023-01-15')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2302.10002', 'Industry Momentum', '[{"name": "Wang L."}]'::jsonb,
        'A cross-sectional analysis of industry momentum strategies focused on the technology sector, showing that momentum persists 6-12 months but reverses thereafter.',
        '["q-fin.PM"]'::jsonb, 'q-fin.PM', 'https://arxiv.org/pdf/2302.10002', '2023-02-20')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2303.10003', 'Market Cycles and Defensive Allocation', '[{"name": "Patel R."}]'::jsonb,
        'A multi-decade study of market cycles showing defensive sectors (healthcare, consumer staples) outperform during contractions while financials lead expansions.',
        '["q-fin.PM", "econ.GN"]'::jsonb, 'q-fin.PM', 'https://arxiv.org/pdf/2303.10003', '2023-03-10')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2304.10004', 'Industry Analysis Framework', '[{"name": "Greene S."}]'::jsonb,
        'A framework for industry analysis combining macroeconomic indicators with sector-level fundamentals to forecast relative sector performance over 6-18 month horizons.',
        '["q-fin.PM"]'::jsonb, 'q-fin.PM', 'https://arxiv.org/pdf/2304.10004', '2023-04-05')""")
    # Noise papers (off-topic)
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2304.99901', 'Quantum Computing Basics', '[{"name": "Author D"}]'::jsonb, 'Introduction to quantum computing.',
        '["quant-ph"]'::jsonb, 'quant-ph', 'https://arxiv.org/pdf/2304.99901', '2023-04-01')""")
    cur.execute("""INSERT INTO scholarly.arxiv_papers (id, title, authors, abstract, categories, primary_category, pdf_url, published)
        VALUES ('2305.99902', 'Ocean Modeling Techniques', '[{"name": "Author E"}]'::jsonb, 'Advanced ocean modeling.',
        '["physics.ao-ph"]'::jsonb, 'physics.ao-ph', 'https://arxiv.org/pdf/2305.99902', '2023-05-15')""")
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