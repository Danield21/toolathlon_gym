#!/usr/bin/env python3
"""Preprocess script for customer-satisfaction-analytics-framework.

Cleans stale email and gcal rows whose subject/summary matches our verifier
ILIKE patterns so prior task runs cannot falsely satisfy the Phase 6 checks.

Snowflake / WooCommerce data is read-only and pre-injected upstream.
"""

from argparse import ArgumentParser
from pathlib import Path
import os


def get_conn():
    import psycopg2
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"),
        port=int(os.environ.get("PGPORT", "5432")),
        dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
        user=os.environ.get("PGUSER", "eigent"),
        password=os.environ.get("PGPASSWORD", "camel"),
    )


def cleanup_phase6_artifacts():
    """Delete stale email/gcal rows matching verifier ILIKE patterns.

    The delete set MUST be a superset of the evaluator's Phase 6 match set
    (email subject/body and event summary, English + Chinese terms), so stale
    rows that could satisfy the checks are always cleared before a fresh run
    and can never cause a false PASS on re-run. Keep this in sync with
    evaluation/main.py check_phase6_distribution().
    """
    EMAIL_KW = ["satisfaction", "customer experience", "findings", "nps", "csat",
                "survey", "feedback",
                "满意度", "客户体验", "调查", "反馈", "净推荐", "洞察"]
    EVENT_KW = ["satisfaction", "review", "customer", "feedback", "findings",
                "满意度", "评审", "回顾", "复盘", "客户", "反馈", "洞察", "发现"]
    try:
        conn = get_conn()
        cur = conn.cursor()
        email_clause = (" OR ".join(["subject ILIKE %s"] * len(EMAIL_KW))
                        + " OR "
                        + " OR ".join(["body_text ILIKE %s"] * len(EMAIL_KW)))
        email_params = [f"%{k}%" for k in EMAIL_KW] + [f"%{k}%" for k in EMAIL_KW]
        cur.execute(f"DELETE FROM email.messages WHERE {email_clause}", email_params)
        event_clause = " OR ".join(["summary ILIKE %s"] * len(EVENT_KW))
        event_params = [f"%{k}%" for k in EVENT_KW]
        cur.execute(f"DELETE FROM gcal.events WHERE {event_clause}", event_params)
        conn.commit()
        conn.close()
        print("[preprocess] Cleaned stale email/gcal rows for Phase 6 checks.")
    except Exception as e:
        print(f"[preprocess] Cleanup skipped (DB not available): {e}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    cleanup_phase6_artifacts()

    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
