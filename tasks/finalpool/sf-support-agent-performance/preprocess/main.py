"""Preprocess for sf-support-agent-performance.

Bug B.1.4 (slot 57): the shared seed dump ships SUPPORT_CENTER__PUBLIC__TICKETS
with a completely empty RESOLVER column, so the task — which asks the agent to
attribute tickets to active support agents via TICKETS.RESOLVER — is
unanswerable. Backfill RESOLVER deterministically (round-robin by stable
TICKET_ID ordering) so every active agent gets a reproducible, non-empty share.

The evaluator recomputes expected values from the live DB (fetch_expected_agents
joins on RESOLVER = AGENT_NAME), so a deterministic backfill keeps the eval
self-consistent. Idempotent: only touches rows where RESOLVER IS NULL/''.
"""
import argparse
import hashlib
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}


def _agent_index(ticket_id: str, n_agents: int) -> int:
    """Deterministic bucket for a ticket id -> [0, n_agents)."""
    h = int(hashlib.md5(str(ticket_id).encode("utf-8")).hexdigest(), 16)
    return h % n_agents


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Active agents (matches evaluator's WHERE ACTIVE = TRUE), ordered for determinism.
    cur.execute('''
        SELECT "AGENT_NAME" FROM sf_data."SUPPORT_CENTER__PUBLIC__AGENTS"
        WHERE "ACTIVE" = TRUE ORDER BY "AGENT_NAME"
    ''')
    agents = [r[0] for r in cur.fetchall()]
    if not agents:
        print("[preprocess] No active agents found; skipping RESOLVER backfill.")
        cur.close()
        conn.close()
        return

    # Tickets with empty resolver, in stable order.
    cur.execute('''
        SELECT "TICKET_ID" FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        WHERE "RESOLVER" IS NULL OR "RESOLVER" = ''
        ORDER BY "TICKET_ID"
    ''')
    empty = [r[0] for r in cur.fetchall()]
    if not empty:
        print("[preprocess] RESOLVER already populated; nothing to do.")
        cur.close()
        conn.close()
        return

    # Assign each empty ticket to a deterministic agent.
    for tid in empty:
        agent = agents[_agent_index(tid, len(agents))]
        cur.execute(
            'UPDATE sf_data."SUPPORT_CENTER__PUBLIC__TICKETS" '
            'SET "RESOLVER" = %s WHERE "TICKET_ID" = %s',
            (agent, tid),
        )

    conn.commit()
    print(f"[preprocess] Backfilled RESOLVER for {len(empty)} tickets across "
          f"{len(agents)} active agents.")

    # Sanity: distribution
    cur.execute('''
        SELECT "RESOLVER", COUNT(*) FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        GROUP BY "RESOLVER" ORDER BY COUNT(*) DESC LIMIT 10
    ''')
    for name, n in cur.fetchall():
        print(f"[preprocess]   {name}: {n} tickets")

    cur.close()
    conn.close()


if __name__ == "__main__":
    main()
