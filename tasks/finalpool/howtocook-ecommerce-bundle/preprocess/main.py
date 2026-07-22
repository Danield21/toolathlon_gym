"""
Preprocess script for howtocook-ecommerce-bundle task.

WooCommerce, Yahoo Finance, and HowToCook are all read-only data sources.
No writable schemas are used in this task, so preprocessing is minimal.

This script performs sanity checks on required read-only data so the task
fails fast if the backing data is missing (rather than silently falling
back to the static groundtruth Excel in evaluation).
"""

import argparse
import os
import sys


DB_CONFIG = dict(
    host=os.environ.get("PGHOST", "localhost"),
    port=int(os.environ.get("PGPORT", "5432")),
    dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
    user=os.environ.get("PGUSER", "eigent"),
    password=os.environ.get("PGPASSWORD", "camel"),
)


def sanity_check_data():
    """Sanity-check read-only data sources the task depends on."""
    try:
        import psycopg2
    except ImportError:
        print("[preprocess][WARN] psycopg2 not available; skipping sanity check.")
        return True
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
    except Exception as e:
        print(f"[preprocess][WARN] Could not connect to PostgreSQL: {e}")
        return True  # non-fatal

    ok = True
    try:
        cur.execute(
            "SELECT COUNT(*) FROM wc.products WHERE categories::text LIKE %s",
            ("%Home Appliances%",),
        )
        hc = cur.fetchone()[0]
        if hc == 0:
            print("[preprocess][ERROR] wc.products has NO Home Appliances entries.")
            ok = False
        else:
            print(f"[preprocess] wc.products Home Appliances count = {hc}")

        cur.execute(
            "SELECT COUNT(*) FROM yf.stock_prices WHERE symbol = 'GC=F'"
        )
        gc = cur.fetchone()[0]
        if gc < 5:
            print(f"[preprocess][ERROR] yf.stock_prices GC=F rows = {gc}, need >=5.")
            ok = False
        else:
            print(f"[preprocess] yf.stock_prices GC=F count = {gc}")
    finally:
        cur.close()
        conn.close()
    return ok


def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    workspace = args.agent_workspace
    os.makedirs(workspace, exist_ok=True)

    print("[preprocess] Running read-only data sanity checks ...")
    ok = sanity_check_data()
    if not ok:
        print("[preprocess][FATAL] Required read-only data missing; aborting so evaluation does not silently fall back to static GT.")
        sys.exit(2)

    print("[preprocess] No writable schemas to inject. Workspace ready.")
    print(f"[preprocess] Agent workspace: {workspace}")
    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
