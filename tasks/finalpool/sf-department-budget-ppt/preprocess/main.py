"""
Preprocess for sf-department-budget-ppt task.

Snowflake is read-only. PDF is pre-generated in initial_workspace.
No writable schemas used. Only stale artifact cleanup in agent_workspace.
"""

import argparse
import os


def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("[preprocess] No DB preprocess needed for this task.")

    # Clean stale artifact files from agent_workspace (if task is re-run)
    if args.agent_workspace and os.path.isdir(args.agent_workspace):
        for fname in ["budget_vs_actuals.xlsx", "budget_vs_actuals.pptx"]:
            fpath = os.path.join(args.agent_workspace, fname)
            if os.path.exists(fpath):
                try:
                    os.remove(fpath)
                    print(f"[preprocess] Removed stale {fpath}")
                except Exception as e:
                    print(f"[preprocess] Could not remove {fpath}: {e}")


if __name__ == "__main__":
    main()
