"""Preprocess for terminal-wc-sf-revenue-reconcile-excel-ppt.
WC and SF data are read-only. Inject noise workspace files as interference.
"""
import argparse
import glob
import os

import openpyxl

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}


def inject_noise_workspace_files(workspace):
    """Inject noise Excel/PPT files as distractors to verify agent uses correct filenames."""
    # Noise 1: Old_Revenue_Report.xlsx with outdated values to tempt agent
    noise_xlsx = os.path.join(workspace, "Old_Revenue_Report.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Legacy"
    ws.append(["Status", "Order_Count", "Total_Revenue"])
    ws.append(["archived", 999, 999999.99])
    ws.append(["legacy", 12, 12345.67])
    wb.save(noise_xlsx)
    print(f"[preprocess] Injected noise file: {noise_xlsx}")

    # Noise 2: Draft_Report.xlsx with misleading content
    draft_xlsx = os.path.join(workspace, "Draft_Report.xlsx")
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "Draft"
    ws2.append(["This is a draft report, do not use"])
    ws2.append(["Placeholder", 0, 0])
    wb2.save(draft_xlsx)
    print(f"[preprocess] Injected noise file: {draft_xlsx}")

    # Noise 3: prior_reconciliation.json with fake data
    noise_json = os.path.join(workspace, "prior_reconciliation.json")
    with open(noise_json, "w") as f:
        f.write('{"note": "previous attempt from 2024, DO NOT USE", "wc_total": 0, "sf_total": 0}')
    print(f"[preprocess] Injected noise file: {noise_json}")


def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    print("[preprocess] WC and SF data are read-only, no DB injection needed.")

    if args.agent_workspace:
        # Remove previous agent outputs (also cleans prior noise)
        for pattern in ["Revenue_Reconciliation.xlsx", "Reconciliation_Presentation.pptx",
                        "reconciliation_*.py", "reconciliation_*.json",
                        "wc_*.json", "sf_*.json",
                        "Old_Revenue_Report.xlsx", "Draft_Report.xlsx",
                        "prior_reconciliation.json"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

        # Inject noise interference files
        if os.path.isdir(args.agent_workspace):
            inject_noise_workspace_files(args.agent_workspace)

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
