"""Sector analyst script - reads financial_data.json and research_findings.json,
maps academic findings to sector performance, and writes sector_analysis.json."""
import json
import os


def main():
    workspace = os.path.dirname(os.path.abspath(__file__))
    fin_path = os.path.join(workspace, "financial_data.json")
    res_path = os.path.join(workspace, "research_findings.json")

    fin = {}
    if os.path.exists(fin_path):
        with open(fin_path) as f:
            fin = json.load(f)

    res = {}
    if os.path.exists(res_path):
        with open(res_path) as f:
            res = json.load(f)

    out = {
        "sector_summary": fin.get("sectors", {}),
        "research_mapping": res.get("findings", []),
    }
    with open(os.path.join(workspace, "sector_analysis.json"), "w") as f:
        json.dump(out, f, indent=2)


if __name__ == "__main__":
    main()
