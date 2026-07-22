"""SLA analyzer.

Reads sla_raw_data.json (combined web benchmark + warehouse ticket data),
performs SLA comparison and gap analysis, then writes sla_results.json.
"""
import json
import os


def load_json(path):
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return {}


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    raw = load_json(os.path.join(here, "sla_raw_data.json"))

    benchmarks = {b.get("priority"): b for b in raw.get("benchmarks", [])}
    rows = []
    for our in raw.get("our_data", []):
        pri = our.get("priority")
        b = benchmarks.get(pri, {})
        our_resp = our.get("avg_response_hrs")
        ind_resp = b.get("avg_response_hrs")
        gap = round((our_resp or 0) - (ind_resp or 0), 2)
        rows.append({
            "priority": pri,
            "ticket_count": our.get("ticket_count"),
            "our_avg_response_hrs": our_resp,
            "industry_avg_response_hrs": ind_resp,
            "response_gap": gap,
            "avg_csat": our.get("avg_csat"),
            "compliance_status": "Compliant" if gap <= 0 else "Non-Compliant",
        })

    out_path = os.path.join(here, "sla_results.json")
    with open(out_path, "w") as f:
        json.dump({"rows": rows}, f, indent=2)


if __name__ == "__main__":
    main()
