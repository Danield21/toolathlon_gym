"""Notion survey processor.

Reads collected survey JSON data, performs gap analysis between internal and
external benchmarks, then writes notion_survey_results.json.
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
    internal = load_json(os.path.join(here, "internal_data.json"))
    external = load_json(os.path.join(here, "external_benchmark.json"))

    items = internal.get("items", [])
    bench = {b.get("item"): b for b in external.get("benchmarks", [])}

    rows = []
    gaps = []
    for it in items:
        internal_v = it.get("value")
        external_v = bench.get(it.get("name"), {}).get("value")
        gap = (internal_v or 0) - (external_v or 0)
        gaps.append(gap)
        rows.append({
            "item": it.get("name"),
            "internal_value": internal_v,
            "external_benchmark": external_v,
            "gap": gap,
        })

    out = {
        "rows": rows,
        "metrics": {
            "total_items": len(rows),
            "avg_gap": (sum(gaps) / len(gaps)) if gaps else 0.0,
        },
    }
    out_path = os.path.join(here, "notion_survey_results.json")
    with open(out_path, "w") as f:
        json.dump(out, f, indent=2)


if __name__ == "__main__":
    main()
