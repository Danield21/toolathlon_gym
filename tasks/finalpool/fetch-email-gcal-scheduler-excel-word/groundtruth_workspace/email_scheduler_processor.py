"""Email scheduler analysis processor.

Reads collected JSON inputs, performs gap analysis between internal and external
benchmark values, and emits aggregated metrics to email_scheduler_results.json.
"""
import json
import sys
import os


def load_data(path: str):
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def analyze(records):
    rows = []
    for rec in records:
        gap = float(rec.get("internal", 0)) - float(rec.get("external", 0))
        rows.append({
            "item": rec.get("item"),
            "internal": rec.get("internal"),
            "external": rec.get("external"),
            "gap": gap,
        })
    return rows


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "input.json"
    if not os.path.exists(src):
        records = []
    else:
        records = load_data(src)
    out = {"items": analyze(records)}
    with open("email_scheduler_results.json", "w", encoding="utf-8") as fh:
        json.dump(out, fh, indent=2)


if __name__ == "__main__":
    main()
