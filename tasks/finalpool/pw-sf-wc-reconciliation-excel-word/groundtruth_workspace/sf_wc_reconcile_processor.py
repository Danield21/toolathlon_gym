#!/usr/bin/env python3
"""SF/WC reconciliation processor.

Reads benchmark and internal regional sales data from JSON files in the
workspace and writes sf_wc_reconcile_results.json comparing internal
revenue against external market size by region.
"""
import json
import os


def load_json(name):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def run():
    benchmark = load_json("benchmark.json")
    internal = load_json("internal_orders.json")
    rows = []
    for region, metrics in internal.get("regions", {}).items():
        bench = benchmark.get("regions", {}).get(region, {})
        rev = metrics.get("revenue", 0)
        market_size_m = bench.get("market_size_m", 0)
        penetration = (rev / 1_000_000.0 / market_size_m * 100
                       if market_size_m else 0)
        rows.append({
            "region": region,
            "order_count": metrics.get("order_count", 0),
            "revenue": rev,
            "market_size_m": market_size_m,
            "market_penetration_pct": round(penetration, 2),
        })
    rows.sort(key=lambda r: (r.get("region") or "").lower())
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "sf_wc_reconcile_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"results": rows}, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
