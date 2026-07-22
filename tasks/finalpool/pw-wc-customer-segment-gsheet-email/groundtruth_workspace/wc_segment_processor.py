#!/usr/bin/env python3
"""WC customer segment / category pricing processor.

Reads benchmark market data and internal product/order data from JSON files
in the workspace and writes wc_segment_results.json comparing internal
average prices against external market averages by category segment.
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
    internal = load_json("internal_products.json")
    rows = []
    bench_cats = benchmark.get("categories", {})
    for cat, metrics in internal.get("categories", {}).items():
        market_avg = bench_cats.get(cat, {}).get("avg_price", 0)
        our_avg = metrics.get("avg_price", 0)
        gap_pct = ((our_avg - market_avg) / market_avg * 100
                   if market_avg else 0)
        rows.append({
            "category": cat,
            "product_count": metrics.get("product_count", 0),
            "our_avg_price": round(our_avg, 2),
            "total_sales": metrics.get("total_sales", 0),
            "market_avg_price": round(market_avg, 2),
            "price_gap_pct": round(gap_pct, 2),
        })
    rows.sort(key=lambda r: (r.get("category") or "").lower())
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "wc_segment_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"results": rows}, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
