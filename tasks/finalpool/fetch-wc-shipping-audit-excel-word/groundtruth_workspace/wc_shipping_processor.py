"""WooCommerce shipping audit processor.

Reads collected JSON inputs from API + WooCommerce sources, computes per-category
gap metrics between internal averages and external benchmark, and emits
wc_shipping_results.json.
"""
import json
import sys
import os
from typing import List, Dict


def load_json(path: str):
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def analyze(internal: List[Dict], external: Dict[str, float]):
    rows = []
    for rec in internal:
        cat = rec.get("category")
        ours = float(rec.get("our_avg_price", 0))
        market = float(external.get(cat, ours))
        gap_pct = ((ours - market) / market * 100) if market else 0.0
        rows.append({
            "category": cat,
            "product_count": rec.get("product_count"),
            "our_avg_price": ours,
            "total_sales": rec.get("total_sales"),
            "market_avg_price": market,
            "price_gap_pct": round(gap_pct, 2),
        })
    return rows


def main():
    internal_path = sys.argv[1] if len(sys.argv) > 1 else "internal.json"
    external_path = sys.argv[2] if len(sys.argv) > 2 else "external.json"
    internal = load_json(internal_path) if os.path.exists(internal_path) else []
    external = load_json(external_path) if os.path.exists(external_path) else {}
    out = {"rows": analyze(internal, external)}
    with open("wc_shipping_results.json", "w", encoding="utf-8") as fh:
        json.dump(out, fh, indent=2)


if __name__ == "__main__":
    main()
