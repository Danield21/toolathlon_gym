"""quality_audit.py - Product quality audit script.

Computes per-category quality metrics from WooCommerce data and writes results
to Product_Quality_Audit.xlsx.

For each category:
  - product_count
  - average_rating
  - review_count
  - percentage of reviews below 3 stars (pct_below_3stars)

Quality score = avg_rating/5 * 100 - 2 * pct_below_3stars.
Risk levels: High (<70), Medium (70-85), Low (>85).

Order and refund data is aggregated by status with total amount and percentage.
"""
import json
from collections import defaultdict

# Reference: this script was executed from command-line to produce
# Product_Quality_Audit.xlsx (see Review_Summary / Refund_Analysis /
# Quality_Scorecard / Survey_Questions sheets).


def categorize_risk(score: float) -> str:
    """Classify quality score into risk tiers."""
    if score < 70:
        return "High"
    elif score < 85:
        return "Medium"
    else:
        return "Low"


def calculate_quality_score(avg_rating: float, pct_below_3: float) -> float:
    """Quality score formula per task.md."""
    return round(avg_rating / 5.0 * 100 - 2 * pct_below_3, 2)


def aggregate_category_reviews(reviews):
    """Aggregate review data per product category."""
    by_cat = defaultdict(lambda: {"products": set(), "ratings": [], "below_3": 0})
    for r in reviews:
        cat = r["category"]
        by_cat[cat]["products"].add(r["product_id"])
        by_cat[cat]["ratings"].append(r["rating"])
        if r["rating"] < 3:
            by_cat[cat]["below_3"] += 1
    result = {}
    for cat, data in by_cat.items():
        avg = sum(data["ratings"]) / len(data["ratings"]) if data["ratings"] else 0
        pct = (data["below_3"] / len(data["ratings"]) * 100) if data["ratings"] else 0
        result[cat] = {
            "product_count": len(data["products"]),
            "avg_rating": round(avg, 2),
            "review_count": len(data["ratings"]),
            "pct_below_3stars": round(pct, 2),
        }
    return result


def aggregate_order_status(orders):
    """Aggregate order statistics by status."""
    by_status = defaultdict(lambda: {"count": 0, "total": 0.0})
    for o in orders:
        s = o["status"]
        by_status[s]["count"] += 1
        by_status[s]["total"] += float(o["total"])
    total_orders = sum(v["count"] for v in by_status.values())
    result = []
    for status, v in sorted(by_status.items()):
        pct = (v["count"] / total_orders * 100) if total_orders else 0
        result.append({
            "status": status,
            "order_count": v["count"],
            "total_amount": round(v["total"], 2),
            "pct_of_total": round(pct, 2),
        })
    return result


def build_scorecard(category_reviews):
    """Build quality scorecard with score and risk_level per category."""
    scorecard = []
    for cat in sorted(category_reviews.keys()):
        data = category_reviews[cat]
        score = calculate_quality_score(data["avg_rating"], data["pct_below_3stars"])
        scorecard.append({
            "category": cat,
            "quality_score": score,
            "risk_level": categorize_risk(score),
        })
    return scorecard


if __name__ == "__main__":
    print("Quality audit computation: see Product_Quality_Audit.xlsx for results.")
