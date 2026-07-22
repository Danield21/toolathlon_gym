"""
correlation_analysis.py

Compute a comparison table showing each stock's YTD performance alongside
the company's revenue trend direction (increasing or decreasing compared
to the prior month). Also calculate the total company revenue across all
segments and the segment with the highest revenue contribution.

This is the groundtruth reference script. Real numbers come from
the SF data warehouse and Yahoo Finance at runtime.
"""
import json

# Stock YTD changes (from Yahoo Finance - example values for reference)
STOCKS = {
    "AMZN":   {"ytd_pct":  4.2},
    "GOOGL":  {"ytd_pct":  2.1},
    "JNJ":    {"ytd_pct": -1.5},
    "JPM":    {"ytd_pct":  3.8},
    "XOM":    {"ytd_pct":  6.7},
    "GC=F":   {"ytd_pct":  8.3},
    "^DJI":   {"ytd_pct":  2.9},
}

# Segment totals from sf_data.SALES_DW__PUBLIC__ORDERS joined with CUSTOMERS
SEGMENT_TOTALS = {
    "Consumer":   {"revenue": 839609.20, "orders": 5423},
    "Enterprise": {"revenue": 793741.69, "orders": 5058},
    "Government": {"revenue": 712686.66, "orders": 4679},
    "SMB":        {"revenue": 702960.78, "orders": 4840},
}

# Recent 12 months revenue (most recent first)
MONTHLY = [
    ("2026-02", 111602.65, 719),
    ("2026-01", 127327.90, 797),
    ("2025-12", 132763.25, 826),
    ("2025-11", 118232.96, 821),
    ("2025-10", 126333.75, 781),
    ("2025-09", 119056.54, 706),
    ("2025-08", 111374.86, 789),
    ("2025-07", 120898.58, 809),
    ("2025-06", 104099.62, 773),
    ("2025-05", 127379.39, 833),
    ("2025-04", 130634.57, 797),
    ("2025-03", 123878.58, 844),
]


def main():
    # Compute revenue trend (direction vs prior month)
    trend = "decreasing" if MONTHLY[0][1] < MONTHLY[1][1] else "increasing"

    # Total company revenue
    total_company_revenue = sum(s["revenue"] for s in SEGMENT_TOTALS.values())

    # Top segment
    top_segment = max(SEGMENT_TOTALS.items(), key=lambda kv: kv[1]["revenue"])[0]

    # Build comparison table
    comparison = []
    for tk, info in STOCKS.items():
        comparison.append({
            "ticker": tk,
            "ytd_pct": info["ytd_pct"],
            "revenue_trend": trend,
        })

    summary = {
        "total_company_revenue": total_company_revenue,
        "top_segment": top_segment,
        "revenue_trend": trend,
        "comparison": comparison,
    }
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
