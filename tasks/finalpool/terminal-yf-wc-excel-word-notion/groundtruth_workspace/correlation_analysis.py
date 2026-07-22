"""
correlation_analysis.py

Compares gold (GC=F) and AMZN historical price trends with the company's
product category margins. Outputs a summary of whether market conditions
suggest expanding or contracting margins.

Reference (groundtruth) script.
"""
import json

# Stock stats (from Yahoo Finance)
GOLD = {"avg_price": 3163.01, "pct_change": 138.7, "volatility": 796.12}
AMZN = {"avg_price": 206.45, "pct_change": 26.2, "volatility": 21.87}

# Product category margins (currently 40% across all)
CATEGORY_MARGINS = {
    "Audio": 40,
    "Cameras": 40,
    "Electronics": 40,
    "Home Appliances": 40,
    "TV & Home Theater": 40,
    "Watches": 40,
}

GOLD_THRESHOLD = 50.0  # %


def main():
    if GOLD["pct_change"] > GOLD_THRESHOLD:
        recommendation = "contracting"
        target_margin = 35
        action = "Monitor costs and consider price adjustment"
    else:
        recommendation = "expanding"
        target_margin = 40
        action = "Maintain current pricing"

    summary = {
        "gold_pct_change": GOLD["pct_change"],
        "amzn_pct_change": AMZN["pct_change"],
        "recommendation": recommendation,
        "target_margin_pct": target_margin,
        "action": action,
    }
    print("Market Condition Analysis:")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
