"""
Forecast builder script.
Reads historical_sales.json and market_outlook.json, computes regional forecasts,
and writes sales_forecast.json.
"""
import json


def main():
    with open("historical_sales.json") as f:
        historical = json.load(f)
    with open("market_outlook.json") as f:
        outlook = json.load(f)

    # Build per-region growth rate lookup
    growth_by_region = {r["region"]: r["growth_rate_pct"] for r in outlook}

    forecasts = []
    for r in historical:
        region = r["region"]
        rate = growth_by_region.get(region, 0)
        forecasted = round(r["current_revenue"] * (1 + rate / 100), 2)
        forecasts.append({
            "region": region,
            "current_orders": r["current_orders"],
            "current_revenue": r["current_revenue"],
            "growth_rate_pct": rate,
            "forecasted_revenue": forecasted,
            "revenue_increase": round(forecasted - r["current_revenue"], 2),
        })

    with open("sales_forecast.json", "w") as f:
        json.dump(forecasts, f, indent=2)


if __name__ == "__main__":
    main()
