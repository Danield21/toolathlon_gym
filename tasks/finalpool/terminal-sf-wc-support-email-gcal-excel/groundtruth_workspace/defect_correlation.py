"""
Defect correlation script.
Reads support_tickets.json and order_data.json, computes defect rate (tickets per 100 orders)
and writes quality_analysis_results.json.
"""
import json


def main():
    with open("support_tickets.json") as f:
        tickets = json.load(f)
    with open("order_data.json") as f:
        orders = json.load(f)

    total_orders = sum(o.get("order_count", 0) for o in orders)
    results = {}
    total_tickets = 0
    by_priority = {}
    for t in tickets:
        p = t.get("priority")
        c = t.get("ticket_count", 0)
        total_tickets += c
        by_priority[p] = c
    if total_orders > 0:
        for p, c in by_priority.items():
            rate = round(c * 100.0 / total_orders, 2)
            results[p] = {"ticket_count": c, "defect_rate_per_100_orders": rate}
    results["totals"] = {"total_tickets": total_tickets, "total_orders": total_orders}
    with open("quality_analysis_results.json", "w") as f:
        json.dump(results, f, indent=2)
    print("Wrote quality_analysis_results.json")


if __name__ == "__main__":
    main()
