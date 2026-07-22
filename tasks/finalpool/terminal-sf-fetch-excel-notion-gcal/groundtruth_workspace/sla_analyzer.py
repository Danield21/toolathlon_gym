"""SLA compliance analyzer.

Computes per-priority SLA compliance rates from the support ticket data
and benchmarks them against industry values.

Compliance formula (per SLA_Framework.pdf):
    compliance_rate = min(100, (sla_target / avg_response_time) * 100)
"""
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}


def compute_priority_compliance():
    """Return {priority: {tickets, avg_rt, target, compliance_pct}}."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute(
        'SELECT t."PRIORITY", COUNT(*) AS tickets, AVG(t."RESPONSE_TIME_HOURS"), '
        '       AVG(t."CUSTOMER_SATISFACTION"), s."RESPONSE_TARGET_HOURS" '
        'FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS" t '
        'JOIN sf_data."SUPPORT_CENTER__PUBLIC__SLA_POLICIES" s '
        '   ON t."PRIORITY" = s."PRIORITY" '
        'GROUP BY t."PRIORITY", s."RESPONSE_TARGET_HOURS"'
    )
    out = {}
    for pri, n, avg_rt, csat, target in cur.fetchall():
        rate = min(100.0, float(target) / float(avg_rt) * 100.0)
        out[str(pri)] = {
            "tickets": int(n),
            "avg_response_time": round(float(avg_rt), 2),
            "avg_satisfaction": round(float(csat), 2) if csat else 0.0,
            "sla_target_hours": float(target),
            "compliance_pct": round(rate, 1),
        }
    cur.close()
    conn.close()
    return out


def compare_with_benchmark(internal, benchmark):
    """Compare overall internal metrics against benchmark.

    benchmark: dict with avg_response_time, customer_satisfaction, sla_compliance.
    Returns rows of (metric, internal_value, benchmark, gap, status).
    """
    rows = []
    rows.append(
        (
            "Average Response Time (hours)",
            internal["avg_response_time"],
            benchmark["avg_response_time"],
            internal["avg_response_time"] - benchmark["avg_response_time"],
            "Below" if internal["avg_response_time"] > benchmark["avg_response_time"] else "Meets",
        )
    )
    rows.append(
        (
            "Customer Satisfaction",
            internal["customer_satisfaction"],
            benchmark["customer_satisfaction"],
            internal["customer_satisfaction"] - benchmark["customer_satisfaction"],
            "Below" if internal["customer_satisfaction"] < benchmark["customer_satisfaction"] else "Meets",
        )
    )
    rows.append(
        (
            "SLA Compliance Rate (%)",
            internal["sla_compliance"],
            benchmark["sla_compliance"],
            internal["sla_compliance"] - benchmark["sla_compliance"],
            "Below" if internal["sla_compliance"] < benchmark["sla_compliance"] else "Meets",
        )
    )
    return rows


if __name__ == "__main__":
    pri = compute_priority_compliance()
    print("Per-priority compliance:")
    for p, r in pri.items():
        print(f"  {p}: {r}")
