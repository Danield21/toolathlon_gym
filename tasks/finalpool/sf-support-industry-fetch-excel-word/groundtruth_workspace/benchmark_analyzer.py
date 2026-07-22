"""Benchmark analyzer for support center performance vs industry standards."""
import json
import os
from pathlib import Path


def main():
    workspace = Path(os.path.dirname(os.path.abspath(__file__)))

    # Read inputs (created by agent in workspace)
    metrics_path = workspace / "support_metrics.json"
    benchmarks_path = workspace / "industry_benchmarks.json"
    if not metrics_path.exists() or not benchmarks_path.exists():
        return

    metrics = json.loads(metrics_path.read_text())
    benchmarks = json.loads(benchmarks_path.read_text())

    comparison = {
        "our_metrics": metrics,
        "industry_benchmarks": benchmarks,
        "comparison": {},
    }

    for key in ["avg_response_time", "resolution_rate", "customer_satisfaction"]:
        if key in metrics and key in benchmarks:
            our_val = metrics[key]
            ind_val = benchmarks[key]
            variance = our_val - ind_val
            # For response_time, lower is better; for others, higher is better
            if key == "avg_response_time":
                status = "Above" if our_val < ind_val else "Below"
            else:
                status = "Above" if our_val > ind_val else "Below"
            comparison["comparison"][key] = {
                "our_value": our_val,
                "industry_avg": ind_val,
                "variance": variance,
                "status": status,
            }

    out_path = workspace / "benchmark_comparison.json"
    out_path.write_text(json.dumps(comparison, indent=2))


if __name__ == "__main__":
    main()
