"""
Evaluation script for arxiv-fetch-terminal-pipeline task.

Checks (aligned with task.md, federated learning topic):
1. paper_data.json with 4 papers (id, title, authors, citations, venue)
2. analysis_pipeline.py and analysis_results.json (average_citations,
   most_cited_paper, papers_per_venue, total_papers)
3. pipeline_report.txt summarizing the findings

Usage:
    python evaluation/main.py --agent_workspace <path> --groundtruth_workspace <path>
"""
import argparse
import json
import os
import re
import sys

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def float_close(a, b, tol=50.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


# Federated learning paper IDs (matches task.md topic and preprocess)
EXPECTED_PAPER_IDS = {"1602.05629", "1806.00582", "2002.05516", "1812.06127"}
EXPECTED_TOTAL_CITATIONS = 12000  # 6500 + 1700 + 800 + 3000
EXPECTED_AVG_CITATIONS = 3000.0
EXPECTED_VENUES = {"AISTATS", "arXiv", "NeurIPS", "MLSys"}
EXPECTED_MOST_CITED_KEYWORDS = ("communication-efficient", "decentralized")


def check_paper_data(agent_workspace):
    """Check paper_data.json: list of 4 papers with id, title, authors, citations, venue."""
    print("\n=== Checking paper_data.json ===")

    paper_data_path = os.path.join(agent_workspace, "paper_data.json")
    if not os.path.isfile(paper_data_path):
        check("paper_data.json exists", False, f"Not found: {paper_data_path}")
        return None

    check("paper_data.json exists", True)

    try:
        with open(paper_data_path, "r") as f:
            papers = json.load(f)
    except Exception as e:
        check("paper_data.json is valid JSON", False, str(e))
        return None
    check("paper_data.json is valid JSON", True)

    check("paper_data.json contains a list", isinstance(papers, list), f"Type: {type(papers)}")
    if not isinstance(papers, list):
        return None

    check("paper_data.json has 4 papers", len(papers) == 4, f"Found {len(papers)} papers")

    # Required fields per task.md: id, title, authors, citations, venue
    required_fields = ["id", "title", "authors", "citations", "venue"]
    for i, paper in enumerate(papers):
        if not isinstance(paper, dict):
            check(f"Paper {i+1} is dict", False, f"Type: {type(paper)}")
            continue
        for field in required_fields:
            check(f"Paper {i+1} has '{field}' field", field in paper,
                  f"Missing; got keys {list(paper.keys())}")

    # All 4 expected paper IDs present
    paper_ids = set()
    for p in papers:
        if isinstance(p, dict):
            paper_ids.add(str(p.get("id", "")).strip())
    overlap = paper_ids & EXPECTED_PAPER_IDS
    check("All 4 federated learning paper IDs in results",
          len(overlap) == 4,
          f"Found {len(overlap)}/4 matching: {overlap}")

    # Total citations equals sum
    total_cit = 0
    for p in papers:
        if isinstance(p, dict):
            try:
                total_cit += int(p.get("citations", 0))
            except (ValueError, TypeError):
                pass
    check("Total citations match expected",
          float_close(total_cit, EXPECTED_TOTAL_CITATIONS, tol=10),
          f"Got {total_cit}, expected {EXPECTED_TOTAL_CITATIONS}")

    return papers


def check_analysis_pipeline(agent_workspace):
    """Check that analysis_pipeline.py was created (per task.md)."""
    print("\n=== Checking analysis_pipeline.py ===")
    pipeline_path = os.path.join(agent_workspace, "analysis_pipeline.py")
    check("analysis_pipeline.py exists", os.path.isfile(pipeline_path),
          f"Not found: {pipeline_path}")


def check_analysis_results(agent_workspace, groundtruth_workspace):
    """Check analysis_results.json against task.md spec."""
    print("\n=== Checking analysis_results.json ===")

    agent_file = os.path.join(agent_workspace, "analysis_results.json")
    if not os.path.isfile(agent_file):
        check("analysis_results.json exists", False, f"Not found: {agent_file}")
        return

    check("analysis_results.json exists", True)

    try:
        with open(agent_file, "r") as f:
            agent_results = json.load(f)
    except Exception as e:
        check("analysis_results.json is valid JSON", False, str(e))
        return
    check("analysis_results.json is valid JSON", True)

    # Required keys per task.md
    for key in ("average_citations", "most_cited_paper", "papers_per_venue", "total_papers"):
        check(f"'{key}' key present in analysis_results.json", key in agent_results)

    # Total papers
    check("total_papers is 4",
          agent_results.get("total_papers") == 4,
          f"Got {agent_results.get('total_papers')}")

    # Average citations within tight tolerance
    avg_cit = agent_results.get("average_citations", 0)
    try:
        avg_val = float(avg_cit)
    except (TypeError, ValueError):
        avg_val = -1
    check("average_citations close to expected (3000)",
          float_close(avg_val, EXPECTED_AVG_CITATIONS, tol=2.0),
          f"Got {avg_cit}, expected ~{EXPECTED_AVG_CITATIONS}")

    # Most cited paper - require keyword from the McMahan paper title
    most_cited = (agent_results.get("most_cited_paper", "") or "").lower()
    has_keyword = any(kw in most_cited for kw in EXPECTED_MOST_CITED_KEYWORDS)
    check("most_cited_paper title is the McMahan FedAvg paper",
          has_keyword,
          f"Got '{most_cited}'")

    # Papers per venue: 4 distinct venues
    ppv = agent_results.get("papers_per_venue", {})
    if isinstance(ppv, dict):
        venue_count = sum(int(v) for v in ppv.values() if str(v).isdigit() or isinstance(v, (int, float)))
        check("papers_per_venue totals to 4 papers",
              venue_count == 4,
              f"Sum across venues = {venue_count}")
        # Expected venues coverage
        venues_lower = {k.lower() for k in ppv.keys()}
        missing = {v for v in EXPECTED_VENUES if v.lower() not in venues_lower}
        check(f"All expected venues present: {EXPECTED_VENUES}",
              len(missing) == 0,
              f"Missing: {missing}")
    else:
        check("papers_per_venue is a dict", False, f"Got type {type(ppv)}")


def check_pipeline_report(agent_workspace):
    """Check pipeline_report.txt exists and references the topic + papers."""
    print("\n=== Checking pipeline_report.txt ===")
    path = os.path.join(agent_workspace, "pipeline_report.txt")
    if not os.path.isfile(path):
        check("pipeline_report.txt exists", False, f"Not found: {path}")
        return
    check("pipeline_report.txt exists", True)
    try:
        with open(path) as f:
            content = f.read()
    except Exception as e:
        check("pipeline_report.txt readable", False, str(e))
        return
    # Word boundary check on topic and basic stats
    text_lower = content.lower()
    check("Report mentions 'federated learning'",
          re.search(r"\bfederated learning\b", text_lower) is not None)
    check("Report mentions paper count (4)",
          ("4" in content) and ("paper" in text_lower))
    check("Report mentions a citation number",
          any(token in content for token in ("6500", "12000", "3000")))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    check_paper_data(args.agent_workspace)
    check_analysis_pipeline(args.agent_workspace)
    check_analysis_results(args.agent_workspace, gt_dir)
    check_pipeline_report(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    print(f"\n=== SUMMARY ===")
    print(f"Results: {PASS_COUNT}/{total} passed, {FAIL_COUNT} failed")

    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
