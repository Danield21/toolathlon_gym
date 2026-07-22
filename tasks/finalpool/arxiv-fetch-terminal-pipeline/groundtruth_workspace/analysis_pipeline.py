"""
Paper analysis pipeline.
Reads paper_data.json, computes statistics, and writes analysis_results.json.
"""
import json
import os
import sys
from collections import Counter


def main():
    workspace = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(workspace, "paper_data.json")
    output_path = os.path.join(workspace, "analysis_results.json")

    if not os.path.exists(input_path):
        print(f"ERROR: {input_path} not found. Please create paper_data.json first.")
        sys.exit(1)

    with open(input_path, "r") as f:
        papers = json.load(f)

    if not isinstance(papers, list) or len(papers) == 0:
        print("ERROR: paper_data.json should contain a non-empty list of papers.")
        sys.exit(1)

    total_papers = len(papers)
    citations = [int(p.get("citations", 0)) for p in papers]
    avg_citations = sum(citations) / total_papers if total_papers > 0 else 0

    # Most cited paper - return the title
    most_cited = max(papers, key=lambda p: int(p.get("citations", 0)))
    most_cited_title = most_cited.get("title", "")

    # Papers per venue
    venue_counts = Counter(p.get("venue", "Unknown") for p in papers)

    results = {
        "average_citations": round(avg_citations, 1),
        "most_cited_paper": most_cited_title,
        "papers_per_venue": dict(venue_counts),
        "total_papers": total_papers,
    }

    with open(output_path, "w") as f:
        json.dump(results, f, indent=2)

    print(f"Analysis complete. Results written to {output_path}")
    print(f"  Total papers: {total_papers}")
    print(f"  Average citations: {avg_citations}")
    print(f"  Most cited: {most_cited_title}")


if __name__ == "__main__":
    main()
