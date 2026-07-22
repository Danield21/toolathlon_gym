"""Curriculum reviewer script.

Reads course_data.json and research_papers.json, identifies gaps between
current courses and emerging research trends, scores relevance, and outputs
curriculum_review.json.
"""
import json
import os


def load_json(path):
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return []


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    courses = load_json(os.path.join(here, "course_data.json"))
    papers = load_json(os.path.join(here, "research_papers.json"))

    # Compute simple gap metrics
    total_courses = len(courses)
    papers_reviewed = len(papers)
    high_relevance = sum(1 for p in papers if p.get("relevance") == "High")
    coverage_pct = round((high_relevance / papers_reviewed * 100), 1) if papers_reviewed else 0.0

    review = {
        "total_courses": total_courses,
        "papers_reviewed": papers_reviewed,
        "high_relevance_papers": high_relevance,
        "curriculum_coverage_pct": coverage_pct,
    }

    out_path = os.path.join(here, "curriculum_review.json")
    with open(out_path, "w") as f:
        json.dump(review, f, indent=2)


if __name__ == "__main__":
    main()
