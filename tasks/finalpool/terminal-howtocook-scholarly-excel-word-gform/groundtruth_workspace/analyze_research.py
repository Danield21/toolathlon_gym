#!/usr/bin/env python3
"""Analyze research findings and aggregate by theme."""
import json
import sys


def analyze(papers):
    """Aggregate papers by primary theme."""
    themes = {}
    for p in papers:
        key = p.get("theme", "general")
        themes.setdefault(key, []).append(p.get("title"))
    return themes


def main():
    if len(sys.argv) < 2:
        print("Usage: analyze_research.py <papers.json>", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1]) as f:
        papers = json.load(f)
    result = analyze(papers)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
