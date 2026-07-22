#!/usr/bin/env python3
"""Build evidence-based menus combining recipes with research findings."""
import json
import sys


def build(categorized_recipes, research_findings):
    """Create 5-day menu plan using evidence-based recipe selections."""
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    menu = []
    all_recipes = []
    for recipes in categorized_recipes.values():
        all_recipes.extend(recipes)
    for i, day in enumerate(days):
        menu.append({
            "day": day,
            "main": all_recipes[i] if i < len(all_recipes) else None,
            "research_rationale": list(research_findings.keys())[:1] if research_findings else [],
        })
    return menu


def main():
    if len(sys.argv) < 3:
        print("Usage: build_menus.py <categorized.json> <research.json>", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1]) as f:
        categorized = json.load(f)
    with open(sys.argv[2]) as f:
        research = json.load(f)
    result = build(categorized, research)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
