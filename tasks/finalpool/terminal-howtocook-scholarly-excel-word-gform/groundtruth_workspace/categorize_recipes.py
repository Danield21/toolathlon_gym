#!/usr/bin/env python3
"""Categorize recipes by type and difficulty."""
import json
import sys


def categorize(recipes):
    """Group recipes by category."""
    out = {}
    for r in recipes:
        cat = r.get("category", "unknown")
        out.setdefault(cat, []).append(r)
    return out


def main():
    if len(sys.argv) < 2:
        print("Usage: categorize_recipes.py <recipes.json>", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1]) as f:
        recipes = json.load(f)
    result = categorize(recipes)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
