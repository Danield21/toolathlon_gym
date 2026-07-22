#!/usr/bin/env python3
"""Validate menus against dietary constraints."""
import json
import sys


def validate(menus):
    """Return list of (day, issue)."""
    issues = []
    for m in menus:
        if m.get("main") is None:
            issues.append((m.get("day"), "no main dish"))
    return issues


def main():
    if len(sys.argv) < 2:
        print("Usage: validate_menus.py <menus.json>", file=sys.stderr)
        sys.exit(1)
    with open(sys.argv[1]) as f:
        menus = json.load(f)
    issues = validate(menus)
    if issues:
        print(json.dumps({"valid": False, "issues": issues}, indent=2))
        sys.exit(1)
    else:
        print(json.dumps({"valid": True}, indent=2))


if __name__ == "__main__":
    main()
