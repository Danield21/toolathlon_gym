#!/usr/bin/env python3
"""Cook event survey processor.

Reads recipe and benchmark data from JSON files in the workspace and
writes cook_survey_results.json with the dishes selected for the event
survey.
"""
import json
import os


def load_json(name):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def run():
    benchmark = load_json("benchmark.json")
    recipes = load_json("recipes.json")
    cal_target = benchmark.get("calorie_target", 600)
    protein_target = benchmark.get("protein_target", 25)
    rows = []
    for r in recipes.get("recipes", []):
        meets = (
            r.get("calories", 0) <= cal_target
            and r.get("protein_g", 0) >= protein_target
        )
        rows.append({
            "name": r.get("name"),
            "category": r.get("category"),
            "calories": r.get("calories"),
            "protein_g": r.get("protein_g"),
            "meets_guidelines": meets,
        })
    rows.sort(key=lambda r: (r.get("name") or "").lower())
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "cook_survey_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({
            "calorie_target": cal_target,
            "protein_target": protein_target,
            "results": rows,
        }, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
