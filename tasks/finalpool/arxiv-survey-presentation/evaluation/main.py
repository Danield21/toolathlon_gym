"""
Evaluation for arxiv-survey-presentation task.
Checks Excel (Survey_Data.xlsx) and PowerPoint (NLP_Survey.pptx).
"""
import argparse
import os
import sys
from pathlib import Path

import openpyxl

PASS_COUNT = 0
FAIL_COUNT = 0

EXPECTED_PAPER_IDS = {"2404.00001", "2404.00002", "2404.00003", "2404.00004", "2404.00005"}


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower().replace(" ", "_") == sheet_name.strip().lower().replace(" ", "_"):
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
        if name.strip().lower().replace("_", " ") == sheet_name.strip().lower().replace("_", " "):
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def find_col(header, names):
    if not header:
        return None
    for i, cell in enumerate(header):
        if cell is None:
            continue
        cl = str(cell).strip().lower().replace(" ", "_")
        for n in names:
            if n.lower().replace(" ", "_") == cl:
                return i
    return None


def check_excel(workspace):
    print("\n=== Checking Excel ===")
    path = os.path.join(workspace, "Survey_Data.xlsx")
    if not os.path.isfile(path):
        record("Excel exists", False, f"Not found: {path}")
        return False
    record("Excel exists", True)

    wb = openpyxl.load_workbook(path, data_only=True)

    # Paper Comparison
    pc_rows = load_sheet_rows(wb, "Paper Comparison") or load_sheet_rows(wb, "Paper_Comparison")
    if pc_rows is None:
        record("Sheet 'Paper Comparison' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Paper Comparison' exists", True)
        data = pc_rows[1:]
        record("Paper Comparison has 5 rows", len(data) == 5, f"Found {len(data)}")

        id_col = find_col(pc_rows[0], ["Paper_ID", "Paper ID", "ID"])
        if id_col is not None:
            found = {str(r[id_col]).strip() for r in data if id_col < len(r) and r[id_col]}
            for pid in EXPECTED_PAPER_IDS:
                record(f"Paper {pid} present", pid in found, f"Found: {found}")

        method_col = find_col(pc_rows[0], ["Method", "method"])
        record("Method column exists", method_col is not None, f"Header: {pc_rows[0]}")

        dataset_col = find_col(pc_rows[0], ["Dataset", "dataset", "Dataset_Used"])
        record("Dataset column exists", dataset_col is not None, f"Header: {pc_rows[0]}")

        # Content check: Method column should cover known NLP methods (stronger: 3+ distinct)
        if method_col is not None:
            methods_text = " ".join(str(r[method_col]).lower() for r in data if method_col < len(r) and r[method_col])
            method_keywords = ["chain", "instruction", "retrieval", "self-consistency", "multimodal", "cot", "prompt"]
            matches = sum(1 for kw in method_keywords if kw in methods_text)
            record("Method column has >=3 NLP method keywords", matches >= 3, f"Matched {matches}, text: {methods_text[:200]}")

        # NEW: Per-paper Method / Dataset specific checks
        if method_col is not None and id_col is not None:
            expected_method = {
                "2404.00001": ["chain", "cot"],
                "2404.00002": ["instruction"],
                "2404.00003": ["retrieval"],
                "2404.00004": ["self-consistency", "consistency"],
                "2404.00005": ["multimodal", "visual"],
            }
            for row in data:
                pid = str(row[id_col]).strip() if id_col < len(row) and row[id_col] else ""
                if pid in expected_method:
                    actual = str(row[method_col] or "").lower() if method_col < len(row) else ""
                    ok = any(kw in actual for kw in expected_method[pid])
                    record(f"Method for {pid} mentions {expected_method[pid][0]}", ok,
                           f"Got '{actual[:80]}'")

        # Content check: Dataset column should cover known datasets (stronger: 3+ distinct)
        if dataset_col is not None:
            ds_text = " ".join(str(r[dataset_col]).lower() for r in data if dataset_col < len(r) and r[dataset_col])
            ds_keywords = ["gsm8k", "mmlu", "hotpotqa", "coco", "imagenet", "flan", "truthfulqa", "squad", "big-bench", "wmt", "natural questions", "vqa"]
            ds_matches = sum(1 for kw in ds_keywords if kw in ds_text)
            record("Dataset column has >=3 known datasets", ds_matches >= 3, f"Matched {ds_matches}, text: {ds_text[:200]}")

        # NEW: Per-paper Dataset check
        if dataset_col is not None and id_col is not None:
            expected_dataset = {
                "2404.00001": ["gsm8k"],
                "2404.00002": ["mmlu"],
                "2404.00003": ["natural questions", "hotpotqa"],
                "2404.00004": ["gsm8k"],
                "2404.00005": ["vqa"],
            }
            for row in data:
                pid = str(row[id_col]).strip() if id_col < len(row) and row[id_col] else ""
                if pid in expected_dataset:
                    actual = str(row[dataset_col] or "").lower() if dataset_col < len(row) else ""
                    ok = any(kw in actual for kw in expected_dataset[pid])
                    record(f"Dataset for {pid} mentions {expected_dataset[pid][0]}", ok,
                           f"Got '{actual[:80]}'")

    # Taxonomy
    tax_rows = load_sheet_rows(wb, "Taxonomy")
    if tax_rows is None:
        record("Sheet 'Taxonomy' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Taxonomy' exists", True)
        data = tax_rows[1:]
        record("Taxonomy has >= 2 categories", len(data) >= 2, f"Found {len(data)}")
        # Check that Category column exists
        cat_col = find_col(tax_rows[0], ["Category", "category", "Topic"])
        record("Taxonomy has Category column", cat_col is not None, f"Header: {tax_rows[0]}")
        if cat_col is not None:
            cats_text = " ".join(str(r[cat_col]).lower() for r in data if cat_col < len(r) and r[cat_col])
            record("Taxonomy references NLP topic keywords",
                   any(kw in cats_text for kw in ["reasoning", "prompt", "tuning", "retrieval", "multimodal", "generation", "training", "architecture", "framework"]),
                   f"Got: {cats_text[:200]}")

    # Summary Statistics
    ss_rows = load_sheet_rows(wb, "Summary Statistics") or load_sheet_rows(wb, "Summary_Statistics")
    if ss_rows is None:
        record("Sheet 'Summary Statistics' exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Sheet 'Summary Statistics' exists", True)
        metrics = {}
        for row in ss_rows[1:]:
            if row and row[0]:
                metrics[str(row[0]).strip().lower().replace(" ", "_")] = row[1] if len(row) > 1 else None

        tp_key = next((k for k in metrics if "total" in k and "paper" in k), None)
        if tp_key:
            try:
                record("Total_Papers = 5", abs(float(metrics[tp_key]) - 5) < 1, f"Got {metrics[tp_key]}")
            except (TypeError, ValueError):
                record("Total_Papers is numeric", False, f"Got {metrics[tp_key]}")

        # Average_Sections check — tighter: must equal 4.0 within ±0.2
        avg_sec_key = next((k for k in metrics if "average" in k and "section" in k), None)
        if avg_sec_key is None:
            avg_sec_key = next((k for k in metrics if "avg" in k and "section" in k), None)
        if avg_sec_key:
            try:
                ok = abs(float(metrics[avg_sec_key]) - 4.0) <= 0.2
                record("Average_Sections == 4.0 (±0.2)", ok, f"Got {metrics[avg_sec_key]}")
            except (TypeError, ValueError):
                record("Average_Sections is numeric", False, f"Got {metrics[avg_sec_key]}")
        else:
            record("Average_Sections metric exists", False, f"Keys: {list(metrics.keys())}")

        # Year_Range check — tighter: must mention 2024
        yr_key = next((k for k in metrics if "year" in k and "range" in k), None)
        if yr_key:
            yr_val = str(metrics[yr_key]).strip() if metrics[yr_key] else ""
            ok = "2024" in yr_val
            record("Year_Range contains 2024", ok, f"Got '{yr_val}'")

        # Most_Common_Category check
        mc_key = next((k for k in metrics if "common" in k or "most" in k), None)
        if mc_key:
            mc_val = str(metrics[mc_key]).strip() if metrics[mc_key] else ""
            record("Most_Common_Category non-empty", len(mc_val) > 0, f"Got {mc_val}")

    return True


def check_pptx(workspace):
    print("\n=== Checking PowerPoint ===")
    path = os.path.join(workspace, "NLP_Survey.pptx")
    if not os.path.isfile(path):
        record("PPTX exists", False, f"Not found: {path}")
        return False
    record("PPTX exists", True)

    try:
        from pptx import Presentation
        prs = Presentation(str(path))
        slides = list(prs.slides)

        record("Has >= 6 slides", len(slides) >= 6, f"Found {len(slides)}")

        all_text = []
        for slide in slides:
            texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        texts.append(p.text)
            all_text.append("\n".join(texts))

        full = "\n".join(all_text).lower()

        # Title slide
        first = all_text[0].lower() if all_text else ""
        has_title = any(kw in first for kw in ["survey", "nlp", "language model", "reasoning", "advances"])
        record("Title slide has survey keywords", has_title, f"First slide: {first[:200]}")

        # Content checks
        paper_methods = ["chain-of-thought", "instruction tuning", "retrieval", "self-consistency", "multimodal"]
        has_papers = any(kw in full for kw in paper_methods)
        record("Mentions paper methods", has_papers)
        # Stronger: require at least 3 distinct methods
        matched = sum(1 for kw in paper_methods if kw in full)
        record("PPT mentions >=3 distinct paper methods", matched >= 3, f"Matched {matched} methods")

        # Summary/conclusion slide
        last = all_text[-1] if all_text else ""
        record("Last slide has content", len(last.strip()) > 10, f"Last slide: {last[:100]}")

        return True
    except Exception as e:
        record("PPTX readable", False, str(e))
        return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace)
    check_pptx(args.agent_workspace)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}")
    sys.exit(0 if FAIL_COUNT == 0 else 1)


if __name__ == "__main__":
    main()
