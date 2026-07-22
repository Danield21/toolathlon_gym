#!/usr/bin/env python3
"""Evaluation for supplier-contract-compliance-audit.

Verifies the agent produced concrete supplier compliance audit deliverables:
  - XLSX compliance assessment (per-supplier scores) and corrective actions
  - DOCX audit findings and summary
  - PDF compliance report (or referenced compliance_report file)
  - Email audit reports to supplier contacts
  - Calendar meetings with under-performing suppliers
"""
import argparse
import json
import os
import sys

import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


# Suppliers from initial_workspace/supplier_list.csv (kept stable for audit topic)
SUPPLIERS = [
    "TechServices Inc",
    "Global Logistics",
    "Quality Manufacturing",
    "Consulting Group Ltd",
    "Support Solutions",
]


def _find_xlsx_by_keywords(workspace, keywords, exclude=None,
                            anti_keywords=None):
    """Score xlsx files by keyword matches; filename match weighs 10x more than content.

    anti_keywords: filename tokens that, when matched, subtract 5 from score
    so a file whose name overwhelmingly belongs to a *different* target
    can't accidentally win. Used to disambiguate compliance_assessment vs
    corrective_actions when both contain 'compliance' content keywords.
    """
    import glob
    import openpyxl
    exclude = exclude or set()
    anti_keywords = anti_keywords or []
    scored = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.xlsx"))):
        if path in exclude:
            continue
        fname_low = os.path.basename(path).lower()
        if fname_low.startswith("~$") or fname_low.startswith("."):
            continue
        score = 0
        for kw in keywords:
            if kw in fname_low:
                score += 10
        for ak in anti_keywords:
            if ak in fname_low:
                score -= 5
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            content_text = " ".join(s.lower() for s in wb.sheetnames)
            for ws in wb.worksheets:
                row_count = 0
                for row in ws.iter_rows(values_only=True):
                    if row_count > 5:
                        break
                    row_count += 1
                    for cell in row:
                        if cell is not None:
                            content_text += " " + str(cell).lower()
            wb.close()
            for kw in keywords:
                if kw in content_text:
                    score += 1
        except Exception:
            pass
        if score > 0:
            scored.append((score, path))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [p for _, p in scored]


def check_xlsx_content(workspace, gt_workspace="."):
    """Locate compliance_assessment + corrective_actions xlsx by filename or content."""
    print("\n=== Check: XLSX files ===")
    import openpyxl

    targets = [
        ("compliance_assessment.xlsx", "compliance assessment xlsx",
         ["compliance_assessment", "compliance assessment", "compliance",
          "assessment", "scorecard"],
         ["corrective", "remediation", "action_plan"]),
        ("corrective_actions.xlsx", "corrective actions xlsx",
         ["corrective_actions", "corrective actions", "corrective",
          "action", "remediation"],
         ["assessment", "scorecard"]),
    ]

    used_paths = set()
    for gt_fname, label, keywords, anti in targets:
        exact = os.path.join(workspace, gt_fname)
        chosen = None
        if os.path.isfile(exact) and exact not in used_paths:
            chosen = exact
        else:
            cands = _find_xlsx_by_keywords(workspace, keywords,
                                           exclude=used_paths,
                                           anti_keywords=anti)
            if cands: chosen = cands[0]
        if chosen is None:
            record(f"xlsx for {label} exists", False,
                   f"No xlsx with keywords {keywords[:3]} found")
            continue
        used_paths.add(chosen)
        record(f"xlsx for {label} exists ({os.path.basename(chosen)})", True)
        fname = os.path.basename(chosen)

        try:
            wb = openpyxl.load_workbook(chosen, data_only=True)
        except Exception as e:
            record(f"xlsx {fname} readable", False, str(e))
            continue
        # All cells flattened to lower-case for content scan
        all_text = ""
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is not None:
                        all_text += str(cell) + "\n"
        all_text_lower = all_text.lower()
        if "compliance" in label and "assessment" in label:
            record(f"{fname} mentions every supplier",
                   all(s.lower() in all_text_lower for s in SUPPLIERS),
                   f"Missing: {[s for s in SUPPLIERS if s.lower() not in all_text_lower]}")
        else:
            mentioned = sum(1 for s in SUPPLIERS if s.lower() in all_text_lower)
            record(f"{fname} mentions at least 4 of 5 suppliers",
                   mentioned >= 4,
                   f"Mentioned {mentioned} suppliers")

        gt_path = os.path.join(gt_workspace, gt_fname)
        if os.path.isfile(gt_path):
            gt_wb = openpyxl.load_workbook(gt_path, data_only=True)
            for gt_sname in gt_wb.sheetnames:
                gt_ws = gt_wb[gt_sname]
                a_ws = None
                for asn in wb.sheetnames:
                    if asn.strip().lower() == gt_sname.strip().lower():
                        a_ws = wb[asn]
                        break
                if a_ws is None and len(wb.sheetnames) == 1 and len(gt_wb.sheetnames) == 1:
                    a_ws = wb[wb.sheetnames[0]]
                if a_ws is None:
                    record(f"GT sheet '{gt_sname}' present in {fname}", False, f"Available: {wb.sheetnames}")
                    continue
                record(f"GT sheet '{gt_sname}' present in {fname}", True)

                a_rows = list(a_ws.iter_rows(values_only=True))
                gt_rows = list(gt_ws.iter_rows(values_only=True))

                def find_header_row(rows):
                    for i, r in enumerate(rows):
                        non_none = [v for v in r if v is not None]
                        if len(non_none) >= 4 and all(isinstance(v, str) for v in non_none):
                            return i
                    return 0

                a_hdr_idx = find_header_row(a_rows)
                gt_hdr_idx = find_header_row(gt_rows)
                a_hdrs = [str(v).strip().lower() if v else "" for v in a_rows[a_hdr_idx]]
                gt_hdrs = [str(v).strip().lower() if v else "" for v in gt_rows[gt_hdr_idx]]
                missing_hdrs = [h for h in gt_hdrs if h and h not in a_hdrs]
                record(f"{fname}/{gt_sname}: required columns present",
                       len(missing_hdrs) == 0, f"missing {missing_hdrs}")

                a_data_count = sum(1 for r in a_rows[a_hdr_idx + 1:] if any(v is not None for v in r))
                record(f"{fname}/{gt_sname}: at least 5 data rows", a_data_count >= 5,
                       f"got {a_data_count}")
            gt_wb.close()
        wb.close()


def _find_docx_by_keywords(workspace, keywords, exclude=None,
                            anti_keywords=None):
    import glob
    from docx import Document
    exclude = exclude or set()
    anti_keywords = anti_keywords or []
    scored = []
    for path in sorted(glob.glob(os.path.join(workspace, "*.docx"))):
        if path in exclude:
            continue
        fname_low = os.path.basename(path).lower()
        if fname_low.startswith("~$") or fname_low.startswith("."):
            continue
        score = 0
        for kw in keywords:
            if kw in fname_low:
                score += 10
        for ak in anti_keywords:
            if ak in fname_low:
                score -= 5
        try:
            doc = Document(path)
            text_low = "\n".join(p.text for p in doc.paragraphs).lower()
            for kw in keywords:
                if kw in text_low:
                    score += 1
        except Exception:
            pass
        if score > 0:
            scored.append((score, path))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [p for _, p in scored]


def check_docx_content(workspace):
    """Locate audit findings + audit summary docx by filename or content keyword."""
    print("\n=== Check: DOCX files ===")
    from docx import Document

    targets = [
        ("audit_findings.docx", "audit findings docx",
         ["audit_findings", "audit findings", "findings", "detailed", "narrative"],
         ["summary", "overview"]),
        ("audit_summary.docx", "audit summary docx",
         ["audit_summary", "audit summary", "summary", "executive", "overview"],
         ["findings_detail", "narrative"]),
    ]

    used_paths = set()
    for gt_fname, label, keywords, anti in targets:
        exact = os.path.join(workspace, gt_fname)
        chosen = None
        if os.path.isfile(exact) and exact not in used_paths:
            chosen = exact
        else:
            cands = _find_docx_by_keywords(workspace, keywords,
                                            exclude=used_paths,
                                            anti_keywords=anti)
            if cands: chosen = cands[0]
        if chosen is None:
            record(f"docx for {label} exists", False,
                   f"No docx with keywords {keywords[:3]} found")
            continue
        used_paths.add(chosen)
        fname = os.path.basename(chosen)
        record(f"docx for {label} exists ({fname})", True)
        try:
            doc = Document(chosen)
        except Exception as e:
            record(f"docx {fname} readable", False, str(e))
            continue
        text = "\n".join(p.text for p in doc.paragraphs)
        text_lower = text.lower()
        for s in SUPPLIERS:
            record(f"{fname} mentions supplier '{s}'",
                   s.lower() in text_lower, "missing")
        record(f"{fname} has >= 800 chars", len(text) >= 800, f"got {len(text)}")
        for kw in ("compliance", "finding", "recommendation"):
            record(f"{fname} mentions '{kw}'",
                   kw in text_lower, f"missing in {fname}")


def check_pdf(workspace):
    print("\n=== Check: PDF compliance report ===")
    pdf_candidates = []
    for f in os.listdir(workspace):
        if f.lower().endswith(".pdf") and "compliance" in f.lower():
            pdf_candidates.append(f)
    record("Compliance PDF report present", len(pdf_candidates) >= 1,
           f"found pdfs: {pdf_candidates}")


def _read_supplier_contacts(workspace):
    """Parse supplier_list.csv to derive expected supplier email domains and
    full email addresses dynamically. Returns ([domains], [emails])."""
    import csv
    sup_csv = os.path.join(workspace, "supplier_list.csv")
    domains, emails = [], []
    if not os.path.isfile(sup_csv):
        return domains, emails
    try:
        with open(sup_csv, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                addr = (row.get("Contact_Email") or "").strip().lower()
                if not addr or "@" not in addr:
                    continue
                emails.append(addr)
                domains.append(addr.split("@", 1)[1])
    except Exception:
        pass
    return domains, emails


def check_email(agent_workspace):
    print("\n=== Check: Audit emails to suppliers ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        record("DB connect", False, str(e))
        return
    cur.execute("SELECT id FROM email.folders WHERE LOWER(name) IN ('sent','drafts')")
    folder_ids = [r[0] for r in cur.fetchall()]
    if not folder_ids:
        record("Email folders for sent/drafts exist", False, "")
        cur.close()
        conn.close()
        return
    placeholders = ",".join(["%s"] * len(folder_ids))
    cur.execute(f"SELECT subject, to_addr, body_text FROM email.messages WHERE folder_id IN ({placeholders})", folder_ids)
    sent = cur.fetchall()
    cur.close()
    conn.close()

    record("At least 1 audit-related email sent",
           any(("audit" in (s or "").lower()) or ("compliance" in (s or "").lower()) for s, _, _ in sent),
           f"found {len(sent)} sent emails")
    # Derive expected supplier domains/emails from supplier_list.csv.
    # Falls back to the supplierN.com convention (which the bundled CSV uses).
    derived_domains, derived_emails = _read_supplier_contacts(agent_workspace)
    if not derived_domains:
        derived_domains = ["supplier1.com", "supplier2.com", "supplier3.com",
                           "supplier4.com", "supplier5.com"]
    matched_domains = set()
    for _, to, _ in sent:
        to_str = json.dumps(to) if not isinstance(to, str) else to
        to_low = to_str.lower()
        # Match either the full email address (preferred) or domain
        for em in derived_emails:
            if em in to_low:
                matched_domains.add(em.split("@", 1)[1])
        for dom in derived_domains:
            if dom in to_low:
                matched_domains.add(dom)
    record("Emails reach at least 2 supplier contacts",
           len(matched_domains) >= 2,
           f"found supplier domains: {matched_domains}")


def check_gcal():
    print("\n=== Check: Calendar meetings with suppliers ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        record("DB connect", False, str(e))
        return
    cur.execute("SELECT id, summary, description FROM gcal.events")
    events = cur.fetchall()
    cur.close()
    conn.close()
    record("At least 1 calendar event created", len(events) >= 1, f"found {len(events)}")
    matched = 0
    for eid, summary, desc in events:
        text = f"{summary or ''} {desc or ''}".lower()
        if any(s.lower() in text for s in SUPPLIERS) or "supplier" in text or "vendor" in text:
            matched += 1
    record("At least 1 calendar event references a supplier/vendor",
           matched >= 1, f"matched {matched}/{len(events)}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--res_log_file", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    ws = args.agent_workspace
    if not os.path.isdir(ws):
        print(f"Agent workspace not found: {ws}")
        sys.exit(1)

    # GT self-test toleration: when agent==groundtruth (V1 GT-only smoke
    # test), the email/gcal checks would fail because no email/event was
    # actually sent.
    is_gt_self_test = (
        args.agent_workspace and args.groundtruth_workspace
        and os.path.abspath(args.agent_workspace) == os.path.abspath(args.groundtruth_workspace)
    )

    check_xlsx_content(ws, args.groundtruth_workspace)
    check_docx_content(ws)
    check_pdf(ws)
    if not is_gt_self_test:
        check_email(ws)
        check_gcal()
    else:
        print("  [SKIP] email and gcal checks (GT self-test mode)")

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    all_ok = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if all_ok else 'FAIL'}")

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "success": all_ok}, f)

    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
