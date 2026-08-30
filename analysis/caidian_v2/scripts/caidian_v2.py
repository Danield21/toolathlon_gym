#!/usr/bin/env python3
"""Caidiăn v2: evidence-preserving oracle prompt/trajectory alignment.

The legacy scorer is intentionally left untouched.  This module fixes the
measurement contract at the extraction layer and exposes a unified IR used by
both the coarse (wave) and fine (node) metrics.
"""

from __future__ import annotations

import csv
import hashlib
import html
import json
import re
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence


VERSION = "2.2.5"
MIN_RAW_STREAM_BYTES = 1000

ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")
EFFICIENT_RE = re.compile(r"To solve this task efficiently.*", re.I | re.S)
NO_DISPATCH_RE = re.compile(
    r"(?:do\s+not|must\s+not|never)\s+dispatch\s+(?:any\s+)?sub-?agents?",
    re.I,
)
WAVE_LINE_RE = re.compile(
    r"(?mi)^[ \t]*(?:\d+\.\s+)?\**Wave\s+(\d+)\s*[\u2014\u2013-].*$"
)
AFTER_WAVE_RE = re.compile(
    r"(?mi)(?:^|(?<=[.!?])\s+)(?:After\s+(?:the\s+)?(?:[a-z-]+\s+){0,3}Wave(?:s|\(s\))?(?:\s+\d+)?|"
    r"Once\s+(?:Wave|all)|"
    r"When\s+all|Following\s+(?:the\s+)?Wave|The\s+main\s+agent\s+then|"
    r"While\s+(?:the\s+)?Wave(?:\s+\d+)?\s+runs|"
    r"In\s+the\s+same\s+(?:orchestration\s+)?response\b[^\n]{0,160}\bmain\s+agent|"
    r"Concurrently,?\s+the\s+main\s+agent|"
    r"Meanwhile,?\s+the\s+main\s+agent|"
    r"Accept\s+only\s+after|"
    r"In\s+the\s+same\s+response\s+that\s+launches\s+Wave)\b"
)
EXPLICIT_WAVE_DEPENDENCY_RE = re.compile(
    r"(?is)\b(?:after|once|following|depends?\s+on|using\s+(?:the\s+)?(?:results?|outputs?|handoffs?)\s+from)"
    r"\b[^\n.!?]{0,160}\bwave\s+(\d+)\b"
)
PARALLEL_WAVE_CONTEXT_RE = re.compile(
    r"(?is)\b(?:while\s+(?:the\s+)?wave(?:\s+\d+)?\s+runs|"
    r"in\s+the\s+same\s+(?:orchestration\s+)?response|concurrently|meanwhile)\b"
)
SWARM_TRUNCATED_PREVIEW_RE = re.compile(
    r"tool output exceeded\s+\d+\s+characters;\s*showing a preview only",
    re.I,
)
OLD_DISPATCH_RE = re.compile(
    r"(?mi)^(\d+)\.\s+Please dispatch (multiple|a single)\s+([a-zA-Z-]+)\s+sub-agents?\b[^:]*:"
)
NUMBERED_ITEM_RE = re.compile(
    r"(?m)^\s*(?:\d+(?:\s*[-–]\s*\d+)?\.|\(\d+\)\.)\s+"
)
MCP_TOOL_RE = re.compile(r"mcp__([a-z0-9_-]+)__([a-z0-9_-]+)", re.I)
TOOL_CALL_NAME_RE = re.compile(r"(?:^|__)(AgentSwarm|Agent)$", re.I)

WORD_NUMBERS = {
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
    "twenty": 20,
    "twenty-one": 21,
    "twenty-two": 22,
    "thirty-one": 31,
}

DOMAIN_PATTERNS: list[tuple[str, tuple[str, ...]]] = [
    ("canvas", ("canvas", "course_id", "course code", "enrollment", "syllabus")),
    ("rail_12306", ("12306", "mcp__rail", "train_no", "telecode", "station code")),
    ("woocommerce", ("woocommerce", "mcp__woo", "woo_", "product-list")),
    ("yahoo-finance", ("yahoo", "ticker", "stock", "yf_", "historical price")),
    ("snowflake", ("snowflake", "mcp__snowflake", "data warehouse", "run_query", ".public.")),
    ("scholarly-arxiv", ("arxiv", "scholarly", "latex", "paper archive")),
    ("external-fetch", ("fetch", "http://", "https://", "localhost:", "endpoint")),
    ("notion", ("notion", "knowledge base", "wiki")),
    ("youtube", ("youtube", "video", "transcript", "playlist")),
    ("playwright-web", ("playwright", "browser", "web page", "snapshot")),
    ("howtocook", ("howtocook", "dish", "menu", "catering")),
    ("emails", ("mcp__emails", "inbox", "sent mail", "recipient", "e-mail")),
    ("calendar", ("calendar", "gcal", "event")),
    ("google-sheet", ("google sheet", "google spreadsheet", "gsheet", "cloud spreadsheet")),
    ("google-form", ("google form", "gform", "survey form")),
    ("salesforce", ("salesforce", "sf_", "support ticket", "sales order")),
    ("filesystem-workspace", ("workspace", "local file", "filesystem")),
    ("terminal", ("python_execute", "terminal", "run_command", "script")),
    ("memory", ("memory mcp", "jsonl memory", "memory store")),
]

WRITE_ROLE_PATTERNS: list[tuple[str, tuple[str, ...]]] = [
    ("gsheet", ("google sheet", "google spreadsheet", "gsheet", "cloud spreadsheet", "shared spreadsheet")),
    ("excel", (".xlsx", "xlsx", "excel", "workbook", "worksheet")),
    ("word", (".docx", "docx", "word document")),
    ("ppt", (".pptx", "pptx", "powerpoint", "slide deck", "presentation")),
    ("pdf", (".pdf", "pdf report", "pdf export")),
    ("notion", ("notion", "knowledge base", "wiki", "database")),
    ("email", ("email", "e-mail", "mail message")),
    ("calendar", ("calendar", "gcal", "event", "meeting")),
    ("gform", ("google form", "gform", "survey form")),
    ("json", (".json", "json packet", "json payload", "json artifact", "json file", "normalized json", "json handoff", "evidence json")),
    ("csv", (".csv", "csv file")),
    ("text", (".txt", "plain text")),
    ("tex", (".tex", "latex source")),
]

READ_RE = re.compile(
    r"\b(read|retrieve|fetch|search|query|inspect|list|enumerate|download|collect|"
    r"snapshot|browse|paginate|open|load|inventory|probe|exhaust)\w*\b",
    re.I,
)
COMPUTE_RE = re.compile(
    r"\b(compute|calculate|aggregate|analy[sz]e|score|reconcile|merge|join|rank|"
    r"compare|synthesi[sz]e|classif|transform|extract|validate|verify|derive|freeze)\w*\b",
    re.I,
)
WRITE_RE = re.compile(
    r"\b(create|write|save|send|schedule|publish|upload|insert|update|upsert|"
    r"build|generate|produce|persist|preserve|add|modify)\w*\b",
    re.I,
)
OWNER_WRITE_RE = re.compile(r"\b(?:sole\s+)?(?:writer|owner|own(?:s|ed|ing)?)\b", re.I)
READ_ONLY_RE = re.compile(
    r"read[- ]only|do not (?:write|create|modify|send|schedule|persist)|writes? nothing",
    re.I,
)
SCRATCH_RE = re.compile(r"runtime[- ]unique|scratch|intermediate|handoff|canonical (?:packet|json)", re.I)

FILE_TARGET_RE = re.compile(
    r"(?<![\w/])(?:[A-Za-z0-9_.-]+/)*[A-Za-z0-9_.-]+\.(?:xlsx|docx|pptx|pdf|json|csv|txt|tex|md|py)\b",
    re.I,
)
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
QUOTED_TARGET_RE = re.compile(
    r"(?i)\b(?:called|named|titled|subject|summary)\s*(?:=|:)?\s*[`\"']([^`\"'\n]{3,100})[`\"']"
)
ARXIV_ID_RE = re.compile(r"\b\d{4}\.\d{4,5}\b")
COURSE_ID_RE = re.compile(r"(?i)\bcourse(?:_id| id)?\s*(?:=|:)?\s*(\d+)\b")
COURSE_CODE_RE = re.compile(r"\b[A-Z]{3}-\d{4}[A-Z]\b")
ROW_RANGE_RE = re.compile(r"(?i)\b(?:rows?|pages?)\s+(\d+(?:\s*[-–]\s*\d+)?)")
POSITION_RE = re.compile(r"(?i)\bposition\s+(\d+)\b")
NEGATED_WRITE_RE = re.compile(
    r"(?i)(?:do not|must not|never|without)\s+(?:\w+\s+){0,3}(?:write|create|send|schedule|modify|persist)"
)
NEGATIVE_CLAUSE_RE = re.compile(
    r"(?i)(?:do\s+not|don't|must\s+not|never|without|no\s+sub-?agent\s+may|"
    r"make\s+no|no\s+(?:workspace\s+|file\s+)?writes?|writes?\s+nothing)\b[^;\n]*"
)
PERSISTENT_TARGET_RE = re.compile(
    r"(?i)\.(?:xlsx|docx|pptx|pdf|json|csv|txt|tex|md|py)\b|google\s+sheet|"
    r"\b(?:notion|knowledge base|calendar|event|e-?mail|workbook|document|presentation|"
    r"database|spreadsheet|sheet|survey|form|page|slides?|json|payload|packet|artifact|file|script|message)\b"
)

REJECT_PATTERNS = re.compile(
    r"(?i)(tool[_ ]?error|invalid (?:argument|request|combination)|rejected|"
    r"cannot (?:use|specify|combine)|must not include|failed to parse|schema validation|"
    r"unknown subagent|missing required (?:field|argument))"
)
REPAIR_RE = re.compile(
    r"^\s*(?:(?:resume|continue|retry|repair|recover|re-run|rerun)\b|"
    r"(?:finish|complete)\s+(?:the\s+)?(?:prior|previous|interrupted)\b)|"
    r"\b(?:previous (?:run|attempt|work)|stopped before|interrupted run)\b",
    re.I | re.S,
)
KNOWN_RAW_DIAGNOSTIC_RE = re.compile(
    r'^unknown format "json" ignored in schema at path "#/properties/(?:icon|cover)"$'
)
BLOCKER_RE = re.compile(
    r"(?i)\b(missing|not found|absent|unreadable|cannot proceed|blocked|blocker|"
    r"stop before delegation|did not exist|does not exist)\b"
)


class TrajectoryIntegrityError(ValueError):
    """The selected trajectory cannot be scored without ignoring corrupt evidence."""


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_text(text: str) -> str:
    return sha256_bytes(text.encode("utf-8"))


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def stable_unique(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        value = str(value).strip()
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


def lower_text(text: str) -> str:
    return ANSI_RE.sub("", text or "").lower()


def semantic_positive_text(text: str) -> str:
    """Remove prohibition clauses so forbidden tools/outputs do not become labels."""
    clean = ANSI_RE.sub("", text or "")
    clean = NEGATIVE_CLAUSE_RE.sub(" ", clean)
    return clean


def infer_agent_type(text: str, declared: str | None = None) -> str:
    low = lower_text(text)
    # A numbered item may override a wave-level default in a mixed-role wave.
    if re.search(r"\bexplor(?:e|er)\s+(?:sub-?agent|agent|owner|reader)\b", low):
        return "explore"
    if re.search(r"\bcoder(?:\s+sub-?agent)?\b", low):
        return "coder"
    if re.search(r"\bplanner\s+(?:sub-?agent|agent)\b|\bplan\s+sub-?agent\b", low):
        return "plan"
    if declared:
        declared = declared.lower()
        if declared in {"explore", "coder", "plan"}:
            return declared
    if re.search(r"\bexplor(?:e|er)\b", low):
        return "explore"
    if re.search(r"\b(?:coder|writer|owner|collector|worker)\b", low):
        return "coder"
    if re.search(r"\bplanner\s+(?:sub-?agent|agent)\b|\bplan\s+sub-?agent\b", low):
        return "plan"
    return "unspecified"


def persistent_write_clauses(text: str) -> list[str]:
    """Return only local clauses that contain both a write verb and a target."""
    positive = lower_text(semantic_positive_text(text))
    clauses = re.split(r"\n+|;|(?<=[.!?])\s+(?=[a-z0-9`])", positive, flags=re.I)
    out: list[str] = []
    for clause in clauses:
        clause = NEGATED_WRITE_RE.sub("", clause)
        write_matches = [match for match in WRITE_RE.finditer(clause) if match.group(0).lower() != "published"]
        owner = OWNER_WRITE_RE.search(clause)
        owner_target = False
        if owner:
            owner_fragment = clause[owner.start():]
            boundary = re.search(
                r"\b(?:reads?|retrieves?|fetches?|queries|searches|computes?|analy[sz]es|issues?|calls?|returns?)\b",
                owner_fragment,
            )
            owner_fragment = owner_fragment[: boundary.start()] if boundary else owner_fragment
            owner_target = bool(PERSISTENT_TARGET_RE.search(owner_fragment))
        if (write_matches and PERSISTENT_TARGET_RE.search(clause)) or owner_target:
            out.append(clause)
    return out


def infer_actions(text: str, agent_type: str = "unspecified") -> list[str]:
    low = lower_text(semantic_positive_text(text))
    actions: list[str] = []
    if READ_RE.search(low) or READ_ONLY_RE.search(low) or agent_type == "explore":
        actions.append("read")
    if COMPUTE_RE.search(low):
        actions.append("compute")
    if persistent_write_clauses(text):
        actions.append("write")
    if agent_type == "coder" and not actions:
        actions.append("compute")
    if not actions:
        actions.append("read")
    return stable_unique(actions)


def infer_domains(text: str) -> list[str]:
    positive = semantic_positive_text(text)
    low = lower_text(positive)
    values: list[str] = []
    for domain, patterns in DOMAIN_PATTERNS:
        if any(pattern in low for pattern in patterns):
            values.append(domain)
    for server, _tool in MCP_TOOL_RE.findall(low):
        normalized = server.replace("-", "_")
        mapping = {
            "google_calendar": "calendar",
            "google_sheet": "google-sheet",
            "google_forms": "google-form",
            "arxiv_local": "scholarly-arxiv",
            "arxiv_latex": "scholarly-arxiv",
            "local": "filesystem-workspace",
        }
        values.append(mapping.get(normalized, normalized.replace("_", "-")))
    if FILE_TARGET_RE.search(positive):
        values.append("filesystem-workspace")
    return stable_unique(values)


def infer_write_roles(text: str, actions: Sequence[str]) -> list[str]:
    if "write" not in actions:
        return []
    owner_fragments: list[str] = []
    verb_fragments: list[str] = []
    for clause in persistent_write_clauses(text):
        owner = OWNER_WRITE_RE.search(clause)
        if owner:
            fragment = clause[owner.start():]
            boundary = re.search(
                r"\b(?:reads?|retrieves?|fetches?|queries|searches|computes?|analy[sz]es|issues?|calls?|returns?)\b",
                fragment,
            )
            owner_fragments.append(fragment[: boundary.start()] if boundary else fragment)
        for match in WRITE_RE.finditer(clause):
            if match.group(0).lower() == "published":
                continue
            verb_fragments.append(clause[match.start(): match.start() + 200])
    owner_low = "\n".join(owner_fragments)
    owner_roles = [role for role, pats in WRITE_ROLE_PATTERNS if any(pat in owner_low for pat in pats)]
    if owner_roles:
        return stable_unique(owner_roles)
    low = "\n".join(verb_fragments)
    roles = [role for role, pats in WRITE_ROLE_PATTERNS if any(pat in low for pat in pats)]
    return stable_unique(roles or ["other-write"])


def infer_role_labels(text: str) -> list[str]:
    low = lower_text(semantic_positive_text(text))
    roles = [role for role, patterns in WRITE_ROLE_PATTERNS if any(pattern in low for pattern in patterns)]
    if "email" in roles and not re.search(
        r"\b(?:send|write|draft)\b[^\n]{0,50}\b(?:email|message)\b|\bemail\s+(?:writer|owner)\b",
        low,
    ):
        roles.remove("email")
    return stable_unique(roles)


def infer_targets(text: str) -> list[str]:
    text = semantic_positive_text(text)
    values: list[str] = []
    for match in FILE_TARGET_RE.findall(text or ""):
        values.append(Path(match).name)
    values.extend(EMAIL_RE.findall(text or ""))
    values.extend(m.group(1).strip() for m in QUOTED_TARGET_RE.finditer(text or ""))
    return stable_unique(values)


def infer_object_ids(text: str) -> list[str]:
    text = semantic_positive_text(text)
    values: list[str] = []
    values.extend(f"paper:{value}" for value in ARXIV_ID_RE.findall(text or ""))
    values.extend(f"course:{value}" for value in COURSE_ID_RE.findall(text or ""))
    values.extend(f"course_code:{value}" for value in COURSE_CODE_RE.findall(text or ""))
    values.extend(f"rows:{value.replace(' ', '')}" for value in ROW_RANGE_RE.findall(text or ""))
    values.extend(f"position:{value}" for value in POSITION_RE.findall(text or ""))
    return stable_unique(values)


def infer_shard_key(
    text: str,
    object_ids: Sequence[str],
    *,
    item_context: bool = False,
) -> str | None:
    """Infer the unit assigned to one agent, not every entity in its payload.

    Pagination parameters and rows/papers mentioned inside a writer's output
    specification are content, not agent partitions.  Typed AgentSwarm items
    are the one place where a concrete object ID is itself assignment evidence.
    """
    low = lower_text(text)
    if item_context:
        if re.search(r"\bteacher\b", low):
            return "teacher"
        if re.search(r"\b(?:page_start|page_end)\b|\bpage\s*[:=]\s*\d+", low):
            return "page"
        if re.search(r"\b(?:paper_id|arxiv_id)\b|\b\d{4}\.\d{4,5}\b", low):
            return "paper"
        if re.search(r"\b(?:course_id|course_code)\b", low):
            return "course"
    if re.search(r"\bteacher(?: enrollment| lookup| work item| owner| agent)s?\b", low):
        return "teacher"
    if re.search(
        r"\b(?:grade|submission|score)?\s*page(?: work item| owner| agent)s?\b|"
        r"\{[^}\n]*\bpage\b[^}\n]*\}",
        low,
    ):
        return "page"
    if re.search(
        r"\b(?:assigned|literal)\s+paper(?:_id| id)?\b|"
        r"\bpaper(?:_id| id)?\s+at\s+position\b|"
        r"\b(?:paper|full-text)\s+(?:reader|owner|agent)\b",
        low,
    ):
        return "paper"
    if re.search(
        r"\b(?:course|per-course)\s+(?:collector|owner|agent)\b|"
        r"\b(?:collector|owner|agent)\s+for\s+course(?:_id| id)?\b|"
        r"\bprocessing\s+course(?:_id| id)?\b|"
        r"\bone\s+(?:canvas\s+)?(?:lms\s+)?course\b|"
        r"\bone\s+per\s+(?:manifest\s+|verified\s+)?course\b",
        low,
    ):
        return "course"
    if item_context and object_ids:
        if any(item.startswith("rows:") for item in object_ids):
            return "page"
        if any(item.startswith("course:") or item.startswith("course_code:") for item in object_ids):
            return "course"
        if any(item.startswith("paper:") for item in object_ids):
            return "paper"
        if any(item.startswith("position:") for item in object_ids):
            return "position"
    for key, pattern in (
        ("course", r"one (?:agent|owner|collector).*?(?:per|for each).*?course"),
        ("paper", r"one (?:agent|owner|reader).*?(?:per|for each).*?paper"),
        ("page", r"one (?:agent|owner).*?(?:per|for each).*?page"),
        ("supplier", r"one (?:agent|owner).*?(?:per|for each).*?supplier"),
        ("destination", r"one (?:agent|owner).*?(?:per|for each).*?destination"),
    ):
        if re.search(pattern, low):
            return key
    return None


def make_node_ir(
    text: str,
    *,
    node_id: str,
    declared_type: str | None = None,
    multiplicity: int = 1,
    item: Any = None,
    signature_text: str | None = None,
) -> dict[str, Any]:
    signature = signature_text.strip() if signature_text and signature_text.strip() else text
    agent_type = infer_agent_type(signature, declared_type)
    actions = infer_actions(text, agent_type)
    object_ids = infer_object_ids(text)
    item_text = ""
    item_object_ids: list[str] = []
    if item is not None:
        item_text = json.dumps(item, ensure_ascii=False) if not isinstance(item, str) else item
        item_object_ids = infer_object_ids(item_text)
        object_ids = stable_unique(object_ids + item_object_ids + [f"item:{item_text[:160]}"])
    full_roles = infer_write_roles(text, actions)
    signature_roles = infer_role_labels(signature) if signature_text and "write" in actions else []
    scoped_roles = [role for role in signature_roles if role != "other-write"]
    return {
        "node_id": node_id,
        "agent_type": agent_type,
        "actions": actions,
        "domains": infer_domains(text),
        "write_roles": stable_unique(scoped_roles or full_roles),
        "targets": infer_targets(text),
        "object_ids": object_ids,
        "shard_key": (
            infer_shard_key(signature, infer_object_ids(signature))
            or (
                infer_shard_key(item_text, item_object_ids, item_context=True)
                if item is not None
                else None
            )
            or infer_shard_key(text, object_ids)
        ),
        "multiplicity": max(1, int(multiplicity)),
        "text": text.strip(),
    }


def parse_count(header: str, node_count: int) -> dict[str, Any]:
    low = lower_text(header)
    exact: int | None = None
    source = "node-count"
    match = re.search(r"\b(?:dispatch(?:es)?|start)\s+(?:exactly\s+)?(\d+)\b", low)
    if match:
        exact = int(match.group(1))
        source = "header-numeric"
    if exact is None:
        match = re.search(r"\b(?:exactly\s+)?(\d+)\s+(?:[a-z-]+\s+){0,4}(?:sub-?agents?|coders?|owners?|writers?|collectors?|explorers?)\b", low)
        if match:
            exact = int(match.group(1))
            source = "header-numeric-role"
    each_of = re.search(r"\beach of the\s+(\d+)\b", low)
    if each_of:
        exact = int(each_of.group(1))
        source = "header-each-of"
    if exact is None:
        for word, number in sorted(WORD_NUMBERS.items(), key=lambda item: -len(item[0])):
            if re.search(rf"\b(?:exactly\s+)?{re.escape(word)}\s+(?:[a-z-]+\s+){{0,4}}(?:sub-?agents?|coders?|owners?|writers?|collectors?|explorers?)\b", low):
                exact = number
                source = "header-word"
                break
    per_item = bool(re.search(r"\b(?:one|owner|agent).*?\b(?:for each|for every|per verified|per frozen)\b", low))
    if re.search(r"\ball\s+concrete\b[^:]{0,100}\bwork items?\b", low):
        per_item = True
    if exact == 1 and per_item and each_of is None:
        exact = None
    dynamic = per_item and exact is None
    max_match = re.search(r"(?:never more than|up to|at most)\s+(\d+)", low)
    if exact is not None:
        minimum = maximum = exact
    elif dynamic:
        minimum, maximum = 1, int(max_match.group(1)) if max_match else None
        source = "header-dynamic"
    else:
        minimum = maximum = max(1, node_count)
    return {
        "min": minimum,
        "max": maximum,
        "exact": exact,
        "dynamic": dynamic,
        "source": source,
    }


def split_numbered_items(body: str) -> tuple[list[str], str]:
    matches = list(NUMBERED_ITEM_RE.finditer(body or ""))
    if not matches:
        return [], ""
    items: list[str] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        item = body[match.start():end].strip()
        if item:
            items.append(item)
    shared = ""
    if items:
        lines = items[-1].splitlines()
        cut = None
        for line_index, line in enumerate(lines[1:], start=1):
            if line.strip() and not line[0].isspace():
                cut = line_index
                break
        if cut is not None:
            shared = "\n".join(lines[cut:]).strip()
            items[-1] = "\n".join(lines[:cut]).strip()
    if shared:
        items = [(item + "\n\n" + shared).strip() for item in items]
    return items, shared


def numbered_item_multiplicity(item: str) -> int:
    match = re.match(r"^\s*(\d+)\s*[-–]\s*(\d+)\.", item or "")
    if not match:
        return 1
    start, end = int(match.group(1)), int(match.group(2))
    return max(1, end - start + 1)


def split_top_level_bullets(body: str) -> list[str]:
    matches = list(re.finditer(r"(?m)^(?P<indent>[ \t]*)[-*]\s+", body or ""))
    if not matches:
        return []
    minimum = min(len(match.group("indent").expandtabs(4)) for match in matches)
    top = [match for match in matches if len(match.group("indent").expandtabs(4)) == minimum]
    items: list[str] = []
    for index, match in enumerate(top):
        end = top[index + 1].start() if index + 1 < len(top) else len(body)
        value = body[match.end():end].strip()
        if value:
            items.append(value)
    return items


def condition_from_prefix(prefix: str) -> dict[str, Any]:
    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", prefix or "") if part.strip()]
    candidates = [
        part
        for part in paragraphs[-4:]
        if re.search(r"(?i)\b(if|only if|unless|provided that|stop before delegation|waves? appl(?:y|ies) only)\b", part)
        and re.search(r"(?i)\b(missing|absent|unreadable|valid|available|present|duplicate|identity|count|source)\b", part)
    ]
    if candidates:
        return {
            "kind": "conditional",
            "text": candidates[-1],
            "eligible": True,
            "eligibility_source": "default-conservative",
            "reason": "No verified override marks this gate as not reached.",
        }
    return {
        "kind": "always",
        "text": "",
        "eligible": True,
        "eligibility_source": "prompt",
        "reason": "Unconditional planned wave.",
    }


def dependency_from_context(context: str, previous_wave_id: str | None) -> str:
    """Keep only dependencies stated by the orchestration prompt."""
    if previous_wave_id is None:
        return "main-agent-prerequisites"
    explicit = EXPLICIT_WAVE_DEPENDENCY_RE.search(context or "")
    if explicit:
        return f"wave_{int(explicit.group(1))}"
    if PARALLEL_WAVE_CONTEXT_RE.search(context or ""):
        return "main-agent-prerequisites"
    if AFTER_WAVE_RE.search(context or "") or re.search(
        r"(?i)\b(?:after|once|when|following|depends?\s+on|using\s+(?:the\s+)?(?:results?|outputs?|handoffs?))\b",
        context or "",
    ):
        return previous_wave_id
    return "main-agent-prerequisites"


def dependency_context_for_wave(header: str, interstitial: str) -> str:
    """Use the current header plus only an explicit transition tail from the prior section."""
    transitions = list(
        re.finditer(
            r"(?im)^[ \t]*(?:after|once|when|following|the\s+main\s+agent\s+then|while|"
        r"in\s+the\s+same|concurrently|meanwhile|accept\s+only\s+after|"
        r"using\s+(?:the\s+)?(?:results?|outputs?|handoffs?)\s+from|depends?\s+on)\b",
            interstitial or "",
        )
    )
    if transitions:
        candidate = (interstitial or "")[transitions[-1].start():].strip()
        strong_transition = bool(
            EXPLICIT_WAVE_DEPENDENCY_RE.search(candidate)
            or re.search(r"(?i)\b(?:waves?(?:\(s\))?|results?|outputs?|handoffs?|receipts?)\b", candidate)
            or re.search(
                r"(?i)\b(?:all|every)\b[^\n.!?]{0,100}\b(?:agents?|sub-?agents?|owners?|collectors?)\b"
                r"[^\n.!?]{0,80}\b(?:return|complete|finish|report|deliver)\w*\b",
                candidate,
            )
            or re.match(r"(?i)^the\s+main\s+agent\s+then\b", candidate)
        )
        tail = candidate if strong_transition else ""
    else:
        tail = ""
    return header + ("\n" + tail if tail else "")


def _allocate_multiplicity(nodes: list[dict[str, Any]], fanout: dict[str, Any], header: str) -> None:
    expected = fanout.get("exact")
    current = sum(max(1, int(node.get("multiplicity", 1))) for node in nodes)
    if not nodes or not isinstance(expected, int) or expected <= current:
        return
    if len(nodes) == 1:
        nodes[0]["multiplicity"] = expected
        return
    residual = expected - current
    generic = make_node_ir(header, node_id="anonymous-template", multiplicity=residual)
    generic["parse_note"] = "Header fan-out exceeds explicit numbered ownership items."
    nodes.append(generic)


def _wave_rollup(wave: dict[str, Any]) -> None:
    nodes = wave.get("nodes") or []
    wave["actions"] = stable_unique(value for node in nodes for value in node["actions"])
    wave["domains"] = stable_unique(value for node in nodes for value in node["domains"])
    wave["write_roles"] = stable_unique(value for node in nodes for value in node["write_roles"])
    wave["targets"] = stable_unique(value for node in nodes for value in node["targets"])
    wave["object_ids"] = stable_unique(value for node in nodes for value in node["object_ids"])
    wave["agent_types"] = stable_unique(node["agent_type"] for node in nodes)
    wave["node_instances"] = sum(int(node.get("multiplicity", 1)) for node in nodes)


def parse_planned_ir(task: str, prompt: str) -> dict[str, Any]:
    match = EFFICIENT_RE.search(prompt or "")
    segment = match.group(0) if match else ""
    segment = re.sub(r"\n\s*\.\.\.\s*\(home=.*$", "", segment, flags=re.S)
    result: dict[str, Any] = {
        "task": task,
        "no_dispatch": bool(NO_DISPATCH_RE.search(segment)),
        "waves": [],
        "parse_warnings": [],
    }
    if not segment:
        result["parse_warnings"].append("No orchestration section found.")
        return result

    headers = list(WAVE_LINE_RE.finditer(segment))
    for index, header_match in enumerate(headers):
        wave_number = int(header_match.group(1))
        header = header_match.group(0).strip()
        start = header_match.end()
        end = headers[index + 1].start() if index + 1 < len(headers) else len(segment)
        body = segment[start:end]
        after = AFTER_WAVE_RE.search(body)
        if after:
            body = body[:after.start()]
        items, _shared = split_numbered_items(body)
        preliminary_fanout = parse_count(header, max(1, len(items)))
        shared_recipe = bool(
            re.search(
                r"(?i)\b(?:same|literal)\s+(?:literal\s+)?recipe\b|"
                r"\bwith\s+only\s+(?:its|the)\s+[^.;\n]{0,40}\s+substituted\b",
                header + "\n" + body,
            )
        )
        if (
            items
            and shared_recipe
            and isinstance(preliminary_fanout.get("exact"), int)
            and preliminary_fanout["exact"] != len(items)
        ):
            payload = (header.split("—", 1)[-1] + "\n" + body).strip()
            items = [payload]
        if not items and (isinstance(preliminary_fanout.get("exact"), int) and preliminary_fanout["exact"] > 1 or preliminary_fanout.get("dynamic")):
            bullets = split_top_level_bullets(body)
            if (
                preliminary_fanout.get("exact") == len(bullets)
                or re.search(r"(?i)all\s+concrete.*work items", header)
            ):
                items = bullets
        if not items:
            payload = (header.split("—", 1)[-1] + "\n" + body).strip()
            items = [payload] if payload else [header]
        header_type = infer_agent_type(header)
        nodes = [
            make_node_ir(
                item,
                node_id=f"wave_{wave_number}.node_{item_index + 1}",
                declared_type=header_type,
                multiplicity=numbered_item_multiplicity(item),
            )
            for item_index, item in enumerate(items)
        ]
        fanout = parse_count(header, len(nodes))
        _allocate_multiplicity(nodes, fanout, header)
        prefix_start = headers[index - 1].end() if index else 0
        prefix = segment[prefix_start:header_match.start()]
        dependency_context_start = headers[index - 1].end() if index else 0
        interstitial = segment[dependency_context_start:header_match.start()]
        wave = {
            "wave_id": f"wave_{wave_number}",
            "ordinal": len(result["waves"]) + 1,
            "dependency": dependency_from_context(
                dependency_context_for_wave(header, interstitial),
                result["waves"][-1]["wave_id"] if result["waves"] else None,
            ),
            "fanout": fanout,
            "condition": condition_from_prefix(prefix),
            "nodes": nodes,
            "header": header,
        }
        _wave_rollup(wave)
        result["waves"].append(wave)

    if not result["waves"]:
        old_headers = list(OLD_DISPATCH_RE.finditer(segment))
        for index, old_match in enumerate(old_headers):
            wave_number = int(old_match.group(1))
            start = old_match.end()
            end = old_headers[index + 1].start() if index + 1 < len(old_headers) else len(segment)
            body = segment[start:end]
            after = AFTER_WAVE_RE.search(body)
            if after:
                body = body[:after.start()]
            items, _shared = split_numbered_items(body)
            if not items and body.strip():
                items = [body.strip()]
            declared = old_match.group(3).lower()
            nodes = [
                make_node_ir(
                    item,
                    node_id=f"wave_{wave_number}.node_{item_index + 1}",
                    declared_type=declared,
                    multiplicity=numbered_item_multiplicity(item),
                )
                for item_index, item in enumerate(items)
            ]
            multiple = old_match.group(2).lower() == "multiple"
            explicit_two_shards = bool(
                re.search(r"(?i)\b(?:two|2)\s+(?:complementary\s+)?shards?\b", body)
            )
            if multiple and len(nodes) == 1 and explicit_two_shards:
                nodes[0]["multiplicity"] = 2
                fanout = {"min": 2, "max": 2, "exact": 2, "dynamic": False, "source": "prompt-shard-count"}
            elif multiple and len(nodes) == 1:
                fanout = {"min": 2, "max": None, "exact": None, "dynamic": True, "source": "legacy-multiple"}
            else:
                exact = 1 if not multiple else max(2, len(nodes))
                fanout = {"min": exact, "max": exact, "exact": exact, "dynamic": False, "source": "legacy-dispatch"}
            dependency_context_start = old_headers[index - 1].end() if index else 0
            interstitial = segment[dependency_context_start:old_match.start()]
            wave = {
                "wave_id": f"wave_{wave_number}",
                "ordinal": len(result["waves"]) + 1,
                "dependency": dependency_from_context(
                    dependency_context_for_wave(old_match.group(0).strip(), interstitial),
                    result["waves"][-1]["wave_id"] if result["waves"] else None,
                ),
                "fanout": fanout,
                "condition": condition_from_prefix(segment[:old_match.start()] if not index else ""),
                "nodes": nodes,
                "header": old_match.group(0).strip(),
            }
            _wave_rollup(wave)
            result["waves"].append(wave)

    if result["no_dispatch"] and result["waves"]:
        result["parse_warnings"].append("Prompt contains both no-dispatch wording and explicit waves; explicit waves retained.")

    return result


def expand_nodes(nodes: Sequence[dict[str, Any]]) -> list[dict[str, Any]]:
    expanded: list[dict[str, Any]] = []
    for node in nodes:
        multiplicity = max(1, int(node.get("multiplicity", 1)))
        for index in range(multiplicity):
            copy = deepcopy(node)
            copy["template_node_id"] = node["node_id"]
            copy["node_id"] = f"{node['node_id']}#{index + 1}" if multiplicity > 1 else node["node_id"]
            copy["multiplicity"] = 1
            if multiplicity > 1:
                # A repeated template's IDs describe the manifest as a whole;
                # they cannot all be literal ownership IDs for every instance.
                copy["manifest_object_ids"] = list(copy.get("object_ids") or [])
                copy["object_ids"] = [f"ordinal:{index + 1}"]
            expanded.append(copy)
    return expanded


FANOUT_DERIVATION_SCHEMAS = {
    "canvas_course_count": ({"kind"}, {"course_name_suffix"}),
    "canvas_grade_page_chunks": ({"kind", "page_size", "max_pages_per_item"}, set()),
    "canvas_grade_pages_plus_course_count": ({"kind", "course_name_suffix", "page_size"}, set()),
    "canvas_quiz_bearing_course_count": ({"kind", "course_name_suffix"}, set()),
}


def _leading_json_value(content: Any, *, task: str, call_id: str, line: int) -> Any:
    text = _content_text(content).lstrip()
    try:
        return json.JSONDecoder().raw_decode(text)[0]
    except (json.JSONDecodeError, IndexError) as exc:
        raise TrajectoryIntegrityError(
            f"{task}: pre-delegation tool result {call_id} at line {line} does not begin with JSON"
        ) from exc


def _predelegation_tool_records(task: str, raw_stream: Path) -> tuple[list[dict[str, Any]], int]:
    """Return paired non-agent tool results strictly before the first delegation call."""
    pending: dict[str, dict[str, Any]] = {}
    seen_call_ids: set[str] = set()
    records: list[dict[str, Any]] = []
    first_delegation_line: int | None = None
    with raw_stream.open(encoding="utf-8", errors="strict") as handle:
        for line_number, line in enumerate(handle, start=1):
            if not line.strip():
                continue
            try:
                payload = json.loads(line)
            except json.JSONDecodeError as exc:
                diagnostic = line.strip()
                if KNOWN_RAW_DIAGNOSTIC_RE.fullmatch(diagnostic):
                    continue
                raise TrajectoryIntegrityError(
                    f"{task}: corrupt raw_stream JSONL at {raw_stream}:{line_number}: {exc.msg}"
                ) from exc
            if not isinstance(payload, dict):
                raise TrajectoryIntegrityError(f"{task}: raw_stream row {line_number} is not a JSON object")
            if payload.get("role") == "assistant":
                calls = payload.get("tool_calls") or []
                if not isinstance(calls, list):
                    raise TrajectoryIntegrityError(f"{task}: tool_calls at line {line_number} is not a list")
                if any(
                    isinstance(call, dict)
                    and TOOL_CALL_NAME_RE.search(str((call.get("function") or {}).get("name") or ""))
                    for call in calls
                ):
                    first_delegation_line = line_number
                    break
                for call in calls:
                    if not isinstance(call, dict):
                        raise TrajectoryIntegrityError(f"{task}: tool call at line {line_number} is not an object")
                    function = call.get("function") or {}
                    name = str(function.get("name") or "")
                    call_id = str(call.get("id") or "")
                    if not call_id or not name:
                        raise TrajectoryIntegrityError(
                            f"{task}: pre-delegation tool call at line {line_number} lacks id or name"
                        )
                    if call_id in seen_call_ids:
                        raise TrajectoryIntegrityError(
                            f"{task}: duplicate pre-delegation tool call ID {call_id} at line {line_number}"
                        )
                    seen_call_ids.add(call_id)
                    pending[call_id] = {
                        "tool_call_id": call_id,
                        "tool_name": name.rsplit("__", 1)[-1],
                        "request_line": line_number,
                        "arguments": _arguments(
                            call,
                            task=task,
                            raw_stream=raw_stream,
                            stream_index=line_number,
                        ),
                    }
            elif payload.get("role") == "tool":
                call_id = str(payload.get("tool_call_id") or "")
                request = pending.pop(call_id, None)
                if request is not None:
                    records.append(
                        {
                            **request,
                            "result_line": line_number,
                            "result_content": payload.get("content"),
                        }
                    )
    if first_delegation_line is None:
        raise TrajectoryIntegrityError(f"{task}: fanout derivation has no delegation boundary")
    return records, first_delegation_line


def _course_manifest(records: Sequence[dict[str, Any]], suffix: str | None, task: str) -> list[dict[str, Any]]:
    candidates = [record for record in records if record["tool_name"] == "canvas_list_courses"]
    if len(candidates) != 1:
        raise TrajectoryIntegrityError(f"{task}: expected exactly one pre-delegation Canvas course-list result")
    courses = _leading_json_value(
        candidates[0]["result_content"],
        task=task,
        call_id=candidates[0]["tool_call_id"],
        line=candidates[0]["result_line"],
    )
    if not isinstance(courses, list):
        raise TrajectoryIntegrityError(f"{task}: Canvas course-list result is not a JSON array")
    if not all(isinstance(course, dict) and isinstance(course.get("id"), int) for course in courses):
        raise TrajectoryIntegrityError(f"{task}: Canvas course manifest has invalid rows")
    if suffix:
        courses = [course for course in courses if str(course.get("name") or "").endswith(suffix)]
    ids = [course["id"] for course in courses]
    if not ids or len(ids) != len(set(ids)):
        raise TrajectoryIntegrityError(f"{task}: Canvas course manifest is empty or has duplicate IDs")
    return courses


def _grade_total_counts(
    records: Sequence[dict[str, Any]], course_ids: set[int], task: str
) -> tuple[dict[int, int], list[int]]:
    totals: dict[int, int] = {}
    lines: list[int] = []
    for record in records:
        if record["tool_name"] != "canvas_get_course_grades":
            continue
        arguments = record["arguments"]
        course_id = arguments.get("course_id")
        if course_id not in course_ids:
            continue
        enrollment_types = arguments.get("type") or []
        if arguments.get("page") != 1 or arguments.get("per_page") != 1 or "StudentEnrollment" not in enrollment_types:
            continue
        result = _leading_json_value(
            record["result_content"],
            task=task,
            call_id=record["tool_call_id"],
            line=record["result_line"],
        )
        pagination = result.get("pagination") if isinstance(result, dict) else None
        total = pagination.get("total_count") if isinstance(pagination, dict) else None
        if isinstance(total, bool) or not isinstance(total, int) or total < 0 or course_id in totals:
            raise TrajectoryIntegrityError(f"{task}: invalid or duplicate discovery count for course {course_id}")
        totals[course_id] = total
        lines.append(record["result_line"])
    if set(totals) != course_ids:
        raise TrajectoryIntegrityError(
            f"{task}: grade discovery covers {sorted(totals)}, expected {sorted(course_ids)}"
        )
    return totals, lines


def derive_predelegation_fanout(
    task: str,
    raw_stream: Path,
    derivation: dict[str, Any],
) -> tuple[int, dict[str, Any]]:
    records, first_delegation_line = _predelegation_tool_records(task, raw_stream)
    kind = derivation["kind"]
    suffix = derivation.get("course_name_suffix")
    courses = _course_manifest(records, suffix, task)
    course_ids = {course["id"] for course in courses}
    course_proof_lines = [
        record["result_line"] for record in records if record["tool_name"] == "canvas_list_courses"
    ]
    proof_lines: list[int]
    component_multiplicities: dict[str, int] | None = None
    if kind == "canvas_course_count":
        exact = len(courses)
        proof_lines = course_proof_lines
    elif kind == "canvas_grade_page_chunks":
        page_size = derivation["page_size"]
        max_pages = derivation["max_pages_per_item"]
        totals, grade_proof_lines = _grade_total_counts(records, course_ids, task)
        proof_lines = course_proof_lines + grade_proof_lines
        exact = sum(((total + page_size - 1) // page_size + max_pages - 1) // max_pages for total in totals.values())
    elif kind == "canvas_grade_pages_plus_course_count":
        page_size = derivation["page_size"]
        totals, grade_proof_lines = _grade_total_counts(records, course_ids, task)
        proof_lines = course_proof_lines + grade_proof_lines
        page_count = sum((total + page_size - 1) // page_size for total in totals.values())
        teacher_count = len(courses)
        component_multiplicities = {"page": page_count, "teacher": teacher_count}
        exact = page_count + teacher_count
    elif kind == "canvas_quiz_bearing_course_count":
        quiz_records: dict[int, dict[str, Any]] = {}
        for record in records:
            if record["tool_name"] != "canvas_list_quizzes":
                continue
            course_id = record["arguments"].get("course_id")
            if course_id in course_ids:
                result = _leading_json_value(
                    record["result_content"],
                    task=task,
                    call_id=record["tool_call_id"],
                    line=record["result_line"],
                )
                if course_id in quiz_records or not isinstance(result, list):
                    raise TrajectoryIntegrityError(f"{task}: invalid or duplicate quiz discovery for course {course_id}")
                quiz_records[course_id] = {**record, "result": result}
        if set(quiz_records) != course_ids:
            raise TrajectoryIntegrityError(f"{task}: quiz discovery does not cover the filtered course manifest")
        exact = sum(bool(record["result"]) for record in quiz_records.values())
        proof_lines = course_proof_lines + [record["result_line"] for record in quiz_records.values()]
    else:  # validated before use
        raise ValueError(f"{task}: unknown fanout derivation kind {kind!r}")
    if exact < 1 or any(line >= first_delegation_line for line in proof_lines):
        raise TrajectoryIntegrityError(f"{task}: invalid or post-delegation fanout derivation evidence")
    proof = {
        "source": "pre-delegation-tool-results",
        "kind": kind,
        "first_delegation_line": first_delegation_line,
        "tool_result_lines": sorted(proof_lines),
        "derived_exact": exact,
    }
    if component_multiplicities is not None:
        proof["component_multiplicities"] = component_multiplicities
    return exact, proof


def validate_override_catalog(
    catalog: dict[str, Any],
    task_names: Sequence[str] | None = None,
) -> None:
    """Validate every manual denominator decision before any score is produced."""
    if not isinstance(catalog, dict):
        raise ValueError("Override catalog must be a JSON object")
    allowed_top = {"version", "policy", "tasks"}
    unknown_top = set(catalog) - allowed_top
    if unknown_top:
        raise ValueError(f"Override catalog has unknown top-level keys: {sorted(unknown_top)}")
    tasks = catalog.get("tasks", {})
    if not isinstance(tasks, dict):
        raise ValueError("Override catalog tasks must be an object")
    if task_names is not None:
        unknown_tasks = set(tasks) - set(task_names)
        if unknown_tasks:
            raise ValueError(f"Override catalog has unknown tasks: {sorted(unknown_tasks)}")
    for task, task_override in tasks.items():
        if not isinstance(task_override, dict):
            raise ValueError(f"{task}: task override must be an object")
        unknown_task_keys = set(task_override) - {"raw_stream_sha256", "waves"}
        if unknown_task_keys:
            raise ValueError(f"{task}: unknown override keys: {sorted(unknown_task_keys)}")
        digest = task_override.get("raw_stream_sha256")
        if not isinstance(digest, str) or not re.fullmatch(r"[0-9a-f]{64}", digest):
            raise ValueError(f"{task}: raw_stream_sha256 must be a lowercase SHA-256 digest")
        waves = task_override.get("waves")
        if not isinstance(waves, dict) or not waves:
            raise ValueError(f"{task}: waves must be a non-empty object")
        for wave_id, override in waves.items():
            if not isinstance(override, dict):
                raise ValueError(f"{task}/{wave_id}: override must be an object")
            unknown_wave_keys = set(override) - {"eligible", "fanout", "reason", "evidence"}
            if unknown_wave_keys:
                raise ValueError(f"{task}/{wave_id}: unknown keys: {sorted(unknown_wave_keys)}")
            if "eligible" not in override and "fanout" not in override:
                raise ValueError(f"{task}/{wave_id}: override must set eligible or fanout")
            if "eligible" in override and not isinstance(override["eligible"], bool):
                raise ValueError(f"{task}/{wave_id}: eligible must be a JSON boolean")
            if not isinstance(override.get("reason"), str) or not override["reason"].strip():
                raise ValueError(f"{task}/{wave_id}: non-empty reason is required")
            evidence = override.get("evidence")
            if not isinstance(evidence, list) or not evidence:
                raise ValueError(f"{task}/{wave_id}: non-empty evidence is required")
            for index, item in enumerate(evidence):
                if not isinstance(item, dict):
                    raise ValueError(f"{task}/{wave_id}: evidence[{index}] must be an object")
                source = item.get("source")
                if source == "raw_stream" and set(item) == {"source", "contains"}:
                    if not isinstance(item["contains"], str) or not item["contains"]:
                        raise ValueError(f"{task}/{wave_id}: evidence[{index}] has empty raw_stream text")
                    continue
                if source == "pre_delegation_tool_results" and set(item) == {"source", "derivation"}:
                    if not isinstance(item["derivation"], str) or not item["derivation"]:
                        raise ValueError(f"{task}/{wave_id}: evidence[{index}] has no derivation name")
                    continue
                raise ValueError(
                    f"{task}/{wave_id}: evidence[{index}] is not a supported evidence object"
                )
            if "fanout" in override:
                fanout = override["fanout"]
                if not isinstance(fanout, dict) or set(fanout) - {"exact", "node_multiplicities", "derivation"}:
                    raise ValueError(f"{task}/{wave_id}: fanout has an invalid schema")
                exact = fanout.get("exact")
                if isinstance(exact, bool) or not isinstance(exact, int) or exact < 1:
                    raise ValueError(f"{task}/{wave_id}: fanout.exact must be a positive integer")
                derivation = fanout.get("derivation")
                if not isinstance(derivation, dict) or not isinstance(derivation.get("kind"), str):
                    raise ValueError(f"{task}/{wave_id}: fanout.derivation is required")
                kind = derivation["kind"]
                if kind not in FANOUT_DERIVATION_SCHEMAS:
                    raise ValueError(f"{task}/{wave_id}: fanout.derivation has an invalid schema")
                required_keys, optional_keys = FANOUT_DERIVATION_SCHEMAS[kind]
                if not required_keys <= set(derivation) or set(derivation) - required_keys - optional_keys:
                    raise ValueError(f"{task}/{wave_id}: fanout.derivation has an invalid schema")
                for key in ("page_size", "max_pages_per_item"):
                    if key in derivation and (
                        isinstance(derivation[key], bool) or not isinstance(derivation[key], int) or derivation[key] < 1
                    ):
                        raise ValueError(f"{task}/{wave_id}: derivation.{key} must be a positive integer")
                if "course_name_suffix" in derivation and (
                    not isinstance(derivation["course_name_suffix"], str) or not derivation["course_name_suffix"]
                ):
                    raise ValueError(f"{task}/{wave_id}: derivation.course_name_suffix must be non-empty")
                structured_evidence = [item for item in evidence if item.get("source") == "pre_delegation_tool_results"]
                if len(structured_evidence) != 1 or structured_evidence[0]["derivation"] != kind:
                    raise ValueError(f"{task}/{wave_id}: fanout requires matching pre-delegation evidence")
                multiplicities = fanout.get("node_multiplicities")
                if multiplicities is not None:
                    if not isinstance(multiplicities, dict) or not multiplicities:
                        raise ValueError(f"{task}/{wave_id}: node_multiplicities must be a non-empty object")
                    if any(
                        not isinstance(key, str)
                        or isinstance(value, bool)
                        or not isinstance(value, int)
                        or value < 1
                        for key, value in multiplicities.items()
                    ):
                        raise ValueError(f"{task}/{wave_id}: node multiplicities must be positive integers")
                    if sum(multiplicities.values()) != exact:
                        raise ValueError(f"{task}/{wave_id}: node multiplicities must sum to fanout.exact")


def apply_verified_overrides(
    planned_ir: dict[str, Any],
    catalog: dict[str, Any],
    raw_stream: Path,
) -> None:
    """Apply only evidence-bound eligibility and runtime-manifest overrides."""
    validate_override_catalog(catalog)
    task = planned_ir["task"]
    task_override = (catalog.get("tasks") or {}).get(task)
    if task_override is None:
        return
    observed_sha = sha256_file(raw_stream)
    if observed_sha != task_override["raw_stream_sha256"]:
        raise ValueError(
            f"{task}: override evidence hash mismatch: expected {task_override['raw_stream_sha256']}, "
            f"observed {observed_sha}"
        )
    raw_text = raw_stream.read_text(encoding="utf-8", errors="strict")
    normalized_raw_text = re.sub(
        r"\s+",
        " ",
        raw_text.replace(r"\n", " ").replace(r"\r", " ").replace(r"\t", " "),
    )
    waves_by_id = {wave["wave_id"]: wave for wave in planned_ir["waves"]}
    unknown_waves = set(task_override["waves"]) - set(waves_by_id)
    if unknown_waves:
        raise ValueError(f"{task}: overrides reference unknown waves: {sorted(unknown_waves)}")
    for wave_id, override in task_override["waves"].items():
        for item in override["evidence"]:
            if item["source"] != "raw_stream":
                continue
            normalized_probe = re.sub(r"\s+", " ", item["contains"])
            if item["contains"] not in raw_text and normalized_probe not in normalized_raw_text:
                raise ValueError(f"{task}/{wave_id}: raw_stream evidence not found: {item['contains']!r}")
        wave = waves_by_id[wave_id]
        if "eligible" in override:
            wave["condition"].update(
                {
                    "eligible": override["eligible"],
                    "eligibility_source": "verified-runtime-evidence",
                    "reason": override["reason"],
                    "evidence": override["evidence"],
                    "evidence_raw_stream_sha256": observed_sha,
                }
            )
        if "fanout" in override:
            if not wave["fanout"].get("dynamic"):
                raise ValueError(f"{task}/{wave_id}: fanout override is allowed only for dynamic planned waves")
            fanout = override["fanout"]
            derived_exact, derivation_proof = derive_predelegation_fanout(
                task, raw_stream, fanout["derivation"]
            )
            if derived_exact != fanout["exact"]:
                raise ValueError(
                    f"{task}/{wave_id}: configured fanout {fanout['exact']} disagrees with "
                    f"pre-delegation derivation {derived_exact}"
                )
            multiplicities = fanout.get("node_multiplicities")
            if multiplicities is None:
                if len(wave["nodes"]) != 1:
                    raise ValueError(f"{task}/{wave_id}: multi-template fanout requires node_multiplicities")
                multiplicities = {wave["nodes"][0]["node_id"]: fanout["exact"]}
            if len(wave["nodes"]) > 1:
                derived_components = derivation_proof.get("component_multiplicities")
                if not isinstance(derived_components, dict):
                    raise ValueError(
                        f"{task}/{wave_id}: multi-template fanout requires derived component multiplicities"
                    )
                expected_multiplicities: dict[str, int] = {}
                seen_components: set[str] = set()
                for node in wave["nodes"]:
                    component = node.get("shard_key")
                    if component not in derived_components or component in seen_components:
                        raise ValueError(
                            f"{task}/{wave_id}: cannot bind planned node {node['node_id']} to a derived component"
                        )
                    seen_components.add(component)
                    expected_multiplicities[node["node_id"]] = derived_components[component]
                if multiplicities != expected_multiplicities:
                    raise ValueError(
                        f"{task}/{wave_id}: node_multiplicities {multiplicities} disagree with "
                        f"derived components {expected_multiplicities}"
                    )
            node_ids = {node["node_id"] for node in wave["nodes"]}
            if set(multiplicities) != node_ids:
                raise ValueError(
                    f"{task}/{wave_id}: node_multiplicities must cover exactly {sorted(node_ids)}"
                )
            for node in wave["nodes"]:
                node["multiplicity"] = multiplicities[node["node_id"]]
            wave["fanout"] = {
                "min": fanout["exact"],
                "max": fanout["exact"],
                "exact": fanout["exact"],
                "dynamic": False,
                "source": "verified-runtime-manifest",
                "reason": override["reason"],
                "evidence": override["evidence"],
                "derivation": fanout["derivation"],
                "derivation_proof": derivation_proof,
                "evidence_raw_stream_sha256": observed_sha,
            }
            _wave_rollup(wave)


def load_source_map(path: Path) -> dict[str, dict[str, str]]:
    rows: dict[str, dict[str, str]] = {}
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        required = {"task", "target_run"}
        if not required.issubset(set(reader.fieldnames or [])):
            raise ValueError(f"SOURCE_MAP missing columns: {sorted(required - set(reader.fieldnames or []))}")
        for row in reader:
            task = (row.get("task") or "").strip()
            if not task:
                continue
            if task in rows:
                raise ValueError(f"Duplicate SOURCE_MAP task: {task}")
            rows[task] = {str(key): str(value or "") for key, value in row.items()}
    return rows


@dataclass(frozen=True)
class RunArtifacts:
    task: str
    run_dir: Path
    raw_stream: Path
    traj_log: Path | None
    run_log: Path | None


def resolve_run_artifacts(dump_root: Path, task: str, source_row: dict[str, str] | None) -> RunArtifacts:
    if source_row is not None:
        target = Path(source_row["target_run"])
        run_dir = target if target.is_absolute() else dump_root / target
    else:
        task_dir = dump_root / task
        run_dirs = [path for path in task_dir.iterdir() if path.is_dir()] if task_dir.is_dir() else []
        if len(run_dirs) != 1:
            raise ValueError(f"{task}: exact source map required; found {len(run_dirs)} run directories")
        run_dir = run_dirs[0]
    if not run_dir.is_dir():
        raise FileNotFoundError(f"{task}: mapped run directory not found: {run_dir}")
    raws = [path for path in run_dir.rglob("raw_stream.jsonl") if path.stat().st_size >= MIN_RAW_STREAM_BYTES]
    if len(raws) != 1:
        raise ValueError(f"{task}: mapped run must contain exactly one non-stub raw_stream.jsonl, found {len(raws)}")
    trajs = list(run_dir.rglob("traj_log.json"))
    run_logs = [path for path in run_dir.rglob("run.log") if path.is_file()]
    return RunArtifacts(
        task=task,
        run_dir=run_dir,
        raw_stream=raws[0],
        traj_log=trajs[0] if len(trajs) == 1 else None,
        run_log=run_logs[0] if len(run_logs) == 1 else None,
    )


def _prompt_from_traj_log(path: Path | None) -> str | None:
    if not path or not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except (OSError, json.JSONDecodeError):
        return None
    task_str = (payload.get("config") or {}).get("task_str") if isinstance(payload, dict) else None
    return task_str if isinstance(task_str, str) and task_str.strip() else None


def _prompt_from_run_log(path: Path | None) -> str | None:
    if not path or not path.exists():
        return None
    text = ANSI_RE.sub("", path.read_text(encoding="utf-8", errors="replace"))
    markers = ["[kimi] launching: kimi -p ", "[runner] launching: kimi -p "]
    for marker in markers:
        start = text.find(marker)
        if start < 0:
            continue
        start += len(marker)
        ends = [position for token in ("\n ... (home=", " ... (home=") if (position := text.find(token, start)) >= 0]
        end = min(ends) if ends else -1
        if end > start:
            candidate = text[start:end].strip()
            if len(candidate) >= 50:
                return candidate
    return None


def extract_runtime_prompt(artifacts: RunArtifacts, reference_prompt: str) -> dict[str, Any]:
    prompt = _prompt_from_traj_log(artifacts.traj_log)
    provenance = "traj_log.config.task_str" if prompt else ""
    source_path = str(artifacts.traj_log) if prompt and artifacts.traj_log else ""
    if not prompt:
        prompt = _prompt_from_run_log(artifacts.run_log)
        provenance = "run.log launch command" if prompt else ""
        source_path = str(artifacts.run_log) if prompt and artifacts.run_log else ""
    if not prompt:
        prompt = reference_prompt
        provenance = "reference-fallback"
        source_path = ""
    return {
        "text": prompt,
        "sha256": sha256_text(prompt),
        "provenance": provenance,
        "source_path": source_path,
        "reference_sha256": sha256_text(reference_prompt),
        "matches_reference": prompt == reference_prompt,
    }


def _tool_name(name: str) -> str | None:
    match = TOOL_CALL_NAME_RE.search(name or "")
    if not match:
        return None
    value = match.group(1).lower()
    return "AgentSwarm" if value == "agentswarm" else "Agent"


def _arguments(
    tool_call: dict[str, Any],
    *,
    task: str,
    raw_stream: Path,
    stream_index: int,
) -> dict[str, Any]:
    raw = (tool_call.get("function") or {}).get("arguments")
    if isinstance(raw, dict):
        return raw
    call_id = str(tool_call.get("id") or "<missing-id>")
    if not isinstance(raw, str) or not raw.strip():
        raise TrajectoryIntegrityError(
            f"{task}: tool call {call_id} arguments at {raw_stream}:{stream_index} "
            "are missing or are not a JSON object"
        )
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise TrajectoryIntegrityError(
            f"{task}: tool call {call_id} has corrupt arguments at "
            f"{raw_stream}:{stream_index}: {exc.msg}"
        ) from exc
    if not isinstance(value, dict):
        raise TrajectoryIntegrityError(
            f"{task}: tool call {call_id} arguments at {raw_stream}:{stream_index} "
            "decode to a non-object JSON value"
        )
    return value


def _content_text(content: Any) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return " ".join(
            str(item.get("text") or item.get("content") or "") if isinstance(item, dict) else str(item)
            for item in content
        )
    if content is None:
        return ""
    return json.dumps(content, ensure_ascii=False, default=str)


def _accepted_call(tool_name: str, result_text: str) -> tuple[bool, str]:
    if not result_text.strip():
        return False, "missing-tool-result"
    low = lower_text(result_text)
    if tool_name == "AgentSwarm":
        accepted = "<agent_swarm_result" in low and bool(
            re.search(r"<subagent\b[^>]*\bagent_id=[\"'][^\"']+[\"']", result_text, re.I)
        )
    else:
        accepted = bool(re.search(r"\bagent_id\s*:\s*[a-z0-9_-]+", low)) and "status:" in low
    # Accepted result envelopes may legitimately contain words such as
    # "invalid" or "rejected" inside the sub-agent's own audit summary.  The
    # harness envelope is authoritative; rejection patterns apply only when no
    # accepted envelope exists.
    if accepted:
        return True, "accepted"
    if REJECT_PATTERNS.search(result_text):
        return False, "rejected-by-harness"
    return (accepted, "accepted" if accepted else "unconfirmed-tool-result")


def _agent_ids(tool_name: str, result_text: str) -> list[str]:
    if tool_name == "AgentSwarm":
        ids = re.findall(r"agent_id=[\"']([^\"']+)[\"']", result_text)
    else:
        ids = re.findall(r"(?i)\bagent_id\s*:\s*([a-z0-9_-]+)", result_text)
    if len(ids) != len(set(ids)):
        raise TrajectoryIntegrityError(f"{tool_name} result repeats an agent ID")
    return ids


def _swarm_result_count(result_text: str) -> int | None:
    match = re.search(r"<summary>(.*?)</summary>", result_text, re.I | re.S)
    if not match:
        return None
    counts = [int(value) for value in re.findall(r"(?i)\b(?:completed|failed)\s*:\s*(\d+)", match.group(1))]
    return sum(counts) if counts else None


def _accepted_agent_ids(tool_name: str, result_text: str, requested_count: int) -> list[str | None]:
    ids = _agent_ids(tool_name, result_text)
    if tool_name != "AgentSwarm":
        return ids
    result_count = _swarm_result_count(result_text)
    accepted_count = result_count if result_count is not None else len(ids)
    if accepted_count > requested_count:
        raise TrajectoryIntegrityError(
            f"AgentSwarm result reports {accepted_count} agents for {requested_count} requested items"
        )
    if len(ids) > accepted_count:
        raise TrajectoryIntegrityError(
            f"AgentSwarm result exposes {len(ids)} agent IDs but summary reports {accepted_count} agents"
        )
    if accepted_count == len(ids):
        return ids
    return ids + [None] * (accepted_count - len(ids))


def _swarm_item_text(item: Any) -> str:
    return item if isinstance(item, str) else json.dumps(item, ensure_ascii=False, default=str)


def _swarm_result_entries(result_text: str) -> list[dict[str, str | None]]:
    entries: list[dict[str, str | None]] = []
    for tag in re.findall(r"<subagent\b([^>]*)/?>", result_text, re.I | re.S):
        attrs = {
            name.lower(): html.unescape(value)
            for name, _quote, value in re.findall(
                r"([a-zA-Z_][\w-]*)\s*=\s*([\"'])(.*?)\2",
                tag,
                re.S,
            )
        }
        agent_id = attrs.get("agent_id")
        if agent_id is not None:
            entries.append({"agent_id": agent_id, "item": attrs.get("item")})
    return entries


def _accepted_swarm_assignments(
    result_text: str,
    requested_items: Sequence[Any],
    accepted_ids: Sequence[str | None],
) -> list[tuple[Any, str | None]]:
    """Pair accepted AgentSwarm IDs with the exact result-envelope items."""
    entries = _swarm_result_entries(result_text)
    accepted_count = len(accepted_ids)
    if len(entries) == accepted_count and all(entry["item"] is None for entry in entries):
        if accepted_count != len(requested_items):
            raise TrajectoryIntegrityError(
                "AgentSwarm accepted only part of the request without exposing every accepted item"
            )
        # A complete result envelope without item attributes preserves request
        # order; this is distinct from an incomplete preview.
        return list(zip(requested_items, [entry["agent_id"] for entry in entries]))
    if len(entries) == accepted_count and any(entry["item"] is None for entry in entries):
        raise TrajectoryIntegrityError(
            "AgentSwarm result inconsistently exposes item attributes"
        )

    requested_text = [_swarm_item_text(item) for item in requested_items]

    if len(entries) < accepted_count:
        if accepted_count != len(requested_items) or not SWARM_TRUNCATED_PREVIEW_RE.search(result_text):
            raise TrajectoryIntegrityError(
                "AgentSwarm result omits accepted entries without an explicit full-accept preview"
            )
        used_preview_indices: set[int] = set()
        bases: set[int] = set()
        visible_pairs: list[tuple[dict[str, str | None], int]] = []
        for position, entry in enumerate(entries):
            numeric_id = re.fullmatch(r"agent-(\d+)", str(entry["agent_id"]))
            if numeric_id is None:
                raise TrajectoryIntegrityError(
                    "AgentSwarm truncated preview cannot infer non-numeric agent IDs"
                )
            if entry["item"] is None:
                requested_index = position
            else:
                candidates = [
                    index
                    for index, candidate in enumerate(requested_text)
                    if index not in used_preview_indices and candidate == str(entry["item"])
                ]
                if not candidates:
                    raise TrajectoryIntegrityError(
                        f"AgentSwarm preview item {entry['item']!r} does not match an unmatched requested item"
                    )
                requested_index = candidates[0]
            if requested_index in used_preview_indices:
                raise TrajectoryIntegrityError("AgentSwarm preview maps multiple entries to one request item")
            used_preview_indices.add(requested_index)
            bases.add(int(numeric_id.group(1)) - requested_index)
            visible_pairs.append((entry, requested_index))
        if len(bases) != 1:
            raise TrajectoryIntegrityError("AgentSwarm preview item/agent mappings imply inconsistent ID bases")
        base = next(iter(bases))
        if base < 0:
            raise TrajectoryIntegrityError("AgentSwarm preview implies a negative agent-ID base")
        inferred_ids = [f"agent-{base + index}" for index in range(accepted_count)]
        for entry, requested_index in visible_pairs:
            if entry["agent_id"] != inferred_ids[requested_index]:
                raise TrajectoryIntegrityError("AgentSwarm preview contradicts a visible item/agent mapping")
        return list(zip(requested_items, inferred_ids))

    if len(entries) != accepted_count:
        raise TrajectoryIntegrityError("AgentSwarm result exposes more entries than accepted agents")

    used: set[int] = set()
    assignments: list[tuple[Any, str | None]] = []
    for entry in entries:
        item_text = str(entry["item"])
        candidates = [
            index
            for index, candidate in enumerate(requested_text)
            if index not in used and candidate == item_text
        ]
        if not candidates:
            raise TrajectoryIntegrityError(
                f"AgentSwarm result item {item_text!r} does not match an unmatched requested item"
            )
        index = candidates[0]
        used.add(index)
        assignments.append((requested_items[index], entry["agent_id"]))
    return assignments


def _render_swarm_prompt(template: str, item: Any) -> str:
    item_text = _swarm_item_text(item)
    rendered = (template or "").replace("{{item}}", item_text).replace("{item}", item_text)
    return rendered + "\nSWARM_ITEM=" + item_text


def extract_actual_ir(task: str, raw_stream: Path) -> dict[str, Any]:
    events: list[dict[str, Any]] = []
    result_by_id: dict[str, str] = {}
    seen_agent_call_ids: set[str] = set()
    seen_result_ids: set[str] = set()
    assistant_index = -1
    parsed_rows = 0
    diagnostic_lines: list[dict[str, Any]] = []
    with raw_stream.open(encoding="utf-8", errors="strict") as handle:
        for stream_index, line in enumerate(handle, start=1):
            if not line.strip():
                continue
            try:
                payload = json.loads(line)
            except json.JSONDecodeError as exc:
                diagnostic = line.strip()
                if KNOWN_RAW_DIAGNOSTIC_RE.fullmatch(diagnostic):
                    diagnostic_lines.append({"line": stream_index, "text": diagnostic})
                    continue
                raise TrajectoryIntegrityError(
                    f"{task}: corrupt raw_stream JSONL at {raw_stream}:{stream_index}: {exc.msg}"
                ) from exc
            if not isinstance(payload, dict):
                raise TrajectoryIntegrityError(f"{task}: raw_stream row {stream_index} is not a JSON object")
            parsed_rows += 1
            role = payload.get("role")
            if role == "assistant":
                assistant_index += 1
                content = _content_text(payload.get("content"))
                for call in payload.get("tool_calls") or []:
                    name = _tool_name((call.get("function") or {}).get("name") or "")
                    if not name:
                        continue
                    call_id = str(call.get("id") or "")
                    if not call_id:
                        raise TrajectoryIntegrityError(
                            f"{task}: Agent/AgentSwarm request at {raw_stream}:{stream_index} has no tool-call ID"
                        )
                    if call_id in seen_agent_call_ids:
                        raise TrajectoryIntegrityError(
                            f"{task}: duplicate Agent/AgentSwarm request ID {call_id} at "
                            f"{raw_stream}:{stream_index}"
                        )
                    seen_agent_call_ids.add(call_id)
                    events.append(
                        {
                            "tool_call_id": call_id,
                            "tool_name": name,
                            "arguments": _arguments(
                                call,
                                task=task,
                                raw_stream=raw_stream,
                                stream_index=stream_index,
                            ),
                            "assistant_index": assistant_index,
                            "stream_index": stream_index,
                            "assistant_content": content,
                        }
                    )
            elif role == "tool":
                call_id = str(payload.get("tool_call_id") or "")
                if call_id:
                    if call_id in seen_result_ids:
                        raise TrajectoryIntegrityError(
                            f"{task}: duplicate tool result ID {call_id} at {raw_stream}:{stream_index}"
                        )
                    seen_result_ids.add(call_id)
                    result_by_id[call_id] = _content_text(payload.get("content"))

    if parsed_rows == 0:
        raise TrajectoryIntegrityError(f"{task}: raw_stream contains no JSON records: {raw_stream}")

    by_message: dict[int, list[dict[str, Any]]] = defaultdict(list)
    rejected: list[dict[str, Any]] = []
    resumes: list[dict[str, Any]] = []
    agent_to_node: dict[str, tuple[int, int]] = {}
    call_records: list[dict[str, Any]] = []

    for event in events:
        result_text = result_by_id.get(event["tool_call_id"], "")
        accepted, acceptance_reason = _accepted_call(event["tool_name"], result_text)
        args = event["arguments"]
        resume_id = args.get("resume") or args.get("resume_id")
        resume_agent_ids = args.get("resume_agent_ids")
        if not isinstance(resume_agent_ids, dict):
            resume_agent_ids = {}
        record = {
            "tool_call_id": event["tool_call_id"],
            "tool_name": event["tool_name"],
            "assistant_index": event["assistant_index"],
            "accepted": accepted,
            "acceptance_reason": acceptance_reason,
            "resume_id": str(resume_id) if resume_id else None,
            "resume_agent_ids": sorted(str(value) for value in resume_agent_ids),
            "description": str(args.get("description") or ""),
            "result_excerpt": result_text[:500],
        }
        call_records.append(record)
        if not accepted:
            rejected.append(record)
            continue
        if resume_agent_ids:
            ids = _accepted_agent_ids(event["tool_name"], result_text, len(resume_agent_ids))
            concrete_ids = [value for value in ids if value is not None]
            unknown_ids = set(concrete_ids) - {str(value) for value in resume_agent_ids}
            if unknown_ids:
                raise TrajectoryIntegrityError(
                    f"{task}: resume AgentSwarm {event['tool_call_id']} returned unknown agent IDs: "
                    f"{sorted(unknown_ids)}"
                )
            record["requested_nodes"] = len(resume_agent_ids)
            record["accepted_nodes"] = len(ids)
            for agent_id in concrete_ids:
                resumes.append(
                    {
                        **record,
                        "resume_id": agent_id,
                        "prompt": str(resume_agent_ids[agent_id]),
                        "linked": False,
                    }
                )
            continue
        if resume_id:
            ids = _accepted_agent_ids(event["tool_name"], result_text, 1)
            if str(resume_id) not in ids:
                raise TrajectoryIntegrityError(
                    f"{task}: resume call {event['tool_call_id']} did not return requested agent ID {resume_id}"
                )
            record["requested_nodes"] = 1
            record["accepted_nodes"] = 1
            resumes.append({**record, "prompt": str(args.get("prompt") or ""), "linked": False})
            continue

        shared = str(args.get("prompt") or args.get("prompt_template") or "")
        description = str(args.get("description") or "")
        items = args.get("items") if event["tool_name"] == "AgentSwarm" and isinstance(args.get("items"), list) else None
        item_values: list[Any] = items if items is not None else [None]
        ids = _accepted_agent_ids(event["tool_name"], result_text, len(item_values))
        if event["tool_name"] == "AgentSwarm":
            assignments = _accepted_swarm_assignments(result_text, item_values, ids)
            item_values = [item for item, _agent_id in assignments]
            ids = [agent_id for _item, agent_id in assignments]
        elif len(ids) != 1:
            raise ValueError(f"{task}: accepted Agent call {event['tool_call_id']} has {len(ids)} agent IDs")
        record["requested_nodes"] = len(items) if items is not None else 1
        record["accepted_nodes"] = len(ids)
        nodes: list[dict[str, Any]] = []
        for item_index, item in enumerate(item_values):
            prompt = _render_swarm_prompt(shared, item) if items is not None else shared
            text = "\n".join(part for part in (description, prompt) if part)
            node = make_node_ir(
                text,
                node_id=f"actual_msg_{event['assistant_index']}.call_{event['tool_call_id']}.item_{item_index + 1}",
                declared_type=str(args.get("subagent_type") or "") or None,
                item=item,
                signature_text=description,
            )
            repair_probe = "\n".join(part for part in (description, prompt[:500]) if part)
            node.update(
                {
                    "tool_call_id": event["tool_call_id"],
                    "agent_id": ids[item_index],
                    "description": description,
                    "is_repair_fresh": bool(REPAIR_RE.search(repair_probe)),
                    "accepted_result": acceptance_reason,
                }
            )
            nodes.append(node)
        by_message[event["assistant_index"]].extend(nodes)

    fresh_waves: list[dict[str, Any]] = []
    for message_index in sorted(by_message):
        nodes = by_message[message_index]
        previous_wave_id = fresh_waves[-1]["wave_id"] if fresh_waves else None
        wave = {
            "wave_id": f"fresh_wave_{len(fresh_waves) + 1}",
            "ordinal": len(fresh_waves) + 1,
            "dependency": previous_wave_id or "main-agent-prerequisites",
            "fanout": {
                "min": len(nodes),
                "max": len(nodes),
                "exact": len(nodes),
                "dynamic": False,
                "source": "accepted-result-envelopes",
            },
            "condition": {
                "kind": "observed",
                "text": "",
                "eligible": True,
                "eligibility_source": "accepted-result-envelopes",
                "reason": "This accepted fresh wave occurred in the exact trajectory.",
            },
            "assistant_index": message_index,
            "nodes": nodes,
            "repair_fresh": any(node.get("is_repair_fresh") for node in nodes),
        }
        _wave_rollup(wave)
        fresh_waves.append(wave)
        wave_index = len(fresh_waves) - 1
        for node_index, node in enumerate(nodes):
            if node.get("agent_id"):
                agent_to_node[str(node["agent_id"])] = (wave_index, node_index)

    for resume in resumes:
        link = agent_to_node.get(str(resume["resume_id"]))
        if link is None:
            continue
        wave_index, node_index = link
        resume["linked"] = True
        resume["linked_wave_id"] = fresh_waves[wave_index]["wave_id"]
        resume["linked_node_id"] = fresh_waves[wave_index]["nodes"][node_index]["node_id"]
        node = fresh_waves[wave_index]["nodes"][node_index]
        node.setdefault("resume_events", []).append(
            {
                "tool_call_id": resume["tool_call_id"],
                "assistant_index": resume["assistant_index"],
                "prompt": resume["prompt"],
            }
        )
        merged = node["text"] + "\nRESUME:\n" + resume["prompt"]
        node["actions"] = stable_unique(node["actions"] + infer_actions(resume["prompt"], node["agent_type"]))
        node["domains"] = stable_unique(node["domains"] + infer_domains(resume["prompt"]))
        node["write_roles"] = stable_unique(node["write_roles"] + infer_write_roles(resume["prompt"], node["actions"]))
        node["targets"] = stable_unique(node["targets"] + infer_targets(resume["prompt"]))
        node["object_ids"] = stable_unique(node["object_ids"] + infer_object_ids(resume["prompt"]))
        node["text"] = merged
        _wave_rollup(fresh_waves[wave_index])

    return {
        "task": task,
        "fresh_waves": fresh_waves,
        "resumes": resumes,
        "rejected_calls": rejected,
        "call_records": call_records,
        "diagnostic_lines": diagnostic_lines,
        "counts": {
            "fresh_waves": len(fresh_waves),
            "fresh_nodes": sum(wave["node_instances"] for wave in fresh_waves),
            "accepted_resumes": len(resumes),
            "linked_resumes": sum(bool(item.get("linked")) for item in resumes),
            "orphan_resumes": sum(not bool(item.get("linked")) for item in resumes),
            "rejected_calls": len(rejected),
            "repair_fresh_waves": sum(bool(wave.get("repair_fresh")) for wave in fresh_waves),
            "repair_fresh_nodes": sum(wave["node_instances"] for wave in fresh_waves if wave.get("repair_fresh")),
            "known_diagnostic_lines": len(diagnostic_lines),
        },
    }


def jaccard(left: Sequence[str], right: Sequence[str]) -> float:
    a, b = set(left), set(right)
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def requirement_coverage(required: Sequence[str], observed: Sequence[str]) -> float:
    needed, actual = set(required), set(observed)
    if not needed:
        return 1.0
    return len(needed & actual) / len(needed)


def node_similarity(planned: dict[str, Any], actual: dict[str, Any]) -> tuple[float, dict[str, Any]]:
    fields: list[tuple[str, float]] = [("actions", 0.35)]
    if planned.get("domains"):
        fields.append(("domains", 0.20))
    constrained_roles = [role for role in planned.get("write_roles", []) if role != "other-write"]
    if constrained_roles:
        fields.append(("write_roles", 0.20))
    if planned.get("targets"):
        fields.append(("targets", 0.10))
    if planned.get("object_ids") and not all(str(value).startswith("ordinal:") for value in planned["object_ids"]):
        fields.append(("object_ids", 0.10))
    if planned.get("shard_key"):
        fields.append(("shard_key", 0.05))
    if planned.get("agent_type") not in {None, "", "unspecified"}:
        fields.append(("agent_type", 0.05))

    planned_writes = "write" in planned.get("actions", [])
    actual_writes = "write" in actual.get("actions", [])
    if planned_writes != actual_writes:
        return 0.0, {"veto": "write-action-mismatch"}
    planned_type = planned.get("agent_type")
    actual_type = actual.get("agent_type")
    if (
        planned_type not in {None, "", "unspecified"}
        and actual_type not in {None, "", "unspecified"}
        and planned_type != actual_type
    ):
        return 0.0, {"veto": "agent-type-disjoint"}
    concrete_actual_roles = [role for role in actual.get("write_roles", []) if role != "other-write"]
    if constrained_roles and concrete_actual_roles and not set(constrained_roles) & set(concrete_actual_roles):
        return 0.0, {"veto": "write-role-disjoint"}
    if planned.get("shard_key") and actual.get("shard_key") and planned["shard_key"] != actual["shard_key"]:
        return 0.0, {"veto": "shard-key-disjoint"}
    concrete_objects = {
        value
        for value in planned.get("object_ids", [])
        if not str(value).startswith(("ordinal:", "position:", "rows:", "item:"))
    }
    if concrete_objects and actual.get("object_ids") and not concrete_objects & set(actual["object_ids"]):
        return 0.0, {"veto": "object-id-disjoint"}

    numerator = denominator = 0.0
    evidence: dict[str, Any] = {}
    for field, weight in fields:
        denominator += weight
        pvalue, avalue = planned.get(field), actual.get(field)
        if field in {"agent_type", "shard_key"}:
            score = 1.0 if pvalue and avalue and pvalue == avalue else 0.0
        else:
            required = constrained_roles if field == "write_roles" else (pvalue or [])
            score = requirement_coverage(required, avalue or [])
        numerator += weight * score
        evidence[field] = score
    score = numerator / denominator if denominator else 0.0
    evidence["score"] = score
    return score, evidence


def node_contract_adherent(
    planned: dict[str, Any],
    actual: dict[str, Any],
    evidence: dict[str, Any],
) -> bool:
    """Require all planned actions and every explicit final-write role."""
    if evidence.get("veto"):
        return False
    if not set(planned.get("actions", [])) <= set(actual.get("actions", [])):
        return False
    planned_roles = {role for role in planned.get("write_roles", []) if role != "other-write"}
    actual_roles = {role for role in actual.get("write_roles", []) if role != "other-write"}
    if planned_roles and not planned_roles <= actual_roles:
        return False
    return True


def maximum_cardinality_matching(
    planned_nodes: Sequence[dict[str, Any]],
    actual_nodes: Sequence[dict[str, Any]],
    *,
    threshold: float,
    strict_contract: bool = False,
) -> list[dict[str, Any]]:
    candidates: dict[int, list[tuple[int, float, dict[str, Any]]]] = defaultdict(list)
    for planned_index, planned in enumerate(planned_nodes):
        for actual_index, actual in enumerate(actual_nodes):
            score, evidence = node_similarity(planned, actual)
            if score >= threshold and (
                not strict_contract or node_contract_adherent(planned, actual, evidence)
            ):
                candidates[planned_index].append((actual_index, score, evidence))
    for planned_index in candidates:
        candidates[planned_index].sort(key=lambda item: (-item[1], item[0]))

    matched_actual: dict[int, int] = {}
    selected_evidence: dict[tuple[int, int], tuple[float, dict[str, Any]]] = {}

    def augment(planned_index: int, seen_actual: set[int]) -> bool:
        for actual_index, score, evidence in candidates.get(planned_index, []):
            if actual_index in seen_actual:
                continue
            seen_actual.add(actual_index)
            previous = matched_actual.get(actual_index)
            if previous is None or augment(previous, seen_actual):
                matched_actual[actual_index] = planned_index
                selected_evidence[(planned_index, actual_index)] = (score, evidence)
                return True
        return False

    planned_order = sorted(range(len(planned_nodes)), key=lambda index: (len(candidates.get(index, [])), index))
    for planned_index in planned_order:
        augment(planned_index, set())

    matches: list[dict[str, Any]] = []
    for actual_index, planned_index in sorted(matched_actual.items(), key=lambda item: (item[1], item[0])):
        score, evidence = selected_evidence[(planned_index, actual_index)]
        matches.append(
            {
                "planned_index": planned_index,
                "actual_index": actual_index,
                "planned_node_id": planned_nodes[planned_index]["node_id"],
                "actual_node_id": actual_nodes[actual_index]["node_id"],
                "score": score,
                "evidence": evidence,
            }
        )
    return matches


def fanout_adherent(planned_wave: dict[str, Any], actual_wave: dict[str, Any]) -> bool:
    count = int(actual_wave.get("node_instances", 0))
    minimum = planned_wave["fanout"].get("min")
    maximum = planned_wave["fanout"].get("max")
    return (minimum is None or count >= int(minimum)) and (maximum is None or count <= int(maximum))


def wave_similarity(planned_wave: dict[str, Any], actual_wave: dict[str, Any]) -> dict[str, Any]:
    template_nodes = expand_nodes(planned_wave["nodes"])
    planned_nodes = template_nodes
    actual_nodes = expand_nodes(actual_wave["nodes"])
    node_matches = maximum_cardinality_matching(planned_nodes, actual_nodes, threshold=0.45)
    contract_matches = maximum_cardinality_matching(
        planned_nodes,
        actual_nodes,
        threshold=0.45,
        strict_contract=True,
    )
    template_matches = maximum_cardinality_matching(template_nodes, actual_nodes, threshold=0.45)
    template_coverage = len({match["planned_index"] for match in template_matches}) / len(template_nodes) if template_nodes else 0.0
    planned_coverage = len(node_matches) / len(planned_nodes) if planned_nodes else 0.0
    actual_coverage = len(node_matches) / len(actual_nodes) if actual_nodes else 0.0
    action_score = requirement_coverage(planned_wave.get("actions", []), actual_wave.get("actions", []))
    domain_score = requirement_coverage(planned_wave.get("domains", []), actual_wave.get("domains", []))
    planned_roles = [role for role in planned_wave.get("write_roles", []) if role != "other-write"]
    role_score = requirement_coverage(planned_roles, actual_wave.get("write_roles", []))
    pcount, acount = len(planned_nodes), len(actual_nodes)
    fanout_score = min(pcount, acount) / max(pcount, acount) if pcount and acount else 0.0
    score = 0.40 * planned_coverage + 0.20 * action_score + 0.20 * fanout_score + 0.10 * domain_score + 0.10 * role_score
    semantic_compatible = bool(node_matches) and template_coverage >= 0.50 and planned_coverage >= 0.34 and action_score >= 0.34 and score >= 0.48
    fanout_ok = fanout_adherent(planned_wave, actual_wave)
    planned_has_write = "write" in planned_wave.get("actions", [])
    actual_has_write = "write" in actual_wave.get("actions", [])
    action_ok = action_score == 1.0 and planned_has_write == actual_has_write
    actual_roles = {role for role in actual_wave.get("write_roles", []) if role != "other-write"}
    role_ok = not planned_roles or set(planned_roles) <= actual_roles
    strict_planned_coverage = len(contract_matches) / len(planned_nodes) if planned_nodes else 0.0
    contract_without_dependency = (
        semantic_compatible
        and template_coverage >= 1.0
        and strict_planned_coverage >= 1.0
        and fanout_ok
        and action_ok
        and role_ok
    )
    return {
        "score": score,
        "semantic_compatible": semantic_compatible,
        "contract_without_dependency": contract_without_dependency,
        # ordered_wave_alignment replaces these two fields after resolving the
        # exact dependency target through the selected monotonic alignment.
        "contract_hit": contract_without_dependency,
        "fanout_adherent": fanout_ok,
        "action_adherent": action_ok,
        "write_role_adherent": role_ok,
        "dependency_adherent": None,
        "planned_node_coverage": planned_coverage,
        "strict_planned_node_coverage": strict_planned_coverage,
        "planned_template_coverage": template_coverage,
        "actual_node_coverage": actual_coverage,
        "action_score": action_score,
        "domain_score": domain_score,
        "write_role_score": role_score,
        "fanout_score": fanout_score,
        "coarse_node_matches": node_matches,
        "contract_node_matches": contract_matches,
}


def _dependency_target_index(wave: dict[str, Any], waves: Sequence[dict[str, Any]]) -> int | None:
    dependency = wave.get("dependency")
    if dependency == "main-agent-prerequisites":
        return None
    for index, candidate in enumerate(waves):
        if candidate.get("wave_id") == dependency:
            return index
    return -1


def _aligned_dependency_adherent(
    planned_index: int,
    actual_index: int,
    planned_waves: Sequence[dict[str, Any]],
    actual_waves: Sequence[dict[str, Any]],
    planned_to_actual: Sequence[int],
) -> bool:
    planned_dependency = _dependency_target_index(planned_waves[planned_index], planned_waves)
    actual_dependency = _dependency_target_index(actual_waves[actual_index], actual_waves)
    if planned_dependency is None:
        return actual_dependency is None
    if (
        actual_dependency is None
        or planned_dependency < 0
        or actual_dependency < 0
        or planned_dependency >= planned_index
        or actual_dependency >= actual_index
    ):
        return False
    return planned_to_actual[planned_dependency] == actual_dependency


def ordered_wave_alignment(planned_waves: Sequence[dict[str, Any]], actual_waves: Sequence[dict[str, Any]]) -> dict[str, Any]:
    similarity = [[wave_similarity(planned, actual) for actual in actual_waves] for planned in planned_waves]
    rows, cols = len(planned_waves), len(actual_waves)
    planned_dependencies = [_dependency_target_index(wave, planned_waves) for wave in planned_waves]
    needed_targets: list[set[int]] = []
    for start in range(rows + 1):
        needed_targets.append(
            {
                dependency
                for dependency in planned_dependencies[start:]
                if isinstance(dependency, int) and dependency >= 0
            }
        )

    memo: dict[
        tuple[int, int, tuple[int, ...]],
        tuple[tuple[int, int, float], tuple[tuple[int, int, bool], ...]],
    ] = {}

    def normalize_bindings(start: int, bindings: tuple[int, ...]) -> tuple[int, ...]:
        needed = needed_targets[start]
        return tuple(value if index in needed else -2 for index, value in enumerate(bindings))

    def solve(
        planned_index: int,
        actual_index: int,
        bindings: tuple[int, ...],
    ) -> tuple[tuple[int, int, float], tuple[tuple[int, int, bool], ...]]:
        bindings = normalize_bindings(planned_index, bindings)
        key = (planned_index, actual_index, bindings)
        if key in memo:
            return memo[key]
        if planned_index >= rows or actual_index >= cols:
            result = ((0, 0, 0.0), ())
            memo[key] = result
            return result

        skipped_planned = list(bindings)
        skipped_planned[planned_index] = -1
        options = [
            solve(planned_index + 1, actual_index, tuple(skipped_planned)),
            solve(planned_index, actual_index + 1, bindings),
        ]
        sim = similarity[planned_index][actual_index]
        if sim["semantic_compatible"]:
            paired_bindings = list(bindings)
            paired_bindings[planned_index] = actual_index
            dependency_ok = _aligned_dependency_adherent(
                planned_index,
                actual_index,
                planned_waves,
                actual_waves,
                paired_bindings,
            )
            tail_objective, tail_path = solve(
                planned_index + 1,
                actual_index + 1,
                tuple(paired_bindings),
            )
            contract_hit = bool(sim["contract_without_dependency"] and dependency_ok)
            options.append(
                (
                    (
                        tail_objective[0] + int(contract_hit),
                        tail_objective[1] + 1,
                        tail_objective[2] + sim["score"],
                    ),
                    ((planned_index, actual_index, dependency_ok),) + tail_path,
                )
            )

        # Primary objective: true contract hits. Semantic diagnostic pair count
        # and score are secondary. Stable ties prefer earlier actual waves.
        result = max(
            options,
            key=lambda candidate: candidate[0]
            + (tuple(-pair[1] for pair in candidate[1]),),
        )
        memo[key] = result
        return result

    _objective, selected_path = solve(0, 0, tuple(-2 for _ in range(rows)))
    pairs: list[dict[str, Any]] = []
    for planned_index, actual_index, dependency_ok in selected_path:
        sim = deepcopy(similarity[planned_index][actual_index])
        sim.update(
            {
                "planned_index": planned_index,
                "actual_index": actual_index,
                "planned_wave_id": planned_waves[planned_index]["wave_id"],
                "actual_wave_id": actual_waves[actual_index]["wave_id"],
                "dependency_adherent": dependency_ok,
                "contract_hit": bool(sim["contract_without_dependency"] and dependency_ok),
            }
        )
        pairs.append(sim)
    planned_paired = {pair["planned_index"] for pair in pairs}
    actual_paired = {pair["actual_index"] for pair in pairs}
    contract_pairs = [pair for pair in pairs if pair["contract_hit"]]
    return {
        "pairs": pairs,
        "contract_pairs": contract_pairs,
        "unpaired_planned": [index for index in range(rows) if index not in planned_paired],
        "unpaired_actual": [index for index in range(cols) if index not in actual_paired],
        "contract_planned_indices": sorted(pair["planned_index"] for pair in contract_pairs),
        "contract_actual_indices": sorted(pair["actual_index"] for pair in contract_pairs),
    }


def fine_metrics(
    planned_waves: Sequence[dict[str, Any]],
    actual_waves: Sequence[dict[str, Any]],
    alignment: dict[str, Any],
) -> dict[str, Any]:
    actual_nodes_by_wave = [expand_nodes(wave["nodes"]) for wave in actual_waves]
    planned_nodes_by_wave = [expand_nodes(wave["nodes"]) for wave in planned_waves]
    matched_planned: set[tuple[int, int]] = set()
    matched_actual: set[tuple[int, int]] = set()
    matches: list[dict[str, Any]] = []
    for pair in alignment["pairs"]:
        planned_index, actual_index = pair["planned_index"], pair["actual_index"]
        local_matches = maximum_cardinality_matching(
            planned_nodes_by_wave[planned_index],
            actual_nodes_by_wave[actual_index],
            threshold=0.58,
            strict_contract=True,
        )
        for match in local_matches:
            matched_planned.add((planned_index, match["planned_index"]))
            matched_actual.add((actual_index, match["actual_index"]))
            matches.append(
                {
                    **match,
                    "planned_wave_index": planned_index,
                    "actual_wave_index": actual_index,
                    "planned_wave_id": planned_waves[planned_index]["wave_id"],
                    "actual_wave_id": actual_waves[actual_index]["wave_id"],
                }
            )
    planned_total = sum(len(nodes) for nodes in planned_nodes_by_wave)
    actual_total = sum(len(nodes) for nodes in actual_nodes_by_wave)
    return {
        "planned_total": planned_total,
        "planned_hit": len(matched_planned),
        "actual_total": actual_total,
        "actual_hit": len(matched_actual),
        "recall": len(matched_planned) / planned_total if planned_total else None,
        "precision": len(matched_actual) / actual_total if actual_total else None,
        "matches": matches,
        "unmatched_planned": [
            {"wave_index": wave_index, "node_id": node["node_id"]}
            for wave_index, nodes in enumerate(planned_nodes_by_wave)
            for node_index, node in enumerate(nodes)
            if (wave_index, node_index) not in matched_planned
        ],
        "unmatched_actual": [
            {"wave_index": wave_index, "node_id": node["node_id"]}
            for wave_index, nodes in enumerate(actual_nodes_by_wave)
            for node_index, node in enumerate(nodes)
            if (wave_index, node_index) not in matched_actual
        ],
    }


def score_task(planned_ir: dict[str, Any], actual_ir: dict[str, Any]) -> dict[str, Any]:
    eligible_waves = [wave for wave in planned_ir["waves"] if wave["condition"].get("eligible", True)]
    ineligible_waves = [wave for wave in planned_ir["waves"] if not wave["condition"].get("eligible", True)]
    unresolved = [wave["wave_id"] for wave in eligible_waves if wave.get("fanout", {}).get("dynamic")]
    if unresolved:
        raise ValueError(
            f"{planned_ir.get('task', '<unknown>')}: eligible dynamic waves lack a verified runtime manifest: "
            f"{unresolved}"
        )
    actual_waves = actual_ir["fresh_waves"]
    alignment = ordered_wave_alignment(eligible_waves, actual_waves)
    fine = fine_metrics(eligible_waves, actual_waves, alignment)
    wave_hit = len(alignment["contract_pairs"])
    actual_hit = len(alignment["contract_actual_indices"])
    counts = actual_ir["counts"]
    overhead_numerator = counts["accepted_resumes"] + counts["repair_fresh_nodes"]
    overhead_denominator = counts["fresh_nodes"] + counts["accepted_resumes"]
    return {
        "planned_wave_total": len(eligible_waves),
        "planned_wave_hit": wave_hit,
        "planned_wave_recall": wave_hit / len(eligible_waves) if eligible_waves else None,
        "ineligible_wave_total": len(ineligible_waves),
        "fresh_wave_total": len(actual_waves),
        "fresh_wave_hit": actual_hit,
        "fresh_wave_precision": actual_hit / len(actual_waves) if actual_waves else None,
        "alignment": alignment,
        "fine": fine,
        "overhead": {
            **counts,
            "overhead_events": overhead_numerator,
            "overhead_denominator": overhead_denominator,
            "overhead_rate": overhead_numerator / overhead_denominator if overhead_denominator else 0.0,
        },
    }


def aggregate_scores(task_records: Sequence[dict[str, Any]]) -> dict[str, Any]:
    totals = {
        "tasks": len(task_records),
        "tasks_scored": 0,
        "tasks_invalid_trajectory": 0,
        "planned_wave_total": 0,
        "planned_wave_hit": 0,
        "ineligible_wave_total": 0,
        "fresh_wave_total": 0,
        "fresh_wave_hit": 0,
        "fine_planned_total": 0,
        "fine_planned_hit": 0,
        "fine_actual_total": 0,
        "fine_actual_hit": 0,
        "accepted_resumes": 0,
        "linked_resumes": 0,
        "orphan_resumes": 0,
        "rejected_calls": 0,
        "repair_fresh_waves": 0,
        "repair_fresh_nodes": 0,
        "fresh_nodes": 0,
        "known_diagnostic_lines": 0,
        "verified_runtime_fanout_waves": 0,
        "prompt_from_traj_log": 0,
        "prompt_from_run_log": 0,
        "prompt_reference_fallback": 0,
        "prompt_reference_mismatch": 0,
    }
    for record in task_records:
        provenance = record["prompt"]["provenance"]
        if provenance == "traj_log.config.task_str":
            totals["prompt_from_traj_log"] += 1
        elif provenance == "run.log launch command":
            totals["prompt_from_run_log"] += 1
        else:
            totals["prompt_reference_fallback"] += 1
        totals["prompt_reference_mismatch"] += int(not record["prompt"]["matches_reference"])
        totals["verified_runtime_fanout_waves"] += sum(
            wave.get("fanout", {}).get("source") == "verified-runtime-manifest"
            for wave in record["planned_ir"]["waves"]
        )
        if record.get("score") is None:
            totals["tasks_invalid_trajectory"] += 1
            continue
        totals["tasks_scored"] += 1
        score = record["score"]
        totals["planned_wave_total"] += score["planned_wave_total"]
        totals["planned_wave_hit"] += score["planned_wave_hit"]
        totals["ineligible_wave_total"] += score["ineligible_wave_total"]
        totals["fresh_wave_total"] += score["fresh_wave_total"]
        totals["fresh_wave_hit"] += score["fresh_wave_hit"]
        totals["fine_planned_total"] += score["fine"]["planned_total"]
        totals["fine_planned_hit"] += score["fine"]["planned_hit"]
        totals["fine_actual_total"] += score["fine"]["actual_total"]
        totals["fine_actual_hit"] += score["fine"]["actual_hit"]
        for key in (
            "accepted_resumes",
            "linked_resumes",
            "orphan_resumes",
            "rejected_calls",
            "repair_fresh_waves",
            "repair_fresh_nodes",
            "fresh_nodes",
            "known_diagnostic_lines",
        ):
            totals[key] += score["overhead"].get(key, 0)
    totals.update(
        {
            "planned_wave_recall": totals["planned_wave_hit"] / totals["planned_wave_total"] if totals["planned_wave_total"] else None,
            "fresh_wave_precision": totals["fresh_wave_hit"] / totals["fresh_wave_total"] if totals["fresh_wave_total"] else None,
            "fine_node_recall": totals["fine_planned_hit"] / totals["fine_planned_total"] if totals["fine_planned_total"] else None,
            "fine_node_precision": totals["fine_actual_hit"] / totals["fine_actual_total"] if totals["fine_actual_total"] else None,
            "resume_repair_overhead_rate": (
                (totals["accepted_resumes"] + totals["repair_fresh_nodes"])
                / (totals["fresh_nodes"] + totals["accepted_resumes"])
                if totals["fresh_nodes"] + totals["accepted_resumes"]
                else 0.0
            ),
        }
    )
    return totals
