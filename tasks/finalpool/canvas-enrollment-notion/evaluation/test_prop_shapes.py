"""Synthetic tests for Notion property reading (both wrapped/unwrapped shapes).

The provider (notion-mcp-server pg-client.ts createPage) stores page
properties pass-through with NO "type" wrapper, e.g.
  {"Name": {"title": [{"text": {"content": "..."}}]}, "Student_Count": {"number": 694}}
while some representations carry a "type" discriminator, e.g.
  {"Name": {"type": "title", "title": [...]}, "Student_Count": {"type": "number", "number": 694}}
The evaluator must read both."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from main import _read_prop_value, _prop_kind, _extract_text

TITLE = "Environmental Economics & Ethics (Spring 2014)"
COUNT = 694

# Unwrapped (actual provider shape)
UNWRAPPED = {
    "Name": {"title": [{"text": {"content": TITLE}}]},
    "Student_Count": {"number": COUNT},
}

# Wrapped (with explicit type discriminator)
WRAPPED = {
    "Name": {"type": "title", "title": [{"text": {"content": TITLE}}]},
    "Student_Count": {"type": "number", "number": COUNT},
}

# rich_text variant (unwrapped)
RICH_UNWRAPPED = {"Notes": {"rich_text": [{"text": {"content": "hello world"}}]}}
# rich_text variant (wrapped)
RICH_WRAPPED = {"Notes": {"type": "rich_text", "rich_text": [{"text": {"content": "hello world"}}]}}

failures = []


def check(label, got, expected):
    ok = got == expected
    print(("PASS" if ok else "FAIL"), label, "->", repr(got), "" if ok else f"(expected {expected!r})")
    if not ok:
        failures.append(label)


# --- kind inference ---
check("prop_kind unwrapped title", _prop_kind(UNWRAPPED["Name"]), "title")
check("prop_kind wrapped title", _prop_kind(WRAPPED["Name"]), "title")
check("prop_kind unwrapped number", _prop_kind(UNWRAPPED["Student_Count"]), "number")
check("prop_kind wrapped number", _prop_kind(WRAPPED["Student_Count"]), "number")
check("prop_kind non-dict -> None", _prop_kind(None), None)
check("prop_kind only-meta -> None", _prop_kind({"id": "x", "type": "title"}), "title")

# --- title reading ---
check("read title unwrapped", _extract_text(_read_prop_value(UNWRAPPED["Name"], "title")), TITLE)
check("read title wrapped", _extract_text(_read_prop_value(WRAPPED["Name"], "title")), TITLE)
check("read title missing-kind -> None", _read_prop_value(UNWRAPPED["Name"], "number"), None)

# --- number reading ---
check("read number unwrapped", _read_prop_value(UNWRAPPED["Student_Count"], "number"), COUNT)
check("read number wrapped", _read_prop_value(WRAPPED["Student_Count"], "number"), COUNT)
check("read number missing-kind -> None", _read_prop_value(UNWRAPPED["Student_Count"], "title"), None)

# --- rich_text reading ---
check("read rich_text unwrapped",
      _extract_text(_read_prop_value(RICH_UNWRAPPED["Notes"], "rich_text")), "hello world")
check("read rich_text wrapped",
      _extract_text(_read_prop_value(RICH_WRAPPED["Notes"], "rich_text")), "hello world")

# --- end-to-end over full props dicts (mirrors evaluator loop) ---
def find_title(props):
    for _, pval in props.items():
        t = _read_prop_value(pval, "title")
        if t is not None:
            return _extract_text(t)
    return ""


def find_count(props):
    for pname, pval in props.items():
        n = _read_prop_value(pval, "number")
        if n is not None:
            return n
    return None


check("e2e title unwrapped", find_title(UNWRAPPED), TITLE)
check("e2e title wrapped", find_title(WRAPPED), TITLE)
check("e2e count unwrapped", find_count(UNWRAPPED), COUNT)
check("e2e count wrapped", find_count(WRAPPED), COUNT)

print()
if failures:
    print(f"FAILED {len(failures)}: {failures}")
    sys.exit(1)
print("All synthetic property-shape tests passed.")
