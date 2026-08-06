"""Convert kimi-code stream-json output into a Toolathlon-style traj dict.

The converter is deliberately defensive: it keeps every raw event under
"raw_events" and extracts best-effort messages / tool_calls for analysis.
"""
import json
import os


def convert_stream(stream_path: str) -> dict:
    messages = []
    tool_calls = []
    raw_events = []

    if not os.path.exists(stream_path):
        return {"messages": [], "tool_calls": [], "raw_events": [],
                "warning": f"stream file missing: {stream_path}"}

    with open(stream_path, encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                ev = json.loads(line)
            except json.JSONDecodeError:
                raw_events.append({"_non_json": line[:2000]})
                continue
            raw_events.append(ev)
            if not isinstance(ev, dict):
                continue
            etype = str(ev.get("type", ""))
            if "message" in etype or ev.get("role"):
                msg = ev.get("message", ev)
                if isinstance(msg, dict):
                    messages.append({
                        "role": msg.get("role", ""),
                        "content": msg.get("content", ""),
                    })
            if "tool" in etype:
                tool_calls.append(ev)

    return {"messages": messages, "tool_calls": tool_calls,
            "raw_events": raw_events}
