"""
Helpers for evaluating Google Calendar events stored in the ``gcal.events``
table. See ``dev_docs/2026-08-13-c2-tz-fix-design.md`` (v3, root-fix) for
background.

Key contract (do not change without reading the design doc):
- ``start_datetime`` / ``end_datetime`` in ``gcal.events`` is a
  ``TIMESTAMPTZ`` column storing the **absolute UTC instant** of the event
  (the Calendar MCP writes UTC via ``toUtcIso`` after the §C.1 fix).
- psycopg2 returns this column as a tz-aware python ``datetime`` whose
  *display* tz follows the PG session ``TimeZone`` setting (case-study
  2026-08-13: compute node default was ``Asia/Shanghai``, so a UTC ``12:00``
  came back as ``20:00+08`` and an evaluator's naive ``sd.hour in {12, 13}``
  silently compared against 20), systematically miscalibrating ~48 gcal
  evaluators across the benchmark.
- The underlying instant is NOT lost: any
  ``astimezone(target_zone)`` on a tz-aware datetime is **session-tz
  independent** because it only trusts the absolute instant.

Therefore: evaluators MUST go through one of the helpers in this module,
NEVER use bare ``sd.hour`` / ``sd.minute`` / ``sd.date()`` /
``strftime('%H:%M')`` on a row read from ``gcal.events``. The lint in
``scripts/lint_gcal_evals.sh`` rejects any stray bare usage so future
evaluators cannot re-introduce the bug.

Two helpers cover the two legitimate "which zone are we comparing against"
semantics seen across the affected evaluators:

- ``get_utc_components(start_dt)``
    Returns ``(date, hour, minute)`` of the event **expressed in UTC**.
    Use when the evaluator explicitly compares the event against a UTC
    target (e.g. task.md says "...UTC"). For an event stored as
    ``08:00 UTC`` you get ``(date, 8, 0)`` regardless of the PG session tz.

- ``get_zone_components(start_dt, zone)``
    Returns ``(date, hour, minute)`` of the event **expressed in ``zone``**.
    ``zone`` may be a str (e.g. ``"America/New_York"``) or a
    ``zoneinfo.ZoneInfo`` instance. DST-aware via zoneinfo.
    Use when task.md declares the event in a specific local timezone
    (e.g. "8 AM Eastern Time (ET) ... use America/New_York", or howtocook
    Chinese recipe events in Asia/Shanghai). For an event stored as
    ``12:00Z`` (= 08:00 EDT) you get ``(date, 8, 0)`` in America/New_York.

Both helpers are session-tz-independent. They work on any compute node
regardless of the PG ``TimeZone`` setting (case-study 2026-08-13 proof:
6 simulated session tz's all produce identical helper output, see
``utils/evaluation/tests/test_gcal_helpers.py``).
"""

from __future__ import annotations

from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from typing import Optional, Tuple, Union

__all__ = [
    "get_utc_components",
    "get_zone_components",
    "to_aware_utc",
    "to_aware_zone",
]

# A zone may be a str ("America/New_York") or a ZoneInfo instance.
ZoneSpec = Union[str, "ZoneInfo"]

# Returned tuple: (date, hour, minute). Each may be None on None input.
Components = Tuple[Optional["datetime.date"], Optional[int], Optional[int]]


def _to_aware(start_dt) -> Optional[datetime]:
    """Coerce a value from psycopg2 / a JSON string / a naive datetime to a
    tz-aware ``datetime``.

    - psycopg2 returns tz-aware datetimes in the PG session tz — pass through.
    - ISO 8601 strings (a rare fallback used by some evaluators that store
      the column as text before parsing) are parsed via ``fromisoformat``
      (with trailing ``Z`` promoted to ``+00:00``).
    - Naive datetimes (should never happen for ``TIMESTAMPTZ`` but, defensively,
      some evaluators may strip tzinfo before reaching us) are assumed UTC
      rather than local time. This is the only safe default: the column is
      ``TIMESTAMPTZ`` so a naive datetime must have been an already-UTC
      instant that lost its tz marker.
    """
    if start_dt is None:
        return None
    if isinstance(start_dt, str):
        s = start_dt.strip()
        if not s:
            return None
        # ISO 8601: tolerate trailing 'Z' (psycopg2/PG use 'Z' for UTC).
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        parsed: "Optional[datetime]" = None
        try:
            parsed = datetime.fromisoformat(s)
        except ValueError:
            # Last resort: try a few common formats.
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
                try:
                    parsed = datetime.strptime(s, fmt)
                    break
                except ValueError:
                    continue
        if parsed is None:
            return None
        # fromisoformat/strptime return a NAIVE datetime when the string has
        # no offset. Per the contract (TIMESTAMPTZ column → any string we
        # receive is an already-UTC instant), assume UTC rather than local.
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed
    if start_dt.tzinfo is None:
        return start_dt.replace(tzinfo=timezone.utc)
    return start_dt


def to_aware_utc(start_dt) -> Optional[datetime]:
    """Return the event instant as a tz-aware ``datetime`` in UTC.

    Use this when you need the full datetime object (not just hour/minute),
    e.g. for arithmetic against a UTC GT instant. session-tz-independent.
    """
    sd = _to_aware(start_dt)
    if sd is None:
        return None
    return sd.astimezone(timezone.utc)


def to_aware_zone(start_dt, zone: ZoneSpec) -> Optional[datetime]:
    """Return the event instant as a tz-aware ``datetime`` in ``zone``.

    Use this when you need the full datetime object expressed in the
    task.md-declared local timezone. DST-aware. session-tz-independent.
    """
    if isinstance(zone, str):
        zone = ZoneInfo(zone)
    sd = _to_aware(start_dt)
    if sd is None:
        return None
    return sd.astimezone(zone)


def get_utc_components(start_dt) -> Components:
    """Return ``(date, hour, minute)`` of the event expressed in UTC.

    Use this when the evaluator compares the event time against a UTC target
    (e.g. task.md says the event is at "... UTC"). session-tz-independent.
    """
    sd = to_aware_utc(start_dt)
    if sd is None:
        return None, None, None
    return sd.date(), sd.hour, sd.minute


def get_zone_components(start_dt, zone: ZoneSpec) -> Components:
    """Return ``(date, hour, minute)`` of the event expressed in ``zone``.

    Use this when the evaluator compares the event time against a *local*
    wall-clock target declared by task.md (e.g. "8 AM Eastern Time",
    "9:00 AM" in an SF company). DST-aware via zoneinfo.
    session-tz-independent.

    ``zone`` may be a str (e.g. ``"America/New_York"``) or a
    ``zoneinfo.ZoneInfo`` instance.
    """
    sd = to_aware_zone(start_dt, zone)
    if sd is None:
        return None, None, None
    return sd.date(), sd.hour, sd.minute