"""Unit tests for ``utils.evaluation.gcal_helpers``.

These tests are the engineering contract for the root-fix in
``dev_docs/2026-08-13-c2-tz-fix-design.md`` (v3): the helper output MUST be
independent of the PG session ``TimeZone`` that psycopg2 used to display the
``TIMESTAMPTZ`` column. If any of these break, the root fix is broken.

Run from the ``toolathlon_gym`` dir with::

    python -m pytest utils/evaluation/tests/test_gcal_helpers.py -v
"""

from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo

import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..")))

from utils.evaluation.gcal_helpers import (
    get_utc_components,
    get_zone_components,
    to_aware_utc,
    to_aware_zone,
)


# ---------- sessiontz-independence (the root-fix proof) ----------

# A real Calendar MCP write: "8 AM ET" (EDT, UTC-4) was converted to UTC by
# toUtcIso and stored as 12:00Z. This is the instant the DB stores.
EV_8AM_EDT = datetime(2013, 10, 13, 12, 0, 0, tzinfo=timezone.utc)

# "8 AM ET" winter: EST UTC-5 -> 13:00Z (DST rolled back).
EV_8AM_EST = datetime(2014, 1, 5, 13, 0, 0, tzinfo=timezone.utc)

# "2 PM UTC": stored as 14:00Z (no local-tz conversion needed).
EV_2PM_UTC = datetime(2026, 3, 9, 14, 0, 0, tzinfo=timezone.utc)


def test_get_utc_components_invariant_across_pg_sessions():
    """get_utc_components must return UTC instant computation,
    regardless of how psycopg2 displayed the TIMESTAMPTZ.

    This is the literal root-fix: if it ever returns session-tz-tied output
    again, the bug is back. We cover 6 plausible session tzs.
    """
    for session_tz in ['UTC', 'Asia/Shanghai', 'America/New_York',
                      'America/Los_Angeles', 'Europe/London', 'Asia/Tokyo']:
        # Simulate psycopg2 returning the same instant in this session tz.
        sd = EV_8AM_EDT.astimezone(ZoneInfo(session_tz))
        assert get_utc_components(sd) == (date(2013, 10, 13), 12, 0), (
            f"UTC helper broken in PG session tz={session_tz}: "
            f"got {get_utc_components(sd)}"
        )


def test_get_zone_components_et_invariant_across_pg_sessions():
    """get_zone_components(ET) must return ET 8 AM, regardless of
    how psycopg2 displayed the TIMESTAMPTZ (same proof, ET side)."""
    for session_tz in ['UTC', 'Asia/Shanghai', 'America/New_York',
                      'America/Los_Angeles', 'Europe/London', 'Asia/Tokyo']:
        sd = EV_8AM_EDT.astimezone(ZoneInfo(session_tz))
        assert get_zone_components(sd, 'America/New_York') == (
            date(2013, 10, 13), 8, 0,
        ), (
            f"ET helper broken in PG session tz={session_tz}: "
            f"got {get_zone_components(sd, 'America/New_York')}"
        )


def test_get_zone_components_pt_invariant_across_pg_sessions():
    """PT side of the same proof (Pacific Time)."""
    # Use a winter date (no DST ambiguity near the boundary): EV_8AM_EST is
    # 2014-01-05 13:00Z. PST (winter) = UTC-8, so 13:00Z = 05:00 PT.
    for session_tz in ['UTC', 'Asia/Shanghai', 'America/New_York',
                      'America/Los_Angeles', 'Europe/London']:
        sd = EV_8AM_EST.astimezone(ZoneInfo(session_tz))
        assert get_zone_components(sd, 'America/Los_Angeles') == (
            date(2014, 1, 5), 5, 0,
        )


def test_get_zone_components_cn_invariant_across_pg_sessions():
    """Asia/Shanghai side of the same proof."""
    for session_tz in ['UTC', 'Asia/Shanghai', 'America/New_York',
                      'America/Los_Angeles']:
        # UTC 12:00 -> Beijing 20:00 (CST UTC+8, no DST)
        sd = EV_8AM_EDT.astimezone(ZoneInfo(session_tz))
        assert get_zone_components(sd, 'Asia/Shanghai') == (
            date(2013, 10, 13), 20, 0,
        )


# ---------- DST awareness ----------

def test_dst_awareness_winter_est():
    """Winter EST: UTC-5. 13:00Z = 08:00 ET."""
    sd = EV_8AM_EST.astimezone(ZoneInfo('Asia/Shanghai'))   # 21:00+08
    assert get_zone_components(sd, 'America/New_York') == (
        date(2014, 1, 5), 8, 0,
    ), "ET helper should be DST-aware (winter EST)"


def test_dst_awareness_summer_edt():
    """Summer EDT: UTC-4. 12:00Z = 08:00 ET."""
    sd = EV_8AM_EDT.astimezone(ZoneInfo('Asia/Shanghai'))   # 20:00+08
    assert get_zone_components(sd, 'America/New_York') == (
        date(2013, 10, 13), 8, 0,
    )


# ---------- string inputs (defensively support) ----------

def test_string_input_iso_with_z():
    assert get_utc_components("2013-10-13T12:00:00Z") == (
        date(2013, 10, 13), 12, 0,
    )
    assert get_zone_components("2013-10-13T12:00:00Z", "America/New_York") == (
        date(2013, 10, 13), 8, 0,
    )


def test_string_input_iso_with_offset():
    assert get_utc_components("2013-10-13T20:00:00+08:00") == (
        date(2013, 10, 13), 12, 0,
    )
    assert get_zone_components("2013-10-13T20:00:00+08:00", "America/New_York") == (
        date(2013, 10, 13), 8, 0,
    )


def test_string_input_naive_fallback():
    """Naive ISO string is interpreted as UTC (per _to_aware)."""
    assert get_utc_components("2013-10-13 12:00:00") == (
        date(2013, 10, 13), 12, 0,
    )


# ---------- None / edge cases ----------

def test_none_input():
    assert get_utc_components(None) == (None, None, None)
    assert get_zone_components(None, "America/New_York") == (None, None, None)
    assert to_aware_utc(None) is None
    assert to_aware_zone(None, "America/New_York") is None


def test_empty_string():
    assert get_utc_components("") == (None, None, None)


def test_naive_datetime_assumed_utc():
    naive = datetime(2013, 10, 13, 12, 0, 0)
    # Naive datetimes are assumed UTC.
    assert get_utc_components(naive) == (date(2013, 10, 13), 12, 0)
    assert get_zone_components(naive, "America/New_York") == (
        date(2013, 10, 13), 8, 0,
    )


def test_zone_accepts_string_or_zoneinfo():
    """helper accepts zone as str or ZoneInfo instance."""
    sd = EV_8AM_EDT.astimezone(ZoneInfo('Asia/Shanghai'))
    s_date, s_h, s_m = get_zone_components(sd, "America/New_York")
    z_date, z_h, z_m = get_zone_components(sd, ZoneInfo("America/New_York"))
    assert (s_date, s_h, s_m) == (z_date, z_h, z_m)


# ---------- to_aware_* variants ----------

def test_to_aware_utc_returns_full_datetime():
    sd = EV_8AM_EDT.astimezone(ZoneInfo('Asia/Shanghai'))
    result = to_aware_utc(sd)
    assert result is not None
    assert result == datetime(2013, 10, 13, 12, 0, 0, tzinfo=timezone.utc)


def test_to_aware_zone_returns_full_datetime_et():
    sd = EV_8AM_EDT.astimezone(ZoneInfo('Asia/Shanghai'))
    result = to_aware_zone(sd, "America/New_York")
    assert result is not None
    assert result.tzinfo == ZoneInfo("America/New_York")
    assert (result.date(), result.hour, result.minute) == (
        date(2013, 10, 13), 8, 0,
    )


if __name__ == "__main__":
    import pytest
    sys.exit(pytest.main([__file__, "-v"]))