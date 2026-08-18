#!/usr/bin/env bash
# scripts/lint_gcal_evals.sh
#
# Guard against re-introducing the §C.2 "TIMESTAMPTZ read tz bug" fixed in
# dev_docs/2026-08-13-c2-tz-fix-design.md (v3, root-fix). The bug: an
# evaluator reads `start_datetime` / `end_datetime` from `gcal.events` and
# then compares the raw `.hour` / `.minute` / `.date()` / `strftime("%H:%M")`
# WITHOUT going through `utils.evaluation.gcal_helpers`. Because psycopg2
# returns the column in the PG session `TimeZone` (not necessarily UTC),
# the comparison silently compares against the wrong wall-clock.
#
# This lint enforces:
#   (1) FATAL: any *.py that reads `gcal.events` AND uses bare `.hour`/
#       `.minute` comparisons AND does NOT import `gcal_helpers` is rejected.
#   (2) FATAL: any *.py that reads `gcal.events` AND takes `.date()` /
#       `strftime("%H` from a start/end column AND does NOT use
#       `gcal_helpers` is rejected (those are also session-tz-affected:
#       `sd.date()` can shift a day, `strftime("%H:%M")` shifts the hour).
#   (3) WARN (→ later forced): any *.py that reads `gcal.events` AND does
#       NOT define an `EXPECTED_TIMEZONE` style spec is flagged. Phase to
#       FATAL once all current evaluators are migrated.
#
# Already-migrated evaluators (using `AT TIME ZONE 'UTC'`, `astimezone`,
# `SET TIME ZONE 'UTC'`, or `gcal_helpers`) are excluded from (1)/(2): the
# historical migrations are accepted as-is even though style differs; the
# move to the unified `gcal_helpers` style is tracked as WARN, not FATAL,
# so we don't break existing PASS'ing cases during the v3 migration.
#
# Usage:
#   bash scripts/lint_gcal_evals.sh            # exit 1 on FATAL; prints WARNs
#   bash scripts/lint_gcal_evals.sh --strict   # also fail on any WARN
#
# Run from the `toolathlon_gym/` dir.

set -u
STRICT=0
if [[ "${1:-}" == "--strict" ]]; then STRICT=1; fi

# Acceptable session-tz-safe patterns: any one of these exempts the file
# from the FATAL bare-comparison check. v3 prefers `gcal_helpers` but keeps
# these historical safe styles valid until a full migration sweep.
SAFE_TZ_PATTERN='gcal_helpers|get_utc_components|get_zone_components|to_aware_utc|to_aware_zone|astimezone|AT TIME ZONE|SET TIME ZONE|zoneinfo|ZoneInfo|timezone\.utc|_dt\.timezone\.utc'

# Bare-comparison patterns that — WITHOUT session-tz-safe helpers — would
# silently miscompare in a non-UTC PG session.
BARE_HOUR_PATTERN='\.hour\s*(!=|==|in|<|>|\s*$|\s*[) ](\s*and|\s*or))'
BARE_MINUTE_PATTERN='\.minute\s*(!=|==|in|<|>)'
BARE_DATE_PATTERN='\.date\(\)'
BARE_STRFTIME_PATTERN='strftime\("%[Hh]'

fatal=0
warn=0

# Find every *.py that reads gcal.events.
while IFS= read -r f; do
  # This file genuinely reads gcal.events (column or table mention).
  reads_gcal=$(grep -cE 'gcal\.events|FROM gcal\.events' "$f" 2>/dev/null || true)
  if [[ "$reads_gcal" -eq 0 ]]; then continue; fi

  uses_safe=$(grep -cE "$SAFE_TZ_PATTERN" "$f" 2>/dev/null || true)

  # FATAL (1): bare .hour/.minute comparison without any safe pattern.
  if [[ "$uses_safe" -eq 0 ]]; then
    if grep -nE "$BARE_HOUR_PATTERN" "$f" >/dev/null 2>&1 \
       || grep -nE "$BARE_MINUTE_PATTERN" "$f" >/dev/null 2>&1; then
      echo "FATAL [gcal-tz-lint] $f"
      echo "   reads gcal.events and uses bare .hour/.minute comparison"
      echo "   but does NOT use gcal_helpers or any session-tz-safe pattern."
      echo "   Fix: see dev_docs/2026-08-13-c2-tz-fix-design.md"
      fatal=$((fatal+1))
    fi

    # FATAL (2): bare .date() or strftime("%H") on these columns.
    if grep -nE "$BARE_DATE_PATTERN" "$f" >/dev/null 2>&1 \
       || grep -nE "$BARE_STRFTIME_PATTERN" "$f" >/dev/null 2>&1; then
      echo "FATAL [gcal-tz-lint] $f"
      echo "   reads gcal.events and extracts .date() / strftime('%H') raw"
      echo "   but does NOT use gcal_helpers or any session-tz-safe pattern."
      echo "   Fix: see dev_docs/2026-08-13-c2-tz-fix-design.md"
      fatal=$((fatal+1))
    fi
  fi

  # WARN (3): reads gcal.events but defines no EXPECTED_TIMEZONE spec.
  # We don't require strict naming yet — define ANY *_TIMEZONE constant or
  # a module-level comment citing the timezone from task.md. The goal: force
  # the maintainer to declare which zone the comparison anchors to.
  has_tz_spec=$(grep -cE 'EXPECTED_TIMEZONE\s*=|TIMEZONE\s*=.*ZoneInfo|TIMEZONE\s*=.*America/|TIMEZONE\s*=.*Asia/|TIMEZONE\s*=.*Europe/' "$f" 2>/dev/null || true)
  if [[ "$has_tz_spec" -eq 0 ]] && [[ "$uses_safe" -gt 0 ]]; then
    # Has safe usage but no spec—historical style. WARn only.
    echo "WARN [gcal-tz-lint] $f"
    echo "   reads gcal.events (protected) but declares no EXPECTED_TIMEZONE spec."
    echo "   v3 strongly recommends adding one; see dev_docs/2026-08-13-c2-tz-fix-design.md"
    warn=$((warn+1))
  fi
done < <(find tasks/finalpool -path '*/evaluation/*.py' -type f 2>/dev/null)

echo "-------------------------------------"
echo "gcal-tz-lint summary: $fatal FATAL / $warn WARN"

if [[ "$fatal" -gt 0 ]]; then exit 1; fi
if [[ "$STRICT" -eq 1 ]] && [[ "$warn" -gt 0 ]]; then exit 1; fi
exit 0