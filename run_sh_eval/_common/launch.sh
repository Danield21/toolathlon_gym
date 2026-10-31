#!/usr/bin/env bash
set +x
set -euo pipefail
umask 077
# Internal dispatcher, called by MODEL/VARIANT/run.sh.
HERE="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
usage() {
  echo 'Usage (from a model/variant directory): bash run.sh [--backend docker|enroot] [--config FILE] [runner options]'
  echo 'The default backend is enroot. Use --backend docker with the MS-Eval Docker images.'
}
if (($# == 0)); then usage >&2; exit 2; fi
if [[ "$1" == --help || "$1" == -h ]]; then usage; exit 0; fi
PRESET="$1"
shift
[[ "$PRESET" =~ ^[A-Za-z0-9_.-]+/(sa|spec|gen|oracle|plan1st)$ ]] || {
  echo 'Invalid preset; expected MODEL/VARIANT' >&2; exit 2;
}
[[ -f "$HERE/$PRESET/launch.py" ]] || { echo "Unknown preset: $PRESET" >&2; exit 2; }
BACKEND=enroot
CONFIG=""
ARGS=()
while (($#)); do
  case "$1" in
    --backend)
      [[ $# -ge 2 ]] || { echo '--backend requires docker or enroot' >&2; exit 2; }
      BACKEND="$2"; shift 2 ;;
    --backend=*) BACKEND="${1#*=}"; shift ;;
    --config)
      [[ $# -ge 2 && -n "$2" ]] || { echo '--config requires a path' >&2; exit 2; }
      CONFIG="$2"; shift 2 ;;
    *) ARGS+=("$1"); shift ;;
  esac
done
case "$BACKEND" in docker|enroot) ;; *) echo 'Backend must be docker or enroot' >&2; exit 2;; esac
if [[ -z "$CONFIG" ]]; then
  CANDIDATE="$HERE/$PRESET/config.env"
  [[ ! -f "$CANDIDATE" ]] || CONFIG="$CANDIDATE"
fi
if [[ -n "$CONFIG" ]]; then
  [[ -f "$CONFIG" ]] || { echo "Missing config: $CONFIG" >&2; exit 2; }
  set -a
  source "$CONFIG"
  set +a
fi
for arg in "${ARGS[@]}"; do
  if [[ "$arg" == --help || "$arg" == -h ]]; then usage; break; fi
done
if [[ "$BACKEND" == docker ]]; then
  exec "${PYTHON_BIN:-python3}" "$HERE/docker/runner.py" "$PRESET" "${ARGS[@]}"
fi
exec "${PYTHON_BIN:-python3}" "$HERE/$PRESET/launch.py" "${ARGS[@]}"
