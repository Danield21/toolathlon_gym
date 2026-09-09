#!/bin/bash
# Merge k3 par47 + needsub dumps into dumps/kimi-k3/auto-subagent
# with the same 167-case split as deepseek-v4-flash-0731.
set -euo pipefail

DUMPS="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps"
DEST="${DUMPS}/kimi-k3"
MODE="auto-subagent"
SPLIT="/lintaoLab2/bowending/project_agent_swarm_benchmark/dev_docs/split_167cases"
SRC_PAR47="${DUMPS}/kimi-code_k3-par47-noeffort"
SRC_NEED="${DUMPS}/kimi-code_k3-needsub-noeffort"
LOG="${DEST}/reorg_auto-subagent_$(date +%Y%m%d-%H%M%S).log"

mkdir -p \
  "${DEST}/${MODE}/no_subagent" \
  "${DEST}/${MODE}/need_subagent/need_subagent_parallel" \
  "${DEST}/${MODE}/need_subagent/need_subagent_no_parallel" \
  "${DEST}/${MODE}/_unlisted"
exec > >(tee -a "$LOG") 2>&1

echo "=== reorg start $(date) host=$(hostname) nproc=$(nproc) ==="
for s in "$SRC_PAR47" "$SRC_NEED"; do
  if [[ ! -d "$s" ]]; then
    echo "FATAL: source missing: $s" >&2
    exit 1
  fi
  echo "[src] $s"
done

rsync_one() {
  local tag="$1" src="$2"
  local incoming="${DEST}/${MODE}/_incoming_${tag}"
  mkdir -p "$incoming"
  echo "[rsync] START ${tag} $(date)"
  rsync -aW --no-compress --human-readable --info=stats2 \
    "${src}/" "${incoming}/"
  echo "[rsync] DONE  ${tag} $(date)"
}

rsync_one par47 "$SRC_PAR47" &
pid1=$!
rsync_one needsub "$SRC_NEED" &
pid2=$!
wait "$pid1"
wait "$pid2"

echo "=== merge + split $(date) ==="
python3 - "$DEST" "$MODE" "$SPLIT" <<'PY'
from pathlib import Path
import shutil
import sys

dest = Path(sys.argv[1])
mode = sys.argv[2]
split = Path(sys.argv[3])
root = dest / mode
inc_a = root / "_incoming_par47"
inc_b = root / "_incoming_needsub"
incoming = root / "_incoming"
if incoming.exists():
    raise SystemExit(f"stale incoming {incoming}")
incoming.mkdir()

def load(name):
    rows = [ln.strip() for ln in (split / name).read_text().splitlines() if ln.strip()]
    return set(rows)

nos, par, ser = load("no_subagent.txt"), load("need_subagent_parallel.txt"), load("need_subagent_no_parallel.txt")
assert len(nos) == 33 and len(par) == 131 and len(ser) == 3
assert not (nos & par or nos & ser or par & ser)

def is_task_dir(p: Path) -> bool:
    return p.is_dir() and p.name not in {
        "no_subagent", "need_subagent", "_unlisted", "_incoming",
        "_incoming_par47", "_incoming_needsub",
    }

def bucket(name: str) -> str:
    if name in nos:
        return "no_subagent"
    if name in par:
        return "need_subagent/need_subagent_parallel"
    if name in ser:
        return "need_subagent/need_subagent_no_parallel"
    return "_unlisted"

def merge_tree(src: Path, tag: str) -> tuple[set[str], list[str]]:
    tasks = set()
    notes = []
    for child in sorted(src.iterdir()):
        target = incoming / child.name
        if is_task_dir(child):
            tasks.add(child.name)
            if target.exists():
                # merge slot dirs; refuse identical slot names
                for slot in child.iterdir():
                    slot_t = target / slot.name
                    if slot_t.exists():
                        raise SystemExit(f"slot collision {child.name}/{slot.name} from {tag}")
                    shutil.move(str(slot), str(slot_t))
                if any(child.iterdir()):
                    raise SystemExit(f"leftover in {child}")
                child.rmdir()
                notes.append(f"merged-task {child.name} from {tag}")
            else:
                shutil.move(str(child), str(target))
        else:
            if target.exists():
                renamed = incoming / f"{tag}__{child.name}"
                if renamed.exists():
                    raise SystemExit(f"root file collision {renamed}")
                shutil.move(str(child), str(renamed))
                notes.append(f"renamed-root {child.name} -> {renamed.name}")
            else:
                shutil.move(str(child), str(target))
    leftover = [p.name for p in src.iterdir()]
    if leftover:
        raise SystemExit(f"{tag} incoming not empty: {leftover}")
    src.rmdir()
    return tasks, notes

tasks_a, notes_a = merge_tree(inc_a, "par47")
tasks_b, notes_b = merge_tree(inc_b, "needsub")
overlap = sorted(tasks_a & tasks_b)
print(f"[merge] par47_tasks={len(tasks_a)} needsub_tasks={len(tasks_b)} overlap={len(overlap)}")
if overlap:
    print("[merge] overlap:", ", ".join(overlap))
for n in notes_a + notes_b:
    print("[merge]", n)

counts = {
    "no_subagent": 0,
    "need_subagent/need_subagent_parallel": 0,
    "need_subagent/need_subagent_no_parallel": 0,
    "_unlisted": 0,
}
for child in sorted(incoming.iterdir()):
    if is_task_dir(child):
        rel = bucket(child.name)
        target = root / rel / child.name
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            raise SystemExit(f"target exists: {target}")
        shutil.move(str(child), str(target))
        counts[rel] += 1
    else:
        target = root / child.name
        if target.exists():
            raise SystemExit(f"target exists: {target}")
        shutil.move(str(child), str(target))

leftover = [p.name for p in incoming.iterdir()]
if leftover:
    raise SystemExit(f"incoming not empty: {leftover}")
incoming.rmdir()
print(f"[split] {mode} {counts} listed={counts['no_subagent']+counts['need_subagent/need_subagent_parallel']+counts['need_subagent/need_subagent_no_parallel']}")
PY

echo "=== verify $(date) ==="
python3 - "$DEST" "$MODE" <<'PY'
from pathlib import Path
import sys
root = Path(sys.argv[1]) / sys.argv[2]

def ndirs(p):
    return len([x for x in p.iterdir() if x.is_dir()]) if p.is_dir() else 0

n_nos = ndirs(root / "no_subagent")
n_par = ndirs(root / "need_subagent/need_subagent_parallel")
n_ser = ndirs(root / "need_subagent/need_subagent_no_parallel")
n_unl = ndirs(root / "_unlisted")
print(f"[verify] no_subagent={n_nos} parallel={n_par} serial={n_ser} unlisted={n_unl} listed={n_nos+n_par+n_ser}")
if n_nos + n_par + n_ser == 0:
    raise SystemExit("no listed tasks")
print("[verify] ok")
PY

echo "[rm] $SRC_PAR47"
rm -rf --one-file-system "$SRC_PAR47"
echo "[rm] $SRC_NEED"
rm -rf --one-file-system "$SRC_NEED"

echo "=== dest tree (depth 3) ==="
find "${DEST}/${MODE}" -mindepth 1 -maxdepth 3 \( -type d -o -name '*.csv' -o -name '*.log' \) | sort | head -n 160
echo "=== reorg done $(date) ==="
