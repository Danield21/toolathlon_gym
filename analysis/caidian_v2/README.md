# Caidian v2 (P0-P4)

This directory is an additive replacement for the legacy Oracle prompt versus
trajectory scorer. The legacy chain remains read-only in `backups/` and is
never imported or overwritten by v2. `legacy_backup_manifest.json` records its
fixed location, counts, modes, and checksums in version control.

## Layout

- `scripts/caidian_v2.py`: extraction, unified IR, ordered wave alignment, and
  one-to-one node scoring.
- `scripts/gen_caidian_v2.py`: 167-task CLI and report writer.
- `config/eligibility_overrides_gt167_20260830.json`: exact-trajectory-hash-bound
  conditional-wave and runtime-manifest denominator decisions.
- `tests/test_caidian_v2.py`: regression tests for the P0-P4 failure modes.
- `guide/16_caidian_v2_methodology.md`: exact metric contract.
- `outputs/`: versioned generated artifacts; no legacy output is written here.
  Every generated report set includes `SHA256SUMS`; the machine-readable meta
  also records the scorer and generator hashes.

## Reproduce the 167-task report on the H800 cluster

```bash
cd /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/analysis/caidian_v2

python3 -m unittest discover -s tests -v

python3 scripts/gen_caidian_v2.py \
  --gt /lintaoLab2/ruanlexing/toolathlon-gym/guide/reference/ground_truth_167_20260829 \
  --dump /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/deepseek-v4-flash-0731/oracle_best_gt167_merged_20260830 \
  --source-map /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/deepseek-v4-flash-0731/oracle_best_gt167_merged_20260830/SOURCE_MAP.tsv \
  --eligibility-overrides config/eligibility_overrides_gt167_20260830.json \
  --legacy-data backups/caidian_legacy_20260830T190032+0800/comparison/caidian_gt167_caidian_data_20260828.json \
  --outdir outputs/caidian_v2_gt167_20260830 \
  --replace-output
```

The command fails closed if the GT task names and exact source manifest do not
match, if a mapped run does not contain exactly one non-stub raw stream, or if
an override's task/wave/schema/hash/evidence does not verify. Known schema-log
diagnostics are retained explicitly; malformed or provider-terminated JSONL is
marked unscorable and contributes to no metric denominator. Runtime fanout
overrides are recomputed from paired structured MCP results strictly before the
first delegation call; delegation prose and actual accepted-agent counts are
not evidence. Output is built in a sibling staging directory and published by
rename; `--replace-output` retains the previous complete output directory
instead of mixing old and new files.
