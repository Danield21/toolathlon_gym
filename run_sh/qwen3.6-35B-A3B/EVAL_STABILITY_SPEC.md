# Qwen3.6-35B-A3B Toolathlon-GYM Evaluation Stability Spec

Status: ready for implementation

Owner: Toolathlon evaluation harness

Scope: Enroot-based Qwen3.6-35B-A3B evaluation on `gnho019:30002`

## 1. Objective

Make the 503-case evaluation distinguish model quality failures from harness or
serving failures. A full run must not begin until model requests are bounded,
completion signaling is unambiguous, required MCP servers start reliably, and
a representative smoke test passes the infrastructure gate.

This work preserves the existing isolation model:

- one Enroot rootfs per task;
- one PostgreSQL port, data directory, and socket directory per task;
- at most five concurrent workers on `login01`;
- the existing Qwen3.6 SGLang server on `gnho019`.

## 2. Observed baseline

Run `qwen36_enroot_503_20260728-195221` showed the following after 27 completed
cases:

| Outcome | Count |
|---|---:|
| OpenAI-compatible API timeout | 16 |
| Agent step timeout at 1200 seconds | 2 |
| Empty model response | 4 |
| Text response without `claim_done` | 2 |
| `claim_done`, evaluator failed | 3 |
| Evaluator passed | 0 |

The PostgreSQL identity checks passed and no port or data-directory collision
was found. The dominant failures are therefore above the Enroot/PostgreSQL
isolation layer.

Confirmed causes:

1. SGLang enables Qwen thinking by default.
2. The client sends no `max_tokens` value.
3. CAMEL uses a 180-second model timeout with three retries, allowing one bad
   request to occupy roughly 12 minutes.
4. The task prompt permits plain text as completion, while `TaskAgent` accepts
   only a `claim_done` tool call.
5. The PG-backed email server rejects an existing task email config when the
   config has no password, although PG mode does not use that password.
6. CAMEL does not recognize the model context size and assumes an effectively
   unlimited value. SGLang then silently truncates at its configured context
   limit (131,072 tokens for this deployment).

## 3. Required behavior

### 3.1 Bounded model requests

The evaluation configuration must support and export these variables:

| Variable | Default | Meaning |
|---|---:|---|
| `MODEL_MAX_TOKENS` | `4096` | Maximum generated tokens per model response |
| `MODEL_TIMEOUT` | `300` | Timeout in seconds for one OpenAI-compatible attempt |
| `MODEL_MAX_RETRIES` | `1` | SDK retries after the initial attempt |
| `MODEL_ENABLE_THINKING` | `0` | Disable Qwen thinking for tool-use evaluation |

`utils/api_model/model_provider.py` must:

1. Parse and validate the four variables.
2. Put `max_tokens` into `model_config_dict`.
3. For OpenAI-compatible Qwen requests, merge the following into
   `model_config_dict` without overwriting unrelated caller settings:

   ```python
   {
       "extra_body": {
           "chat_template_kwargs": {
               "enable_thinking": False
           }
       }
   }
   ```

4. Pass `timeout=MODEL_TIMEOUT` and `max_retries=MODEL_MAX_RETRIES` to
   `ModelFactory.create`.
5. Log the effective model request configuration once per task, excluding API
   keys.
6. Reject invalid integer, boolean, or non-positive values before MCP startup.

Thinking must be controlled per request. Changing this default must not require
restarting SGLang.

### 3.2 Explicit context budget

The agent must not rely on CAMEL's unknown-model fallback or SGLang's automatic
truncation.

Add these variables:

| Variable | Default | Meaning |
|---|---:|---|
| `AGENT_TOKEN_LIMIT` | `122880` | Maximum conversation tokens retained by CAMEL |
| `AGENT_STEP_TIMEOUT` | `1200` | Whole-agent-step timeout in seconds |
| `AGENT_TOOL_TIMEOUT` | `120` | One tool execution timeout in seconds |
| `AGENT_RATE_LIMIT_RETRIES` | `5` | Bounded CAMEL rate-limit retries |

`utils/roles/task_agent.py` must pass these values to `ChatAgent`. The explicit
token limit leaves 8,192 tokens below the 131,072-token server context window,
which covers the 4,096-token response budget and protocol overhead.

Context summarization remains disabled for benchmark reproducibility. If the
explicit token budget is exceeded, the task must terminate with a classified
`context_limit` failure instead of depending on silent server truncation.

After this guard is verified, `--allow-auto-truncate` must be removed from the
Qwen SGLang launch command, or be disabled by default behind
`SGLANG_ALLOW_AUTO_TRUNCATE=0`. This serving change requires an intentional
SGLang restart and must be performed after the request-side changes are tested.

### 3.3 Unambiguous completion protocol

`claim_done` is the only successful completion signal.

Before constructing `ChatAgent`, `TaskAgent` must normalize task system prompts:

- replace `local-claim_done` with `claim_done`;
- remove language saying that a response without a tool call counts as
  completion;
- append this invariant if it is not already present:

  > A plain-text response never marks the task complete. After all required
  > work is finished, you must call the `claim_done` tool. Do not call it while
  > any required action is incomplete.

The runner must not reinterpret arbitrary completion-like text as success. If
the model returns no messages, or returns text without `claim_done`, the task
must end with a protocol-specific failure reason rather than a generic failure.

Required failure reasons:

- `empty_model_response`
- `missing_claim_done`
- `max_iterations`
- `context_limit`
- `model_timeout`
- `agent_step_timeout`

### 3.4 PG-only Email MCP configuration

The email MCP server in this evaluation always uses PostgreSQL-backed IMAP and
SMTP implementations. Passwords for real external mail servers are therefore
not required.

Change the email config loading API to distinguish PG mode from real-mail mode:

```python
load_email_config(config_file, require_password: bool = True)
```

The PG-backed server must call it with `require_password=False`. In this mode:

- `email` is retained from the task config when present;
- `name` is retained when present;
- a missing password becomes the internal sentinel value `unused`;
- malformed JSON and a missing email remain startup errors;
- no real IMAP or SMTP connection is attempted.

The case `arxiv-reading-plan-excel-gcal-email` is the regression test because
its config intentionally contains an email identity but no password.

### 3.5 Enroot runtime refresh

The Enroot template is a snapshot. Every worker must receive all changed
runtime files before `enroot start`; otherwise host fixes will not affect the
evaluation.

`run_eval_parallel.sh` must refresh at least:

- `utils/api_model/model_provider.py`
- `utils/roles/task_agent.py`
- `utils/mcp/tool_servers.py`
- `local_servers/emails-mcp/src/emails_mcp/server.py`
- `local_servers/emails-mcp/src/emails_mcp/config/settings.py`

The destination for email files is under the worker rootfs
`/opt/local_servers/emails-mcp/src/emails_mcp/`. Any failed copy is a
`rootfs_fail` and the worker must not start.

All model and agent variables defined in this spec must be explicitly included
in the Enroot `ENV_ARGS` list. The launcher must print their effective values,
excluding secrets.

### 3.6 Failure classification and summaries

Extend `summary_parallel_<run_id>.csv` with a final `failure_reason` column:

```text
task,status,exit_code,output_dir,pg_port,duration_s,failure_reason
```

Allowed values:

- empty for a successful evaluator result;
- `evaluation_failed` when the agent called `claim_done` but evaluation failed;
- the agent reasons listed in section 3.3;
- `mcp_startup_failed`;
- `pg_fail`;
- `rootfs_fail`;
- `runner_error` for an otherwise unclassified harness exception.

Classification must be emitted directly by the runner where possible. Log
text matching is permitted only as a compatibility fallback.

The case study report must separately show:

1. infrastructure/protocol failures;
2. cases that reached the evaluator;
3. evaluator pass/fail counts.

This prevents evaluator failures from being mixed with invalid runs.

## 4. Validation plan

### 4.1 Unit and configuration checks

Required tests:

1. `model_provider` produces `max_tokens=4096` and
   `enable_thinking=False` with defaults.
2. Boolean parsing accepts `0/1`, `true/false`, and rejects unknown values.
3. Existing `extra_body` keys survive the thinking configuration merge.
4. Invalid timeout, retry, token, and context values fail before task startup.
5. PG email config without password loads with password `unused`.
6. Real-mail config loading still requires a password when
   `require_password=True`.
7. A task prompt containing the legacy plain-response completion sentence is
   normalized to require `claim_done`.

### 4.2 PostgreSQL isolation regression

Run the existing two-worker PG-only test. It must produce two
`pg_test_success` rows and exactly one isolation marker per database.

```bash
MAX_CONCURRENT=2 PG_TEST_ONLY=1 \
  bash run_sh/qwen3.6-35B-A3B/run_eval_parallel.sh \
  12306-beijing-shanghai-trip-notion-gcal-word \
  12306-canvas-fieldtrip-gcal-word-email
```

### 4.3 Five-case full-stack smoke

Run these cases first at concurrency 1, then repeat at concurrency 5:

```text
12306-howtocook-team-trip-catering-excel-gcal
academic-conference-planning-coordination
academic-literature-review
arxiv-reading-plan-excel-gcal-email
arxiv-research-landscape-report
```

The smoke infrastructure gate passes only if:

- 0 model API timeouts;
- 0 agent step timeouts;
- 0 empty model responses;
- 0 Email MCP startup failures;
- 0 PostgreSQL identity, port, or startup failures;
- all five cases call `claim_done` and reach their evaluator;
- no request is observed decoding until client deletion;
- each task records a non-empty, valid failure classification if its evaluator
  does not pass.

Evaluator pass rate is reported but is not an infrastructure gate, because it
measures model quality.

### 4.4 Twenty-case canary

After both smoke runs pass, run a fixed 20-case canary at concurrency 5. The
503-case run may start only when:

- there are no model, step, MCP-startup, PostgreSQL, or rootfs failures;
- at least 19 of 20 cases reach the evaluator;
- P95 task duration is below 15 minutes;
- no task exceeds 20 minutes;
- evaluator results and infrastructure failures are reported separately.

## 5. Full-run guardrails

The full run uses `MAX_CONCURRENT=5` only after the canary gate passes.

After the first 20 completed cases, the orchestrator must stop scheduling new
cases if either condition is true:

- any infrastructure failure category exceeds 2 cases;
- fewer than 18 cases reached the evaluator.

Already-running workers may finish and clean up normally. The orchestrator must
write `circuit_breaker_triggered` and the triggering counts to its log and exit
non-zero after cleanup.

No automatic retry of an entire case is allowed in the scored run. A retried
case must go to a separate dump root and be labeled as diagnostic.

## 6. Rollout sequence

1. Stop the currently invalid 503-case run and preserve its dumps as a failed
   diagnostic run.
2. Implement sections 3.1 through 3.6.
3. Run unit/configuration checks.
4. Run the PostgreSQL isolation regression.
5. Run five cases at concurrency 1.
6. Run the same five cases at concurrency 5.
7. Run the fixed 20-case canary.
8. Restart SGLang without automatic truncation and repeat the five-case smoke.
9. Start a fresh 503-case evaluation in a new dump directory.

The old dump directory must never be reused or merged into the new scored run.

## 7. Definition of done

This spec is complete when:

- all required code and launcher changes are implemented;
- unit/configuration checks pass;
- both five-case smoke runs pass the infrastructure gate;
- the 20-case canary meets its gate;
- the PostgreSQL isolation regression remains clean;
- effective request settings are visible in every task log;
- the full-run summary separates harness failures from evaluator failures;
- a fresh 503-case run can be started without reusing invalid results.
