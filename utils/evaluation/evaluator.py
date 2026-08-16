from typing import Dict, Any, List, Optional
from utils.roles.task_agent import TaskStatus
from utils.data_structures.task_config import TaskConfig
from utils.general.helper import run_command, read_json, write_json
import logging
import os
from pathlib import Path

# Infrastructure-only directories the harness/framework writes inside the
# agent workspace during a run. These are NOT task deliverables and must never
# be counted as visible artifacts or interfere with evaluation:
#   .tool_results           — kimi-code oversized tool-output offload (P0-1 fix)
#   .overlong_tool_outputs  — local_tools_server save_overlong_output
#   .python_tmp             — python_execute scratch scripts
# All three start with '.', so the hidden-file filter below already excludes
# them; this set documents the contract explicitly for future scanners.
INFRA_WORKSPACE_DIRS = frozenset({".tool_results", ".overlong_tool_outputs", ".python_tmp"})

class TaskEvaluator:
    """Task evaluator"""

    @staticmethod
    def _has_visible_workspace_artifacts(agent_workspace: str) -> bool:
        workspace = Path(agent_workspace)
        if not workspace.is_dir():
            return False
        for path in workspace.rglob("*"):
            if not path.is_file():
                continue
            rel_parts = path.relative_to(workspace).parts
            # Hidden files/dirs (including INFRA_WORKSPACE_DIRS) and bytecode
            # caches are framework/runtime byproducts, not task deliverables.
            if any(part.startswith(".") or part == "__pycache__" for part in rel_parts):
                continue
            return True
        return False
    
    @staticmethod
    async def evaluate_one(dump_line: Dict[str, Any]) -> Dict[str, Any]:
        """
        Single task evaluation
        Expected content to be checked:
        - user response: check all outputs from user side
        - response: check all outputs from llm
        - tool calls: check all tool calls from llm
        - tool outputs: check all tool outputs
        ====== The following checks need to be started from config ======
        - local status: check files in specific workspace directory (e.g. saved some things, modified some things etc)
        - remote status: manually call MCP server to check if remote status is normally modified [not sure if possible]
        Use the above content to determine whether the task execution is successful or not
        """
        task_config = TaskConfig.from_dict(dump_line['config'])
        task_status = dump_line['status']
        # Prepare information for evaluation
        res_log_file = task_config.log_file
        agent_workspace = task_config.agent_workspace
        groundtruth_workspace = task_config.evaluation.groundtruth_workspace
        eval_command = task_config.evaluation.evaluation_command
        launch_time = task_config.launch_time
        print(f"launch time in eval is {launch_time}")

        # Official completion still requires claim_done/SUCCESS. However, for
        # real-time evaluation we must not discard already-produced artifacts:
        # when a failed/no-claim run has visible workspace deliverables, run the
        # normal evaluator while the mock services and DB are still alive. The
        # result remains annotated as lifecycle-incomplete so downstream reports
        # can keep success and artifact quality as separate axes.
        artifact_only_eval = False
        if task_status != TaskStatus.SUCCESS.value:
            if eval_command is not None and TaskEvaluator._has_visible_workspace_artifacts(agent_workspace):
                artifact_only_eval = True
                print(
                    f"[eval] task status is {task_status}; running artifact-only "
                    "evaluation because workspace deliverables exist."
                )
            else:
                return {
                    "pass": None,
                    "details": f"Task status: {task_status}, only SUCCESS counts as pass; pass is null",
                    "task_status": task_status,
                    "claim_done_required": True,
                    "artifact_eval_attempted": False,
                }

        if task_status != TaskStatus.SUCCESS.value and not artifact_only_eval:
            return {
                "pass": None,
                "details": f"Task status: {task_status}, only SUCCESS counts as pass; pass is null"
            }

        # Evaluate all content. For artifact_only_eval=True this is a
        # supplemental quality signal for an incomplete lifecycle, not proof
        # that the agent followed the required claim_done protocol.
        if eval_command is not None:
            # try:
            # Strip weekday name (e.g. "Saturday") from launch_time — same as
            # utils/roles/task_agent.py does for preprocess. Otherwise every
            # evaluation/main.py that calls datetime.fromisoformat() crashes
            # (c4 batch: wc-vip / wc-product-review "Invalid isoformat string").
            lt_clean = " ".join((launch_time or "").split()[:2])  # keep "YYYY-MM-DD HH:MM:SS"
            args = f"--res_log_file {res_log_file} --agent_workspace {agent_workspace} --groundtruth_workspace {groundtruth_workspace} --launch_time \"{lt_clean}\""
            command = f"{eval_command} {args}"
            output, error, returncode = await run_command(command, debug=True)
            print("== Evaluation STDOUT ==")
            print(output)
            print("== Evaluation STDERR ==")
            print(error)
            if returncode != 0:
                return {
                    "pass": False,
                    "failure": output,
                    "task_status": task_status,
                    "claim_done_required": True,
                    "artifact_eval_attempted": artifact_only_eval,
                }
                
        # Finally, it's successful
        return {
            "pass": True,
            "details": (
                "All evaluation checks passed, and task status is success"
                if not artifact_only_eval
                else "Artifact-only evaluation passed, but task status was not success; claim_done lifecycle remains incomplete"
            ),
            "task_status": task_status,
            "claim_done_required": True,
            "artifact_eval_attempted": artifact_only_eval,
        }
    
    @staticmethod
    async def evaluate_from_log_file(log_file_path: str, allow_resume: bool = False) -> Dict[str, Any]:
        """Evaluate task from log file"""
        try:            
            if not os.path.exists(log_file_path):
                return {
                    "pass": False,
                    "failure": "log_file_not_found",
                    "details": f"Log file not found: {log_file_path}"
                }
            # if allow_resume AND we can load pre exist eval res, we just load it
            eval_file_path = os.path.join(os.path.dirname(log_file_path),"eval_res.json")
            if allow_resume and os.path.exists(eval_file_path):
                eval_res = read_json(eval_file_path)
                return eval_res
            # otherwise, we do real eval and store the eval result
            dump_line = read_json(log_file_path)
            eval_res = await TaskEvaluator.evaluate_one(dump_line)
            write_json(eval_res, eval_file_path)
            return eval_res
            
        except Exception as e:
            logging.error(f"Error evaluating from log file {log_file_path}: {e}")
            return {
                "pass": False,
                "failure": "evaluation_error",
                "details": str(e)
            }
    
    @staticmethod
    async def batch_evaluate(run_results: List[Dict[str, Any]], allow_resume: bool=False) -> List[Dict[str, Any]]:
        """Batch evaluate task results"""
        eval_results = []
        
        for run_result in run_results:
            eval_result = {
                "task_config_path": run_result["task_config_path"],
                "task_id": run_result.get("task_id", "unknown"),
            }
            
            if not run_result.get("success", False):
                eval_result["evaluation"] = {
                    "pass": False,
                    "failure": "task_execution_failed",
                    "details": run_result.get("error", "Unknown error")
                }
            else:
                log_file = run_result.get("log_file")
                if log_file:
                    eval_result["evaluation"] = await TaskEvaluator.evaluate_from_log_file(log_file, allow_resume = allow_resume)
                else:
                    eval_result["evaluation"] = {
                        "pass": False,
                        "failure": "no_log_file",
                        "details": "No log file generated"
                    }
            
            eval_result["pass"] = eval_result["evaluation"]["pass"]
            eval_results.append(eval_result)
        
        return eval_results
