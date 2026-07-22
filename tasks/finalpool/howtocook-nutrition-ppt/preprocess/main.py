import argparse
import os


def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()
    print("No preprocess needed for HowToCook recipe task")

    # Cleanup stale outputs in agent workspace (skip for GT)
    if args.agent_workspace:
        ws_abs = os.path.abspath(args.agent_workspace)
        if "groundtruth" not in ws_abs.lower():
            for fname in ("Nutrition_Comparison.xlsx", "Healthy_Eating_Guide.pptx"):
                stale = os.path.join(args.agent_workspace, fname)
                if os.path.exists(stale):
                    try:
                        os.remove(stale)
                        print(f"[preprocess] Removed stale {stale}")
                    except Exception as e:
                        print(f"[preprocess] Could not remove {stale}: {e}")


if __name__ == "__main__":
    main()
