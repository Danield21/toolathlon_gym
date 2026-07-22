"""
Preprocess for sales-target-reconciliation task.
- Unpacks mock dashboard HTML from files/mock_dashboard.tar.gz
- Starts HTTP server at http://localhost:30160
NOTE: Does NOT inject data into Snowflake — uses existing SALES_DW data.
NOTE: PDF is pre-generated in initial_workspace/.
      Groundtruth Excel is pre-generated in groundtruth_workspace/.
      Mock HTML is pre-generated in files/mock_dashboard.tar.gz.
"""
import argparse
import asyncio
import os
import shutil
import tarfile


async def run_command(cmd: str):
    proc = await asyncio.create_subprocess_shell(
        cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE
    )
    await proc.wait()


async def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")

    # 0. Clean any stale output from previous runs
    if args.agent_workspace:
        stale = os.path.join(args.agent_workspace, "Q1_Sales_Review.xlsx")
        if os.path.isfile(stale):
            try:
                os.remove(stale)
                print(f"  Removed stale {stale}")
            except OSError as e:
                print(f"  Warning: could not remove {stale}: {e}")

    # 1. Unpack mock dashboard HTML from files/
    print("Setting up mock analytics dashboard ...")
    tmp_dir = os.path.join(task_root, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_dashboard.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"  -> Extracted {tar_path} to {tmp_dir}")

    # 2. Start HTTP server
    mock_dir = os.path.join(tmp_dir, "mock_dashboard")
    port = 30160
    await run_command(f"kill -9 $(lsof -ti:{port}) 2>/dev/null")
    await asyncio.sleep(0.5)
    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory {mock_dir} "
        f"> {mock_dir}/server.log 2>&1 &"
    )
    await asyncio.sleep(1)
    # Health check: wait up to 5s for the server to respond on /
    import urllib.request
    healthy = False
    for _ in range(10):
        try:
            with urllib.request.urlopen(
                f"http://localhost:{port}/", timeout=0.5
            ) as resp:
                if resp.status == 200:
                    healthy = True
                    break
        except Exception:
            await asyncio.sleep(0.5)
    if healthy:
        print(f"  -> Mock dashboard HEALTHY at http://localhost:{port}")
    else:
        print(f"  -> Mock dashboard launched but health check failed at "
              f"http://localhost:{port}")

    print("\nPreprocessing completed successfully!")


if __name__ == "__main__":
    asyncio.run(main())
