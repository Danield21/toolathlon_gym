"""
Preprocess script for sf-workforce-diversity-review task.

1. Extracts mock_pages.tar.gz and starts HTTP server on port 30211.
2. sf_data is read-only -- do NOT modify.
"""
import argparse
import asyncio
import os
import shutil
import tarfile
import urllib.request


async def wait_until_ready(url, tries=20, interval=0.5):
    """Poll the mock portal until it answers, so the agent never races a
    half-started server."""
    for _ in range(tries):
        try:
            with urllib.request.urlopen(url, timeout=2) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            pass
        await asyncio.sleep(interval)
    return False


async def setup_mock_server():
    """Extract mock_pages.tar.gz and start HTTP server on port 30211."""
    print("[preprocess] Setting up mock diversity benchmark portal...")

    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files_dir = os.path.join(task_root, "files")
    tmp_dir = os.path.join(task_root, "tmp")

    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)
    print(f"[preprocess] Extracted {tar_path} to {tmp_dir}")

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    port = 30211

    # Clear any leftover listener so the http.server below binds cleanly.
    # (Concurrent instances serve identical content, so killing a stale one is
    # harmless; the health check below ensures the new server is up.)
    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {port} --directory {mock_dir} "
        f"> {mock_dir}/server.log 2>&1 &"
    )

    url = f"http://localhost:{port}/"
    if await wait_until_ready(url):
        print(f"[preprocess] Mock portal running at {url}")
    else:
        print(f"[preprocess] WARNING: mock portal at {url} did not answer within "
              f"the retry window; the agent may need to retry its first request.")


async def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    await setup_mock_server()
    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
