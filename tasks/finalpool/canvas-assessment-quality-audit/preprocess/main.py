"""
Preprocess script for canvas-assessment-quality-audit task.
Starts mock HTTP server on port 30221.
Canvas is read-only.
"""
import argparse
import asyncio
import os
import shutil

PORT = 30221


def _prepare_serve_dir(task_root):
    """Locate (and if needed create) the directory served by the mock portal.

    Preferred layout is files/mock_pages shipped with the task; copy it into
    tmp/mock_pages so the http.server log can be written next to the content.
    Falls back to an empty-but-created tmp/mock_pages with a placeholder index
    so the server never fails to start on a missing directory.
    """
    serve_dir = os.path.join(task_root, "tmp", "mock_pages")
    if os.path.isdir(serve_dir):
        return serve_dir
    files_dir = os.path.join(task_root, "files", "mock_pages")
    os.makedirs(os.path.dirname(serve_dir), exist_ok=True)
    if os.path.isdir(files_dir):
        shutil.copytree(files_dir, serve_dir)
        return serve_dir
    os.makedirs(serve_dir, exist_ok=True)
    fallback = os.path.join(serve_dir, "index.html")
    if not os.path.isfile(fallback):
        with open(fallback, "w", encoding="utf-8") as fh:
            fh.write(
                "<html><body><h1>Assessment Quality Standards</h1>"
                "<p>Portal content unavailable.</p></body></html>"
            )
    return serve_dir


async def main():
    # No writable schemas to DELETE - read-only data sources
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Start mock HTTP server
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    serve_dir = _prepare_serve_dir(task_root)

    kill_proc = await asyncio.create_subprocess_shell(
        f"kill -9 $(lsof -ti:{PORT}) 2>/dev/null",
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    await kill_proc.wait()
    await asyncio.sleep(0.5)

    await asyncio.create_subprocess_shell(
        f"nohup python3 -m http.server {PORT} --directory {serve_dir} "
        f"> {serve_dir}/server.log 2>&1 &"
    )
    await asyncio.sleep(1)
    print(f"[preprocess] Mock server running at http://localhost:{PORT}")
    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
