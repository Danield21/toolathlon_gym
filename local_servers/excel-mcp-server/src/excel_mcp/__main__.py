import sys
import traceback

import typer

from .server import run_sse, run_stdio, run_streamable_http

app = typer.Typer(help="Excel MCP Server")


def _log_lifecycle(message: str) -> None:
    """Log lifecycle messages without writing to the MCP stdout channel."""
    try:
        if sys.stderr is not None and not sys.stderr.closed:
            print(message, file=sys.stderr)
    except (BrokenPipeError, OSError, ValueError):
        # The parent may close stdio before the MCP server finishes unwinding.
        pass

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    try:
        run_sse()
    except KeyboardInterrupt:
        _log_lifecycle("\nShutting down server...")
    except Exception as e:
        _log_lifecycle(f"\nError: {e}")
        if sys.stderr is not None and not sys.stderr.closed:
            traceback.print_exc(file=sys.stderr)
    finally:
        _log_lifecycle("Service stopped.")

@app.command()
def streamable_http():
    """Start Excel MCP Server in streamable HTTP mode"""
    try:
        run_streamable_http()
    except KeyboardInterrupt:
        _log_lifecycle("\nShutting down server...")
    except Exception as e:
        _log_lifecycle(f"\nError: {e}")
        if sys.stderr is not None and not sys.stderr.closed:
            traceback.print_exc(file=sys.stderr)
    finally:
        _log_lifecycle("Service stopped.")

@app.command()
def stdio():
    """Start Excel MCP Server in stdio mode"""
    try:
        run_stdio()
    except KeyboardInterrupt:
        _log_lifecycle("\nShutting down server...")
    except Exception as e:
        _log_lifecycle(f"\nError: {e}")
        if sys.stderr is not None and not sys.stderr.closed:
            traceback.print_exc(file=sys.stderr)
    finally:
        _log_lifecycle("Service stopped.")

if __name__ == "__main__":
    app()
