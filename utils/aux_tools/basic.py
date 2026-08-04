"""Basic local tools: claim_done, sleep."""
import asyncio


class ClaimDoneSignal(BaseException):
    """Internal control signal used to stop CAMEL immediately after claim_done."""


def make_claim_done(done_flag: list):
    """Return a claim_done tool that terminates the current CAMEL step."""
    async def claim_done() -> str:
        """Call this tool when the task is fully completed."""
        done_flag[0] = True
        # CAMEL catches ordinary Exception subclasses as tool failures and then
        # asks the model for another response.  A private BaseException signal
        # reaches TaskAgent directly, avoiding an unnecessary generation after
        # the sole completion signal has already fired.
        raise ClaimDoneSignal()
    return claim_done


async def sleep(seconds: float = 1) -> str:
    """Sleep for the given number of seconds.

    Args:
        seconds: Number of seconds to sleep (default 1).
    """
    await asyncio.sleep(seconds)
    return f"Slept {seconds} seconds."
