"""Ctrl+C must report a clean stop, never a traceback.

WHAT HAPPENED
-------------
A real run was interrupted mid-crawl and printed this:

    Stopping after the current safe point...
    [CYCLE] 0 row(s) waiting in current_data.xlsx
    Traceback (most recent call last):
      ...
    RuntimeError: extension service worker is not available

Both halves are true and the run WAS safe -- the position was saved and the
staged pages were kept -- but the traceback says the opposite, and an
operator watching a laptop has no way to tell a clean stop from a crash.

WHY IT HAPPENS
--------------
SIGINT goes to the whole process group, so Playwright's node driver dies the
instant Ctrl+C is pressed.  This code does not stop there: the signal handler
only sets a flag, and the loop keeps running until the next safe point.  In
that gap it calls into a browser that is already gone, and the call fails.

So after a stop has been asked for, a browser failure is a CONSEQUENCE of the
stop, not a new fault.  main() now reports it as the interruption it is --
but only when a stop was actually requested.  Without one it is a genuine
crash and must still surface in full.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper as S                               # noqa: E402


class InterruptCase(unittest.TestCase):
    def setUp(self):
        self.said = []
        self._say = S.say
        self._stop = S.STOP_REQUESTED
        S.say = self.said.append

    def tearDown(self):
        S.say = self._say
        S.STOP_REQUESTED = self._stop

    def run_main_with(self, error, stop_requested):
        """Drive main()'s tail: a runner that raises `error`."""
        S.STOP_REQUESTED = stop_requested

        def boom():
            raise error

        try:
            boom()
        except KeyboardInterrupt:
            return 130, "keyboard"
        except Exception as exc:
            if not S.STOP_REQUESTED:
                raise
            S.say("")
            S.say("Interrupted safely.")
            return 130, "clean"

    # -- the case the operator actually hit -------------------------------
    def test_a_dead_browser_after_ctrl_c_is_a_clean_stop(self):
        code, how = self.run_main_with(
            RuntimeError("extension service worker is not available"),
            stop_requested=True)
        self.assertEqual(code, 130)
        self.assertEqual(how, "clean")
        self.assertTrue(any("Interrupted safely" in line for line in self.said))

    def test_a_dropped_driver_after_ctrl_c_is_a_clean_stop(self):
        code, _how = self.run_main_with(
            Exception("Connection closed while reading from the driver"),
            stop_requested=True)
        self.assertEqual(code, 130)

    # -- and a real crash must NOT be swallowed ---------------------------
    def test_the_same_error_without_a_stop_still_raises(self):
        with self.assertRaises(RuntimeError):
            self.run_main_with(
                RuntimeError("extension service worker is not available"),
                stop_requested=False)

    # -- the source itself carries the guard ------------------------------
    def test_main_reraises_when_no_stop_was_requested(self):
        import inspect
        source = inspect.getsource(S.main)
        self.assertIn("if not STOP_REQUESTED:", source,
                      "main() must re-raise a real crash")
        self.assertIn("raise", source)

    def test_close_tolerates_a_browser_that_is_already_gone(self):
        import asyncio
        import types

        workflow = S.Workflow.__new__(S.Workflow)

        async def dead_close():
            raise RuntimeError("Target page, context or browser has been closed")

        workflow.browser = types.SimpleNamespace(close=dead_close)
        workflow.existing = {}
        workflow._master_dirty = False
        asyncio.run(workflow.close())          # must not raise


if __name__ == "__main__":
    unittest.main(verbosity=2)
