"""Terminal progress rendering: live countdowns, cooldown block, phase cards.

All presentation, no policy.  Time and sleep are mocked so nothing waits, and
the tests prove the two things that matter for safety:

  * the countdown REPRESENTS the existing delay and never adds sleeping, and a
    Ctrl+C during it is still noticed promptly; and
  * ticks are terminal-only -- a non-interactive stream gets no carriage-return
    animation and the log never fills with per-second lines.

A StringIO stands in for a redirected stream (isatty False); FakeTTY stands in
for a real terminal (isatty True) so the '\\r' path is exercised without a tty.
"""

import asyncio
import io
import logging
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper                                           # noqa: E402



class FakeTTY(io.StringIO):
    """A StringIO that claims to be an interactive terminal."""

    def isatty(self):
        return True


class RecordingHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.messages = []

    def emit(self, record):
        self.messages.append(record.getMessage())


# ======================= sleep_interruptible + on_tick ======================
class SleepTickCase(unittest.TestCase):
    """on_tick is a display hook: it must not change the wait at all."""

    def naps_and_ticks(self, seconds, on_tick=None, check_stop=None, step=1.0):
        naps, ticks = [], []

        async def fake(secs):
            naps.append(secs)

        def tick(remaining):
            ticks.append(remaining)
            if on_tick:
                on_tick(remaining)
        asyncio.run(student_scraper.sleep_interruptible(
            seconds, check_stop=check_stop, step=step, sleep=fake,
            on_tick=tick))
        return naps, ticks

    # 4. the countdown introduces no extra sleeping
    def test_on_tick_adds_no_sleep(self):
        with_tick, ticks = self.naps_and_ticks(5)
        self.assertEqual(len(with_tick), 5)
        self.assertAlmostEqual(sum(with_tick), 5.0)
        # identical to a run with no tick at all
        naps = []

        async def fake(secs):
            naps.append(secs)
        asyncio.run(student_scraper.sleep_interruptible(5, sleep=fake))
        self.assertEqual(len(naps), len(with_tick))

    def test_on_tick_sees_each_remaining(self):
        _naps, ticks = self.naps_and_ticks(3)
        self.assertEqual(ticks, [3.0, 2.0, 1.0])       # before each step

    # 5. Ctrl+C during the wait aborts, even with a countdown attached
    def test_ctrl_c_still_interruptible(self):
        class Boom(Exception):
            pass
        calls = {"n": 0}

        def check():
            calls["n"] += 1
            if calls["n"] >= 2:
                raise Boom()
        with self.assertRaises(Boom):
            self.naps_and_ticks(60, check_stop=check)


# ================================ Countdown =================================
class CountdownCase(unittest.TestCase):
    def drive(self, cd, seconds, step=1.0):
        """Feed the countdown the same remaining values the sleep would."""
        cd.start()
        remaining = float(seconds)
        while remaining > 0:
            cd.tick(remaining)
            remaining -= step
        cd.finish()

    # 1/2/3. search / page / profile countdown all render the tagged line
    def test_tags_render_on_tty(self):
        for tag in ("SEARCH", "PAGE", "PROFILE"):
            tty = FakeTTY()
            cd = student_scraper.Countdown(5, tag, stream=tty)
            self.drive(cd, 5)
            out = tty.getvalue()
            self.assertIn("[PACING][%s]" % tag, out)
            self.assertIn("00:05 remaining", out)
            self.assertIn("00:01 remaining", out)
            self.assertIn("Done. Continuing...", out)
            self.assertIn("\r", out)                   # in-place updates

    # 8. a non-interactive stream gets NO carriage-return animation
    def test_non_tty_no_carriage_return(self):
        plain = io.StringIO()
        cd = student_scraper.Countdown(5, "SEARCH", stream=plain)
        self.drive(cd, 5)
        out = plain.getvalue()
        self.assertNotIn("\r", out)
        self.assertNotIn("remaining", out)             # no per-tick spam
        self.assertIn("waiting", out)                  # simple start
        self.assertIn("done", out)                     # simple finish

    # 9. terminal countdown ticks are NOT written to any log
    def test_ticks_not_logged(self):
        handler = RecordingHandler()
        logger = logging.getLogger("test_progress_ticks")
        logger.handlers.clear()
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        tty = FakeTTY()
        cd = student_scraper.Countdown(5, "SEARCH", stream=tty, logger=logger)
        self.drive(cd, 5)
        # exactly two log records: started + finished, and neither is a tick
        self.assertEqual(len(handler.messages), 2)
        self.assertTrue(any("started" in m for m in handler.messages))
        self.assertTrue(any("finished" in m for m in handler.messages))
        self.assertFalse(any("remaining" in m for m in handler.messages))

    def test_zero_seconds_never_draws_negative(self):
        tty = FakeTTY()
        cd = student_scraper.Countdown(0.4, "SEARCH", stream=tty)
        cd.start()
        cd.tick(0.4)                                   # int() -> 0, skipped
        cd.finish()
        self.assertNotIn("-", tty.getvalue())


# ============================= CooldownDisplay ==============================
class CooldownDisplayCase(unittest.TestCase):
    # 6. the cooldown block shows level, reason, retry-at and HH:MM:SS remaining
    def test_displays_remaining(self):
        tty = FakeTTY()
        cd = student_scraper.CooldownDisplay(
            2, "GitHub search rate limit", "2026-09-28 18:45:00", 1800,
            stream=tty)
        cd.start()
        cd.tick(1800)
        out = tty.getvalue()
        self.assertIn("[RATE LIMIT COOLDOWN]", out)
        self.assertIn("Level      : 2", out)
        self.assertIn("Reason     : GitHub search rate limit", out)
        self.assertIn("Retry at   : 2026-09-28 18:45:00", out)
        self.assertIn("Remaining  : 00:30:00", out)

    # 7. a long cooldown redraws far less often than once a second
    def test_long_cooldown_lower_refresh(self):
        self.assertEqual(student_scraper.refresh_for(30), 1)      # normal pacing
        self.assertEqual(student_scraper.refresh_for(300), 10)    # <= 10 min
        self.assertEqual(student_scraper.refresh_for(1800), 30)   # > 10 min
        tty = FakeTTY()
        cd = student_scraper.CooldownDisplay(1, "x", "later", 7200, stream=tty)
        cd.start()
        remaining = 7200.0
        while remaining > 0:                               # step every second
            cd.tick(remaining)
            remaining -= 1.0
        draws = tty.getvalue().count("\r")
        # a two-hour wait stepping every second draws ~7200/30, never 7200
        self.assertLess(draws, 300)
        self.assertGreater(draws, 200)

    # 8. non-TTY cooldown prints the block once, no animation
    def test_non_tty_prints_once(self):
        plain = io.StringIO()
        cd = student_scraper.CooldownDisplay(1, "x", "later", 600, stream=plain)
        cd.start()
        for r in (600, 300, 60, 1):
            cd.tick(r)
        cd.finish()
        out = plain.getvalue()
        self.assertNotIn("\r", out)
        self.assertEqual(out.count("Remaining  :"), 1)     # once, not per tick
        self.assertIn("Cooldown finished.", out)

    # 9. cooldown ticks are not logged either
    def test_cooldown_ticks_not_logged(self):
        handler = RecordingHandler()
        logger = logging.getLogger("test_progress_cooldown")
        logger.handlers.clear()
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        tty = FakeTTY()
        cd = student_scraper.CooldownDisplay(3, "x", "later", 600, stream=tty,
                                      logger=logger)
        cd.start()
        for r in range(600, 0, -1):
            cd.tick(float(r))
        cd.finish()
        self.assertEqual(len(handler.messages), 2)         # started + finished
        self.assertFalse(any("Remaining" in m for m in handler.messages))


# ================================ phase cards ================================
class CardCase(unittest.TestCase):
    # 10. rank card shows the right rank / search / query
    def test_rank_card(self):
        card = student_scraper.rank_card(5, 80, 39, 40,
                                  "Software Developer Intern Amritsar", 77, 0)
        self.assertIn("[RANK 05/80]", card)
        self.assertIn("Search      : 39/40", card)
        self.assertIn("Query       : Software Developer Intern Amritsar", card)
        self.assertIn("Students    : 77", card)
        self.assertIn("Deferred    : 0", card)

    def test_profile_card(self):
        card = student_scraper.profile_card(2, 9, "example-user", 37, 40)
        self.assertIn("[PROFILE]", card)
        self.assertIn("Candidate   : 2/9", card)
        self.assertIn("Username    : example-user", card)
        self.assertIn("Search      : 37/40", card)

    def test_crawl_card(self):
        card = student_scraper.crawl_card(21, 40, 2, 3, 20, 24)
        self.assertIn("[CRAWL]", card)
        self.assertIn("Search      : 21/40", card)
        self.assertIn("Page        : 2/3", card)
        self.assertIn("Collected   : 20", card)
        self.assertIn("Expected    : 24", card)

    def test_retry_banner_search(self):
        text = student_scraper.retry_banner("search", query="Graduate Trainee Amritsar")
        self.assertIn("Retrying same operation", text)
        self.assertIn("Retry search:", text)
        self.assertIn("Graduate Trainee Amritsar", text)

    def test_retry_banner_profile(self):
        """A profile rate limit retries the whole SEARCH, so say so.

        The profile is named as the cause, but the banner must not imply that
        only that one profile is redone.
        """
        text = student_scraper.retry_banner(
            "profile", query='in:B.Tech CS location:"Mohali"',
            username="vu241fa04074-cpu")
        self.assertIn("Retry search", text)
        self.assertIn("a profile was rate-limited", text)
        self.assertIn('in:B.Tech CS location:"Mohali"', text)
        self.assertIn("Profile: vu241fa04074-cpu", text)

    def test_retry_banner_pagination(self):
        text = student_scraper.retry_banner("pagination",
                                     query="University Student Mohali", page=3)
        self.assertIn("Retry pagination:", text)
        self.assertIn("Search: University Student Mohali", text)
        self.assertIn("Page: 3", text)

    def test_ranges_compact(self):
        self.assertEqual(student_scraper._ranges([1, 2, 3, 4]), "1-4")
        self.assertEqual(student_scraper._ranges([5, 6, 7, 20]), "5-7, 20")
        self.assertEqual(student_scraper._ranges([]), "-")

    def test_device_card(self):
        card = student_scraper.device_card(
            owned=list(range(1, 21)), completed=[1, 2, 3, 4],
            current_rank=5, total_ranks=80, rank_progress="39/40",
            students=465)
        self.assertIn("DEVICE PROGRESS", card)
        self.assertIn("Owned ranks       : 1-20", card)
        self.assertIn("Finished ranks    : 4/20", card)
        self.assertIn("Current rank      : 5/80", card)
        self.assertIn("Rank progress     : 39/40", card)
        self.assertIn("Students collected: 465", card)
        self.assertIn("Pending:", card)
        self.assertIn("5-20", card)


if __name__ == "__main__":
    unittest.main(verbosity=2)
