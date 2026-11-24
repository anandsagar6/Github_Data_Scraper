"""Conservative pacing + escalating rate-limit cooldown (all time mocked).

Normal pacing puts a random delay between different completed searches and
between result pages.  A genuine GitHub rate-limit/challenge starts a persisted cooldown of a
FRESH RANDOM RATE_LIMIT_COOLDOWN_MIN..MAX minutes that a restart cannot
bypass and that resets when the blocked operation finally works.  There is
no ladder and no level that stops the run: every hit draws its own wait and
the same search is retried afterwards, for as many hits as it takes.  No
test waits for real time -- sleeps and 'now' are injected, and the tests
that need an exact duration pin the draw with fixed_cooldown().

A GitHub 5xx is NOT that.  It means GitHub is broken, not that GitHub is
restricting us, so it too waits and retries for as long as the outage lasts
-- it never stops the job.  OutageCase covers that.
"""

import asyncio
import datetime
import inspect
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                         # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (RateLimitCooldown, cooldown_seconds, parse_retry_after, sleep_interruptible)   # noqa: E402
import sandbox                                         # noqa: E402


def at(minute):
    return datetime.datetime(2026, 9, 28, 12, minute, 0)


def fixed_cooldown(case, minutes):
    """Pin the random cooldown draw to exactly `minutes` for one test.

    cooldown_seconds() reads these two module globals on every call, so a
    test that needs an exact duration squeezes MIN and MAX onto the same
    value instead of reaching into `random`.
    """
    low = student_scraper.RATE_LIMIT_COOLDOWN_MIN
    high = student_scraper.RATE_LIMIT_COOLDOWN_MAX
    student_scraper.RATE_LIMIT_COOLDOWN_MIN = minutes
    student_scraper.RATE_LIMIT_COOLDOWN_MAX = minutes
    case.addCleanup(setattr, student_scraper, "RATE_LIMIT_COOLDOWN_MAX", high)
    case.addCleanup(setattr, student_scraper, "RATE_LIMIT_COOLDOWN_MIN", low)


def band():
    """The production cooldown band in SECONDS, as (low, high)."""
    return (student_scraper.RATE_LIMIT_COOLDOWN_MIN * 60,
            student_scraper.RATE_LIMIT_COOLDOWN_MAX * 60)


def stop_at_the_end_of_the_ladder(case):
    """Pin C.STOP_WHEN_COOLDOWN_EXHAUSTED=True for one test.

    The project default is False -- an unattended rank waits at the top of
    the ladder rather than ending overnight.  The tests below are about the
    OTHER mode, so they say so instead of inheriting whichever default is
    currently configured.
    """
    saved = C.STOP_WHEN_COOLDOWN_EXHAUSTED
    C.STOP_WHEN_COOLDOWN_EXHAUSTED = True
    case.addCleanup(setattr, C, "STOP_WHEN_COOLDOWN_EXHAUSTED", saved)


class CooldownEscalationCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="cool_")
        self.path = os.path.join(self.dir, "rate_limit_state.json")

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.dir, ignore_errors=True)

    def cd(self):
        return RateLimitCooldown(path=self.path)

    # 3-6.  Every hit waits a fresh random draw inside the band, the level
    # keeps climbing so the log can show a limit that is not clearing, and
    # nothing ever exhausts.  Twenty hits is the point: an unattended run
    # must survive the twentieth rate limit exactly like the first.
    def test_every_hit_draws_a_fresh_wait_inside_the_band(self):
        low, high = band()
        c = self.cd()
        for expected_level in range(1, 21):
            self.assertEqual(c.escalate("search", "rate limit", now=at(0)),
                             expected_level)
            wait = c.remaining(now=at(0))
            self.assertTrue(low <= wait <= high,
                            "hit %d waited %ds, outside the %d..%ds band"
                            % (expected_level, wait, low, high))
            self.assertFalse(c.exhausted(),
                             "hit %d must not stop an unattended run"
                             % expected_level)

    # ...and the draw really is a draw, not a constant dressed up as one.
    def test_the_wait_actually_varies(self):
        c = self.cd()
        seen = set()
        for _ in range(60):
            c.escalate("search", "rate limit", now=at(0))
            seen.add(c.remaining(now=at(0)))
        self.assertGreater(len(seen), 1,
                           "60 hits all waited %r -- the cooldown is not "
                           "random" % seen)

    # a one-line pin so the band cannot be widened by accident
    def test_the_band_is_sane(self):
        low, high = band()
        self.assertGreater(low, 0, "a zero-minute cooldown is a busy loop")
        self.assertLessEqual(low, high, "MIN must not exceed MAX")

    # 7. a further failure after level 4 -> exhausted (caller stops the device)
    def test_fifth_failure_exhausts(self):
        stop_at_the_end_of_the_ladder(self)
        c = self.cd()
        for _ in range(4):
            c.escalate("search", "rate limit", now=at(0))
        level = c.escalate("search", "rate limit", now=at(0))
        self.assertEqual(level, 5)
        self.assertTrue(c.exhausted())

    # 8. a successful retry resets the escalation
    def test_success_resets(self):
        c = self.cd()
        c.escalate("profile", "rate limit", now=at(0))
        c.escalate("profile", "rate limit", now=at(0))
        self.assertEqual(c.level, 2)
        self.assertTrue(c.clear())
        self.assertEqual(c.level, 0)
        self.assertFalse(c.active(now=at(0)))

    # 9/10. the cooldown persists across a restart and cannot be bypassed
    def test_persists_and_blocks_restart(self):
        # pinned to 15 minutes so the window is observable across a restart;
        # the duration is a random draw in production, which is tested above
        fixed_cooldown(self, 15)
        c = self.cd()
        for _ in range(3):
            c.escalate("search", "rate limit", now=at(0))
        fresh = self.cd()                                   # simulate restart
        self.assertTrue(fresh.active(now=at(5)))            # still cooling at 12:05
        self.assertEqual(fresh.level, 3)
        self.assertGreater(fresh.remaining(now=at(5)), 0)
        self.assertFalse(fresh.active(now=at(20)))          # expired by 12:20

    # a server Retry-After is GitHub's own reset time, so it wins OUTRIGHT
    def test_retry_after_is_taken_as_given(self):
        """In BOTH directions: GitHub knows its reset time, our ladder guesses.

        max(policy, header) used to override the server the wrong way -- a
        "retry after 60 seconds" still waited the full policy minimum.
        """
        self.assertTrue(C.TRUST_RETRY_AFTER)
        # longer than our ladder -> respected
        c = self.cd()
        c.escalate("search", "rate limit", retry_after=40 * 60, now=at(0))
        self.assertEqual(c.remaining(now=at(0)), 40 * 60)
        # SHORTER than our ladder -> also respected
        c2 = RateLimitCooldown(path=os.path.join(self.dir, "b.json"))
        for _ in range(3):                      # level 3 = 15 min policy
            c2.escalate("search", "rate limit", now=at(0))
        c2.escalate("search", "rate limit", retry_after=60, now=at(0))
        self.assertEqual(c2.remaining(now=at(0)), 60)

    def test_never_shortening_can_be_restored(self):
        """TRUST_RETRY_AFTER=False keeps the old max(policy, header)."""
        fixed_cooldown(self, 7)
        original = C.TRUST_RETRY_AFTER
        C.TRUST_RETRY_AFTER = False
        try:
            c = RateLimitCooldown(path=os.path.join(self.dir, "c.json"))
            c.escalate("search", "rate limit", retry_after=1, now=at(0))
            # our 7-minute draw wins over GitHub's 1 second, as before
            self.assertEqual(c.remaining(now=at(0)), 7 * 60)
        finally:
            C.TRUST_RETRY_AFTER = original

    def test_a_zero_retry_after_still_waits_at_least_a_second(self):
        """Retry-After: 0 must not become a busy loop."""
        c = RateLimitCooldown(path=os.path.join(self.dir, "d.json"))
        c.escalate("search", "rate limit", retry_after=0, now=at(0))
        self.assertGreaterEqual(c.remaining(now=at(0)), 1)

    def test_cooldown_seconds_and_retry_after_parse(self):
        low, high = band()
        self.assertEqual(cooldown_seconds(0), 0)        # no cooldown at all
        self.assertEqual(cooldown_seconds(-3), 0)
        for level in (1, 4, 9, 99):                     # level selects nothing
            self.assertTrue(low <= cooldown_seconds(level) <= high)
        self.assertEqual(parse_retry_after("120"), 120)
        self.assertIsNone(parse_retry_after(""))
        self.assertIsNone(parse_retry_after(None))


class NeverStopOnRateLimitCase(unittest.TestCase):
    """C.STOP_WHEN_COOLDOWN_EXHAUSTED=False -- the project default.

    An unattended rank must not end its run because GitHub throttled it for
    an hour: it waits at the top of the ladder and keeps retrying the same
    search, exactly as it already does for a GitHub 5xx outage.
    """

    def setUp(self):
        self.saved = C.STOP_WHEN_COOLDOWN_EXHAUSTED
        C.STOP_WHEN_COOLDOWN_EXHAUSTED = False
        self.addCleanup(setattr, C, "STOP_WHEN_COOLDOWN_EXHAUSTED",
                        self.saved)
        # a sandbox + a fake sleep, so the loop below can be driven for real
        self.dir = tempfile.mkdtemp(prefix="neverstop_")
        self.addCleanup(shutil.rmtree, self.dir, True)
        self.ctx = sandbox.enter(os.path.join(self.dir, "runtime"))
        self.addCleanup(sandbox.restore)
        self.said = []
        self.slept = []
        orig_say = student_scraper.say
        orig_sleep = student_scraper.sleep_interruptible
        self.addCleanup(setattr, student_scraper, "say", orig_say)
        self.addCleanup(setattr, student_scraper, "sleep_interruptible",
                        orig_sleep)
        student_scraper.say = self.said.append

        async def fake_sleep(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
            self.slept.append(seconds)
        student_scraper.sleep_interruptible = fake_sleep

    def cd(self):
        return student_scraper.RateLimitCooldown()

    def workflow(self):
        wf = student_scraper.Workflow.__new__(student_scraper.Workflow)
        wf.context = self.ctx
        wf.cooldown = student_scraper.RateLimitCooldown()
        wf.options = type("O", (), {})()
        wf.state = None
        return wf

    # THE WHOLE POINT, driven through the real code path: thirty rate limits
    # in a row, each one answered with "wait a random MIN..MAX minutes, then
    # retry the SAME search".  _cooldown_and_wait returning True is what
    # makes the caller `continue` instead of raising Stopped, so thirty Trues
    # is thirty retries and no stop.
    def test_thirty_rate_limits_in_a_row_all_retry(self):
        low, high = band()
        wf = self.workflow()
        for hit in range(1, 31):
            limit = student_scraper.RateLimited("GitHub rate limit",
                                               source="search")
            kept_going = asyncio.run(wf._cooldown_and_wait(limit))
            self.assertTrue(kept_going,
                            "hit %d stopped the run; it must retry" % hit)
            self.assertEqual(wf.cooldown.level, hit,
                             "the level must still climb, for the log")
            # one second of slack at the bottom: escalate() stores an
            # absolute wake-up time and remaining() int()-truncates the
            # fraction of a second that has passed since, so a 300s draw is
            # slept as 299s.  That predates the random band and is harmless.
            waited = self.slept[-1]
            self.assertTrue(low - 1 <= waited <= high,
                            "hit %d slept %ss, outside the %s..%ss band"
                            % (hit, waited, low, high))
        self.assertFalse(wf.cooldown.exhausted())
        # and it really did re-draw rather than repeat one number
        self.assertGreater(len(set(self.slept)), 1,
                           "every cooldown was identical -- not random")

    # a success in between wipes the escalation, so a run that recovers is
    # not carrying thirty hits' worth of history into the next limit
    def test_a_success_between_limits_resets_the_level(self):
        wf = self.workflow()
        for _ in range(3):
            asyncio.run(wf._cooldown_and_wait(
                student_scraper.RateLimited("rate limit", source="search")))
        self.assertEqual(wf.cooldown.level, 3)
        wf.cooldown.clear()
        self.assertEqual(wf.cooldown.level, 0)
        asyncio.run(wf._cooldown_and_wait(
            student_scraper.RateLimited("rate limit", source="search")))
        self.assertEqual(wf.cooldown.level, 1)

    def test_the_ladder_never_exhausts(self):
        c = self.cd()
        for _ in range(C.MAX_COOLDOWN_LEVEL + 6):
            c.escalate("search", "rate limit", now=at(0))
            self.assertFalse(c.exhausted(),
                             "an unattended run must not stop itself")

    def test_the_wait_never_grows_past_the_band(self):
        low, high = band()
        c = self.cd()
        for hit in range(C.MAX_COOLDOWN_LEVEL + 20):
            c.escalate("search", "rate limit", now=at(0))
            self.assertTrue(low <= c.remaining(now=at(0)) <= high,
                            "hit %d escaped the band -- the wait must not "
                            "grow without bound either" % (hit + 1))

    def test_a_success_still_resets_the_ladder(self):
        c = self.cd()
        for _ in range(C.MAX_COOLDOWN_LEVEL + 2):
            c.escalate("search", "rate limit", now=at(0))
        c.clear()
        self.assertEqual(c.level, 0)
        self.assertFalse(c.exhausted())

    def test_an_outage_is_unaffected(self):
        """Outages never exhausted before this flag existed either."""
        c = self.cd()
        for _ in range(C.MAX_COOLDOWN_LEVEL + 2):
            c.escalate("search", "502", now=at(0), kind="outage")
        self.assertFalse(c.exhausted())
        self.assertEqual(c.kind, "outage")


class SleepInterruptibleCase(unittest.TestCase):
    def test_sleeps_in_steps(self):
        naps = []

        async def fake(seconds):
            naps.append(seconds)
        asyncio.run(sleep_interruptible(3, sleep=fake, step=1.0))
        self.assertEqual(len(naps), 3)             # 1 + 1 + 1
        self.assertAlmostEqual(sum(naps), 3.0)

    def test_zero_seconds_does_not_sleep(self):
        naps = []

        async def fake(seconds):
            naps.append(seconds)
        asyncio.run(sleep_interruptible(0, sleep=fake))
        self.assertEqual(naps, [])

    # 15. Ctrl+C during a delay aborts cleanly (check_stop raises)
    def test_check_stop_aborts(self):
        class Boom(Exception):
            pass

        def check():
            raise Boom()

        async def fake(seconds):
            return None
        with self.assertRaises(Boom):
            asyncio.run(sleep_interruptible(60, check_stop=check, sleep=fake))


class WorkflowPacingCase(unittest.TestCase):
    """The run-loop pacing + cooldown wiring, with sleep and the crawl mocked."""

    def setUp(self):
        import student_scraper as S
        self.S = S
        self.dir = tempfile.mkdtemp(prefix="wfpace_")
        self.ctx = sandbox.enter(os.path.join(self.dir, "runtime"))
        self.slept = []
        self._orig_sleep = student_scraper.sleep_interruptible

        async def fake_sleep(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
            self.slept.append(seconds)
        student_scraper.sleep_interruptible = fake_sleep

    def tearDown(self):
        sandbox.restore()
        student_scraper.sleep_interruptible = self._orig_sleep
        shutil.rmtree(self.dir, ignore_errors=True)

    def workflow(self):
        wf = self.S.Workflow.__new__(self.S.Workflow)
        wf.context = self.ctx
        wf.cooldown = student_scraper.RateLimitCooldown()
        wf.options = type("O", (), {})()
        return wf

    # 1. a completed search triggers a 30-60s delay
    def test_search_delay_range(self):
        wf = self.workflow()
        asyncio.run(wf._search_delay())
        self.assertEqual(len(self.slept), 1)
        self.assertTrue(student_scraper.SEARCH_DELAY_MIN <= self.slept[0]
                        <= student_scraper.SEARCH_DELAY_MAX)

    # 2. a page transition triggers a page delay
    def test_page_delay_range(self):
        import student_scraper as page_crawler
        collector = page_crawler.PageByPageCollector(
None, object(), say=lambda *_a: None)
        asyncio.run(collector._page_delay())
        self.assertEqual(len(self.slept), 1)
        self.assertTrue(student_scraper.PAGE_DELAY_MIN <= self.slept[0] <= student_scraper.PAGE_DELAY_MAX)

    # a between-profiles transition triggers a profile delay
    def test_profile_delay_range(self):
        wf = self.workflow()
        asyncio.run(wf._profile_delay())
        self.assertEqual(len(self.slept), 1)
        self.assertTrue(student_scraper.PROFILE_DELAY_MIN <= self.slept[0]
                        <= student_scraper.PROFILE_DELAY_MAX)

    # a ZERO-RESULT search is one page load, so it gets the shorter pause
    def test_zero_result_search_uses_short_delay(self):
        wf = self.workflow()
        wf._last_search_empty = True
        asyncio.run(wf._search_delay())
        self.assertEqual(len(self.slept), 1)
        self.assertTrue(student_scraper.ZERO_RESULT_DELAY_MIN <= self.slept[0]
                        <= student_scraper.ZERO_RESULT_DELAY_MAX)
        # and it really is shorter than a productive search's pause
        self.assertLess(self.slept[0], student_scraper.SEARCH_DELAY_MIN)

    # a search that actually collected keeps the FULL pause
    def test_productive_search_uses_full_delay(self):
        wf = self.workflow()
        wf._last_search_empty = False
        asyncio.run(wf._search_delay())
        self.assertTrue(student_scraper.SEARCH_DELAY_MIN <= self.slept[0]
                        <= student_scraper.SEARCH_DELAY_MAX)

    # an unset flag must behave like a productive search (never shorten by
    # accident on a code path that forgot to set it)
    def test_missing_flag_defaults_to_full_delay(self):
        wf = self.workflow()
        self.assertFalse(hasattr(wf, "_last_search_empty"))
        asyncio.run(wf._search_delay())
        self.assertTrue(student_scraper.SEARCH_DELAY_MIN <= self.slept[0]
                        <= student_scraper.SEARCH_DELAY_MAX)

    # the exact normal-pacing values the operator asked for: these are
    # find_naukri_job.py's, so each assertion names the naukri constant it
    # came from.  A change here is a change to how hard this hits GitHub.
    def test_config_values(self):
        # Through sandbox.pristine(), because this case runs inside a sandbox
        # and a sandbox collapses the pacing so the suite does not sleep for
        # real -- these are the PRODUCTION numbers, which is the point.
        real = sandbox.pristine
        # naukri HX_KEYWORD_GAP_SECONDS 15, jittered up to 35%
        self.assertEqual((real("SEARCH_DELAY_MIN"), real("SEARCH_DELAY_MAX")),
                         (15, 20))
        # naukri LISTING_PAGE_DELAY_*_SECONDS
        self.assertEqual((real("PAGE_DELAY_MIN"), real("PAGE_DELAY_MAX")),
                         (60, 180))
        # naukri JOB_CARD_DELAY_*_SECONDS
        self.assertEqual((real("PROFILE_DELAY_MIN"),
                          real("PROFILE_DELAY_MAX")), (5, 25))
        # naukri HX_CITY_GAP_SECONDS 5, jittered up to 35%
        self.assertEqual((real("ZERO_RESULT_DELAY_MIN"),
                          real("ZERO_RESULT_DELAY_MAX")), (5, 7))
        # naukri HX_JITTER_FRACTION / HX_CYCLE_COOLDOWN_MINUTES
        self.assertEqual(student_scraper.JITTER_FRACTION, 0.35)
        self.assertEqual(real("LAP_COOLDOWN_MINUTES"), 45)
        # the rate-limit cooldown band, ours, drawn fresh on every hit
        self.assertEqual((real("RATE_LIMIT_COOLDOWN_MIN"),
                          real("RATE_LIMIT_COOLDOWN_MAX")), (5, 30))
        self.assertEqual(C.MAX_COOLDOWN_LEVEL, 4)

    # an empty search must still be the CHEAP one of the two
    def test_an_empty_search_is_cheaper_than_a_productive_one(self):
        self.assertLess(student_scraper.ZERO_RESULT_DELAY_MAX, student_scraper.SEARCH_DELAY_MIN)

    # every range must be a range, or random.uniform() raises
    def test_every_range_is_the_right_way_round(self):
        for low, high in ((student_scraper.SEARCH_DELAY_MIN, student_scraper.SEARCH_DELAY_MAX),
                          (student_scraper.ZERO_RESULT_DELAY_MIN, student_scraper.ZERO_RESULT_DELAY_MAX),
                          (student_scraper.PAGE_DELAY_MIN, student_scraper.PAGE_DELAY_MAX),
                          (student_scraper.PROFILE_DELAY_MIN, student_scraper.PROFILE_DELAY_MAX)):
            self.assertLessEqual(low, high)
            self.assertGreater(low, 0)

    # 11 (search rate limit): cooldown escalates and the SAME op is retried
    def test_cooldown_and_wait_retries(self):
        wf = self.workflow()
        limit = self.S.RateLimited("search rate limit", source="search")
        retry = asyncio.run(wf._cooldown_and_wait(limit))
        self.assertTrue(retry)                       # caller retries same unit
        self.assertEqual(wf.cooldown.level, 1)
        self.assertEqual(len(self.slept), 1)         # waited the cooldown
        self.assertGreater(self.slept[0], 0)

    # after level 4, cooldown_and_wait refuses (device stops)
    def test_cooldown_and_wait_exhausts(self):
        stop_at_the_end_of_the_ladder(self)
        wf = self.workflow()
        for _ in range(4):
            wf.cooldown.escalate("search", "rate limit")
        limit = self.S.RateLimited("search rate limit", source="search")
        retry = asyncio.run(wf._cooldown_and_wait(limit))
        self.assertFalse(retry)                      # exhausted -> stop
        self.assertTrue(wf.cooldown.exhausted())

    # 10. an active persisted cooldown is respected on (re)start, then waited
    def test_respect_active_cooldown_waits(self):
        wf = self.workflow()
        wf.cooldown.escalate("search", "rate limit")  # active now
        ok = asyncio.run(wf._respect_active_cooldown())
        self.assertTrue(ok)                          # proceeds after waiting
        self.assertEqual(len(self.slept), 1)

    def test_respect_active_cooldown_stops_when_exhausted(self):
        stop_at_the_end_of_the_ladder(self)
        wf = self.workflow()
        for _ in range(5):
            wf.cooldown.escalate("search", "rate limit")
        ok = asyncio.run(wf._respect_active_cooldown())
        self.assertFalse(ok)                         # exhausted -> device stops
        self.assertEqual(self.slept, [])             # no wait, just stop

    # no cooldown -> proceed immediately, no wait
    def test_no_cooldown_proceeds(self):
        wf = self.workflow()
        ok = asyncio.run(wf._respect_active_cooldown())
        self.assertTrue(ok)
        self.assertEqual(self.slept, [])


class DetectionUnchangedCase(unittest.TestCase):
    # 13/14/16. detection helpers unchanged; a non-throttle failure is not a
    # rate limit, and the throttle markers still classify correctly
    def test_is_throttle_reason(self):
        from student_scraper import is_throttle_reason
        self.assertTrue(is_throttle_reason("GitHub rate limit / challenge"))
        self.assertTrue(is_throttle_reason("returned HTTP 429"))
        self.assertFalse(is_throttle_reason("404 not found"))
        self.assertFalse(is_throttle_reason("navigation timeout"))


class OutageCase(unittest.TestCase):
    """A GitHub 5xx waits it out and never stops the job.

    Before this, a 503 on the search page raised Stopped on the FIRST
    occurrence -- an overnight run ended over a hiccup that cleared in
    minutes.  A 429/403 still escalates and still stops: that one is GitHub
    telling us to back off, and is respected exactly as before.
    """

    def setUp(self):
        import student_scraper as S
        self.S = S
        self.dir = tempfile.mkdtemp(prefix="outage_")
        self.ctx = sandbox.enter(os.path.join(self.dir, "runtime"))
        self.said = []
        self.slept = []
        self._say = S.say
        self._orig_sleep = student_scraper.sleep_interruptible
        S.say = self.said.append

        async def fake_sleep(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
            self.slept.append(seconds)
        student_scraper.sleep_interruptible = fake_sleep

    def tearDown(self):
        sandbox.restore()
        self.S.say = self._say
        student_scraper.sleep_interruptible = self._orig_sleep
        shutil.rmtree(self.dir, ignore_errors=True)

    def workflow(self):
        wf = self.S.Workflow.__new__(self.S.Workflow)
        wf.context = self.ctx
        wf.cooldown = student_scraper.RateLimitCooldown()
        wf.options = type("O", (), {})()
        wf.state = None
        return wf

    def text(self):
        return "\n".join(self.said)

    # ------------------------------------------------- classification ----
    def test_a_5xx_is_an_outage(self):
        from student_scraper import is_outage_reason
        for code in sorted(C.OUTAGE_STATUS):
            self.assertTrue(is_outage_reason("GitHub returned HTTP %d" % code),
                            code)
        self.assertTrue(is_outage_reason("502 Bad Gateway"))
        self.assertTrue(is_outage_reason("Service Unavailable"))

    def test_a_rate_limit_is_not_an_outage(self):
        from student_scraper import is_outage_reason
        for reason in ("GitHub returned HTTP 429", "GitHub returned HTTP 403",
                       "secondary rate limit", "404 not found",
                       "navigation timeout"):
            self.assertFalse(is_outage_reason(reason), reason)

    def test_a_5xx_behind_githubs_generic_screen_is_still_an_outage(self):
        """The SPA wraps a 503 in the same wording it uses for a 429."""
        from student_scraper import is_outage_reason, is_throttle_reason
        wrapped = "Unable to load page. Status: 503 Please reload page"
        self.assertTrue(is_outage_reason(wrapped))
        # it also matches the throttle markers, which is why the outage
        # check has to run first
        self.assertTrue(is_throttle_reason(wrapped))

    def test_a_wrapped_429_is_still_only_a_rate_limit(self):
        from student_scraper import is_outage_reason
        self.assertFalse(is_outage_reason(
            "Unable to load page. Status: 429 Please reload page"))

    def test_a_5xx_search_page_classifies_as_an_outage_reason(self):
        """End to end: the classifier's own words must read as an outage."""
        from student_scraper import PageState, classify_search_page
        from student_scraper import is_outage_reason
        state = {"title": "Error 503",
                 "bodyHead": "Unable to load page. Status: 503",
                 "bodyTail": ""}
        page_state, why = classify_search_page(state)
        self.assertIn(page_state, (PageState.HTTP_ERROR, PageState.BLOCKED))
        self.assertTrue(is_outage_reason(why), why)

    def test_the_search_branch_checks_for_an_outage_before_stopping(self):
        source = inspect.getsource(self.S.Workflow.collect)
        self.assertIn("is_outage_reason", source)
        self.assertLess(source.index("is_outage_reason"),
                        source.index("raise Stopped"))

    # ------------------------------------------------------ never stops ---
    def test_an_outage_never_exhausts(self):
        c = student_scraper.RateLimitCooldown()
        for _ in range(10):
            c.escalate("search", "GitHub returned HTTP 503",
                       now=at(0), kind="outage")
        self.assertEqual(c.level, 10)
        self.assertFalse(c.exhausted())          # would have stopped at 5

    def test_a_rate_limit_still_exhausts(self):
        """The conservative ladder is untouched for 429/403."""
        stop_at_the_end_of_the_ladder(self)
        c = student_scraper.RateLimitCooldown()
        for _ in range(5):
            c.escalate("search", "rate limit", now=at(0))
        self.assertTrue(c.exhausted())

    def test_the_outage_kind_survives_a_restart(self):
        """A restart mid-outage must not turn into a stop."""
        c = student_scraper.RateLimitCooldown()
        for _ in range(6):
            c.escalate("search", "HTTP 503", now=at(0), kind="outage")
        reloaded = student_scraper.RateLimitCooldown()
        self.assertEqual(reloaded.kind, "outage")
        self.assertFalse(reloaded.exhausted())

    def test_the_wait_stays_inside_the_band(self):
        low, high = band()
        for level in range(1, 40):
            self.assertTrue(low <= cooldown_seconds(level) <= high)

    def test_cooldown_and_wait_keeps_retrying_an_outage(self):
        wf = self.workflow()
        for _ in range(6):
            wf.cooldown.escalate("search", "HTTP 503", kind="outage")
        limit = self.S.RateLimited("GitHub returned HTTP 503",
                                   source="search", kind="outage")
        self.assertTrue(asyncio.run(wf._cooldown_and_wait(limit)))
        self.assertIn("[OUTAGE]", self.text())
        self.assertNotIn("[DEVICE STOPPED]", self.text())
        self.assertEqual(len(self.slept), 1)

    def test_a_restart_during_an_outage_proceeds_instead_of_stopping(self):
        wf = self.workflow()
        for _ in range(6):
            wf.cooldown.escalate("search", "HTTP 503", kind="outage")
        self.assertTrue(asyncio.run(wf._respect_active_cooldown()))
        self.assertNotIn("[DEVICE STOPPED]", self.text())

    def test_a_rate_limit_restart_past_the_ladder_still_stops(self):
        stop_at_the_end_of_the_ladder(self)
        wf = self.workflow()
        for _ in range(5):
            wf.cooldown.escalate("search", "rate limit")
        self.assertFalse(asyncio.run(wf._respect_active_cooldown()))
        self.assertIn("[DEVICE STOPPED]", self.text())

    def test_a_rate_limited_search_is_not_relabelled_an_outage(self):
        wf = self.workflow()
        limit = self.S.RateLimited("secondary rate limit", source="search")
        self.assertEqual(limit.kind, "rate-limit")
        asyncio.run(wf._cooldown_and_wait(limit))
        self.assertIn("[RATE LIMIT]", self.text())
        self.assertNotIn("[OUTAGE]", self.text())


if __name__ == "__main__":
    unittest.main(verbosity=2)
