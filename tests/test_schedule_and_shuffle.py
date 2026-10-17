"""The random keyword order, the night stop and the work/break cycle.

    keywords shuffled once per lap, and a restart resumes the same lap in
        the same order rather than reshuffling
    no collecting inside the night window, browser closed for the wait
    after a full work period a random break, counting WORKING time only
    every switch has an off position that restores the old behaviour

Nothing here touches the real clock, the real browser or the real project
folder: the cycle file is a temp path, the browser is a stub, the pauses are
seconds, and the work clock is driven by a fake monotonic().
"""

import asyncio
import datetime
import json
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                       # noqa: E402
import student_scraper as S                              # noqa: E402


# Ten keywords: these tests are about ORDER, and the real 3,450 would only
# make them slower.
KEYWORDS = ["kw%02d" % n for n in range(1, 11)]


class FakeMonotonic:
    """A clock the test moves by hand, so a two-hour wait takes no time."""

    def __init__(self, start=1000.0):
        self.now = float(start)

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += float(seconds)


class Base(unittest.TestCase):
    """Every case edits the settings, so every case puts them back."""

    # The knobs live on student_scraper (below its imports), not in config;
    # only the state-file path is config's.
    KNOBS = ("NIGHT_PAUSE_ENABLED", "NIGHT_START_HOUR", "NIGHT_END_HOUR",
             "KEYWORD_SHUFFLE_ENABLED", "JITTER_FRACTION",
             "LAP_COOLDOWN_MINUTES", "BROWSER_RESTART_PAUSE_SECONDS",
             "WORK_PERIOD_MIN_HOURS", "WORK_PERIOD_MAX_HOURS",
             "BREAK_MIN_MINUTES", "BREAK_MAX_MINUTES",
             "SEARCH_DELAY_MIN", "SEARCH_DELAY_MAX",
             "ZERO_RESULT_DELAY_MIN", "ZERO_RESULT_DELAY_MAX",
             "PAGE_DELAY_MIN", "PAGE_DELAY_MAX",
             "PROFILE_DELAY_MIN", "PROFILE_DELAY_MAX")

    def setUp(self):
        self._knobs = {name: getattr(S, name) for name in self.KNOBS}
        self._cycle_file = C.KEYWORD_CYCLE_FILE
        self.root = tempfile.mkdtemp(prefix="kwcycle_")
        C.KEYWORD_CYCLE_FILE = os.path.join(self.root, "keyword_cycle.json")

    def tearDown(self):
        for name, value in self._knobs.items():
            setattr(S, name, value)
        C.KEYWORD_CYCLE_FILE = self._cycle_file


class ShuffleCase(Base):
    def test_it_is_a_permutation_and_not_the_file_order(self):
        order = S.shuffle_keywords(KEYWORDS, 1)
        self.assertEqual(sorted(order), sorted(KEYWORDS))
        self.assertNotEqual(order, KEYWORDS, "the shuffle did nothing")

    def test_the_same_lap_always_gives_the_same_order(self):
        """The whole point: a restart must not reshuffle."""
        self.assertEqual(S.shuffle_keywords(KEYWORDS, 4),
                         S.shuffle_keywords(KEYWORDS, 4))

    def test_a_new_lap_and_a_different_rank_both_give_new_orders(self):
        """So laps differ, and five laptops do not send the same keyword at
        the same moment."""
        lap1 = S.shuffle_keywords(KEYWORDS, 1)
        self.assertNotEqual(lap1, S.shuffle_keywords(KEYWORDS, 2))
        self.assertNotEqual(lap1, S.shuffle_keywords(KEYWORDS, 1, this_rank=4))

    def test_switched_off_gives_the_file_order_back(self):
        S.KEYWORD_SHUFFLE_ENABLED = False
        self.assertEqual(S.keyword_order(5, KEYWORDS), KEYWORDS)


class CycleFileCase(Base):
    def test_reading_does_not_create_the_file(self):
        """banner() and the report modes read this; a read that writes is how
        a read-only path leaves state in the project folder."""
        self.assertEqual(S.read_keyword_cycle(), S.KEYWORD_CYCLE_START)
        self.assertFalse(os.path.exists(C.KEYWORD_CYCLE_FILE))

    def test_opening_the_next_lap_advances_and_persists(self):
        S.save_keyword_cycle(1, KEYWORDS)
        self.assertEqual(S.open_next_keyword_cycle(KEYWORDS), 2)
        self.assertEqual(S.read_keyword_cycle(), 2)
        with open(C.KEYWORD_CYCLE_FILE, encoding="utf-8") as handle:
            self.assertEqual(json.load(handle)["keyword_cycle"], 2)

    def test_an_unreadable_file_costs_the_order_not_the_run(self):
        with open(C.KEYWORD_CYCLE_FILE, "w", encoding="utf-8") as handle:
            handle.write("{ not json")
        self.assertEqual(S.read_keyword_cycle(), S.KEYWORD_CYCLE_START)


class SearchListCase(Base):
    """The shuffle must change the ORDER and nothing else: current_row.json
    is a plain index into this list."""

    def key(self, searches):
        return sorted((city.canonical, spelling, keyword)
                      for city, spelling, keyword in searches)

    def test_the_same_searches_in_a_different_order(self):
        plain = S.build_search_list(1, 400, say_total=False)
        shuffled = S.build_search_list(
            1, 400, say_total=False,
            keywords=S.shuffle_keywords(S.load_base_keywords(), 2))
        self.assertEqual(len(shuffled), len(plain))
        self.assertEqual(self.key(shuffled), self.key(plain),
                         "the shuffle changed WHICH searches run")
        self.assertNotEqual([kw for _, _, kw in shuffled],
                            [kw for _, _, kw in plain])

    def test_no_keywords_argument_is_the_old_behaviour(self):
        self.assertEqual(S.build_search_list(1, 400, say_total=False),
                         S.build_search_list(1, 400, say_total=False,
                                             keywords=S.load_base_keywords()))


class NightWindowCase(Base):
    def at(self, hour, minute=0):
        return datetime.datetime(2026, 10, 17, hour, minute)

    def setUp(self):
        super().setUp()
        S.NIGHT_PAUSE_ENABLED = True
        S.NIGHT_START_HOUR, S.NIGHT_END_HOUR = 0, 6

    def test_it_waits_inside_the_window_and_not_outside_it(self):
        self.assertEqual(S.night_seconds_remaining(self.at(1, 30)),
                         int(4.5 * 3600))
        self.assertEqual(S.night_seconds_remaining(self.at(7)), 0)

    def test_a_window_may_cross_midnight(self):
        S.NIGHT_START_HOUR, S.NIGHT_END_HOUR = 23, 5
        self.assertEqual(S.night_seconds_remaining(self.at(23, 30)),
                         int(5.5 * 3600))
        self.assertEqual(S.night_seconds_remaining(self.at(4)), 3600)
        self.assertEqual(S.night_seconds_remaining(self.at(22)), 0)

    def test_start_equals_end_is_no_window(self):
        S.NIGHT_START_HOUR = S.NIGHT_END_HOUR = 5
        self.assertEqual(S.night_seconds_remaining(self.at(5, 30)), 0)

    def test_switched_off_is_silent_at_every_hour(self):
        """It is checked HERE and nowhere else, so off must mean no message,
        no wait and no browser restart anywhere."""
        S.NIGHT_PAUSE_ENABLED = False
        for hour in range(24):
            self.assertEqual(S.night_seconds_remaining(self.at(hour)), 0)

    def test_every_knob_sits_below_the_imports(self):
        """They are edited per laptop, so they are the first thing in the
        file after the imports -- not 600 lines into a tuning file."""
        with open(S.__file__.replace(".pyc", ".py"), encoding="utf-8") as fh:
            source = fh.read()
        top = source[:source.index("# -- per-laptop override")]
        self.assertIn("import config as C", top)
        self.assertIn("rank = 1", top)
        for name in self.KNOBS:
            self.assertIn("%s =" % name, top, "%s moved away" % name)
            self.assertFalse(hasattr(C, name), "%s is in config" % name)


class WorkClockCase(Base):
    def test_a_pause_is_not_working_time(self):
        clock = FakeMonotonic()
        with mock.patch.object(S.time, "monotonic", clock):
            S.reset_work_timer()
            clock.advance(10 * 60)            # ten minutes of collecting
            S.work_clock_pause()
            clock.advance(90 * 60)            # ninety minutes of waiting
            S.work_clock_resume()
            self.assertAlmostEqual(S.work_seconds(), 10 * 60, delta=1)

    def test_pauses_nest(self):
        """Two guards can be waiting at once; the clock must stop once."""
        S.reset_work_timer()
        S.work_clock_pause()
        S.work_clock_pause()
        S.work_clock_resume()
        self.assertIsNone(S._WORK_MARK, "the clock restarted too early")
        S.work_clock_resume()
        self.assertIsNotNone(S._WORK_MARK, "the clock never restarted")

    def test_the_period_and_the_break_are_redrawn_in_range(self):
        S.WORK_PERIOD_MIN_HOURS, S.WORK_PERIOD_MAX_HOURS = 2, 6
        S.BREAK_MIN_MINUTES, S.BREAK_MAX_MINUTES = 30, 60
        periods = {S.reset_work_timer() for _ in range(40)}
        breaks = {S.pick_break_seconds() for _ in range(40)}
        self.assertTrue(all(2 <= hours <= 6 for hours in periods), periods)
        self.assertTrue(all(1800 <= s <= 3600 for s in breaks), breaks)
        self.assertGreater(len(periods), 1, "every period was the same")
        self.assertGreater(len(breaks), 1, "every break was the same")

    def test_zero_hours_disables_the_cycle(self):
        S.WORK_PERIOD_MIN_HOURS = S.WORK_PERIOD_MAX_HOURS = 0
        self.assertEqual(S.reset_work_timer(), 0.0)
        S._WORK_BANKED = 99 * 3600
        self.assertIsNone(S.due_pause())


class DuePauseCase(Base):
    def setUp(self):
        super().setUp()
        S.NIGHT_PAUSE_ENABLED = False
        S.WORK_PERIOD_MIN_HOURS = S.WORK_PERIOD_MAX_HOURS = 2
        S.BREAK_MIN_MINUTES, S.BREAK_MAX_MINUTES = 30, 60
        S.reset_work_timer()

    def test_nothing_is_due_at_the_start_of_a_period(self):
        self.assertIsNone(S.due_pause())

    def test_a_full_period_earns_a_break(self):
        S._WORK_BANKED = 2 * 3600
        reason, seconds = S.due_pause()
        self.assertIn("work period reached", reason)
        self.assertTrue(1800 <= seconds <= 3600, seconds)

    def test_cooldowns_do_not_earn_a_break(self):
        """The clock counts collecting, not waiting."""
        clock = FakeMonotonic()
        with mock.patch.object(S.time, "monotonic", clock):
            S.reset_work_timer()
            S.work_clock_pause()
            clock.advance(2 * 3600)           # two hours of cooldown
            S.work_clock_resume()
            self.assertIsNone(S.due_pause())
            clock.advance(2 * 3600)           # two hours of collecting
            self.assertIsNotNone(S.due_pause())

    def test_the_night_stop_wins_over_a_due_break(self):
        """A period ending at 23:58 must not take both."""
        S._WORK_BANKED = 99 * 3600
        S.NIGHT_PAUSE_ENABLED = True
        S.NIGHT_START_HOUR, S.NIGHT_END_HOUR = 0, 23
        reason, _ = S.due_pause(datetime.datetime(2026, 10, 17, 2))
        self.assertTrue(reason.startswith("night stop"), reason)


class FakeBrowser:
    def __init__(self):
        self.playwright, self.context, self._page = "pw", "ctx", "page"
        self._context_closed = False
        self.closed = 0

    async def close(self):
        self.closed += 1


def stub_workflow():
    """A Workflow built the way the other Workflow tests build one."""
    workflow = S.Workflow.__new__(S.Workflow)
    workflow.browser = FakeBrowser()
    workflow.scraper, workflow.profiles = "scraper", "profiles"
    workflow._master_dirty = False
    workflow.searches = ["the-previous-list"]
    workflow.keyword_cycle = S.KEYWORD_CYCLE_START
    workflow.restarts = 0
    workflow.revalidated = 0

    async def restart_browser():
        workflow.restarts += 1
        return "fresh-page"

    async def live_page(page=None):
        workflow.revalidated += 1
        return page

    workflow.restart_browser, workflow.live_page = restart_browser, live_page
    return workflow


class ScheduledPauseCase(Base):
    def setUp(self):
        super().setUp()
        S.NIGHT_PAUSE_ENABLED = False
        S.WORK_PERIOD_MIN_HOURS = S.WORK_PERIOD_MAX_HOURS = 2
        S.BROWSER_RESTART_PAUSE_SECONDS = 600
        S.reset_work_timer()

    def test_nothing_due_leaves_the_loop_alone(self):
        workflow = stub_workflow()
        self.assertEqual(asyncio.run(workflow._schedule_gate("page")), "page")
        self.assertEqual(workflow.browser.closed, 0)
        self.assertEqual(workflow.restarts, 0)

    def test_a_short_pause_keeps_the_browser(self):
        workflow = stub_workflow()
        page = asyncio.run(workflow._long_pause("page", 1, "short break"))
        self.assertEqual((workflow.browser.closed, workflow.restarts), (0, 0))
        self.assertEqual(page, "page")
        self.assertEqual(workflow.revalidated, 1, "page not re-validated")

    def test_a_long_pause_closes_chromium_and_relaunches(self):
        S.BROWSER_RESTART_PAUSE_SECONDS = 1
        workflow = stub_workflow()
        page = asyncio.run(workflow._long_pause("page", 1, "long break"))
        self.assertEqual(workflow.browser.closed, 1)
        self.assertIsNone(workflow.browser.context)
        self.assertTrue(workflow.browser._context_closed)
        self.assertIsNone(workflow.scraper, "a dead scraper was kept")
        self.assertEqual((workflow.restarts, page), (1, "fresh-page"))

    def test_a_pause_ends_the_period_and_leaves_the_clock_running(self):
        """A leaked pause would stop the clock for the rest of the run, and
        then no later break would ever come due."""
        S._WORK_BANKED = 99 * 3600
        asyncio.run(stub_workflow()._long_pause("page", 1, "break"))
        self.assertIsNone(S.due_pause())
        self.assertEqual(S._WORK_PAUSE_DEPTH, 0)
        self.assertIsNotNone(S._WORK_MARK, "the work clock stayed paused")

    def test_zero_seconds_does_nothing_at_all(self):
        workflow = stub_workflow()
        self.assertEqual(asyncio.run(workflow._long_pause("page", 0, "no")),
                         "page")
        self.assertEqual((workflow.browser.closed, workflow.restarts,
                          workflow.revalidated), (0, 0, 0))


class LapCase(Base):
    def test_the_lap_wrap_reshuffles_and_rebuilds_the_list(self):
        S.KEYWORD_SHUFFLE_ENABLED = True
        S.save_keyword_cycle(1)
        workflow = stub_workflow()
        first = workflow.next_keyword_cycle()
        self.assertEqual(workflow.keyword_cycle, 2)
        self.assertEqual(S.read_keyword_cycle(), 2,
                         "the lap was not persisted before its first search")
        self.assertEqual(len(first), len(S.build_search_list(say_total=False)))
        workflow.searches = first
        second = workflow.next_keyword_cycle()
        self.assertNotEqual([kw for _, _, kw in first],
                            [kw for _, _, kw in second],
                            "two laps ran in the same order")

    def test_with_the_shuffle_off_the_lap_wrap_does_nothing(self):
        S.KEYWORD_SHUFFLE_ENABLED = False
        S.save_keyword_cycle(4)
        workflow = stub_workflow()
        self.assertEqual(workflow.next_keyword_cycle(), ["the-previous-list"])
        self.assertEqual(S.read_keyword_cycle(), 4,
                         "the lap advanced with the shuffle off")

    def test_a_failure_to_reshuffle_keeps_the_previous_order(self):
        S.KEYWORD_SHUFFLE_ENABLED = True
        workflow = stub_workflow()
        with mock.patch.object(S, "keyword_order",
                               side_effect=RuntimeError("unreadable")):
            self.assertEqual(workflow.next_keyword_cycle(),
                             ["the-previous-list"])

    def test_the_lap_cooldown_closes_chromium(self):
        """45 minutes is well past the pause length that closes the browser,
        so a lap boundary must not hold an idle Chromium."""
        self.assertGreaterEqual(int(S.LAP_COOLDOWN_MINUTES) * 60,
                                int(S.BROWSER_RESTART_PAUSE_SECONDS))


class JitterCase(Base):
    """sleep_jitter(): a computed wait (a retry backoff) spread out."""

    def setUp(self):
        super().setUp()
        self.slept = []

        async def fake_sleep(seconds, check_stop=None, *a, **k):
            self.slept.append((seconds, check_stop))

        patch = mock.patch.object(S, "sleep_interruptible", fake_sleep)
        patch.start()
        self.addCleanup(patch.stop)

    def test_the_delay_is_the_base_plus_up_to_the_fraction(self):
        S.JITTER_FRACTION = 0.35
        drawn = {round(asyncio.run(S.sleep_jitter(10)), 6) for _ in range(40)}
        self.assertTrue(all(10 <= d <= 13.5 for d in drawn), drawn)
        self.assertGreater(len(drawn), 1, "the backoff was not jittered")

    def test_a_zero_fraction_is_the_base_exactly(self):
        S.JITTER_FRACTION = 0.0
        self.assertEqual(asyncio.run(S.sleep_jitter(7)), 7.0)

    def test_nothing_to_wait_for_is_not_a_wait(self):
        for base in (0, -5, None):
            self.assertEqual(asyncio.run(S.sleep_jitter(base)), 0.0)
        self.assertEqual(self.slept, [])

    def test_it_stays_interruptible(self):
        """A backoff is a wait like any other: Ctrl+C must reach it."""
        def stop():
            return None
        asyncio.run(S.sleep_jitter(5, stop))
        self.assertIs(self.slept[0][1], stop)


class BannerCase(Base):
    def test_it_reports_the_order_the_schedule_and_the_pacing(self):
        """Two copies differ by their keyword order and pacing now, so both
        have to be on screen at startup."""
        said = []
        searches = S.build_search_list(1, 400, say_total=False)[:3]
        with mock.patch.object(S, "say", said.append):
            S.banner(searches, {})
        printed = "\n".join(said)
        for expected in ("Keyword order", "Night stop", "Work / break",
                         "%g-%g s" % (S.SEARCH_DELAY_MIN, S.SEARCH_DELAY_MAX),
                         "%g-%g s" % (S.PAGE_DELAY_MIN, S.PAGE_DELAY_MAX),
                         "%g-%g s" % (S.PROFILE_DELAY_MIN,
                                      S.PROFILE_DELAY_MAX),
                         "%d min" % int(S.LAP_COOLDOWN_MINUTES)):
            self.assertIn(expected, printed)

    def test_it_cannot_stop_a_run(self):
        """It is the first thing a run does; a broken keyword file must cost
        one line of the banner, not the run."""
        said = []
        searches = S.build_search_list(1, 400, say_total=False)[:3]
        with mock.patch.object(S, "say", said.append), \
             mock.patch.object(S, "keyword_order",
                               side_effect=RuntimeError("unreadable")):
            S.banner(searches, {})
        self.assertIn("unavailable", "\n".join(said))

    def test_the_work_line_reads_right_before_and_after_the_first_draw(self):
        """Reading the drawn period alone said "disabled" when the cycle was
        switched on but the clock had not started."""
        S.WORK_PERIOD_MIN_HOURS, S.WORK_PERIOD_MAX_HOURS = 2, 6
        S._WORK_PERIOD_HOURS = 0.0
        self.assertIn("2.0-6.0 h", S.work_break_status())
        S.WORK_PERIOD_MIN_HOURS = S.WORK_PERIOD_MAX_HOURS = 3
        S.reset_work_timer()
        self.assertIn("3.0 h of work this period", S.work_break_status())


if __name__ == "__main__":
    unittest.main(verbosity=2)
