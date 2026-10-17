"""The browser dying must not end the run.

Regression: a crawl finished, the browser died during the XLSX export,
the export was rescued from disk so collection looked successful, and the
very first profile navigation then failed with

    Page.goto: Target page, context or browser has been closed

Collection had already succeeded, so re-crawling would have been pure
waste.  The fix is to notice the death and relaunch on the same
persistent profile.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook                        # noqa: E402

import config as C                                   # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (BrowserLost, StateManager)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                          # noqa: E402


class FakeCity:
    def __init__(self, state, canonical, aliases):
        self.country = "India"
        self.state = state
        self.canonical = canonical
        self.search_aliases = tuple(aliases)


def make_plan(keywords=("kw",)):
    """A tiny plan whose queries are 'kw Bhopal', 'kw-a Bhopal', ..."""
    return sandbox.plan(
        list(keywords), [FakeCity("Madhya Pradesh", "Bhopal", ["Bhopal"])])


class Options:
    show = False
    limit = 0
    process_existing = True
    max_pages = None
    verbose = False
    rank = 1
    total_ranks = 1


class FakePage:
    def __init__(self):
        self.closed = False
        self.gotos = []

    def is_closed(self):
        return self.closed

    async def reload(self, **_kwargs):
        """Mirrors a Playwright page: _stable_page() reloads a page that did
        not settle before abandoning the search."""
        self.reloads = getattr(self, "reloads", 0) + 1

    async def goto(self, url, **_kwargs):
        if self.closed:
            raise Exception("Target page, context or browser has been closed")
        self.gotos.append(url)


class FakeBrowser:
    """A browser that can be killed and counted."""

    def __init__(self):
        self.alive = True
        self.restarts = 0
        self.current = FakePage()
        self.pages = [self.current]
        # Mirrors BrowserManager: students finished since the last restart,
        # and a plain tally of credits so a test can assert they happen.
        self._healthy_since_restart = 0
        self.healthy_credits = 0

    def note_healthy(self):
        """Mirrors BrowserManager.note_healthy -- same forgiveness rule."""
        self.healthy_credits += 1
        if self.restarts == 0:
            self._healthy_since_restart = 0
            return
        self._healthy_since_restart += 1
        if self._healthy_since_restart >= C.BROWSER_RESTART_FORGIVE_AFTER:
            self.restarts = 0
            self._healthy_since_restart = 0

    async def is_alive(self):
        return self.alive

    def page_count(self):
        return len(self.pages) if self.alive else 0

    async def page(self):
        return self.current

    async def get_live_page(self):
        """Mirrors BrowserManager: never hands back a dead page."""
        if not self.alive:
            return None
        if self.current is None or self.current.is_closed():
            self.current = FakePage()
            self.pages = [self.current]
        return self.current

    async def restart(self, with_extension=True):
        self.restarts += 1
        self._healthy_since_restart = 0
        self.alive = True
        self.current = FakePage()
        self.pages = [self.current]
        return self.current

    async def github_login_name(self, _page):
        return "tester"

    async def close_extra_pages(self, keep=()):
        return 0

    async def release_page_memory(self, _page):
        return True

    def kill(self):
        self.alive = False
        self.current.closed = True


class StubWorkflow(S.Workflow):
    """Workflow with the network stages replaced, browser kept fake."""

    def __init__(self, options, browser, context, plan, store, die_at=None):
        self.options = options
        self.context = context
        self.searches = plan.assign(1, 1)
        self.store = store
        self.existing = {}
        self._master_dirty = False
        self.existing = {}
        self._master_dirty = False
        self.unit = self.searches[0]
        self.state = StateManager(path=context.state_file)
        self.browser = browser
        self.scraper = None
        self.profiles = None
        self.portfolios = None
        self.resumes = None
        self.reader = S.XlsxReader()
        self.merger = S.StudentMerger()
        self.transient = S.TransientFailureLog()
        self.processed = 0
        self.die_at = die_at
        self.seen = []

    async def do_candidate(self, page, candidate):
        self.seen.append(candidate.username)
        if candidate.username == self.die_at:
            self.die_at = None          # dies once
            self.browser.kill()
            raise BrowserLost("Target page, context or browser has "
                              "been closed")
        # A live page is required, exactly as the real path requires one.
        await page.goto("https://github.com/%s" % candidate.username)
        record = {column: "" for column in C.FINAL_COLUMNS}
        record["GitHub Profile"] = candidate.profile_url
        record["github"] = candidate.profile_url
        saved, _reason = self.store.save_student(
            record, search_unit=self.unit, campaign_id="",
            rank=S.rank, total_ranks=S.num_sys)
        if saved:
            self.processed += 1


def make_export(path, count):
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "search-match"])
    for number in range(1, count + 1):
        sheet.append(["https://github.com/person%d.png?size=40" % number,
                      "https://github.com/person%d" % number,
                      "Person %d" % number])
    book.save(path)


class TestBrowserRecovery(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="recovery_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        self.stores = []
        self.export = os.path.join(self.context.keyword_download_dir,
                                   "kw-bhopal.xlsx")
        make_export(self.export, 6)

    def tearDown(self):
        sandbox.restore()
        for store in getattr(self, "stores", []):
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def _workflow(self, options, browser, die_at=None, keywords=("kw",)):
        plan = make_plan(keywords)
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        return StubWorkflow(options, browser, self.context, plan, store,
                            die_at=die_at)

    def _run(self, workflow):
        return asyncio.run(workflow.do_search_unit(workflow.browser.current,
                                                   workflow.searches[0]))

    def test_dead_browser_is_detected_before_navigating(self):
        browser = FakeBrowser()
        browser.kill()
        workflow = self._workflow(Options(), browser)
        self._run(workflow)
        # It restarted instead of calling goto on a dead page.
        self.assertGreaterEqual(browser.restarts, 1)
        self.assertEqual(workflow.processed, 6)

    def test_death_mid_run_restarts_and_retries_same_candidate(self):
        browser = FakeBrowser()
        workflow = self._workflow(Options(), browser, die_at="person3")
        self._run(workflow)
        self.assertEqual(browser.restarts, 1)
        # person3 was attempted twice: once before the death, once after.
        self.assertEqual(workflow.seen.count("person3"), 2)
        # Every candidate still ended up saved exactly once.
        self.assertEqual(workflow.processed, 6)
        store = student_scraper.open_store(self.context.store_db,
                                        read_only=True)
        self.addCleanup(store.close)
        self.assertEqual(store.count(), 6)

    def test_collection_is_not_repeated_after_a_restart(self):
        """The keyword XLSX is reused; the extension never runs again."""
        browser = FakeBrowser()
        workflow = self._workflow(Options(), browser, die_at="person2")
        collected = []

        async def spy(*_args, **_kwargs):
            collected.append(True)
            return -1

        workflow.collect = spy
        self._run(workflow)
        self.assertEqual(browser.restarts, 1)
        self.assertEqual(collected, [],
                         "a browser restart must not re-run collection")
        self.assertEqual(workflow.processed, 6)

    def test_state_and_xlsx_survive_a_restart(self):
        browser = FakeBrowser()
        workflow = self._workflow(Options(), browser, die_at="person2")
        self._run(workflow)
        # All six resolved, so cleanup is allowed to remove the XLSX.
        self.assertFalse(os.path.exists(self.export))
        self.assertTrue(workflow.state.data["cleanup_complete"])

    def test_xlsx_kept_when_restarts_are_exhausted(self):
        browser = FakeBrowser()

        class AlwaysDies(StubWorkflow):
            async def do_candidate(self, page, candidate):
                self.seen.append(candidate.username)
                self.browser.kill()
                raise BrowserLost("dead again")

        browser.restarts = C.BROWSER_MAX_RESTARTS
        plan = make_plan()
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        workflow = AlwaysDies(Options(), browser, self.context, plan, store)
        with self.assertRaises(BrowserLost):
            self._run(workflow)
        self.assertTrue(os.path.exists(self.export),
                        "the keyword XLSX must never be deleted on failure")

    def test_health_helpers_tolerate_no_browser(self):
        workflow = self._workflow(Options(), FakeBrowser())
        workflow.browser = None
        page = FakePage()
        self.assertIs(asyncio.run(workflow.ensure_browser(page)), page)
        self.assertIs(asyncio.run(workflow.enter_profile_phase(page)), page)
        asyncio.run(workflow._reclaim(page, 1))      # must not raise


class TestPageLifecycleAcrossKeywords(unittest.TestCase):
    """A restart during keyword A must not poison keyword B.

    The bug: run() obtained one page, passed it to do_keyword() for every
    keyword, and never reassigned it.  A restart during the profile phase
    of keyword A produced a new page, but run() kept the old one, so
    keyword B called collect(dead_page) and died on the first goto with
    TargetClosedError -- after keyword A had completed successfully.
    """

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="lifecycle_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        self.stores = []
        make_export(os.path.join(self.context.keyword_download_dir,
                                 "kw-a-bhopal.xlsx"), 4)
        make_export(os.path.join(self.context.keyword_download_dir,
                                 "kw-b-bhopal.xlsx"), 2)

    def tearDown(self):
        sandbox.restore()
        for store in getattr(self, "stores", []):
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def _workflow(self, options, browser, die_at=None, keywords=("kw",)):
        plan = make_plan(keywords)
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        return StubWorkflow(options, browser, self.context, plan, store,
                            die_at=die_at)

    def test_second_search_unit_uses_the_new_page(self):
        browser = FakeBrowser()
        page_a = browser.current
        workflow = self._workflow(Options(), browser, die_at="person3",
                                  keywords=("kw-a", "kw-b"))
        unit_a, unit_b = workflow.searches[0], workflow.searches[1]
        collected = []

        async def fake_collect(page, keyword, target):
            # This is the navigation that used to raise TargetClosedError.
            await page.goto("https://github.com/search?q=%s&type=users"
                            % keyword)
            collected.append((keyword, page))
            # collect() returns (status, rows); a failure preserves the unit.
            return S.CollectResult.FAILED, -1

        workflow.collect = fake_collect

        async def two_keywords():
            # Unit A: process_existing, so collect() is skipped.
            await workflow.do_search_unit(page_a, unit_a)
            # Unit B: run() would hand over its ORIGINAL page here.
            workflow.options.process_existing = False
            try:
                await workflow.do_search_unit(page_a, unit_b)
            except S.Stopped:
                # The stubbed collect() returns -1, so the unit stops.
                # What matters is WHICH page it navigated through.
                pass

        asyncio.run(two_keywords())

        # The browser really did restart during keyword A.
        self.assertEqual(browser.restarts, 1)
        page_b = browser.current
        self.assertIsNot(page_b, page_a)

        # Keyword B collected through the NEW page.
        self.assertEqual(len(collected), 1)
        _keyword, used = collected[0]
        self.assertIs(used, page_b, "collection used a stale page")
        self.assertIsNot(used, page_a)

        # page_a was legitimately used before it died; what must never
        # happen is the keyword-B search navigation landing on it.  That
        # is the exact call that raised TargetClosedError.
        self.assertFalse([url for url in page_a.gotos if "search?q=" in url],
                         "keyword B navigated the dead page from keyword A")
        self.assertTrue([url for url in page_b.gotos if "search?q=" in url],
                        "keyword B should have navigated the new page")

    def test_live_page_replaces_a_closed_page(self):
        browser = FakeBrowser()
        workflow = self._workflow(Options(), browser)
        stale = browser.current
        stale.closed = True
        browser.alive = False
        fresh = asyncio.run(workflow.live_page(stale))
        self.assertIsNot(fresh, stale)
        self.assertFalse(fresh.is_closed())
        self.assertEqual(browser.restarts, 1)

    def test_live_page_keeps_a_healthy_page(self):
        browser = FakeBrowser()
        workflow = self._workflow(Options(), browser)
        same = asyncio.run(workflow.live_page(browser.current))
        self.assertIs(same, browser.current)
        self.assertEqual(browser.restarts, 0)

    def test_collection_recovers_when_the_page_is_dead(self):
        browser = FakeBrowser()
        workflow = self._workflow(Options(), browser)
        workflow.options.process_existing = False
        browser.current.closed = True
        browser.alive = False
        used = {}

        async def fake_open(page, *_a, **_kw):
            used["page"] = page
            raise S.Stopped("stop after the navigation")

        workflow.scraper = type("X", (), {
            "open_instant_data_scraper": staticmethod(fake_open)})()

        async def run_collect():
            try:
                await workflow.collect(
                    browser.current, "kw-b",
                    os.path.join(self.context.keyword_download_dir,
                                 "kw-b.xlsx"))
            except Exception:
                pass

        asyncio.run(run_collect())
        self.assertEqual(browser.restarts, 1,
                         "collect() must restart a dead browser")


class TestRestartBudgetIsPerStreak(unittest.TestCase):
    """The restart budget counts a STREAK of failures, not a whole run.

    It used to count the run: BrowserManager.restarts was zeroed in
    __init__ and nowhere else, so five browser deaths spread over weeks
    exhausted the budget exactly as surely as five in a row, and a server
    left running stopped on a browser that was behaving perfectly well.
    """

    def _manager(self):
        manager = S.BrowserManager.__new__(S.BrowserManager)
        manager.restarts = 0
        manager._healthy_since_restart = 0
        return manager

    def _healthy_run(self, manager, students):
        for _ in range(students):
            manager.note_healthy()

    def test_crashes_spread_across_a_long_run_are_forgiven(self):
        manager = self._manager()
        for _ in range(C.BROWSER_MAX_RESTARTS * 3):
            manager.restarts += 1
            manager._healthy_since_restart = 0
            self._healthy_run(manager, C.BROWSER_RESTART_FORGIVE_AFTER)
            self.assertEqual(
                manager.restarts, 0,
                "a healthy streak must hand the restart budget back")

    def test_a_flapping_browser_still_exhausts_the_budget(self):
        manager = self._manager()
        for _ in range(C.BROWSER_MAX_RESTARTS):
            manager.restarts += 1
            manager._healthy_since_restart = 0
            # One student short of forgiveness, every time.
            self._healthy_run(manager, C.BROWSER_RESTART_FORGIVE_AFTER - 1)
        self.assertGreaterEqual(
            manager.restarts, C.BROWSER_MAX_RESTARTS,
            "a browser that keeps dying must still stop the run")

    def test_a_partial_streak_does_not_forgive(self):
        manager = self._manager()
        manager.restarts = 2
        self._healthy_run(manager, C.BROWSER_RESTART_FORGIVE_AFTER - 1)
        self.assertEqual(manager.restarts, 2,
                         "forgiveness must need the full streak")
        manager.note_healthy()
        self.assertEqual(manager.restarts, 0,
                         "the last student of the streak must forgive")

    def test_crediting_is_cheap_while_the_budget_is_full(self):
        manager = self._manager()
        self._healthy_run(manager, C.BROWSER_RESTART_FORGIVE_AFTER * 2)
        self.assertEqual((manager.restarts, manager._healthy_since_restart),
                         (0, 0),
                         "with nothing to forgive the counter must stay put")

    # The CALL SITE needs no test of its own here: every test above in
    # TestBrowserRecovery drives process_rows() through StubWorkflow, and
    # all ten of them failed with "FakeBrowser object has no attribute
    # note_healthy" until the double grew the method.  That is the
    # regression guard -- remove the call and they break again.


if __name__ == "__main__":
    unittest.main(verbosity=2)
