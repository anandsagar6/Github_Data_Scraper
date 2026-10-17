"""A multi-page crawl must PROVE it paginated before it is called complete.

THE LIVE LOSS THIS FILE EXISTS FOR
----------------------------------
    [SEARCH] 10 result(s) and a next page
    [SCRAPER] pagination selector (stable): a[rel="next"]
    [SCRAPER] Start Crawling clicked
    [SCRAPER] last page reached (results present, no Next) - 10 row(s)
    [SCRAPER] collected 10 row(s) over 1 page(s)
    [OK] 10 rows

"University Student Visakhapatnam" has 19 users over 2 pages.  Ten were
collected, the search unit was accepted as complete, and NINE PEOPLE WERE
SILENTLY LOST.

The captured state at the decision point proves what happened:

    url                  ...&type=users&p=2   the browser DID navigate
    nextDisabled         True                 p2 really is the last page
    resultsListChildren  9                    nine users ON SCREEN
    ext_pages            1                    only one page ever scraped
    ext_error            Table not found      the extension lost p2
    ext_instructions     "Crawling stopped. Please download data or
                          continue crawling."

    page header          "Users 19 (19) results"
    pagination           "Previous 1 2 Next"
    10 collected + 9 uncollected = 19.  Exactly.

From the page alone, "the last page, with results on it" is
INDISTINGUISHABLE from "the extension failed to scrape the last page".  Only
the crawl's own history can tell them apart, which is what CrawlProgress is.

WHY ONE CHECK IS NOT ENOUGH
---------------------------
A bare "did any page transition happen?" test would have PASSED this case --
a transition did happen.  Two independent guarantees are required:

    transitions >= 1            the crawler actually moved
    ext_pages >= browser page   every page reached was also SCRAPED
"""

import asyncio
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                      # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (CrawlOutcome, CrawlProgress, classify_page_state, page_number)   # noqa: E402
import student_scraper as S                              # noqa: E402
from test_single_page import (FakeBrowser, FakePage,      # noqa: E402
                              FakePopup, Options, StubScraper,
                              make_export, page_state, scraper)
from test_zero_results import (ZeroResultsCase,            # noqa: E402
                               zero_state)

SEARCH = "https://github.com/search?q=University+Student+Visakhapatnam&type=users"


# ================================================== 1. the page number ======
class TestPageNumber(unittest.TestCase):
    def test_no_page_parameter_is_page_one(self):
        self.assertEqual(page_number(SEARCH), 1)
        self.assertEqual(page_number(""), 1)
        self.assertEqual(page_number(None), 1)

    def test_the_p_parameter_is_the_page(self):
        self.assertEqual(page_number(SEARCH + "&p=2"), 2)
        self.assertEqual(page_number(SEARCH + "&p=17"), 17)

    def test_a_nonsense_page_parameter_is_page_one(self):
        self.assertEqual(page_number(SEARCH + "&p=abc"), 1)
        self.assertEqual(page_number(SEARCH + "&p=0"), 1)


# ============================================ 2. the completeness contract ==
class TestCrawlProgress(unittest.TestCase):
    def progress(self, had_next=True, url=SEARCH):
        return CrawlProgress(had_next=had_next, start_url=url)

    # -- the exact live failure -------------------------------------------
    def test_the_live_failure_is_refused(self):
        """Browser reached p=2, extension scraped 1 page: rows are missing."""
        p = self.progress()
        p.note_stats(1, 10)
        p.note_url(SEARCH + "&p=2")
        p.note_stats(1, 10)
        ok, why = p.verify_complete()
        self.assertFalse(ok)
        self.assertIn("reached page 2", why)
        self.assertIn("never collected", why)

    def test_a_bare_transition_check_would_not_have_caught_it(self):
        """Documents why the second guarantee exists."""
        p = self.progress()
        p.note_stats(1, 10)
        p.note_url(SEARCH + "&p=2")
        p.note_stats(1, 10)
        self.assertGreaterEqual(p.transitions, 1,
                               "a transition DID happen in the live failure")
        self.assertFalse(p.verify_complete()[0],
                         "yet the crawl is provably incomplete")

    # -- requirement 1: had Next + navigation happened -> success ---------
    def test_had_next_and_pagination_happened_is_complete(self):
        p = self.progress()
        p.note_stats(1, 10)
        p.note_url(SEARCH + "&p=2")
        p.note_stats(2, 19)
        ok, why = p.verify_complete()
        self.assertTrue(ok, why)
        self.assertEqual(p.transitions, 1)

    # -- requirement 2: had Next + zero transitions -> failure ------------
    def test_had_next_and_zero_transitions_is_refused(self):
        p = self.progress()
        p.note_stats(1, 10)
        ok, why = p.verify_complete()
        self.assertFalse(ok)
        self.assertEqual(p.transitions, 0)
        self.assertIn("never left page 1", why)

    # -- requirement 4: three pages -> two transitions --------------------
    def test_a_three_page_crawl_records_two_transitions(self):
        p = self.progress()
        p.note_stats(1, 10)
        p.note_url(SEARCH + "&p=2")
        p.note_stats(2, 20)
        p.note_url(SEARCH + "&p=3")
        p.note_stats(3, 26)
        self.assertEqual(p.transitions, 2)
        self.assertEqual(p.max_browser_page, 3)
        self.assertEqual(p.ext_pages, 3)
        self.assertTrue(p.verify_complete()[0])

    def test_a_ten_page_crawl_records_nine_transitions(self):
        p = self.progress()
        p.note_stats(1, 10)
        for number in range(2, 11):
            p.note_url("%s&p=%d" % (SEARCH, number))
            p.note_stats(number, number * 10)
        self.assertEqual(p.transitions, 9)
        self.assertTrue(p.verify_complete()[0])

    # -- single page and zero results are unaffected ----------------------
    def test_a_single_page_with_no_next_is_complete(self):
        p = self.progress(had_next=False)
        p.note_stats(1, 7)
        self.assertTrue(p.verify_complete()[0])

    def test_no_next_and_no_movement_is_complete(self):
        """A one-page result set must not be forced to paginate."""
        p = self.progress(had_next=False)
        p.note_stats(1, 10)
        self.assertEqual(p.transitions, 0)
        self.assertTrue(p.verify_complete()[0])

    # -- assorted ---------------------------------------------------------
    def test_the_same_url_seen_twice_is_not_a_transition(self):
        p = self.progress()
        p.note_stats(1, 10)
        for _poll in range(5):
            p.note_url(SEARCH)
        self.assertEqual(p.transitions, 0)

    def test_going_backwards_is_not_a_transition(self):
        p = self.progress()
        p.note_url(SEARCH + "&p=2")
        p.note_url(SEARCH)
        self.assertEqual(p.transitions, 1)
        self.assertEqual(p.max_browser_page, 2)

    def test_starting_mid_pagination_counts_from_there(self):
        p = self.progress(url=SEARCH + "&p=3")
        self.assertEqual(p.start_page, 3)
        p.note_stats(3, 30)
        p.note_url(SEARCH + "&p=4")
        p.note_stats(4, 38)
        self.assertEqual(p.transitions, 1)
        self.assertTrue(p.verify_complete()[0])

    def test_counters_never_go_backwards(self):
        p = self.progress()
        p.note_stats(3, 30)
        p.note_stats(1, 10)        # a stale read must not lower them
        self.assertEqual(p.ext_pages, 3)
        self.assertEqual(p.ext_rows, 30)

    def test_describe_names_every_number(self):
        p = self.progress()
        p.note_stats(2, 19)
        p.note_url(SEARCH + "&p=2")
        text = p.describe()
        for fragment in ("pages visited", "transitions", "extension pages",
                         "rows"):
            self.assertIn(fragment, text)


# ================================= 3. an uninspectable page is not an answer =
class TestStalePageCannotClaimCompletion(unittest.TestCase):
    """Requirement 5: a stale handle must never produce 'last page reached'."""

    def test_a_page_error_is_table_lost_not_end_of_results(self):
        outcome, why = classify_page_state(
            {"page_error": "Target page, context or browser has been closed"})
        self.assertEqual(outcome, CrawlOutcome.TABLE_LOST)
        self.assertNotEqual(outcome, CrawlOutcome.END_OF_RESULTS)
        self.assertIn("could not be inspected", why)

    def test_a_page_error_is_not_exportable(self):
        outcome, _why = classify_page_state({"page_error": "detached"})
        self.assertNotIn(outcome, CrawlOutcome.EXPORTABLE)

    def test_an_empty_capture_never_reads_as_the_last_page(self):
        """Before the fix, all-None fields fell through to END_OF_RESULTS."""
        outcome, _why = classify_page_state(
            {"page_error": "Execution context was destroyed",
             "hasNext": None, "resultsListChildren": None,
             "profileLinkCount": None})
        self.assertNotEqual(outcome, CrawlOutcome.END_OF_RESULTS)

    def test_a_real_last_page_is_still_end_of_results(self):
        """The healthy path is unchanged."""
        outcome, _why = classify_page_state(
            {"title": "User search results", "bodyHead": "users", "bodyTail": "",
             "hasNext": False, "nextDisabled": None, "profileLinkCount": 4,
             "resultsListPresent": True, "resultsListChildren": 4,
             "noResults": False, "saysError": False, "blocked_marker": None,
             "ext_pages": 3, "ext_rows": 24, "ext_error": "Table not found"})
        self.assertEqual(outcome, CrawlOutcome.END_OF_RESULTS)

    def test_incomplete_is_not_exportable_as_complete(self):
        self.assertNotIn(CrawlOutcome.INCOMPLETE, CrawlOutcome.EXPORTABLE)


# ================================== 4. the detector loop, end to end ========
class LoopPopup(FakePopup):
    """A popup that plays back a scripted crawl, poll by poll."""

    def __init__(self, script, error_at_end="", **kwargs):
        super().__init__(**kwargs)
        self.script = list(script)       # [(pages, rows, url), ...]
        self.step = -1
        self.error_at_end = error_at_end
        self.has_start = True

    def advance(self, page):
        if self.step + 1 < len(self.script):
            self.step += 1
        pages, rows, url = self.script[max(self.step, 0)]
        self.rows = rows
        self._pages = pages
        page.url = url
        finished = self.step >= len(self.script) - 1
        self.error = self.error_at_end if finished else ""
        # The extension re-enables Start and clears Stop when it halts.
        self.has_start = True
        return finished

    async def inner_text(self, selector, timeout=None):
        if selector == C.SEL_STATS:
            return "Pages scraped: %d Rows collected: %d" % (self._pages,
                                                             self.rows)
        if selector == C.SEL_ERROR:
            return self.error
        if selector == C.SEL_INSTRUCTIONS:
            return "Crawling stopped. Please download data or continue crawling."
        return ""

    def locator(self, selector):
        from test_single_page import FakeLocator
        if selector == C.SEL_STATS:
            return FakeLocator(count=1, text="Pages scraped: %d Rows "
                                             "collected: %d"
                                             % (self._pages, self.rows))
        if selector == C.SEL_ERROR:
            return FakeLocator(count=1 if self.error else 0, text=self.error)
        if selector == C.SEL_INSTRUCTIONS:
            return FakeLocator(
                count=1,
                text="Crawling stopped. Please download data or continue "
                     "crawling.")
        if selector == C.SEL_STOP_CRAWL:
            return FakeLocator(count=1 if self.crawling else 0)
        if selector == C.SEL_START_CRAWL:
            return FakeLocator(count=1 if self.has_start else 0)
        return FakeLocator(count=0)

    crawling = False


class TestDetectorLoop(unittest.TestCase):
    """Drive the real wait_for_crawl_complete() over a scripted crawl."""

    def run_loop(self, script, had_next, page_state_at_end=None,
                 error_at_end="Table not found"):
        instance, said = scraper()
        popup = LoopPopup(script, error_at_end=error_at_end)
        popup._pages, popup.rows = script[0][0], script[0][1]
        page = FakePage(state=page_state_at_end or page_state(
            hasNext=False, nextByShape=0, nextByText=0,
            resultsListChildren=4, profileLinkCount=4,
            resultProfileLinkCount=4))
        page.url = script[0][2]

        # The popup advances one script step per poll.
        original_read = instance.read_stats

        async def read_stats(target):
            popup.advance(page)
            return await original_read(target)

        instance.read_stats = read_stats
        saved_poll, saved_stable = C.CRAWL_POLL_SECONDS, C.CRAWL_STABLE_SECONDS
        C.CRAWL_POLL_SECONDS, C.CRAWL_STABLE_SECONDS = 0.0, 0
        try:
            outcome, pages, rows = asyncio.run(
                instance.wait_for_crawl_complete(popup, page,
                                                 had_next=had_next))
        finally:
            C.CRAWL_POLL_SECONDS = saved_poll
            C.CRAWL_STABLE_SECONDS = saved_stable
        return outcome, pages, rows, said, instance.progress

    def test_the_live_failure_is_now_incomplete(self):
        """2 real pages, extension scraped 1 -> INCOMPLETE, not success."""
        outcome, pages, rows, said, progress = self.run_loop(
            script=[(1, 10, SEARCH),
                    (1, 10, SEARCH + "&p=2")],
            had_next=True,
            page_state_at_end=page_state(
                hasNext=True, nextByShape=1, nextDisabled=True,
                resultsListChildren=9, profileLinkCount=11,
                resultProfileLinkCount=9))
        self.assertEqual(outcome, CrawlOutcome.INCOMPLETE)
        self.assertNotIn(outcome, CrawlOutcome.EXPORTABLE)
        self.assertTrue(any("INCOMPLETE" in line for line in said), said)
        self.assertTrue(any("never collected" in line for line in said), said)
        self.assertEqual(progress.max_browser_page, 2)
        self.assertEqual(progress.ext_pages, 1)

    def test_a_healthy_two_page_crawl_finishes(self):
        outcome, pages, rows, said, progress = self.run_loop(
            script=[(1, 10, SEARCH),
                    (2, 19, SEARCH + "&p=2")],
            had_next=True)
        self.assertIn(outcome, CrawlOutcome.EXPORTABLE)
        self.assertEqual(rows, 19)
        self.assertEqual(progress.transitions, 1)
        self.assertEqual(progress.ext_pages, 2)

    def test_a_three_page_crawl_logs_two_transitions(self):
        outcome, pages, rows, said, progress = self.run_loop(
            script=[(1, 10, SEARCH),
                    (2, 20, SEARCH + "&p=2"),
                    (3, 26, SEARCH + "&p=3")],
            had_next=True)
        self.assertIn(outcome, CrawlOutcome.EXPORTABLE)
        self.assertEqual(rows, 26)
        self.assertEqual(progress.transitions, 2)
        transitions = [line for line in said if "] transition " in line]
        self.assertEqual(len(transitions), 2, transitions)
        self.assertIn("page 1 -> 2", transitions[0])
        self.assertIn("page 2 -> 3", transitions[1])

    def test_a_crawler_that_never_navigates_is_incomplete(self):
        outcome, _pages, _rows, said, progress = self.run_loop(
            script=[(1, 10, SEARCH), (1, 10, SEARCH)],
            had_next=True)
        self.assertEqual(outcome, CrawlOutcome.INCOMPLETE)
        self.assertEqual(progress.transitions, 0)
        self.assertTrue(any("never left page 1" in line for line in said),
                        said)

    def test_a_single_page_crawl_is_unaffected(self):
        outcome, _pages, rows, said, progress = self.run_loop(
            script=[(1, 7, SEARCH)], had_next=False)
        self.assertIn(outcome, CrawlOutcome.EXPORTABLE)
        self.assertEqual(rows, 7)
        self.assertEqual(progress.transitions, 0)
        self.assertEqual(
            [line for line in said if "] transition " in line], [])

    def test_the_before_start_line_reports_the_pre_crawl_facts(self):
        _outcome, _pages, _rows, said, _progress = self.run_loop(
            script=[(1, 10, SEARCH), (2, 19, SEARCH + "&p=2")],
            had_next=True)
        before = [line for line in said if "before start" in line]
        self.assertEqual(len(before), 1, said)
        self.assertIn("Next present: yes", before[0])

    def test_logging_stays_concise_on_a_long_crawl(self):
        script = [(1, 10, SEARCH)]
        for number in range(2, 13):
            script.append((number, number * 10, "%s&p=%d" % (SEARCH, number)))
        _outcome, _pages, _rows, said, progress = self.run_loop(
            script=script, had_next=True)
        self.assertEqual(progress.transitions, 11)
        # One line per transition, plus before-start and final.  Not per poll.
        self.assertLessEqual(len(said), 11 + 6, said)


# ============================= 5. collect() treats INCOMPLETE as a failure ===
class TestCollectIncomplete(ZeroResultsCase):
    """The IDS-crawl route and its completion guard.

    Multi-page searches now go page-by-page, so this route is reached only in
    the residual case it was always meant for: the page classifies as ONE page
    of results, but ensure_next_selector() still finds a usable next selector
    (GitHub renders a present-but-inert Next on a last page).  The guard has to
    hold there too, so these tests drive it with a single-page page state and a
    PAGINATED verdict.
    """

    def workflow_for(self, outcome, rows=10):
        page = FakePage(state=page_state())          # SINGLE page of results
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        from student_scraper import StateManager
        workflow = S.Workflow.__new__(S.Workflow)
        workflow.options = Options()
        workflow.context = self.context
        workflow.searches = self.assigned
        workflow.store = store
        workflow.existing = {}
        workflow._master_dirty = False
        workflow.unit = self.unit
        workflow.state = StateManager(path=self.context.state_file)
        workflow.browser = FakeBrowser(page)
        workflow.scraper = StubScraper(S.NextControl.PAGINATED, FakePopup(),
                                       rows=rows, outcome=outcome)
        workflow.profiles = None
        workflow.portfolios = None
        workflow.resumes = None
        workflow.merger = S.StudentMerger()
        workflow.reader = S.XlsxReader()
        workflow.transient = S.TransientFailureLog()
        workflow.processed = 0
        workflow.keyword_number = 1
        workflow.keyword_total = len(self.assigned)
        return workflow, page

    def test_incomplete_is_reported_as_failed_not_as_rows(self):
        workflow, page = self.workflow_for(CrawlOutcome.INCOMPLETE)
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(rows, -1)

    def test_incomplete_keeps_the_rows_it_did_get_as_a_partial(self):
        """Real rows are never thrown away; they are just not 'complete'."""
        workflow, page = self.workflow_for(CrawlOutcome.INCOMPLETE)
        asyncio.run(workflow.collect(page, self.unit.search_query,
                                     self.target))
        partial = self.target.replace(".xlsx", ".partial.xlsx")
        self.assertIn(partial, workflow.scraper.exports)
        self.assertFalse(os.path.exists(self.target),
                         "the real target must NOT be written")

    def test_an_incomplete_search_unit_preserves_and_stops(self):
        workflow, page = self.workflow_for(CrawlOutcome.INCOMPLETE)
        with self.assertRaises(S.Stopped):
            asyncio.run(workflow.do_search_unit(page, self.unit))
        self.assertFalse(workflow.state.collection_complete)
        self.assertFalse(workflow.state.data["cleanup_complete"])
        self.assertEqual(workflow.store.count(), 0)

    def test_a_finished_crawl_still_exports_normally(self):
        workflow, page = self.workflow_for(CrawlOutcome.FINISHED)
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 10)
        self.assertEqual(workflow.scraper.exports, [self.target])

    def test_a_multi_page_search_now_goes_page_by_page(self):
        """Multi-page no longer uses the IDS cross-page crawl at all."""
        taken = {}

        async def fake_page_by_page(page, target, first_state):
            taken["called"] = True
            return S.CollectResult.RESULTS, 19

        workflow, page = self.workflow_for(CrawlOutcome.FINISHED)
        page.state = page_state(hasNext=True, nextByShape=1)   # MULTI page
        workflow._collect_page_by_page = fake_page_by_page
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 19)
        self.assertTrue(taken.get("called"),
                        "a multi-page state must dispatch to page-by-page")
        self.assertEqual(workflow.scraper.crawls, 0,
                         "Start Crawling must not be used across pages")

    def test_the_completion_guard_still_receives_the_pre_crawl_fact(self):
        """had_next is False on this route, and must be passed explicitly."""
        seen = {}

        class Recording(StubScraper):
            async def wait_for_crawl_complete(self, popup, page,
                                              max_pages=None, had_next=False):
                seen["had_next"] = had_next
                self.crawls += 1
                return CrawlOutcome.FINISHED, 1, 10

        workflow, page = self.workflow_for(CrawlOutcome.FINISHED)
        workflow.scraper = Recording(S.NextControl.PAGINATED, FakePopup(),
                                     rows=10)
        asyncio.run(workflow.collect(page, self.unit.search_query,
                                     self.target))
        self.assertIn("had_next", seen, "the argument must be passed")
        self.assertFalse(seen["had_next"])

    def test_a_single_page_state_never_reaches_the_crawl_detector(self):
        """No pagination proof is needed where there is no pagination.

        A single-page result set takes finish_single_page(), so
        wait_for_crawl_complete() -- and its had_next requirement -- is not
        consulted at all.
        """
        calls = []

        class Recording(StubScraper):
            async def wait_for_crawl_complete(self, popup, page,
                                              max_pages=None, had_next=False):
                calls.append(had_next)
                return CrawlOutcome.FINISHED, 1, 7

        workflow, page = self.workflow_for(CrawlOutcome.FINISHED)
        page.state = page_state()          # single page, no Next
        workflow.scraper = Recording(S.NextControl.SINGLE_PAGE, FakePopup(),
                                     rows=7)
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 7)
        self.assertEqual(calls, [],
                         "a single page must not go through the crawl loop")
        self.assertEqual(workflow.scraper.single_pages, 1)


# =========================== 6. exported rows match the extension's count ===
class TestExportedRowsMatch(ZeroResultsCase):
    def test_the_xlsx_row_count_matches_the_final_row_count(self):
        from test_zero_results import TestCollectZeroResults  # noqa: F401
        page = FakePage(state=page_state())
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        from student_scraper import StateManager
        workflow = S.Workflow.__new__(S.Workflow)
        workflow.options = Options()
        workflow.context = self.context
        workflow.searches = self.assigned
        workflow.store = store
        workflow.existing = {}
        workflow._master_dirty = False
        workflow.unit = self.unit
        workflow.state = StateManager(path=self.context.state_file)
        workflow.browser = FakeBrowser(page)
        workflow.scraper = StubScraper(S.NextControl.PAGINATED, FakePopup(),
                                       rows=19,
                                       outcome=CrawlOutcome.FINISHED)
        workflow.profiles = None
        workflow.portfolios = None
        workflow.resumes = None
        workflow.merger = S.StudentMerger()
        workflow.reader = S.XlsxReader()
        workflow.transient = S.TransientFailureLog()
        workflow.processed = 0
        workflow.keyword_number = 1
        workflow.keyword_total = 1

        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 19, "collect() reports the exported count")
        from student_scraper import validate_export
        self.assertEqual(validate_export(self.target), 19,
                         "the XLSX on disk must hold exactly those rows")
        candidates = workflow.reader.read(self.target,
                                          self.unit.search_query)
        self.assertEqual(len(candidates), 19)


# ================== 7. browser death before and after a proven export =======
class TestBrowserDeath(ZeroResultsCase):
    def build(self, scraper_obj):
        page = FakePage(state=page_state())
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        from student_scraper import StateManager
        workflow = S.Workflow.__new__(S.Workflow)
        workflow.options = Options()
        workflow.context = self.context
        workflow.searches = self.assigned
        workflow.store = store
        workflow.existing = {}
        workflow._master_dirty = False
        workflow.unit = self.unit
        workflow.state = StateManager(path=self.context.state_file)
        workflow.browser = FakeBrowser(page)
        workflow.scraper = scraper_obj
        workflow.profiles = None
        workflow.portfolios = None
        workflow.resumes = None
        workflow.merger = S.StudentMerger()
        workflow.reader = S.XlsxReader()
        workflow.transient = S.TransientFailureLog()
        workflow.processed = 0
        workflow.keyword_number = 1
        workflow.keyword_total = 1
        return workflow, page

    def test_death_after_a_proven_export_still_recovers_the_file(self):
        """Requirement 6: the existing disk-recovery path is untouched."""
        class Recovering(StubScraper):
            async def export_xlsx(self, popup, destination):
                # save_as failed; the bytes were rescued from disk.
                self.exports.append(destination)
                make_export(destination, self.rows)
                return self.rows

        workflow, page = self.build(
            Recovering(S.NextControl.PAGINATED, FakePopup(), rows=19,
                       outcome=CrawlOutcome.FINISHED))
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 19)
        self.assertTrue(os.path.exists(self.target))

    def test_death_before_pagination_is_proven_does_not_export_as_complete(self):
        """Requirement 7: an unproven crawl must not become a complete XLSX."""
        class DiesMidCrawl(StubScraper):
            async def wait_for_crawl_complete(self, popup, page,
                                              max_pages=None, had_next=False):
                self.crawls += 1
                # The renderer died; the page cannot be inspected, so the
                # crawl is INCOMPLETE rather than "last page reached".
                return CrawlOutcome.INCOMPLETE, 1, 10

        workflow, page = self.build(
            DiesMidCrawl(S.NextControl.PAGINATED, FakePopup(), rows=10))
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(rows, -1)
        self.assertFalse(os.path.exists(self.target))

    def test_zero_results_is_unaffected_by_any_of_this(self):
        """Requirement 9."""
        workflow, page = self.build(
            StubScraper(S.NextControl.SINGLE_PAGE, FakePopup()))
        page.state = zero_state()
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(status, S.CollectResult.ZERO_RESULTS)
        self.assertEqual(rows, 0)
        self.assertEqual(workflow.scraper.exports, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
