"""A result set with ONE page is complete, not broken.

THE LIVE BUG THIS FILE EXISTS FOR
---------------------------------
    [SCRAPER] GitHub showed no "Next" control (single page of results)
    [FAIL] the extension has no way to reach page 2
    [KEEP] collection did not complete; nothing removed

The diagnosis was right and the verdict was wrong.  Instant Data Scraper
renders #startScraping only once a "Next" selector has been marked, so on a
one-page result set that button never appears.  ensure_next_selector()
returned a plain bool, its no-Next branch returned
`self._has(popup, SEL_START_CRAWL)` -- necessarily False -- and collect()
could only read that single False as "cannot reach page 2".

Three genuinely different situations were collapsed into one bool:

    1  GitHub has no Next control        -> a valid SINGLE PAGE
    2  GitHub has Next, extension can use it -> crawl
    3  GitHub has Next, extension cannot use it, OR the page is
       blocked / errored / empty         -> failure, preserve everything

They are now three distinct verdicts, decided from GitHub's own DOM rather
than from the extension failing to do something.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook                          # noqa: E402

import config as C                                     # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (CrawlOutcome, NextControl, StateManager, classify_pagination)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                            # noqa: E402


# =============================================================== fixtures ===
def page_state(**overrides):
    """A PAGE_STATE_JS result for a healthy one-page Users search."""
    state = {
        "url": "https://github.com/search?q=College+Student+Visakhapatnam"
               "&type=users",
        "title": "Search · College Student Visakhapatnam · GitHub",
        "bodyChars": 4000,
        "bodyHead": "12 users College Student Visakhapatnam",
        "bodyTail": "ProTip! Press / to search",
        "hasNext": False,
        "nextDisabled": None,
        "nextHref": "",
        "nextByShape": 0,
        "nextByText": 0,
        "paginationPresent": False,
        "profileLinkCount": 12,
        "resultProfileLinkCount": 12,
        "resultsListPresent": True,
        "resultsListChildren": 12,
        "noResults": False,
        "zeroUsersCopy": False,
        "zeroUsersCount": False,
        "saysError": False,
        "blocked_marker": None,
        "isUsersSearch": True,
        "mainRegionPresent": True,
        "loginForm": False,
        "loginCopy": False,
    }
    state.update(overrides)
    return state


class FakeLocator:
    def __init__(self, count=0, visible=True, text="", on_click=None,
                 attributes=None):
        self._count = count
        self._visible = visible
        self._text = text
        self._on_click = on_click
        self._attributes = attributes or {}
        self.clicks = 0

    async def get_attribute(self, name, timeout=None):
        return self._attributes.get(name)

    async def count(self):
        return self._count

    @property
    def first(self):
        return self

    async def is_visible(self):
        return self._visible

    async def inner_text(self, timeout=None):
        return self._text

    async def fill(self, value, timeout=None):
        return None

    async def click(self, timeout=None):
        self.clicks += 1
        if self._on_click is not None:
            self._on_click()


class FakePage:
    """Just enough Playwright surface for the pagination decision."""

    def __init__(self, state=None, next_selectors=(), next_clickable=True):
        self.state = state if state is not None else page_state()
        self.url = self.state["url"]
        self._next_selectors = set(next_selectors)
        self._next_clickable = next_clickable
        self.next_clicks = 0
        self.closed = False
        self.gotos = []

    def is_closed(self):
        return self.closed

    async def goto(self, url, **_kwargs):
        self.gotos.append(url)

    async def bring_to_front(self):
        return None

    async def evaluate(self, script, arg=None):
        # The only script a page is asked to run here is PAGE_STATE_JS.
        return dict(self.state)

    def locator(self, selector):
        if selector in self._next_selectors:
            def _click():
                self.next_clicks += 1
            if not self._next_clickable:
                def _click():                        # noqa: F811
                    raise RuntimeError("element detached")
            return FakeLocator(count=1, visible=True, on_click=_click)
        if selector == "body":
            return FakeLocator(count=1, text=self.state["bodyHead"])
        return FakeLocator(count=0)


class FakePopup:
    """The extension popup.

    has_start models the real behaviour precisely: Instant Data Scraper shows
    #startScraping only once a Next selector is marked, which is exactly why a
    single-page result set cannot be crawled and must be exported directly.
    """

    def __init__(self, has_start=False, has_next_button=True, rows=12,
                 error="", accepts_learning=True, learned_selector=""):
        self.has_start = has_start
        self.has_next_button = has_next_button
        self.rows = rows
        self.error = error
        self.accepts_learning = accepts_learning
        self.learned_selector = learned_selector
        self.storage = {}
        self.xlsx_clicks = 0
        self.start_clicks = 0
        self.locate_clicks = 0
        self.reloads = 0
        self.closed = False

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    async def reload(self, **_kwargs):
        self.reloads += 1

    async def wait_for_selector(self, selector, timeout=None):
        if selector == C.SEL_START_CRAWL and not self.has_start:
            raise RuntimeError("no start button")
        return FakeLocator(count=1)

    async def evaluate(self, script, arg=None):
        if isinstance(arg, list) and len(arg) == 2:
            self.storage[arg[0]] = arg[1]
            return None
        if isinstance(arg, str):
            return self.storage.get(arg, "")
        return ""

    async def click(self, selector, timeout=None):
        if selector == C.SEL_XLSX:
            self.xlsx_clicks += 1
        elif selector == C.SEL_START_CRAWL:
            self.start_clicks += 1
        elif selector == C.SEL_NEXT_BUTTON:
            self.locate_clicks += 1
            if self.accepts_learning:
                self.has_start = True
                if self.learned_selector:
                    self.storage[
                        C.NEXT_SELECTOR_STORAGE_KEY % "github.com"] = \
                        self.learned_selector

    async def inner_text(self, selector, timeout=None):
        if selector == C.SEL_STATS:
            return ("Pages scraped: 1 Rows collected: %d" % self.rows)
        return ""

    def locator(self, selector):
        if selector == C.SEL_START_CRAWL:
            return FakeLocator(count=1 if self.has_start else 0)
        if selector == C.SEL_NEXT_BUTTON:
            return FakeLocator(count=1 if self.has_next_button else 0)
        if selector == C.SEL_STATS:
            return FakeLocator(count=1, text="Pages scraped: 1 Rows "
                                             "collected: %d" % self.rows)
        if selector == C.SEL_ERROR:
            return FakeLocator(count=1 if self.error else 0, text=self.error)
        if selector in (C.SEL_INSTRUCTIONS, C.SEL_STOP_CRAWL,
                        C.SEL_WRONG_TABLE):
            return FakeLocator(count=0)
        return FakeLocator(count=0)


class StubState:
    def __init__(self):
        self.selectors = {}

    def next_selector_for(self, host):
        return self.selectors.get(host, "")

    def remember_next_selector(self, host, selector):
        self.selectors[host] = selector


def scraper(popup_state=None):
    said = []
    instance = student_scraper.InstantDataScraper(browser=None, state=StubState(),
                                      say=said.append)
    return instance, said


# ============================================ 1. the pure decision function ==
class TestClassifyPagination(unittest.TestCase):
    """GitHub's DOM decides, not the extension's failure."""

    def test_no_next_control_with_results_is_a_single_page(self):
        verdict, why = classify_pagination(page_state())
        self.assertEqual(verdict, NextControl.SINGLE_PAGE)
        self.assertIn("no next page", why)

    def test_a_rel_next_link_means_paginated(self):
        verdict, _why = classify_pagination(
            page_state(hasNext=True, nextByShape=1))
        self.assertEqual(verdict, NextControl.PAGINATED)

    def test_an_aria_label_next_is_also_paginated(self):
        """hasNext only looks for a[rel=next]; the broader probe catches this."""
        verdict, _why = classify_pagination(page_state(nextByShape=1))
        self.assertEqual(verdict, NextControl.PAGINATED)

    def test_a_next_link_found_only_by_its_text_is_paginated(self):
        verdict, _why = classify_pagination(
            page_state(nextByText=1, paginationPresent=True))
        self.assertEqual(verdict, NextControl.PAGINATED)

    def test_a_disabled_next_on_the_last_page_is_a_single_page(self):
        verdict, _why = classify_pagination(
            page_state(hasNext=True, nextByShape=1, nextDisabled=True))
        self.assertEqual(verdict, NextControl.SINGLE_PAGE)

    def test_no_next_and_no_results_is_a_failure_not_a_single_page(self):
        verdict, why = classify_pagination(
            page_state(profileLinkCount=0, resultProfileLinkCount=0,
                       resultsListChildren=0, resultsListPresent=False))
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertIn("no user results", why)

    def test_a_blocked_page_is_a_failure_however_few_next_links(self):
        verdict, why = classify_pagination(
            page_state(blocked_marker="rate limit"))
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertIn("rate limit", why)

    def test_an_http_error_page_is_a_failure(self):
        verdict, why = classify_pagination(page_state(title="Error 429"))
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertIn("429", why)

    def test_a_github_error_page_is_a_failure(self):
        verdict, _why = classify_pagination(page_state(saysError=True))
        self.assertEqual(verdict, NextControl.FAILED)

    def test_visible_results_beat_empty_state_text(self):
        """If users are on the page, no wording may override them.

        A results page can carry the phrase "0 users" in a sidebar count for
        some other search type, so the count of actual results is asked FIRST
        and wins.
        """
        verdict, _why = classify_pagination(page_state(noResults=True))
        self.assertEqual(verdict, NextControl.SINGLE_PAGE)

    def test_a_genuinely_empty_page_never_reaches_the_single_page_path(self):
        """ZERO_RESULTS is handled by collect() before IDS is opened."""
        verdict, _why = classify_pagination(
            page_state(noResults=True, profileLinkCount=0,
                       resultProfileLinkCount=0, resultsListChildren=0,
                       resultsListPresent=False))
        self.assertEqual(verdict, NextControl.FAILED)

    def test_three_profile_links_are_enough_evidence_of_results(self):
        """The threshold, and that it is the NON-CHROME count that matters."""
        self.assertEqual(
            classify_pagination(page_state(resultProfileLinkCount=3,
                                           profileLinkCount=3,
                                           resultsListChildren=0))[0],
            NextControl.SINGLE_PAGE)
        self.assertEqual(
            classify_pagination(page_state(resultProfileLinkCount=2,
                                           profileLinkCount=2,
                                           resultsListChildren=0,
                                           resultsListPresent=False))[0],
            NextControl.FAILED)

    def test_chrome_links_alone_are_not_results(self):
        """A signed-in page's own nav/avatar links must not count.

        profileLinkCount counts every single-segment profile link on the page,
        including the viewer's own avatar and footer links.  Only the
        chrome-excluded count decides.
        """
        verdict, _why = classify_pagination(
            page_state(profileLinkCount=9,        # nav, avatar, footer
                       resultProfileLinkCount=0,  # nothing in the results area
                       resultsListChildren=0, resultsListPresent=False))
        self.assertEqual(verdict, NextControl.FAILED)


# ==================================== 2. ensure_next_selector's three ways ===
class TestEnsureNextSelector(unittest.TestCase):
    def run_it(self, page, popup):
        instance, said = scraper()
        verdict = asyncio.run(instance.ensure_next_selector(popup, page))
        return verdict, said, instance

    def test_no_next_control_returns_single_page(self):
        """The live bug: this used to return False and fail the search."""
        page = FakePage()
        popup = FakePopup(has_start=False)
        verdict, said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.SINGLE_PAGE)
        self.assertTrue(any("single page of results" in line
                            for line in said), said)
        self.assertTrue(any("single-page mode" in line for line in said), said)

    def test_single_page_teaches_no_pagination_selector(self):
        page = FakePage()
        popup = FakePopup(has_start=False)
        verdict, _said, instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.SINGLE_PAGE)
        self.assertEqual(popup.locate_clicks, 0,
                         "no Locate flow may run when there is no Next")
        self.assertEqual(page.next_clicks, 0)
        self.assertEqual(instance.state.selectors, {},
                         "nothing may be remembered as a Next selector")

    def test_a_stable_selector_that_works_gives_paginated(self):
        page = FakePage(state=page_state(hasNext=True, nextByShape=1),
                        next_selectors={'a[rel="next"]'})
        popup = FakePopup(has_start=True)
        verdict, said, instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.PAGINATED)
        self.assertEqual(instance.state.selectors.get("github.com"),
                         'a[rel="next"]')
        self.assertTrue(any("pagination selector" in line for line in said),
                        said)

    def test_an_extension_that_already_has_start_gives_paginated(self):
        page = FakePage(state=page_state(hasNext=True, nextByShape=1))
        popup = FakePopup(has_start=True)
        verdict, _said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.PAGINATED)

    def test_next_present_but_extension_cannot_learn_it_fails(self):
        """Case 3: a real next page the extension refuses to accept."""
        page = FakePage(state=page_state(hasNext=True, nextByShape=1,
                                         paginationPresent=True),
                        next_selectors={'a:has-text("Next")'})
        popup = FakePopup(has_start=False, accepts_learning=False)
        verdict, said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertTrue(any("did not accept" in line for line in said), said)

    def test_next_present_but_unclickable_fails_not_single_page(self):
        """Never guess 'single page' for a many-page set."""
        page = FakePage(state=page_state(hasNext=True, nextByShape=1),
                        next_selectors={'a[rel="next"]'},
                        next_clickable=False)
        popup = FakePopup(has_start=False)
        verdict, said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertTrue(any("could not be clicked" in line for line in said),
                        said)
        self.assertFalse(any("single page" in line for line in said), said)

    def test_next_present_but_no_locate_control_fails(self):
        page = FakePage(state=page_state(hasNext=True, nextByShape=1))
        popup = FakePopup(has_start=False, has_next_button=False)
        verdict, said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertTrue(any("no Locate control" in line for line in said),
                        said)

    def test_a_blocked_page_fails_and_is_not_single_page(self):
        """Case 4: must never be mistaken for single-page success."""
        page = FakePage(state=page_state(title="Error 429",
                                         bodyHead="Unable to load page. "
                                                  "Status: 429"))
        popup = FakePopup(has_start=False)
        verdict, said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.FAILED)
        self.assertTrue(any("not a usable results page" in line
                            for line in said), said)
        self.assertFalse(any("single-page mode" in line for line in said),
                         said)

    def test_an_empty_page_fails_and_is_not_single_page(self):
        page = FakePage(state=page_state(profileLinkCount=0,
                                         resultProfileLinkCount=0,
                                         resultsListChildren=0,
                                         resultsListPresent=False))
        popup = FakePopup(has_start=False)
        verdict, _said, _instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.FAILED)

    def test_the_learned_selector_is_remembered(self):
        page = FakePage(state=page_state(hasNext=True, nextByShape=1),
                        next_selectors={'a:has-text("Next")'})
        popup = FakePopup(has_start=False, accepts_learning=True,
                          learned_selector="div>a.learned")
        verdict, _said, instance = self.run_it(page, popup)
        self.assertEqual(verdict, NextControl.PAGINATED)
        self.assertEqual(instance.state.selectors.get("github.com"),
                         "div>a.learned")


# ============================================= 3. the single-page completion =
class TestFinishSinglePage(unittest.TestCase):
    def test_it_reports_one_page_and_the_row_count(self):
        instance, said = scraper()
        popup = FakePopup(rows=12)
        outcome, pages, rows = asyncio.run(
            instance.finish_single_page(popup, FakePage()))
        self.assertEqual(outcome, CrawlOutcome.SINGLE_PAGE)
        self.assertEqual(pages, 1)
        self.assertEqual(rows, 12)
        self.assertTrue(any("single page complete" in line for line in said),
                        said)

    def test_single_page_is_exportable(self):
        self.assertIn(CrawlOutcome.SINGLE_PAGE, CrawlOutcome.EXPORTABLE)

    def test_start_crawling_is_never_clicked(self):
        instance, _said = scraper()
        popup = FakePopup(rows=12)
        asyncio.run(instance.finish_single_page(popup, FakePage()))
        self.assertEqual(popup.start_clicks, 0)

    def test_no_table_is_a_failure_not_a_zero_row_success(self):
        instance, said = scraper()
        popup = FakePopup(rows=0)
        # wait_for_table legitimately waits EXTENSION_TIMEOUT for the content
        # script to answer; shortened here so the suite does not sit for 30s
        # proving a timeout that production needs to be long.
        saved = C.EXTENSION_TIMEOUT
        C.EXTENSION_TIMEOUT = 300
        self.addCleanup(setattr, C, "EXTENSION_TIMEOUT", saved)
        outcome, pages, rows = asyncio.run(
            instance.finish_single_page(popup, FakePage()))
        self.assertEqual(outcome, CrawlOutcome.NEVER_STARTED)
        self.assertEqual((pages, rows), (0, 0))
        self.assertNotIn(outcome, CrawlOutcome.EXPORTABLE)


# ================================================== 4. collect() end to end ==
def make_export(path, count):
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "search-match"])
    for number in range(1, count + 1):
        sheet.append(["https://github.com/person%d.png?size=40" % number,
                      "https://github.com/person%d" % number,
                      "Person %d" % number])
    book.save(path)
    book.close()


class FakeBrowser:
    def __init__(self, page):
        self.current = page
        self.restarts = 0
        # Mirrors BrowserManager: students finished since the last restart.
        # process_rows() credits the restart budget after every student, so
        # a double that omits this raises AttributeError there.  Imported by
        # test_exact_count, test_retry_search and test_zero_results too.
        self._healthy_since_restart = 0

    def note_healthy(self):
        """Mirrors BrowserManager.note_healthy -- same forgiveness rule."""
        if self.restarts == 0:
            self._healthy_since_restart = 0
            return
        self._healthy_since_restart += 1
        if self._healthy_since_restart >= C.BROWSER_RESTART_FORGIVE_AFTER:
            self.restarts = 0
            self._healthy_since_restart = 0

    async def is_alive(self):
        return True

    async def close_extra_pages(self, keep=()):
        return 0

    async def release_page_memory(self, page):
        return True

    async def get_live_page(self):
        return self.current

    def page_count(self):
        return 1


class StubScraper:
    """The real decision, with the browser work replaced by canned answers."""

    def __init__(self, verdict, popup, rows=12, outcome=None):
        self.verdict = verdict
        self.popup = popup
        self.rows = rows
        self.outcome = outcome
        self.exports = []
        self.crawls = 0
        self.single_pages = 0
        self.had_next = None

    async def capture_page_state(self, page, reason=""):
        return dict(page.state)

    def log_state(self, state):
        return None

    async def open_instant_data_scraper(self, page):
        return self.popup

    async def ensure_next_selector(self, popup, page):
        return self.verdict

    async def finish_single_page(self, popup, page):
        self.single_pages += 1
        return CrawlOutcome.SINGLE_PAGE, 1, self.rows

    async def start_crawling(self, popup):
        return True

    async def establish_results_table(self, popup, page, expected_ids):
        return getattr(self, "table_ok", True)

    async def stabilize_results_table(self, popup, page, expected_ids):
        # The real method establishes the table then requires it to hold
        # stable; the fake collapses that to the same table_ok knob the
        # establishment tests already use.  Stability itself is unit-tested
        # against the real InstantDataScraper in test_table_establishment.py.
        return getattr(self, "table_ok", True)

    async def export_page_from_preview(self, popup, destination, expected_ids):
        # Production writes the page XLSX from the verified preview; the fake
        # models that as producing the same file its export produced.  The
        # real writer/reader are exercised in test_preview_export.py.
        return await self.export_xlsx(popup, destination)

    async def preview_identities(self, popup, expected_ids=None):
        # by default the preview matches whatever the export will contain
        return getattr(self, "_preview_ids", set())

    def dom_result_identities(self, state):
        from student_scraper import InstantDataScraper
        return InstantDataScraper.dom_result_identities(state)

    def export_identities(self, path):
        from student_scraper import InstantDataScraper
        return InstantDataScraper.export_identities(path)

    async def wait_for_crawl_complete(self, popup, page, max_pages=None,
                                      had_next=False):
        self.crawls += 1
        self.had_next = had_next
        return (self.outcome or CrawlOutcome.FINISHED), 4, self.rows

    async def export_xlsx(self, popup, destination):
        self.exports.append(destination)
        make_export(destination, self.rows)
        return self.rows


class Options:
    show = False
    limit = 0
    process_existing = False
    max_pages = None
    verbose = False
    rank = 1
    total_ranks = 1


class FakeCity:
    def __init__(self, state, canonical, aliases):
        self.country = "India"
        self.state = state
        self.canonical = canonical
        self.search_aliases = tuple(aliases)


class CollectCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="singlepage_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan(
            ["College Student"],
            [FakeCity("Andhra Pradesh", "Visakhapatnam", ["Visakhapatnam"])])
        self.context = sandbox.enter(self.runtime)
        self.unit = self.plan.assign(1, 1)[0]
        self.target = os.path.join(self.context.keyword_download_dir,
                                   "%s.xlsx" % self.unit.slug)
        self.stores = []

    def tearDown(self):
        sandbox.restore()
        for store in self.stores:
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def workflow(self, verdict, popup=None, rows=12, outcome=None,
                 page=None):
        page = page or FakePage()
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        workflow = S.Workflow.__new__(S.Workflow)
        workflow.options = Options()
        workflow.context = self.context
        workflow.searches = self.plan.assign(1, 1)
        workflow.store = store
        workflow.existing = {}
        workflow._master_dirty = False
        workflow.unit = self.unit
        workflow.state = StateManager(path=self.context.state_file)
        workflow.browser = FakeBrowser(page)
        workflow.scraper = StubScraper(verdict, popup or FakePopup(),
                                       rows=rows, outcome=outcome)
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

    def collect(self, workflow, page):
        """(status, rows) -- collect() reports WHY, not just how many."""
        return asyncio.run(workflow.collect(page, self.unit.search_query,
                                            self.target))


class TestCollectSinglePage(CollectCase):
    def test_a_single_page_result_set_collects_successfully(self):
        workflow, page = self.workflow(NextControl.SINGLE_PAGE)
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 12, "a one-page result set must NOT fail")
        self.assertTrue(os.path.exists(self.target))

    def test_the_xlsx_is_downloaded_exactly_once(self):
        workflow, page = self.workflow(NextControl.SINGLE_PAGE)
        self.collect(workflow, page)
        self.assertEqual(workflow.scraper.exports, [self.target])
        self.assertEqual(len(workflow.scraper.exports), 1)

    def test_no_crawl_is_started_for_a_single_page(self):
        workflow, page = self.workflow(NextControl.SINGLE_PAGE)
        self.collect(workflow, page)
        self.assertEqual(workflow.scraper.crawls, 0)
        self.assertEqual(workflow.scraper.single_pages, 1)

    def test_candidate_processing_starts_normally_afterwards(self):
        """The whole point: a single page feeds the profile phase."""
        workflow, page = self.workflow(NextControl.SINGLE_PAGE, rows=4)
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 4)
        candidates = workflow.reader.read(self.target,
                                          self.unit.search_query)
        self.assertEqual(len(candidates), 4)
        self.assertEqual(candidates[0].username, "person1")

    def test_a_multi_page_result_set_still_crawls(self):
        """The existing pagination path must be untouched."""
        workflow, page = self.workflow(NextControl.PAGINATED)
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 12)
        self.assertEqual(workflow.scraper.crawls, 1)
        self.assertEqual(workflow.scraper.single_pages, 0)
        self.assertEqual(len(workflow.scraper.exports), 1)

    def test_a_failed_verdict_preserves_everything(self):
        make_export(self.target, 3)          # a previous attempt's file
        before = open(self.target, "rb").read()
        workflow, page = self.workflow(NextControl.FAILED)
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(rows, -1)
        self.assertEqual(workflow.scraper.exports, [],
                         "a failure must not export anything")
        self.assertTrue(os.path.exists(self.target))
        self.assertEqual(open(self.target, "rb").read(), before)
        self.assertFalse(workflow.state.collection_complete)

    def test_a_single_page_with_no_table_preserves_everything(self):
        class NoTable(StubScraper):
            async def finish_single_page(self, popup, page):
                self.single_pages += 1
                return CrawlOutcome.NEVER_STARTED, 0, 0

        workflow, page = self.workflow(NextControl.SINGLE_PAGE)
        workflow.scraper = NoTable(NextControl.SINGLE_PAGE, FakePopup())
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(rows, -1)
        self.assertEqual(workflow.scraper.exports, [])
        self.assertFalse(workflow.state.collection_complete)


# =================================== 5. the saved live-run state still works =
class TestSavedStateResumes(unittest.TestCase):
    """The failed run's campaign and SearchUnit must survive the fix.

    The state on disk was written by the run that hit this bug.  Running the
    same command must resume THAT campaign and THAT search, not start a new
    one -- nothing about the fix touches the plan or its fingerprints.
    """

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="resume_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan()
        self.context = sandbox.enter(self.runtime)

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def test_a_state_saved_by_the_failed_run_resumes_the_same_search(self):
        """The failed run's search identity must survive the fix."""
        assigned = self.plan.assign(1, 1)
        first = StateManager(path=self.context.state_file)
        first.begin_unit(0, assigned[0])
        expected = assigned[0].search_query
        self.assertEqual(first.search_query, expected)
        self.assertFalse(first.collection_complete)

        # A later run, same configuration: no refusal, same search in hand.
        again = StateManager(path=self.context.state_file)
        self.assertEqual(again.search_query, expected)
        self.assertEqual(again.unit_index, 0)

    def test_the_fix_does_not_change_how_much_work_there_is(self):
        """A code fix must not silently change the search list."""
        import student_scraper as keyword_loader
        import student_scraper as locations
        self.assertEqual(
            len(keyword_loader.load_base_keywords())
            * locations.search_spelling_count(),
            3004950)


if __name__ == "__main__":
    unittest.main(verbosity=2)
