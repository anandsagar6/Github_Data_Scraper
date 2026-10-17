"""A valid search with zero users is a COMPLETED search unit, not a failure.

THE LIVE BUG THIS FILE EXISTS FOR
---------------------------------
    [SEARCH] GitHub Users page opened
    [SCRAPER] this is not a usable results page: GitHub reported no users
    [FAIL] this search could not be collected; see the reason above
    [KEEP] collection did not complete; nothing removed

"College Student Visakhapatnam" has no GitHub users.  That is an answer, not
an error -- but it was reported as a collection failure, preserved as
unresolved, and retried identically on every run, so one legitimately empty
query blocked all 3,060 searches of the rank forever.

TWO DEFECTS, BOTH FIXED HERE
----------------------------
1.  TWO divergent zero-checks that disagreed about the same page.
    collect() tested the ASCII substring "We couldn't find any users", which
    never matches GitHub's typographic apostrophe, so it fell through.  The
    extension-side regex
        /We couldn.t find any users|no users matched|0 results/i
    matched (its "." accepts any apostrophe) -- but that regex ALSO matched
    "10 results", "100 results" and "230 results" as substrings, so a page
    FULL of users could report itself empty.  There is now ONE classifier.

2.  Zero results had no completion path.  Even when detected, the old code
    marked collection complete pointing at an XLSX that was never created,
    then read that missing file, then failed to delete it.

ORDER OF CHECKS IS THE OTHER HALF OF THE FIX
--------------------------------------------
A login page, a 429, a challenge page and an error page all render zero user
results.  "How many users?" is therefore asked LAST, after "is this page even
real?", and visible results always beat empty-state wording.
"""

import asyncio
import json
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                     # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (PageState, StateManager, classify_search_page)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                            # noqa: E402
from test_single_page import (FakeBrowser, FakePage,    # noqa: E402
                              FakePopup, Options, StubScraper,
                              page_state)


def zero_state(**overrides):
    """A valid Users search page that GitHub says has no users."""
    base = page_state(
        bodyHead="We couldn’t find any users matching "
                 "'College Student Visakhapatnam'",
        profileLinkCount=0, resultProfileLinkCount=0,
        resultsListPresent=False, resultsListChildren=0,
        hasNext=False, nextByShape=0, nextByText=0,
        noResults=True, zeroUsersCopy=True)
    base.update(overrides)
    return base


# ============================================== 1. classification, in order ==
class TestClassifySearchPage(unittest.TestCase):
    def test_a_valid_page_with_zero_users_is_zero_results(self):
        verdict, why = classify_search_page(zero_state())
        self.assertEqual(verdict, PageState.ZERO_RESULTS)
        self.assertIn("no users", why)

    def test_the_typographic_apostrophe_is_matched(self):
        """The exact copy GitHub renders, with U+2019."""
        for apostrophe in ("’", "'", "´"):
            body = "We couldn%st find any users matching your search" \
                   % apostrophe
            verdict, _why = classify_search_page(
                zero_state(bodyHead=body))
            self.assertEqual(verdict, PageState.ZERO_RESULTS, apostrophe)

    def test_a_page_with_real_results_is_never_zero_results(self):
        for verdict in (classify_search_page(page_state())[0],
                        classify_search_page(
                            page_state(hasNext=True, nextByShape=1))[0]):
            self.assertIn(verdict, PageState.COLLECTABLE)
            self.assertNotEqual(verdict, PageState.ZERO_RESULTS)

    def test_visible_results_beat_empty_state_wording(self):
        """Step 6 before step 7: users present wins, whatever the text says."""
        verdict, _why = classify_search_page(page_state(noResults=True))
        self.assertEqual(verdict, PageState.RESULTS_SINGLE_PAGE)

    def test_single_and_multi_page_results_are_distinguished(self):
        self.assertEqual(classify_search_page(page_state())[0],
                         PageState.RESULTS_SINGLE_PAGE)
        self.assertEqual(
            classify_search_page(page_state(hasNext=True, nextByShape=1))[0],
            PageState.RESULTS_MULTI_PAGE)
        self.assertEqual(
            classify_search_page(page_state(nextByText=1,
                                            paginationPresent=True))[0],
            PageState.RESULTS_MULTI_PAGE)

    def test_a_disabled_next_is_the_last_page_not_multi_page(self):
        self.assertEqual(
            classify_search_page(page_state(hasNext=True, nextByShape=1,
                                            nextDisabled=True))[0],
            PageState.RESULTS_SINGLE_PAGE)

    # -- the four impostors, each of which also has zero results -----------
    def test_a_login_page_is_login_required_not_zero_results(self):
        for overrides in ({"loginForm": True}, {"loginCopy": True}):
            verdict, why = classify_search_page(zero_state(**overrides))
            self.assertEqual(verdict, PageState.LOGIN_REQUIRED, overrides)
            self.assertIn("sign-in", why)

    def test_a_login_page_wins_even_over_visible_results(self):
        """An auth wall is never a results page, however it renders."""
        verdict, _why = classify_search_page(page_state(loginForm=True))
        self.assertEqual(verdict, PageState.LOGIN_REQUIRED)

    def test_a_rate_limited_page_is_blocked_not_zero_results(self):
        verdict, why = classify_search_page(
            zero_state(blocked_marker="secondary rate limit"))
        self.assertEqual(verdict, PageState.BLOCKED)
        self.assertIn("rate limit", why)

    def test_an_http_error_page_is_http_error_not_zero_results(self):
        verdict, why = classify_search_page(zero_state(title="Error 429"))
        self.assertEqual(verdict, PageState.HTTP_ERROR)
        self.assertIn("429", why)

    def test_a_status_line_in_the_body_is_also_an_http_error(self):
        verdict, _why = classify_search_page(
            zero_state(title="", bodyHead="Unable to load page. Status: 503"))
        self.assertEqual(verdict, PageState.HTTP_ERROR)

    def test_a_github_error_page_is_github_error_not_zero_results(self):
        verdict, _why = classify_search_page(zero_state(saysError=True))
        self.assertEqual(verdict, PageState.GITHUB_ERROR)

    def test_a_non_users_page_is_unknown_invalid(self):
        verdict, why = classify_search_page(
            page_state(isUsersSearch=False,
                       url="https://github.com/search?q=x&type=repositories"))
        self.assertEqual(verdict, PageState.UNKNOWN_INVALID)
        self.assertIn("not a GitHub Users search", why)

    def test_no_results_and_no_empty_state_message_is_unknown_invalid(self):
        """Silence is not a zero-result answer."""
        verdict, why = classify_search_page(
            zero_state(noResults=False, zeroUsersCopy=False))
        self.assertEqual(verdict, PageState.UNKNOWN_INVALID)
        self.assertIn("did not render as expected", why)

    def test_an_uninspectable_page_is_unknown_invalid(self):
        verdict, _why = classify_search_page({"page_error": "detached frame"})
        self.assertEqual(verdict, PageState.UNKNOWN_INVALID)

    def test_the_state_sets_are_disjoint_and_cover_everything(self):
        every = {value for name, value in vars(PageState).items()
                 if name.isupper() and isinstance(value, str)}
        grouped = (PageState.COLLECTABLE | PageState.COMPLETE_EMPTY
                   | PageState.FAILURES)
        self.assertEqual(every, grouped)
        self.assertEqual(PageState.COLLECTABLE & PageState.FAILURES, set())
        self.assertEqual(PageState.COMPLETE_EMPTY & PageState.FAILURES, set())
        self.assertEqual(PageState.COLLECTABLE & PageState.COMPLETE_EMPTY,
                         set())


# ============================== 2. the substring bug that is now impossible ==
class TestCountsAreNotSubstrings(unittest.TestCase):
    """The old /0 results/ alternative matched "10 results"."""

    def test_a_results_count_ending_in_zero_is_not_emptiness(self):
        for body in ("Showing 10 results for College Student",
                     "100 results", "230 results", "20 users", "190 users",
                     "10 users College Student Visakhapatnam"):
            verdict, _why = classify_search_page(page_state(bodyHead=body))
            self.assertIn(verdict, PageState.COLLECTABLE, body)

    def test_an_actual_zero_count_is_emptiness(self):
        verdict, _why = classify_search_page(
            zero_state(bodyHead="0 users  Code 0  Repositories 0",
                       zeroUsersCopy=False, noResults=True))
        self.assertEqual(verdict, PageState.ZERO_RESULTS)


# ================================================== 3. collect() behaviour ===
class ZeroResultsCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="zero_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan(
            ["College Student", "Fresher"],
            [FakeCity("Andhra Pradesh", "Visakhapatnam", ["Visakhapatnam"])])
        self.context = sandbox.enter(self.runtime)
        self.assigned = self.plan.assign(1, 1)
        self.unit = self.assigned[0]
        self.target = os.path.join(self.context.keyword_download_dir,
                                   "%s.xlsx" % self.unit.slug)
        self.stores = []

    def tearDown(self):
        sandbox.restore()
        for store in self.stores:
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def workflow(self, state, rows=12, verdict=None):
        page = FakePage(state=state)
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
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
        workflow.scraper = StubScraper(
            verdict or S.NextControl.SINGLE_PAGE, FakePopup(), rows=rows)
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


class FakeCity:
    def __init__(self, state, canonical, aliases):
        self.country = "India"
        self.state = state
        self.canonical = canonical
        self.search_aliases = tuple(aliases)


class TestASlowSearchPageDoesNotCrashTheRun(ZeroResultsCase):
    """The first navigation in collect() must not escape as a traceback.

    OBSERVED 2026-10-08 (Script3, rank 3): Page.goto on the search URL
    exceeded 45s and the TimeoutError unwound through collect(),
    do_search_unit(), _run() and main(), killing the process -- 48 collected
    students were left out of the visible workbook and the operator got a
    stack trace instead of "run the same command to resume".
    """

    def _slow_goto(self, page, fails):
        """Make page.goto time out `fails` times, then behave normally."""
        original = page.goto
        state = {"calls": 0}

        async def goto(url, **kwargs):
            state["calls"] += 1
            if state["calls"] <= fails:
                raise Exception("Page.goto: Timeout 45000ms exceeded.")
            return await original(url, **kwargs)

        page.goto = goto
        return state

    def test_a_search_page_that_opens_on_a_retry_still_works(self):
        workflow, page = self.workflow(zero_state())
        calls = self._slow_goto(page, fails=1)
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(calls["calls"], 2, "the navigation must be retried")
        self.assertEqual(status, S.CollectResult.ZERO_RESULTS,
                         "a retry that succeeds must reach the real verdict")
        self.assertEqual(rows, 0)

    def test_a_search_page_that_never_opens_fails_cleanly(self):
        workflow, page = self.workflow(zero_state())
        calls = self._slow_goto(page, fails=99)
        status, rows = asyncio.run(
            workflow.collect(page, self.unit.search_query, self.target))
        self.assertEqual(calls["calls"], C.SEARCH_NAV_ATTEMPTS,
                         "it must stop after the configured attempts")
        self.assertEqual(status, S.CollectResult.FAILED,
                         "a dead search page must FAIL, never raise")
        self.assertEqual(rows, -1)


class TestCollectZeroResults(ZeroResultsCase):
    def collect(self, workflow, page):
        return asyncio.run(workflow.collect(page, self.unit.search_query,
                                            self.target))

    def test_zero_results_is_reported_as_zero_results_not_failure(self):
        workflow, page = self.workflow(zero_state())
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.ZERO_RESULTS)
        self.assertEqual(rows, 0)

    def test_the_extension_is_never_opened_for_zero_results(self):
        """Do not pay for Instant Data Scraper to learn there is nothing."""
        workflow, page = self.workflow(zero_state())
        self.collect(workflow, page)
        self.assertEqual(workflow.scraper.single_pages, 0)
        self.assertEqual(workflow.scraper.crawls, 0)

    def test_no_export_is_attempted_and_no_xlsx_is_created(self):
        workflow, page = self.workflow(zero_state())
        self.collect(workflow, page)
        self.assertEqual(workflow.scraper.exports, [])
        self.assertFalse(os.path.exists(self.target))
        self.assertEqual(os.listdir(self.context.keyword_download_dir), [])

    def test_a_login_page_fails_and_does_not_report_zero_results(self):
        workflow, page = self.workflow(zero_state(loginForm=True))
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(rows, -1)

    def test_a_rate_limited_search_page_enters_cooldown_not_zero(self):
        # A rate-limit/challenge on the search page is no longer a hard stop:
        # it raises RateLimited so the run loop cools down and retries the SAME
        # search.  It still must NOT be reported as zero results.
        workflow, page = self.workflow(
            zero_state(blocked_marker="secondary rate limit"))
        with self.assertRaises(S.RateLimited):
            self.collect(workflow, page)

    def test_a_rate_limit_http_error_enters_cooldown(self):
        workflow, page = self.workflow(zero_state(title="Error 429"))
        with self.assertRaises(S.RateLimited):
            self.collect(workflow, page)

    def test_a_non_throttle_block_still_stops_the_run(self):
        # A genuine non-rate-limit block is still a hard, safe stop.
        workflow, page = self.workflow(
            zero_state(blocked_marker="access to this site has been"))
        with self.assertRaises(S.Stopped):
            self.collect(workflow, page)

    def test_a_github_error_page_fails(self):
        workflow, page = self.workflow(zero_state(saysError=True))
        status, _rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.FAILED)

    def test_an_unrecognisable_page_fails(self):
        workflow, page = self.workflow(
            zero_state(noResults=False, zeroUsersCopy=False))
        status, _rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.FAILED)

    def test_valid_single_page_results_still_collect(self):
        workflow, page = self.workflow(page_state())
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 12)
        self.assertEqual(workflow.scraper.single_pages, 1)

    def test_valid_multi_page_results_go_page_by_page(self):
        """Multi-page collection is still collection -- by a new route."""
        workflow, page = self.workflow(page_state(hasNext=True,
                                                  nextByShape=1),
                                       verdict=S.NextControl.PAGINATED)

        async def fake_page_by_page(target_page, target, first_state):
            return S.CollectResult.RESULTS, 19

        workflow._collect_page_by_page = fake_page_by_page
        status, rows = self.collect(workflow, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 19)
        self.assertEqual(workflow.scraper.crawls, 0,
                         "the IDS cross-page crawl is no longer used")


# ========================= 4. the search unit COMPLETES and the rank moves ===
class TestSearchUnitCompletes(ZeroResultsCase):
    def run_unit(self, workflow, page):
        return asyncio.run(workflow.do_search_unit(page, self.unit))

    def test_a_zero_result_unit_completes_successfully(self):
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)              # must NOT raise Stopped
        self.assertTrue(workflow.state.collection_complete)
        self.assertEqual(workflow.state.data["collection"]["rows"], 0)
        self.assertTrue(workflow.state.data["processing"]["complete"])
        self.assertTrue(workflow.state.data["cleanup_complete"])
        self.assertEqual(workflow.state.data["phase"], "done")

    def test_a_zero_result_unit_asks_for_the_short_pause(self):
        """The pause after a unit is sized from the load that unit caused.

        A zero-result search is a single results-page load, so it flags the
        shorter ZERO_RESULT pause instead of the full between-searches one.
        """
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)
        self.assertTrue(workflow._last_search_empty)

    def test_no_xlsx_filename_is_recorded_for_a_unit_that_has_none(self):
        """The old code stored a path to a file it never created."""
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)
        self.assertEqual(workflow.state.collection_file, "")
        self.assertFalse(os.path.exists(self.target))

    def test_no_student_row_is_saved(self):
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)
        self.assertEqual(workflow.store.count(), 0)
        self.assertEqual(workflow.processed, 0)

    def test_the_limit_is_not_consumed(self):
        """--limit 3 means three STUDENTS, not three searches visited."""
        workflow, page = self.workflow(zero_state())
        workflow.options.limit = 3
        self.run_unit(workflow, page)
        self.assertEqual(workflow.processed, 0,
                         "a zero-result search must not spend the limit")

    def test_no_temporary_candidate_data_is_left_behind(self):
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)
        self.assertEqual(student_scraper.stale_student_files(), [])

    def test_debris_from_an_earlier_attempt_is_swept(self):
        with open(self.context.temp_student_file, "w",
                  encoding="utf-8") as handle:
            handle.write("half-collected student from a previous run")
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)
        self.assertFalse(os.path.exists(self.context.temp_student_file))

    def test_the_next_search_unit_starts_normally(self):
        workflow, page = self.workflow(zero_state())
        self.run_unit(workflow, page)
        workflow.state.advance_unit()
        self.assertEqual(workflow.state.unit_index, 1)
        self.assertEqual(workflow.state.units_done, 1)
        # advance_unit() clears the per-unit identity and moves the index on.
        # current_row.json is what the run actually resumes from; this state
        # file only ever describes the search currently in hand.
        self.assertEqual(workflow.state.search_query, "")

        # ...and search 2 runs, with its own clean per-unit blocks.
        workflow.state.begin_unit(1, self.assigned[1])
        self.assertEqual(workflow.state.search_query,
                         self.assigned[1].search_query)
        self.assertFalse(workflow.state.collection_complete)
        self.assertFalse(workflow.state.data["cleanup_complete"])

    def test_a_real_failure_still_preserves_and_stops(self):
        """The general rule is untouched."""
        workflow, page = self.workflow(zero_state(saysError=True))
        with self.assertRaises(S.Stopped):
            self.run_unit(workflow, page)
        self.assertFalse(workflow.state.collection_complete)
        self.assertFalse(workflow.state.data["cleanup_complete"])

    def test_results_units_still_go_through_the_profile_phase(self):
        """A unit WITH candidates must still read its XLSX and save rows.

        The network stages are stubbed; what is under test is that a results
        unit takes the profile path at all, which the zero-result branch must
        not have short-circuited.
        """
        workflow, page = self.workflow(page_state(), rows=3)
        seen = []

        async def fake_candidate(_page, candidate):
            seen.append(candidate.username)
            record = {column: "" for column in C.FINAL_COLUMNS}
            record["GitHub Profile"] = candidate.profile_url
            record["github"] = candidate.profile_url
            saved, _reason = workflow.store.save_student(
                record, search_unit=workflow.unit,
                campaign_id="", rank=1, total_ranks=1)
            if saved:
                workflow.processed += 1

        workflow.do_candidate = fake_candidate
        self.run_unit(workflow, page)
        self.assertEqual(seen, ["person1", "person2", "person3"])
        self.assertTrue(workflow.state.data["cleanup_complete"])
        self.assertEqual(workflow.store.count(), 3)
        self.assertEqual(workflow.state.collection_file, self.target,
                         "a results unit DOES record its XLSX")
        self.assertFalse(os.path.exists(self.target),
                         "a fully resolved unit deletes its XLSX")


# ===================================== 5. the saved live-run state survives ==
class TestSavedStateStillResumes(unittest.TestCase):
    """campaign_20260927_001 / College Student Visakhapatnam must resume."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="zresume_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan()
        self.context = sandbox.enter(self.runtime)

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def test_the_failed_runs_state_resumes_the_same_search_unit(self):
        """The state the live failure left on disk must still be usable.

        It was written by the campaign-era version, so it carries `campaign`
        and `assignment` blocks this version knows nothing about.  Those are
        dropped; the search currently in hand is what matters and is kept.
        """
        saved = {
            "version": 4,
            "campaign": {
                "campaign_id": "campaign_20260927_001",
                "plan_fingerprint": "1b1591be3e5e237b992deb36e9a94d3c",
            },
            "assignment": {"rank": 1, "total_ranks": 1,
                           "assigned_units": 3004950},
            "position": {"unit_index": 0, "units_done": 0},
            "search_unit": {
                "country": "India", "state": "Andhra Pradesh",
                "canonical_city": "Visakhapatnam",
                "search_city": "Visakhapatnam",
                "base_keyword": "B.Tech Computer Science",
                "search_query": "in:B.Tech Computer Science location:Visakhapatnam",
                "slug": "btech-computer-science-visakhapatnam",
            },
            "phase": "search",
            "collection": {"complete": False, "xlsx_file": "", "rows": 0},
            "processing": {"next_row_index": 0, "total_rows": 0,
                           "complete": False},
            "current_candidate": {"username": "", "source_row": -1},
            "cleanup_complete": False,
            "next_selectors": {},
        }
        with open(self.context.state_file, "w", encoding="utf-8") as handle:
            json.dump(saved, handle)

        state = StateManager(path=self.context.state_file)
        self.assertEqual(state.search_query,
                         "in:B.Tech Computer Science location:Visakhapatnam")
        self.assertEqual(state.data["search_unit"]["canonical_city"],
                         "Visakhapatnam")
        self.assertFalse(state.collection_complete)
        self.assertEqual(state.unit_index, 0)

    def test_the_whole_search_list_is_still_three_million(self):
        """A detection fix must not change how much work there is."""
        import student_scraper as keyword_loader
        import student_scraper as locations
        self.assertEqual(
            len(keyword_loader.load_base_keywords())
            * locations.search_spelling_count(),
            3004950)


if __name__ == "__main__":
    unittest.main(verbosity=2)
