"""Exact-count integrity: a page GitHub says holds N users is not complete
until N USABLE identities are collected -- single page AND multi page.

THE LIVE BUG THIS FILE EXISTS FOR (Case 3)
------------------------------------------
    College Student Vijayawada
    [SEARCH] 5 result(s), no next page
    single page complete: 3 rows
    [OK] 3 rows                      <-- WRONG: GitHub said 5

Single-page collection accepted whatever IDS exported without checking it
against GitHub's stated total.  Multi-page collection already enforced that
count; single page was a hole in the same rule.  Two counts matter:

    raw     = rows IDS exported
    usable  = unique GitHub identities parsed from them  (a row with no
              username/profile URL does not count toward completeness)
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook              # noqa: E402

import config as C                                        # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (StateManager, ValidationMode)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                               # noqa: E402
from test_single_page import (FakeBrowser, FakePage,       # noqa: E402
                              FakePopup, Options, StubScraper,
                              page_state)


def export_with_users(path, usernames, unusable=0):
    """A page export with `usernames` valid rows plus `unusable` junk rows."""
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "search-match"])
    for name in usernames:
        sheet.append(["https://github.com/%s.png" % name,
                      "https://github.com/%s" % name, name.title()])
    for i in range(unusable):
        sheet.append(["", "not-a-github-link-%d" % i, "junk"])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    book.save(path)
    book.close()
    return path


class SinglePageCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="exact_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan(
            ["College Student"],
            [type("Ct", (), {"country": "India", "state": "Andhra Pradesh",
                             "canonical": "Vijayawada",
                             "search_aliases": ("Vijayawada",)})()])
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

    def workflow(self, total_users, export_usernames, unusable=0):
        state = page_state(reportedTotalUsers=str(total_users),
                           resultsListChildren=total_users,
                           profileLinkCount=total_users,
                           resultProfileLinkCount=total_users,
                           hasNext=False, nextByShape=0)
        page = FakePage(state=state)
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)

        rows_written = len(export_usernames) + unusable

        class OnesPage(StubScraper):
            async def finish_single_page(inner, popup, pg):
                inner.single_pages += 1
                return S.CrawlOutcome.SINGLE_PAGE, 1, rows_written

            async def export_xlsx(inner, popup, destination,
                                  download_dir=None):
                export_with_users(destination, export_usernames, unusable)
                inner.exports.append(destination)
                return rows_written

        wf = S.Workflow.__new__(S.Workflow)
        wf.options = Options()
        wf.context = self.context
        wf.searches = self.plan.assign(1, 1)
        wf.store = store
        wf.existing = {}
        wf._master_dirty = False
        wf.unit = self.unit
        wf.state = StateManager(path=self.context.state_file)
        wf.browser = FakeBrowser(page)
        wf.scraper = OnesPage(S.NextControl.SINGLE_PAGE, FakePopup(),
                              rows=rows_written)
        wf.reader = S.XlsxReader()
        wf.merger = S.StudentMerger()
        wf.transient = S.TransientFailureLog()
        wf.processed = 0
        wf.profiles = wf.portfolios = wf.resumes = None
        wf.keyword_number = 1
        wf.keyword_total = 1
        return wf, page

    def collect(self, wf, page):
        return asyncio.run(wf.collect(page, self.unit.search_query,
                                      self.target))

    @staticmethod
    def _stub_candidates(wf):
        """Replace the network stages so do_search_unit can run offline."""
        async def fake_candidate(_page, candidate):
            record = {col: "" for col in C.FINAL_COLUMNS}
            record["GitHub Profile"] = candidate.profile_url
            record["github"] = candidate.profile_url
            saved, _reason = wf.store.save_student(
                record, search_unit=wf.unit, campaign_id="",
                rank=1, total_ranks=1)
            if saved:
                wf.processed += 1
        wf.do_candidate = fake_candidate

    # -- 1. exact match succeeds -------------------------------------------
    def test_five_stated_five_usable_succeeds(self):
        wf, page = self.workflow(5, ["a", "b", "c", "d", "e"])
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 5)
        self.assertTrue(os.path.exists(self.target))
        self.assertEqual(len(load_workbook(self.target).active[1]), 3)

    # -- 2. shortfall is incomplete (Case 3) -------------------------------
    def test_five_stated_three_exported_is_incomplete(self):
        wf, page = self.workflow(5, ["a", "b", "c"])
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(rows, -1)
        self.assertFalse(os.path.exists(self.target),
                         "no complete keyword XLSX on a shortfall")
        self.assertTrue(os.path.exists(
            self.target.replace(".xlsx", ".partial.xlsx")),
            "the shortfall export is preserved for diagnosis")

    # -- 3. five rows but two unusable is incomplete -----------------------
    def test_five_stated_five_rows_two_unusable_is_incomplete(self):
        wf, page = self.workflow(5, ["a", "b", "c"], unusable=2)
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.FAILED,
                         "3 usable of 5 stated is incomplete even at 5 rows")
        self.assertFalse(os.path.exists(self.target))

    # -- 6/7/8. no XLSX, no advance, staging preserved ---------------------
    def test_a_shortfall_does_not_advance_the_search_unit(self):
        wf, page = self.workflow(5, ["a", "b", "c"])
        self._stub_candidates(wf)
        with self.assertRaises(S.Stopped):
            asyncio.run(wf.do_search_unit(page, self.unit))
        self.assertFalse(wf.state.collection_complete)
        self.assertFalse(wf.state.data["cleanup_complete"])
        self.assertEqual(wf.state.units_done, 0)
        self.assertEqual(wf.store.count(), 0)
        self.assertTrue(os.path.exists(
            self.target.replace(".xlsx", ".partial.xlsx")))

    # -- 10. normal complete single page unchanged -------------------------
    def test_a_complete_single_page_still_advances(self):
        wf, page = self.workflow(3, ["a", "b", "c"])
        self._stub_candidates(wf)
        asyncio.run(wf.do_search_unit(page, self.unit))
        self.assertTrue(wf.state.collection_complete)
        self.assertEqual(wf.store.count(), 3)

    # exact count only enforced when GitHub states one ---------------------
    def test_no_stated_total_accepts_whatever_is_there(self):
        wf, page = self.workflow_no_total(["a", "b", "c"])
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 3)

    def workflow_no_total(self, usernames):
        # page_state() default carries no exact total
        wf, page = self.workflow(3, usernames)
        page.state = page_state(reportedTotalUsers=None, hasNext=False,
                                nextByShape=0,
                                resultsListChildren=len(usernames),
                                profileLinkCount=len(usernames),
                                resultProfileLinkCount=len(usernames))
        return wf, page

    def test_an_abbreviated_total_is_not_enforced(self):
        wf, page = self.workflow(3, ["a", "b", "c"])
        page.state = page_state(reportedTotalUsers="9.1k", hasNext=False,
                                nextByShape=0, resultsListChildren=3,
                                profileLinkCount=3, resultProfileLinkCount=3)
        status, _rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.RESULTS)

    # -- 9. zero-results unchanged (never reaches the guard) ---------------
    def test_zero_results_is_unaffected(self):
        from test_zero_results import zero_state
        page = FakePage(state=zero_state())
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        wf = S.Workflow.__new__(S.Workflow)
        wf.options = Options(); wf.context = self.context; wf.plan = self.plan
        wf.searches = self.plan.assign(1, 1)
        wf.state = StateManager(path=self.context.state_file)
        wf.browser = FakeBrowser(page)
        wf.scraper = StubScraper(S.NextControl.SINGLE_PAGE, FakePopup())
        wf.reader = S.XlsxReader(); wf.merger = S.StudentMerger()
        wf.transient = S.TransientFailureLog()
        wf.processed = 0; wf.profiles = wf.portfolios = wf.resumes = None
        wf.keyword_number = 1; wf.keyword_total = 1
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.ZERO_RESULTS)
        self.assertEqual(rows, 0)


class TestVerifyExportCountHelper(unittest.TestCase):
    """The shared decision, unit-tested directly."""

    def helper(self, total, usable):
        wf = S.Workflow.__new__(S.Workflow)
        state = {"reportedTotalUsers": None if total is None else str(total)}
        return wf._verify_export_count(state, usable, usable, 1)

    def test_exact_match_ok(self):
        ok, _why = self.helper(5, 5)
        self.assertTrue(ok)

    def test_shortfall_refused(self):
        ok, why = self.helper(5, 3)
        self.assertFalse(ok)
        self.assertIn("5 users", why)
        self.assertIn("2 missing", why)

    def test_excess_refused(self):
        ok, why = self.helper(5, 6)
        self.assertFalse(ok)
        self.assertIn("changed", why)

    def test_no_total_accepts(self):
        self.assertTrue(self.helper(None, 3)[0])

    def test_abbreviated_total_accepts(self):
        wf = S.Workflow.__new__(S.Workflow)
        ok, _why = wf._verify_export_count({"reportedTotalUsers": "9.1k"},
                                           3, 3, 1)
        self.assertTrue(ok)


# ============================ multi-page shortfalls (4, 5) ==================
class TestMultiPageExactCount(unittest.TestCase):
    """The multi-page guard already existed; these pin the live cases."""

    def verify(self, total, unique, pages):
        return student_scraper.verify_collection(
            unique, total, ValidationMode.EXACT_TOTAL, pages,
            student_scraper.expected_page_count(total), last_page_had_next=False)

    def test_case1_ten_plus_three_of_fourteen_is_incomplete(self):
        # Engineering Student Visakhapatnam: 10 + 3 usable, GitHub 14
        ok, why = self.verify(14, 13, 2)
        self.assertFalse(ok)
        self.assertIn("missing", why)

    def test_case2_ten_ten_three_of_twentyfive_is_incomplete(self):
        # University Student Vijayawada: 10 + 10 + 3 usable, GitHub 25
        ok, why = self.verify(25, 23, 3)
        self.assertFalse(ok)
        self.assertIn("missing", why)

    def test_full_counts_pass(self):
        self.assertTrue(self.verify(14, 14, 2)[0])
        self.assertTrue(self.verify(25, 25, 3)[0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
