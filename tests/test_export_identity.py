"""The export must match the DOM by IDENTITY SET, not just by count.

THE LIVE BUG THIS FILE EXISTS FOR (Engineering Student Visakhapatnam p2)
------------------------------------------------------------------------
The extension's preview (#hot) and its XLSX export are both built from the
SAME model (mn.data, through sr()), verified in the extension source:

    preview:  Or() -> sr(mn.data)          -> setGrid -> Handsontable
    export:   fl() -> cr(mn.data)=sr(...)  -> XLSX

so they cannot differ at one instant -- but mn.data is mutated
asynchronously by "Try another table" (getTableData + server processing),
so a table validated a moment ago can shift before the XLSX click.  Live:

    DOM page 2 = {Suresh240618, Dorothy-21, Sri1234rama, Nasirkhan6}
    preview establishment = 4/4 exact
    XLSX export           = Suresh240618 + "images" + page-chrome rows

Two guards defend this, and BOTH are tested here:

    PROTECTION 1 (pre-export)  re-read the preview immediately before the
                               click; refuse if it no longer matches the DOM.
    PROTECTION 2 (post-export) the exported file's identities must EXACTLY
                               equal the DOM identities -- a set, not a count.
                               This is what rejects the Engineering p2 export.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook                             # noqa: E402

import config as C                                        # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (InstantDataScraper, StateManager)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                               # noqa: E402
from test_single_page import (FakeBrowser, FakePage,       # noqa: E402
                              FakePopup, Options, StubScraper,
                              page_state)

ENG_P2 = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]


def write_identities(path, usernames, garbage=()):
    """A page export with valid github rows for `usernames` plus garbage."""
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "search-match"])
    for name in usernames:
        sheet.append(["https://github.com/%s.png" % name,
                      "https://github.com/%s" % name, name.title()])
    for text in garbage:
        sheet.append(["", text, "chrome"])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    book.save(path)
    book.close()
    return path


# =========================== the export_identities helper directly ==========
class TestExportIdentities(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="expid_")

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_reads_usable_identities_from_an_export(self):
        path = write_identities(os.path.join(self.dir, "e.xlsx"), ENG_P2)
        ids = InstantDataScraper.export_identities(path)
        self.assertEqual(ids, {"suresh240618", "dorothy-21", "sri1234rama",
                               "nasirkhan6"})

    def test_the_engineering_bad_export_yields_only_one_real_user(self):
        # 1 real user + page-chrome rows, like the live bad export
        path = write_identities(
            os.path.join(self.dir, "bad.xlsx"), ["Suresh240618"],
            garbage=["Filter Users 14 results", "Previous 1 2 Next",
                     "images"])
        ids = InstantDataScraper.export_identities(path)
        self.assertEqual(ids, {"suresh240618"},
                         "page-chrome and 'images' are not GitHub identities")
        self.assertNotEqual(ids, {n.lower() for n in ENG_P2})


# =============================== single-page pre/post guards =================
class SinglePageGuardCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="idguard_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan(
            ["Engineering Student"],
            [type("Ct", (), {"country": "India", "state": "Andhra Pradesh",
                             "canonical": "Visakhapatnam",
                             "search_aliases": ("Visakhapatnam",)})()])
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

    def workflow(self, dom_users, export_users, export_garbage=(),
                 preview_users=None, total=None):
        """preview_users defaults to dom_users (preview matches);
        set it different to simulate the table shifting before the click."""
        state = page_state(
            reportedTotalUsers=str(total if total is not None
                                   else len(dom_users)),
            resultProfileUsers=list(dom_users),
            resultsListChildren=len(dom_users),
            profileLinkCount=len(dom_users),
            resultProfileLinkCount=len(dom_users),
            hasNext=False, nextByShape=0)
        page = FakePage(state=state)
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        preview = preview_users if preview_users is not None else dom_users
        eg = export_garbage

        class Sc(StubScraper):
            async def finish_single_page(inner, popup, pg):
                inner.single_pages += 1
                return S.CrawlOutcome.SINGLE_PAGE, 1, len(export_users) + len(eg)

            async def establish_results_table(inner, popup, page, expected):
                return True

            async def preview_identities(inner, popup, expected_ids=None):
                import student_scraper as _id
                return {_id.normalize_username(u) for u in preview}

            async def export_xlsx(inner, popup, destination,
                                  download_dir=None):
                write_identities(destination, export_users, eg)
                inner.exports.append(destination)
                return len(export_users) + len(eg)

        wf = S.Workflow.__new__(S.Workflow)
        wf.options = Options(); wf.context = self.context; wf.plan = self.plan
        wf.searches = self.plan.assign(1, 1)
        wf.state = StateManager(path=self.context.state_file)
        wf.browser = FakeBrowser(page)
        wf.scraper = Sc(S.NextControl.SINGLE_PAGE, FakePopup(),
                        rows=len(export_users) + len(eg))
        wf.reader = S.XlsxReader(); wf.merger = S.StudentMerger()
        wf.transient = S.TransientFailureLog()
        wf.processed = 0; wf.profiles = wf.portfolios = wf.resumes = None
        wf.keyword_number = 1; wf.keyword_total = 1
        return wf, page

    def collect(self, wf, page):
        return asyncio.run(wf.collect(page, self.unit.search_query,
                                      self.target))

    # 1. preview exact + XLSX exact -> success
    def test_preview_and_export_both_exact_succeeds(self):
        wf, page = self.workflow(ENG_P2, ENG_P2)
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 4)
        self.assertTrue(os.path.exists(self.target))

    # 2. preview exact but XLSX identities differ -> incomplete
    def test_preview_exact_but_export_identities_differ_is_incomplete(self):
        # export has Suresh + three OTHER users (right count, wrong identities)
        wf, page = self.workflow(
            ENG_P2, ["Suresh240618", "someoneelse", "another", "third"])
        status, _rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertFalse(os.path.exists(self.target))
        self.assertTrue(os.path.exists(
            self.target.replace(".xlsx", ".partial.xlsx")))

    # 3. the exact live case: preview 4/4, export = 1 user + chrome (+"images")
    def test_the_engineering_p2_bad_export_is_rejected(self):
        wf, page = self.workflow(
            ENG_P2, ["Suresh240618"],
            export_garbage=["Filter Users 14 results", "Previous Next",
                            "images"])
        status, _rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.FAILED,
                         "an export missing 3 of 4 users must be rejected")
        self.assertFalse(os.path.exists(self.target))

    # 4. preview changes before XLSX click -> refuse export
    def test_preview_shift_before_click_refuses_export(self):
        # preview now shows a different set than the DOM -> pre-export guard
        wf, page = self.workflow(ENG_P2, ENG_P2,
                                 preview_users=["Suresh240618"])
        exports_before = []

        status, _rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.FAILED)
        self.assertEqual(wf.scraper.exports, [],
                         "export must not run when the preview shifted")

    # 8/9. set equality, not counts
    def test_same_count_wrong_identities_rejected(self):
        wf, page = self.workflow(["a", "b", "c"], ["a", "x", "y"])
        status, _rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.FAILED)

    def test_correct_identities_in_any_order_accepted(self):
        wf, page = self.workflow(["a", "b", "c"], ["c", "a", "b"])
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 3)

    # 10. a genuinely correct 5/5 page still succeeds (the passing pages)
    def test_a_clean_five_user_page_still_succeeds(self):
        users = ["u1", "u2", "u3", "u4", "u5"]
        wf, page = self.workflow(users, users)
        status, rows = self.collect(wf, page)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 5)


# =============================== page-by-page guards ========================
class PageGuardCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pgguard_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        self.slug = "engineering-student-visakhapatnam"
        self.said = []

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def _export_current_page(self, dom_users, export_users, export_garbage=(),
                             preview_users=None):
        """Drive PageByPageCollector._export_current_page with a fake scraper."""
        staging = self.context.staging_for(self.slug)
        os.makedirs(staging, exist_ok=True)
        preview = preview_users if preview_users is not None else dom_users
        eg = export_garbage

        class Sc:
            def __init__(self):
                self.exports = []

            async def open_instant_data_scraper(self, page):
                return FakePopup()

            async def wait_for_table(self, popup, timeout=None):
                return True

            async def read_stats(self, popup):
                return 1, len(export_users) + len(eg)

            async def establish_results_table(self, popup, page, expected):
                return True

            async def stabilize_results_table(self, popup, page, expected):
                return True

            async def export_page_from_preview(self, popup, destination,
                                               expected):
                write_identities(destination, export_users, eg)
                self.exports.append(destination)
                return len(export_users) + len(eg)

            async def preview_identities(self, popup, expected_ids=None):
                import student_scraper as _id
                return {_id.normalize_username(u) for u in preview}

            def dom_result_identities(self, state):
                return InstantDataScraper.dom_result_identities(state)

            def export_identities(self, path):
                return InstantDataScraper.export_identities(path)

            async def export_xlsx(self, popup, destination, download_dir=None):
                write_identities(destination, export_users, eg)
                self.exports.append(destination)
                return len(export_users) + len(eg)

        collector = student_scraper.PageByPageCollector(
            None, Sc(), self.said.append)
        collector.staging = staging
        import student_scraper as _id
        expected = {_id.normalize_username(u) for u in dom_users}
        return collector, expected

    def run_export(self, **kw):
        collector, expected = self._export_current_page(**kw)
        return asyncio.run(collector._export_current_page(
            FakePage(state=page_state()), 2, expected)), collector

    def test_a_matching_page_export_succeeds(self):
        result, collector = self.run_export(dom_users=ENG_P2,
                                            export_users=ENG_P2)
        self.assertEqual(result.rows, 4)
        self.assertTrue(os.path.exists(
            student_scraper.page_path(collector.staging, 2)))

    def test_the_engineering_p2_bad_page_export_is_rejected(self):
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            self.run_export(dom_users=ENG_P2, export_users=["Suresh240618"],
                            export_garbage=["Filter", "Previous Next",
                                            "images"])
        self.assertIn("do not match", str(caught.exception))

    def test_a_preview_shift_before_page_export_is_rejected(self):
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            self.run_export(dom_users=ENG_P2, export_users=ENG_P2,
                            preview_users=["Suresh240618"])
        self.assertIn("changed after validation", str(caught.exception))

    def test_same_count_wrong_identities_rejected_per_page(self):
        with self.assertRaises(student_scraper.PageCrawlError):
            self.run_export(dom_users=["a", "b", "c"],
                            export_users=["a", "x", "y"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
