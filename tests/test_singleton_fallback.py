"""Small exact single-page result sets IDS can't infer -> use verified DOM.

THE LIVE BLOCKERS THIS FILE EXISTS FOR
--------------------------------------
    MCA Student Visakhapatnam            -> exact 1  (IDS can't establish)
    Student Developer Visakhapatnam      -> exact 2  (IDS can't establish)
    Final Year Student Vijayawada        -> exact 3  (IDS can't establish)

A small single-page result set has too little repetition for IDS to infer a
table.  When GitHub states an EXACT total between 1 and 10 AND the results DOM
shows exactly that many valid UNIQUE github identities, the scraper writes one
row per identity from the ALREADY-VERIFIED DOM identities rather than losing
them.  This carries ONLY candidate identities (profile URLs) -- the profile
pipeline still collects every detail.  Normal IDS is always tried first; this
runs only when IDS fails and all conditions hold, and never for multi-page,
abbreviated/unknown totals, or counts above 10.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper                                           # noqa: E402
from student_scraper import (CrawlOutcome, InstantDataScraper, NextControl)   # noqa: E402

import student_scraper as S                                # noqa: E402
from test_single_page import FakeBrowser, FakePage, page_state  # noqa: E402


def st(users, total):
    """A minimal PAGE_STATE_JS-shaped dict for the direct fallback tests."""
    return {"reportedTotalUsers": str(total),
            "resultProfileUsers": list(users)}


def full_state(users, total):
    """A complete single-page state that classify_search_page accepts."""
    return page_state(reportedTotalUsers=str(total),
                      resultProfileUsers=list(users),
                      resultsListChildren=len(users),
                      profileLinkCount=len(users),
                      resultProfileLinkCount=len(users),
                      hasNext=False, nextByShape=0)


def workflow(scraper=None):
    wf = S.Workflow.__new__(S.Workflow)
    wf.scraper = scraper or InstantDataScraper(browser=None, state=None,
                                               say=lambda *_a: None)
    wf.reader = S.XlsxReader()
    return wf


def expected_of(users):
    return InstantDataScraper.dom_result_identities(
        {"resultProfileUsers": list(users)})


def names(n):
    return ["user%02d" % i for i in range(n)]


class TestExactSinglePageDomFallback(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="domfb_")
        self.target = os.path.join(self.dir, "kw.xlsx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def fallback(self, state, expected=None, scraper=None,
                 outcome=CrawlOutcome.SINGLE_PAGE):
        wf = workflow(scraper)
        exp = expected if expected is not None else \
            InstantDataScraper.dom_result_identities(state)
        return wf._exact_single_page_dom_fallback(state, outcome, self.target,
                                                  "kw", exp)

    def _works(self, count):
        users = names(count)
        res = self.fallback(st(users, count))
        self.assertEqual(res, (S.CollectResult.RESULTS, count))
        self.assertEqual(InstantDataScraper.export_identities(self.target),
                         expected_of(users))
        self.assertEqual(len(student_scraper.XlsxReader().read(self.target, "kw")),
                         count)

    # 1-5. exact 1 / 2 / 3 / 5 / 10 -> fallback works
    def test_exact_one(self):
        self._works(1)

    def test_exact_two(self):
        self._works(2)

    def test_exact_three(self):
        self._works(3)

    def test_exact_five(self):
        self._works(5)

    def test_exact_ten(self):
        self._works(10)

    # 6. exact count=3 but DOM=2 -> incomplete
    def test_count_three_dom_two_is_incomplete(self):
        res = self.fallback(st(["a", "b"], 3))
        self.assertEqual(res, (S.CollectResult.FAILED, -1))
        self.assertFalse(os.path.exists(self.target))

    # 7. exact count=3 but DOM=4 -> incomplete
    def test_count_three_dom_four_is_incomplete(self):
        res = self.fallback(st(["a", "b", "c", "d"], 3))
        self.assertEqual(res, (S.CollectResult.FAILED, -1))
        self.assertFalse(os.path.exists(self.target))

    # 8. a duplicate DOM identity reduces the unique count -> incomplete
    def test_duplicate_identity_is_incomplete(self):
        # GitHub says 3, DOM lists a,b,a -> only 2 unique -> mismatch
        res = self.fallback(st(["a", "b", "a"], 3))
        self.assertEqual(res, (S.CollectResult.FAILED, -1))
        self.assertFalse(os.path.exists(self.target))

    # 9. an invalid identity -> incomplete (count matches but handle unusable)
    def test_invalid_identity_is_incomplete(self):
        # GitHub states 3, but one DOM handle is not a usable github username,
        # so only 2 valid unique identities exist -> the counts disagree.
        state = st(["alice", "bob", "not a handle"], 3)
        expected = InstantDataScraper.dom_result_identities(state)  # {alice,bob}
        res = self.fallback(state, expected=expected)
        self.assertEqual(res, (S.CollectResult.FAILED, -1))
        self.assertFalse(os.path.exists(self.target))

    # 10. count = 11 -> no fallback (declines)
    def test_count_eleven_declines(self):
        users = names(11)
        res = self.fallback(st(users, 11))
        self.assertIsNone(res)
        self.assertFalse(os.path.exists(self.target))

    # 11. a multi-page / non-single-page outcome -> no fallback
    def test_multi_page_declines(self):
        res = self.fallback(st(["a", "b"], 2),
                            outcome=CrawlOutcome.FINISHED)
        self.assertIsNone(res)

    # abbreviated / unknown totals -> no fallback
    def test_abbreviated_total_declines(self):
        res = self.fallback(st(["a"], "9.1k"))
        self.assertIsNone(res)

    # 13. readback must equal the DOM identity SET exactly.  The shared
    #     write_dom_fallback verifies the file it wrote; force the writer to
    #     emit the wrong identities and confirm the guard rejects it.
    def test_readback_mismatch_is_rejected(self):
        import student_scraper as IDS_MOD
        state = st(["alice", "bob", "carol"], 3)
        expected = InstantDataScraper.dom_result_identities(state)
        original = IDS_MOD.InstantDataScraper.write_dom_candidates_xlsx

        def wrong(destination, _usernames):
            return original(destination, ["someoneelse", "another", "third"])
        IDS_MOD.InstantDataScraper.write_dom_candidates_xlsx = staticmethod(
            wrong)
        try:
            ok, count, reason = IDS_MOD.write_dom_fallback(
                state, self.target, expected, 3)
        finally:
            IDS_MOD.InstantDataScraper.write_dom_candidates_xlsx = original
        self.assertFalse(ok)
        self.assertIn("read back", reason)

    # 14. the written file has exactly one identity per row
    def test_one_identity_per_row(self):
        from openpyxl import load_workbook
        users = names(3)
        self.fallback(st(users, 3))
        book = load_workbook(self.target, data_only=True)
        sheet = book.active
        rows = [["" if c is None else str(c) for c in r]
                for r in sheet.iter_rows(values_only=True)]
        book.close()
        self.assertEqual(len(rows), 1 + 3)             # header + one per user
        per_row = [InstantDataScraper._identities_from_cells(r)
                   for r in rows[1:]]
        self.assertTrue(all(len(ids) == 1 for ids in per_row), per_row)
        self.assertEqual(set().union(*per_row), expected_of(users))

    # 15. the fallback rows carry ONLY candidate identity, no profile data
    def test_rows_have_no_profile_data(self):
        self.fallback(st(["abc", "def"], 2))
        cands = student_scraper.XlsxReader().read(self.target, "kw")
        self.assertEqual({c.username for c in cands}, {"abc", "def"})
        for cand in cands:
            self.assertEqual(cand.profile_url,
                             "https://github.com/%s" % cand.username)
            self.assertEqual(cand.xlsx_name, "")
            self.assertEqual(cand.xlsx_bio, "")
            self.assertEqual(cand.xlsx_location, "")


# ============================ through collect() (wiring) ====================
class IntegrationScraper:
    """Enough of the scraper for collect()'s single-page path."""

    def __init__(self, dom_users, total, stabilize_ok):
        self.dom_users = dom_users
        self.total = total
        self.stabilize_ok = stabilize_ok
        self.dom_written = False
        self.preview_exported = False

    async def capture_page_state(self, page, reason=""):
        return dict(page.state)

    def log_state(self, state):
        return None

    async def open_instant_data_scraper(self, page):
        return object()

    async def ensure_next_selector(self, popup, page):
        return NextControl.SINGLE_PAGE

    async def finish_single_page(self, popup, page):
        return CrawlOutcome.SINGLE_PAGE, 1, len(self.dom_users)

    def dom_result_identities(self, state):
        return InstantDataScraper.dom_result_identities(state)

    def dom_usernames(self, state):
        return InstantDataScraper.dom_usernames(state)

    def write_dom_candidates_xlsx(self, target, usernames):
        self.dom_written = True
        return InstantDataScraper.write_dom_candidates_xlsx(target, usernames)

    def export_identities(self, path):
        return InstantDataScraper.export_identities(path)

    async def stabilize_results_table(self, popup, page, expected):
        return self.stabilize_ok

    async def preview_identities(self, popup, expected_ids=None):
        import student_scraper as _id
        return {_id.normalize_username(u) for u in self.dom_users}

    async def export_page_from_preview(self, popup, target, expected):
        self.preview_exported = True
        InstantDataScraper.write_dom_candidates_xlsx(target, self.dom_users)
        return len(self.dom_users)


class TestFallbackThroughCollect(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="domfbint_")
        self.target = os.path.join(self.dir, "kw.xlsx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _collect(self, scraper):
        page = FakePage(state=full_state(scraper.dom_users, scraper.total))
        wf = workflow(scraper)
        wf.browser = FakeBrowser(page)
        wf.options = types.SimpleNamespace(max_pages=100)
        wf.context = types.SimpleNamespace(rank=1)
        return asyncio.run(wf.collect(page, "Final Year Student Vijayawada",
                                      self.target))

    # IDS fails on an exact-3 single page -> DOM fallback collects all three
    def test_ids_fails_dom_fallback_collects_three(self):
        sc = IntegrationScraper(["aaa", "bbb", "ccc"], 3, stabilize_ok=False)
        status, rows = self._collect(sc)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 3)
        self.assertFalse(sc.preview_exported)      # not the normal IDS path
        self.assertTrue(os.path.exists(self.target))
        self.assertEqual(InstantDataScraper.export_identities(self.target),
                         {"aaa", "bbb", "ccc"})

    # 12. when IDS establishes, the normal path is preferred (no fallback)
    def test_ids_works_uses_normal_path(self):
        sc = IntegrationScraper(["aaa", "bbb", "ccc"], 3, stabilize_ok=True)
        status, rows = self._collect(sc)
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertTrue(sc.preview_exported)       # normal IDS path was used


if __name__ == "__main__":
    unittest.main(verbosity=2)
