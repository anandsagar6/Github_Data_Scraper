"""A small page inside a multi-page crawl that IDS can't establish is recovered
from its exact DOM identities -- the SAME fallback used for a whole single-page
search -- instead of discarding the pages already staged.

THE LIVE CASE THIS FILE EXISTS FOR
----------------------------------
    B.Tech Student Guntur -- GitHub exact total 12
    Page 1: 10 users, IDS exact -> page_001.xlsx
    Page 2: 2 users, IDS cannot establish the table
Before: the whole search failed with page 1 already staged.
After:  page 2 is written from its exact DOM identities, merged 10 + 2 = 12,
        which matches GitHub's stated total.  A wrong/ambiguous DOM never
        applies, and a merged total that does not equal the stated total still
        STOPs (staging preserved, no final keyword XLSX).
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper                                           # noqa: E402
from student_scraper import (InstantDataScraper, page_number)   # noqa: E402

import sandbox                                         # noqa: E402
from openpyxl import load_workbook                          # noqa: E402
from test_single_page import FakeLocator, page_state        # noqa: E402

SEARCH = "https://github.com/search?q=B.Tech+Student+Guntur&type=users"


class Page:
    """A GitHub tab paginating through scripted pages, each exposing its DOM
    result users (resultProfileUsers) so the identity path runs."""

    def __init__(self, pages, total, counts=None):
        self.pages = pages            # [[usernames], ...]
        self.counts = counts or [len(p) for p in pages]   # card count per page
        self.total = total
        self.index = 0
        self.closed = False

    @property
    def url(self):
        return SEARCH if self.index == 0 else "%s&p=%d" % (SEARCH,
                                                           self.index + 1)

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    @property
    def has_next(self):
        return self.index < len(self.pages) - 1

    def locator(self, selector):
        if selector == 'a[rel="next"]' and self.has_next:
            def _click():
                self.index += 1
            return FakeLocator(count=1, visible=True, on_click=_click)
        return FakeLocator(count=0)

    def state(self):
        users = self.pages[self.index]
        count = self.counts[self.index]
        return page_state(
            url=self.url, hasNext=self.has_next,
            nextByShape=1 if self.has_next else 0,
            nextDisabled=None if self.has_next else True,
            resultsListChildren=count, profileLinkCount=count,
            resultProfileLinkCount=count, resultProfileUsers=list(users),
            reportedTotalUsers=str(self.total))


class Popup:
    def __init__(self, users):
        self.users = users
        self.closed = False

    def is_closed(self):
        return self.closed

    async def close(self):
        self.closed = True

    async def bring_to_front(self):
        return None


def write_page(destination, usernames):
    """A staged page export IDS would have produced (avatar + profile URL)."""
    from openpyxl import Workbook
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "name"])
    for name in usernames:
        sheet.append(["https://github.com/%s.png" % name,
                      "https://github.com/%s" % name, name.title()])
    book.save(destination)
    book.close()


class Scraper:
    """Only the surface PageByPageCollector uses; stabilize fails on chosen
    pages so the DOM fallback is exercised."""

    last_export_recovered = False

    def __init__(self, stabilize_fail_on=()):
        self.stabilize_fail_on = set(stabilize_fail_on)
        self.preview_exports = []

    async def wait_for_stable_results(self, page, timeout=None):
        return True, page.state()

    async def open_instant_data_scraper(self, page):
        return Popup(page.pages[page.index])

    async def wait_for_table(self, popup, timeout=None):
        return True

    async def read_stats(self, popup):
        return 1, len(popup.users)

    def dom_result_identities(self, state):
        return InstantDataScraper.dom_result_identities(state)

    async def stabilize_results_table(self, popup, page, expected_ids):
        return page_number(page.url) not in self.stabilize_fail_on

    async def preview_identities(self, popup, expected_ids=None):
        return set(expected_ids or ())

    async def export_page_from_preview(self, popup, destination, expected_ids):
        self.preview_exports.append(page_number_of(destination))
        write_page(destination, popup.users)
        return len(popup.users)

    def export_identities(self, path):
        return InstantDataScraper.export_identities(path)


def page_number_of(destination):
    return os.path.basename(destination)


class MultiPageFallbackCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="mpf_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        self.target = os.path.join(self.context.keyword_download_dir,
                                   "btech-student-guntur.xlsx")
        self._orig_sleep = student_scraper.sleep_interruptible

        async def _noop(*_a, **_k):
            return None
        student_scraper.sleep_interruptible = _noop     # no real pacing waits

    def tearDown(self):
        sandbox.restore()
        student_scraper.sleep_interruptible = self._orig_sleep
        shutil.rmtree(self.root, ignore_errors=True)

    def collect(self, page, scraper):
        collector = student_scraper.PageByPageCollector(
            None, scraper, say=lambda *_a: None)
        asyncio.run(collector.collect(page, "btech-student-guntur"))
        return collector

    def collect_and_merge(self, pages, total, stabilize_fail_on=(),
                          counts=None):
        page = Page(pages, total, counts=counts)
        scraper = Scraper(stabilize_fail_on=stabilize_fail_on)
        collector = self.collect(page, scraper)
        rows, why = collector.merge_into(self.target)
        return rows, why, collector

    # ---- helpers -------------------------------------------------------
    def staging(self):
        return student_scraper.staging_dir("btech-student-guntur")

    def staged_files(self):
        s = self.staging()
        return sorted(n for n in os.listdir(s)) if os.path.isdir(s) else []

    P1 = ["u%02d" % i for i in range(10)]
    P2 = ["v0", "v1"]

    # 1. page1=10 IDS exact, page2=2 IDS fails -> DOM fallback -> merged 12
    def test_small_page_dom_fallback_completes_twelve(self):
        rows, _why, _c = self.collect_and_merge(
            [self.P1, self.P2], total=12, stabilize_fail_on={2})
        self.assertEqual(rows, 12)
        self.assertTrue(os.path.exists(self.target))
        # the merged workbook holds exactly the 12 unique identities
        self.assertEqual(
            len(InstantDataScraper.export_identities(self.target)), 12)

    # 6. the fallback page has exactly one identity per row
    def test_fallback_page_one_identity_per_row(self):
        page = Page([self.P1, self.P2], total=12)
        scraper = Scraper(stabilize_fail_on={2})
        self.collect(page, scraper)
        page2 = student_scraper.page_path(self.staging(), 2)
        book = load_workbook(page2, data_only=True)
        body = [r for r in book.active.iter_rows(min_row=2, values_only=True)]
        book.close()
        ids = [InstantDataScraper._identities_from_cells(
            ["" if c is None else str(c) for c in r]) for r in body]
        self.assertEqual(len(ids), 2)
        self.assertTrue(all(len(i) == 1 for i in ids), ids)
        self.assertEqual(set().union(*ids), {"v0", "v1"})

    # 8. normal IDS multi-page path is unchanged (no fallback when IDS works)
    def test_normal_multipage_path_unchanged(self):
        page = Page([self.P1, self.P2], total=12)
        scraper = Scraper(stabilize_fail_on=())     # IDS establishes both pages
        collector = self.collect(page, scraper)
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 12)
        # both pages went through the verified-preview export, not the fallback
        self.assertEqual(sorted(scraper.preview_exports),
                         ["page_001.xlsx", "page_002.xlsx"])

    # 2. page2 DOM=2 but the DOM yields 3 identities -> STOP
    def test_page_count_vs_identities_mismatch_stops(self):
        # card count says 2, but resultProfileUsers lists 3 -> ambiguous
        page = Page([self.P1, ["v0", "v1", "v2"]], total=13, counts=[10, 2])
        scraper = Scraper(stabilize_fail_on={2})
        with self.assertRaises(student_scraper.PageCrawlError):
            self.collect(page, scraper)
        self.assertFalse(os.path.exists(self.target))       # no final XLSX
        self.assertIn("page_001.xlsx", self.staged_files())  # staging preserved

    # 3. page2 has a duplicate identity -> STOP
    def test_duplicate_identity_stops(self):
        page = Page([self.P1, ["v0", "v0"]], total=12, counts=[10, 2])
        scraper = Scraper(stabilize_fail_on={2})
        with self.assertRaises(student_scraper.PageCrawlError):
            self.collect(page, scraper)
        self.assertFalse(os.path.exists(self.target))

    # 4. page2 expected 2 but only 1 valid DOM identity -> STOP
    def test_only_one_valid_identity_stops(self):
        page = Page([self.P1, ["v0", "not a handle"]], total=12, counts=[10, 2])
        scraper = Scraper(stabilize_fail_on={2})
        with self.assertRaises(student_scraper.PageCrawlError):
            self.collect(page, scraper)
        self.assertFalse(os.path.exists(self.target))

    # 5. final merged total != GitHub's exact total -> STOP (no final XLSX)
    def test_final_total_mismatch_stops(self):
        page = Page([self.P1, self.P2], total=13, counts=[10, 2])  # states 13
        scraper = Scraper(stabilize_fail_on={2})
        collector = self.collect(page, scraper)             # staging = 10 + 2
        with self.assertRaises(student_scraper.PageCrawlError):
            collector.merge_into(self.target)
        self.assertFalse(os.path.exists(self.target))
        # 9. staging preserved on failure
        self.assertEqual(self.staged_files(),
                         ["page_001.xlsx", "page_002.xlsx"])

    # 10. no final keyword XLSX exists until every page validated & merged
    def test_no_final_xlsx_until_validated(self):
        page = Page([self.P1, ["v0", "v1", "v2"]], total=13, counts=[10, 2])
        scraper = Scraper(stabilize_fail_on={2})
        with self.assertRaises(student_scraper.PageCrawlError):
            self.collect(page, scraper)
        self.assertFalse(os.path.exists(self.target))

    # 7. the shared single-page fallback core still works (reuse, not a copy)
    def test_single_page_core_still_works(self):
        from student_scraper import write_dom_fallback
        state = page_state(resultProfileUsers=["a", "b", "c"],
                           reportedTotalUsers="3")
        expected = InstantDataScraper.dom_result_identities(state)
        dest = os.path.join(self.root, "single.xlsx")
        ok, count, reason = write_dom_fallback(state, dest, expected, 3)
        self.assertTrue(ok, reason)
        self.assertEqual(count, 3)
        self.assertEqual(InstantDataScraper.export_identities(dest), expected)


if __name__ == "__main__":
    unittest.main(verbosity=2)
