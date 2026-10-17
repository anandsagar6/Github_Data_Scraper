"""Dynamic-result reconciliation: GitHub's header total can change mid-crawl.

Live case (University Student Mohali): GitHub states 52, but six DOM-validated
pages (10+10+10+10+10+3) yield 53 unique, 0 duplicates, 0 unusable, pagination
ending naturally.  The strict count check would trap the rank forever.  So,
after such a reconcilable mismatch, ONE controlled reconciliation runs:

  * reopen page 1, re-read GitHub's total; accept only if it now equals the
    collected unique set;
  * otherwise run ONE clean verification re-crawl; if two fully-validated
    crawls produce the SAME identity set it is a STABLE mismatch, accepted
    ONLY with --accept-stable-pagination-mismatch (default STRICT stops).

Page-level DOM validation, dedupe and one-identity-per-row are never weakened.
"""

import asyncio
import copy
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper                                           # noqa: E402
from student_scraper import (InstantDataScraper, page_number)   # noqa: E402

import sandbox                                         # noqa: E402
from test_single_page import FakeLocator, page_state        # noqa: E402

SEARCH = "https://github.com/search?q=University+Student+Mohali&type=users"

# 10 + 10 + 10 + 10 + 10 + 3 = 53 unique identities across six pages.
# Hyphens only -- GitHub usernames never contain underscores.
PAGES = [["p%d-u%02d" % (p, i) for i in range(n)]
         for p, n in enumerate([10, 10, 10, 10, 10, 3])]


def _pnum(url):
    return page_number(url)


class Page:
    def __init__(self, pages, total):
        self.pages = pages
        self.pages_v2 = copy.deepcopy(pages)   # what a re-crawl sees
        self.total = total                     # GitHub's current header total
        self.index = 0
        self.variant = 0                       # 0 first crawl, 1 after goto
        self.closed = False

    @property
    def url(self):
        return SEARCH if self.index == 0 else "%s&p=%d" % (SEARCH,
                                                           self.index + 1)

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    async def goto(self, url, **_kw):
        # Reopening the search (reconcile step 1) resets to the requested page
        # and switches to the re-crawl variant.
        self.index = _pnum(url) - 1
        self.variant = 1

    @property
    def _active(self):
        return self.pages_v2 if self.variant else self.pages

    @property
    def has_next(self):
        return self.index < len(self._active) - 1

    def locator(self, selector):
        if selector == 'a[rel="next"]' and self.has_next:
            def _click():
                self.index += 1
            return FakeLocator(count=1, visible=True, on_click=_click)
        return FakeLocator(count=0)

    def state(self):
        users = self._active[self.index]
        return page_state(
            url=self.url, hasNext=self.has_next,
            nextByShape=1 if self.has_next else 0,
            nextDisabled=None if self.has_next else True,
            resultsListChildren=len(users), profileLinkCount=len(users),
            resultProfileLinkCount=len(users), resultProfileUsers=list(users),
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
    """IDS succeeds on every page (the mismatch is the total, not a bad page)."""

    last_export_recovered = False

    async def wait_for_stable_results(self, page, timeout=None):
        return True, page.state()

    async def open_instant_data_scraper(self, page):
        return Popup(page._active[page.index])

    async def wait_for_table(self, popup, timeout=None):
        return True

    async def read_stats(self, popup):
        return 1, len(popup.users)

    def dom_result_identities(self, state):
        return InstantDataScraper.dom_result_identities(state)

    async def stabilize_results_table(self, popup, page, expected_ids):
        return True

    async def preview_identities(self, popup, expected_ids=None):
        return set(expected_ids or ())

    async def export_page_from_preview(self, popup, destination, expected_ids):
        write_page(destination, popup.users)
        return len(popup.users)

    def export_identities(self, path):
        return InstantDataScraper.export_identities(path)


class ReconcileCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="recon_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        self.target = os.path.join(self.context.keyword_download_dir,
                                   "university-student-mohali.xlsx")
        self._orig_sleep = student_scraper.sleep_interruptible

        async def _noop(*_a, **_k):
            return None
        student_scraper.sleep_interruptible = _noop

    def tearDown(self):
        sandbox.restore()
        student_scraper.sleep_interruptible = self._orig_sleep
        shutil.rmtree(self.root, ignore_errors=True)

    def crawl(self, page, scraper):
        collector = student_scraper.PageByPageCollector(
            None, scraper, say=lambda *_a: None)
        asyncio.run(collector.collect(page, "university-student-mohali"))
        unique, dupes, unusable, ok, why = collector.merge_result(self.target)
        return collector, unique, dupes, unusable, ok

    def staged(self):
        s = student_scraper.staging_dir("university-student-mohali")
        return sorted(os.listdir(s)) if os.path.isdir(s) else []

    # 9. a normal exact-total crawl is unchanged (52 pages -> 52, accepted)
    def test_normal_exact_total_unchanged(self):
        page = Page([PAGES[0], PAGES[1], PAGES[2], PAGES[3], PAGES[4],
                     PAGES[5][:1]], total=51)      # 10*5+1 = 51 == header 51
        collector, unique, dupes, unusable, ok = self.crawl(page, Scraper())
        self.assertTrue(ok)
        self.assertEqual(unique, 51)
        self.assertFalse(collector.is_exact_total_mismatch(unique, dupes,
                                                           unusable))

    # the 52-vs-53 mismatch is recognised as reconcilable
    def test_mismatch_is_reconcilable(self):
        page = Page(PAGES, total=52)
        collector, unique, dupes, unusable, ok = self.crawl(page, Scraper())
        self.assertFalse(ok)
        self.assertEqual(unique, 53)
        self.assertTrue(collector.is_exact_total_mismatch(unique, dupes,
                                                          unusable))

    # 8. duplicates / unusable rows -> NOT reconcilable (normal stop)
    def test_duplicates_not_reconcilable(self):
        pages = copy.deepcopy(PAGES)
        pages[1][0] = pages[0][0]        # a duplicate across pages -> 52 unique
        page = Page(pages, total=53)     # header 53, so 52 unique is a shortfall
        collector, unique, dupes, unusable, ok = self.crawl(page, Scraper())
        self.assertFalse(ok)
        self.assertGreater(dupes, 0)
        self.assertFalse(collector.is_exact_total_mismatch(unique, dupes,
                                                          unusable))

    def _first_ids(self):
        return InstantDataScraper.export_identities(self.target)

    # 1. final GitHub total now equals the collected unique set -> reconciled
    def test_reconciled_when_final_total_matches(self):
        page = Page(PAGES, total=52)
        collector, unique, _d, _u, _ok = self.crawl(page, Scraper())
        page.total = 53                              # GitHub now shows 53
        accepted, info = asyncio.run(collector.reconcile(
            page, self.target, unique, self._first_ids(), accept_stable=False))
        self.assertTrue(accepted)
        self.assertEqual(info["kind"], "reconciled")
        self.assertEqual(info["final"], 53)

    # 2/3/4. final still 52, a second crawl yields the SAME 53 -> stable, and
    # strict (no flag) does NOT accept
    def test_stable_mismatch_not_accepted_in_strict(self):
        page = Page(PAGES, total=52)
        collector, unique, _d, _u, _ok = self.crawl(page, Scraper())
        page.total = 52                              # GitHub still shows 52
        accepted, info = asyncio.run(collector.reconcile(
            page, self.target, unique, self._first_ids(), accept_stable=False))
        self.assertEqual(info["kind"], "stable")     # recognised as stable
        self.assertEqual(info["second"], 53)
        self.assertFalse(accepted)                   # strict default stops

    # 5. explicit flag accepts the stable mismatch
    def test_stable_mismatch_accepted_with_flag(self):
        page = Page(PAGES, total=52)
        collector, unique, _d, _u, _ok = self.crawl(page, Scraper())
        page.total = 52
        accepted, info = asyncio.run(collector.reconcile(
            page, self.target, unique, self._first_ids(), accept_stable=True))
        self.assertEqual(info["kind"], "stable")
        self.assertTrue(accepted)

    # 6. the verification re-crawl produces a DIFFERENT set -> not accepted
    def test_changed_set_not_accepted(self):
        page = Page(PAGES, total=52)
        # the re-crawl variant swaps one user on the last page
        page.pages_v2[5] = ["different-x", "different-y", "different-z"]
        collector, unique, _d, _u, _ok = self.crawl(page, Scraper())
        page.total = 52
        accepted, info = asyncio.run(collector.reconcile(
            page, self.target, unique, self._first_ids(), accept_stable=True))
        self.assertEqual(info["kind"], "changed")
        self.assertFalse(accepted)

    # 7. a page-level DOM mismatch during the crawl STOPs before any reconcile
    def test_page_level_mismatch_stops_crawl(self):
        class WrongExport(Scraper):
            async def export_page_from_preview(self, popup, destination,
                                               expected_ids):
                # write a DIFFERENT identity than the page's DOM -> PROTECTION 2
                write_page(destination, ["someoneelse"])
                return 1
        page = Page(PAGES, total=53)
        collector = student_scraper.PageByPageCollector(
            None, WrongExport(), say=lambda *_a: None)
        with self.assertRaises(student_scraper.PageCrawlError):
            asyncio.run(collector.collect(page, "university-student-mohali"))
        self.assertFalse(os.path.exists(self.target))

    # 10. staging is preserved after a rejected reconcile (nothing deleted)
    def test_staging_preserved_after_reject(self):
        page = Page(PAGES, total=52)
        collector, unique, _d, _u, _ok = self.crawl(page, Scraper())
        page.total = 52
        asyncio.run(collector.reconcile(
            page, self.target, unique, self._first_ids(), accept_stable=False))
        # the verifier re-crawl re-populated staging; page files are still there
        self.assertTrue(len(self.staged()) >= 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
