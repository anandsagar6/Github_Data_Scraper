"""Page-by-page collection: one fresh IDS per page, ONE verified keyword XLSX.

WHY THIS ARCHITECTURE
---------------------
The diagnostic proved GitHub REPLACES the results DOM when it paginates, and
that the popup stays bound to the page it was opened on -- while a FRESH popup
on page 2 detects all 9 users there.  So the extension is reinitialised per
page, each page is exported to rank-local staging, and the staged pages are
merged into ONE keyword XLSX only after every page has succeeded.

A popup holds its rows only while it lives, so reinitialising for page 2
necessarily discards page 1.  Exporting each page before replacing the popup
is forced by the extension, not chosen.

ATOMICITY, UNCHANGED
--------------------
page 1 succeeds + page 2 fails  ->  NO keyword XLSX, staging preserved,
                                    collection.complete stays False.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook             # noqa: E402

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (StateManager, ValidationMode, expected_page_count, expected_total_users)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                              # noqa: E402
from test_single_page import FakeLocator, page_state      # noqa: E402

SEARCH = "https://github.com/search?q=University+Student+Visakhapatnam&type=users"


# ============================================================== fake world ==
def write_page_export(path, usernames):
    """An Instant Data Scraper style page export."""
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "search-match",
                  "user-location"])
    for name in usernames:
        sheet.append(["https://github.com/%s.png?size=40" % name,
                      "https://github.com/%s" % name,
                      name.title(), "Visakhapatnam"])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    book.save(path)
    book.close()
    return path


def rows_of(path):
    book = load_workbook(path, read_only=True)
    sheet = book.active
    rows = [r for i, r in enumerate(sheet.iter_rows(values_only=True))
            if i and any(v not in (None, "") for v in r)]
    book.close()
    return rows


class FakeCrawlPage:
    """A GitHub tab that paginates through a scripted set of pages."""

    def __init__(self, pages, total=None, next_broken=False,
                 stall_page=None):
        # pages: [[usernames], ...] in page order
        self.pages = pages
        self.index = 0
        self.total = total
        self.next_broken = next_broken
        self.stall_page = stall_page
        self.closed = False
        self.next_clicks = 0

    # -- the bits the collector uses ---------------------------------------
    @property
    def url(self):
        return SEARCH if self.index == 0 else "%s&p=%d" % (SEARCH,
                                                           self.index + 1)

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    async def reload(self, **_kwargs):
        """Mirrors a Playwright page: _stable_page() reloads a page that did
        not settle before abandoning the search."""
        self.reloads = getattr(self, "reloads", 0) + 1

    async def wait_for_function(self, script, arg=None, timeout=None):
        return True

    @property
    def has_next(self):
        return self.index < len(self.pages) - 1

    def locator(self, selector):
        if selector == 'a[rel="next"]' and self.has_next:
            def _click():
                self.next_clicks += 1
                if self.next_broken:
                    raise RuntimeError("element detached")
                self.index += 1
            return FakeLocator(count=1, visible=True, on_click=_click)
        return FakeLocator(count=0)

    # -- what the page probe would report ----------------------------------
    def state(self):
        users = self.pages[self.index]
        stalled = (self.stall_page is not None
                   and self.index + 1 == self.stall_page)
        return page_state(
            url=self.url,
            hasNext=self.has_next,
            nextByShape=1 if self.has_next else 0,
            nextDisabled=None if self.has_next else True,
            resultsListChildren=0 if stalled else len(users),
            profileLinkCount=0 if stalled else len(users),
            resultProfileLinkCount=0 if stalled else len(users),
            reportedTotalUsers=None if self.total is None else str(self.total),
            skeletonCount=1 if stalled else 0,
            readyState="complete",
            domSignature=1000 + self.index,
        )


class FakePagePopup:
    """A popup bound to whatever page it was opened on."""

    def __init__(self, usernames, page_url, fail_table=False,
                 fail_export=False):
        self.usernames = usernames
        self.page_url = page_url
        self.fail_table = fail_table
        self.fail_export = fail_export
        self.closed = False
        self.url = ("chrome-extension://abc/src/popup.html?tabid=7&url=%s"
                    % page_url)

    def is_closed(self):
        return self.closed

    async def close(self):
        self.closed = True

    async def bring_to_front(self):
        return None


class FakeCrawlScraper:
    """Only the surface PageByPageCollector uses."""

    def __init__(self, page, fail_table_on=None, fail_export_on=None,
                 stability_fails_on=None, die_on=None, rows_override=None):
        self.page = page
        self.fail_table_on = fail_table_on
        self.fail_export_on = fail_export_on
        self.stability_fails_on = stability_fails_on
        self.die_on = die_on
        self.rows_override = rows_override or {}
        self.popups = []
        self.exports = []
        self.stability_calls = 0

    async def wait_for_stable_results(self, page, timeout=None, **kwargs):
        self.stability_calls += 1
        state = page.state()
        if self.stability_fails_on == page.index + 1:
            return False, state
        return True, state

    async def open_instant_data_scraper(self, page):
        number = page.index + 1
        if self.die_on == number:
            from student_scraper import BrowserLost
            raise BrowserLost("browser died opening page %d" % number)
        popup = FakePagePopup(page.pages[page.index], page.url,
                             fail_table=(self.fail_table_on == number),
                             fail_export=(self.fail_export_on == number))
        self.popups.append(popup)
        return popup

    async def wait_for_table(self, popup, timeout=None):
        return not popup.fail_table

    async def establish_results_table(self, popup, page, expected_ids):
        return getattr(self, "table_ok", True)

    async def stabilize_results_table(self, popup, page, expected_ids):
        return getattr(self, "table_ok", True)

    async def export_page_from_preview(self, popup, destination,
                                       expected_ids):
        return await self.export_xlsx(popup, destination)

    def dom_result_identities(self, state):
        from student_scraper import InstantDataScraper
        return InstantDataScraper.dom_result_identities(state)

    async def read_stats(self, popup):
        rows = self.rows_override.get(popup.page_url, len(popup.usernames))
        return 1, rows

    last_export_recovered = False

    async def export_xlsx(self, popup, destination, download_dir=None):
        if popup.fail_export:
            return -1
        write_page_export(destination, popup.usernames)
        self.exports.append(destination)
        return len(popup.usernames)


PAGE1 = ["alpha", "bravo", "charlie", "delta", "echo",
         "foxtrot", "golf", "hotel", "india", "juliet"]
PAGE2 = ["kilo", "lima", "mike", "november", "oscar",
         "papa", "quebec", "romeo", "sierra"]
PAGE3 = ["tango", "uniform", "victor"]


class CrawlCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pbp_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        # A SECOND copy -- a different checkout, so a different folder.
        self.other = sandbox.other_folder(
            os.path.join(self.root, "other_copy"))
        self.slug = "university-student-visakhapatnam"
        self.target = os.path.join(self.context.keyword_download_dir,
                                   "%s.xlsx" % self.slug)
        self.said = []
        import student_scraper as ratelimit
        self._orig_sleep = ratelimit.sleep_interruptible

        async def _noop(*_a, **_k):
            return None
        ratelimit.sleep_interruptible = _noop     # no real page-pacing waits

    def tearDown(self):
        sandbox.restore()
        import student_scraper as ratelimit
        ratelimit.sleep_interruptible = self._orig_sleep
        shutil.rmtree(self.root, ignore_errors=True)

    def collector(self, scraper):
        return student_scraper.PageByPageCollector(
            None, scraper, self.said.append)

    def run_crawl(self, pages, total=None, **kwargs):
        page = FakeCrawlPage(pages, total=total,
                             next_broken=kwargs.pop("next_broken", False),
                             stall_page=kwargs.pop("stall_page", None))
        scraper = FakeCrawlScraper(page, **kwargs)
        collector = self.collector(scraper)
        return page, scraper, collector

    @property
    def staging(self):
        return self.context.staging_for(self.slug)


# ============================================ 1-2. the happy multi-page path =
class TestMultiPageCollection(CrawlCase):
    def test_two_pages_of_ten_and_nine_merge_to_nineteen(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        rows, why = collector.merge_into(self.target)
        self.assertEqual(rows, 19)
        self.assertEqual(len(rows_of(self.target)), 19)
        self.assertIn("19", why)

    def test_three_pages_all_merge(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                 total=22)
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 22)
        self.assertEqual(len(collector.pages), 3)
        self.assertEqual([result.number for result in collector.pages],
                         [1, 2, 3])

    def test_pages_are_staged_in_order_with_padded_names(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                  total=22)
        asyncio.run(collector.collect(page, self.slug))
        names = sorted(os.listdir(self.staging))
        self.assertEqual(names, ["page_001.xlsx", "page_002.xlsx",
                                 "page_003.xlsx"])

    def test_the_merged_rows_are_in_page_order(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        merged = [row[1] for row in rows_of(self.target)]
        self.assertTrue(merged[0].endswith("/alpha"))
        self.assertTrue(merged[9].endswith("/juliet"))
        self.assertTrue(merged[10].endswith("/kilo"))

    def test_the_merged_export_is_parsable_by_the_reader(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        candidates = S.XlsxReader().read(self.target, "University Student "
                                                      "Visakhapatnam")
        self.assertEqual(len(candidates), 19)
        self.assertEqual(candidates[0].username, "alpha")
        self.assertEqual(candidates[0].xlsx_location, "Visakhapatnam")


# ================================== 3-5. one fresh popup per page, bound ====
class TestFreshPopupPerPage(CrawlCase):
    def test_one_popup_is_opened_per_page(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                 total=22)
        asyncio.run(collector.collect(page, self.slug))
        self.assertEqual(len(scraper.popups), 3,
                         "exactly one fresh popup per page")

    def test_every_popup_is_closed_before_the_next_page(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                 total=22)
        asyncio.run(collector.collect(page, self.slug))
        self.assertTrue(all(popup.closed for popup in scraper.popups),
                        "a popup must never survive a page change")

    def test_no_popup_is_reused(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        self.assertEqual(len({id(popup) for popup in scraper.popups}), 2)

    def test_each_popup_is_bound_to_its_own_page_url(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                 total=22)
        asyncio.run(collector.collect(page, self.slug))
        bound = [popup.page_url for popup in scraper.popups]
        self.assertEqual(bound[0], SEARCH)
        self.assertIn("p=2", bound[1])
        self.assertIn("p=3", bound[2])
        for popup in scraper.popups:
            self.assertIn(popup.page_url, popup.url,
                          "the popup must carry the CURRENT page url")

    def test_the_dom_is_waited_for_on_every_page(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                 total=22)
        asyncio.run(collector.collect(page, self.slug))
        self.assertGreaterEqual(scraper.stability_calls, 3,
                                "stability must be confirmed per page")


# ======================================================= 6, 12. failures ====
class TestFailuresPreserve(CrawlCase):
    def test_page_two_table_failure_creates_no_final_xlsx(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19,
                                                  fail_table_on=2)
        with self.assertRaises(student_scraper.PageCrawlError):
            asyncio.run(collector.collect(page, self.slug))
        self.assertFalse(os.path.exists(self.target))
        self.assertTrue(os.path.exists(
            student_scraper.page_path(self.staging, 1)),
            "page 1 staging must be preserved")

    def test_page_two_export_failure_creates_no_final_xlsx(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19,
                                                   fail_export_on=2)
        with self.assertRaises(student_scraper.PageCrawlError):
            asyncio.run(collector.collect(page, self.slug))
        self.assertFalse(os.path.exists(self.target))
        self.assertEqual(sorted(os.listdir(self.staging)),
                         ["page_001.xlsx"])

    def test_an_unstable_page_two_fails_rather_than_guessing(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19,
                                                   stability_fails_on=2)
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            asyncio.run(collector.collect(page, self.slug))
        self.assertIn("stabilise", str(caught.exception))
        self.assertFalse(os.path.exists(self.target))

    def test_a_next_that_cannot_be_clicked_fails(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19,
                                                   next_broken=True)
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            asyncio.run(collector.collect(page, self.slug))
        self.assertIn("could not be clicked", str(caught.exception))

    def test_browser_death_between_pages_preserves_staging(self):
        """With no restart available, a death preserves staging and stops.

        (The restart-and-continue path is covered in TestBrowserRestart,
        which supplies a browser manager the collector can relaunch.)
        """
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19,
                                                   die_on=2)
        from student_scraper import BrowserLost
        with self.assertRaises((student_scraper.PageCrawlError, BrowserLost)):
            asyncio.run(collector.collect(page, self.slug))
        self.assertTrue(os.path.exists(
            student_scraper.page_path(self.staging, 1)),
            "page 1 staging must survive a page-2 death")
        self.assertFalse(os.path.exists(self.target))

    def test_staging_is_never_removed_on_failure(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19,
                                                   fail_table_on=2)
        with self.assertRaises(student_scraper.PageCrawlError):
            asyncio.run(collector.collect(page, self.slug))
        self.assertTrue(os.path.isdir(self.staging))
        self.assertTrue(os.listdir(self.staging))


# ================================================= 7-9. count verification ==
class TestVerification(CrawlCase):
    def verify(self, **kwargs):
        defaults = {"unique_rows": 19, "expected_total": 19,
                    "mode": ValidationMode.EXACT_TOTAL, "pages_done": 2,
                    "expected_pages": 2, "last_page_had_next": False}
        defaults.update(kwargs)
        return student_scraper.verify_collection(**defaults)

    def test_exact_total_matched_is_accepted(self):
        ok, why = self.verify()
        self.assertTrue(ok, why)

    def test_exact_total_nineteen_collected_eighteen_is_refused(self):
        ok, why = self.verify(unique_rows=18)
        self.assertFalse(ok)
        self.assertIn("1 are missing", why)

    def test_exact_total_nineteen_collected_twenty_is_refused(self):
        ok, why = self.verify(unique_rows=20)
        self.assertFalse(ok)
        self.assertIn("changed while paginating", why)

    def test_page_completeness_accepts_when_next_ran_out(self):
        ok, why = self.verify(expected_total=None,
                             mode=ValidationMode.PAGE_COMPLETENESS,
                             expected_pages=None, unique_rows=250,
                             pages_done=25)
        self.assertTrue(ok, why)

    def test_page_completeness_refuses_while_next_remained(self):
        ok, why = self.verify(expected_total=None,
                             mode=ValidationMode.PAGE_COMPLETENESS,
                             expected_pages=None, last_page_had_next=True)
        self.assertFalse(ok)
        self.assertIn("still offered a next page", why)

    def test_a_short_crawl_is_refused_under_page_completeness(self):
        ok, why = self.verify(expected_total=None,
                             mode=ValidationMode.PAGE_COMPLETENESS,
                             expected_pages=10, pages_done=3)
        self.assertFalse(ok)
        self.assertIn("3 of about 10", why)

    def test_a_shortfall_deletes_the_merged_file(self):
        """A file that failed verification must not be left behind."""
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=25)
        asyncio.run(collector.collect(page, self.slug))
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            collector.merge_into(self.target)
        self.assertIn("missing", str(caught.exception))
        self.assertFalse(os.path.exists(self.target),
                         "an unverified merge must not survive")

    def test_the_validation_mode_is_recorded(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        self.assertEqual(collector.validation_mode,
                         ValidationMode.EXACT_TOTAL)
        self.assertEqual(collector.state_block("a1")["validation"],
                         ValidationMode.EXACT_TOTAL)

    def test_an_abbreviated_total_uses_page_completeness(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2],
                                                   total="9.1k")
        asyncio.run(collector.collect(page, self.slug))
        self.assertEqual(collector.validation_mode,
                         ValidationMode.PAGE_COMPLETENESS)
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 19)

    def test_no_stated_total_uses_page_completeness(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=None)
        asyncio.run(collector.collect(page, self.slug))
        self.assertEqual(collector.validation_mode,
                         ValidationMode.PAGE_COMPLETENESS)
        self.assertTrue(collector.merge_into(self.target)[0])


# ============================================ 9. cross-page duplicates ======
class TestDeduplication(CrawlCase):
    def test_a_user_on_both_pages_becomes_one_row(self):
        overlap = PAGE1[:2]
        page, _scraper, collector = self.run_crawl(
            [PAGE1, overlap + PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 19, "the two repeats collapse")
        seen = [row[1] for row in rows_of(self.target)]
        self.assertEqual(len(seen), len(set(seen)))

    def test_dedupe_is_case_insensitive(self):
        page, _scraper, collector = self.run_crawl(
            [["Alpha", "bravo"], ["ALPHA", "charlie"]], total=3)
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 3)

    def test_merge_reports_what_it_dropped(self):
        first = write_page_export(
            student_scraper.page_path(self.staging, 1), ["alpha", "bravo"])
        second = write_page_export(
            student_scraper.page_path(self.staging, 2), ["bravo", "charlie"])
        unique, duplicates, unusable = student_scraper.merge_staged_pages(
            [first, second], self.target)
        self.assertEqual((unique, duplicates, unusable), (3, 1, 0))

    def test_rows_with_no_github_identity_are_dropped_not_kept(self):
        book = Workbook()
        sheet = book.active
        sheet.append(["a", "b"])
        sheet.append(["not a url", "neither"])
        sheet.append(["https://github.com/real", "Real"])
        path = student_scraper.page_path(self.staging, 1)
        os.makedirs(self.staging, exist_ok=True)
        book.save(path)
        book.close()
        unique, _dupes, unusable = student_scraper.merge_staged_pages(
            [path], self.target)
        self.assertEqual(unique, 1)
        self.assertEqual(unusable, 1)

    def test_columns_are_aligned_by_name_across_pages(self):
        os.makedirs(self.staging, exist_ok=True)
        first = student_scraper.page_path(self.staging, 1)
        book = Workbook(); sheet = book.active
        sheet.append(["href", "name"])
        sheet.append(["https://github.com/alpha", "Alpha"])
        book.save(first); book.close()
        second = student_scraper.page_path(self.staging, 2)
        book = Workbook(); sheet = book.active
        sheet.append(["name", "href", "extra"])          # reordered + wider
        sheet.append(["Bravo", "https://github.com/bravo", "x"])
        book.save(second); book.close()

        student_scraper.merge_staged_pages([first, second], self.target)
        book = load_workbook(self.target)
        sheet = book.active
        header = [cell.value for cell in sheet[1]]
        rows = {row[header.index("href")]: row[header.index("name")]
                for row in rows_of(self.target)}
        book.close()
        self.assertEqual(rows["https://github.com/alpha"], "Alpha")
        self.assertEqual(rows["https://github.com/bravo"], "Bravo")


# ======================================== 13. retries never mix staging =====
class TestRetryStaging(CrawlCase):
    def test_a_retry_archives_the_previous_attempt(self):
        stale = write_page_export(student_scraper.page_path(self.staging, 7),
                                  ["stale1", "stale2"])
        self.assertTrue(os.path.exists(stale))
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        self.assertEqual(sorted(os.listdir(self.staging)),
                         ["page_001.xlsx", "page_002.xlsx"],
                         "page_007 from the old attempt must not be here")
        archives = [name for name in os.listdir(self.context.crawl_staging_dir)
                    if ".attempt_" in name]
        self.assertEqual(len(archives), 1)
        self.assertTrue(collector.attempt_archive)

    def test_stale_pages_never_reach_the_merge(self):
        write_page_export(student_scraper.page_path(self.staging, 3),
                          ["ghost1", "ghost2", "ghost3"])
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 19, "ghosts must not be merged in")
        names = [row[1] for row in rows_of(self.target)]
        self.assertFalse(any("ghost" in name for name in names))

    def test_two_retries_keep_separate_archives(self):
        for _attempt in range(2):
            write_page_export(student_scraper.page_path(self.staging, 1), ["x"])
            page, _scraper, collector = self.run_crawl([PAGE1], total=10)
            asyncio.run(collector.collect(page, self.slug))
        archives = [name for name in os.listdir(self.context.crawl_staging_dir)
                    if ".attempt_" in name]
        self.assertEqual(len(archives), 2)

    def test_an_empty_staging_directory_is_not_archived(self):
        os.makedirs(self.staging, exist_ok=True)
        page, _scraper, collector = self.run_crawl([PAGE1], total=10)
        asyncio.run(collector.collect(page, self.slug))
        self.assertEqual(collector.attempt_archive, "")

    def test_the_merge_refuses_a_page_count_mismatch(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        write_page_export(student_scraper.page_path(self.staging, 9), ["extra"])
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            collector.merge_into(self.target)
        self.assertIn("staged file", str(caught.exception))


# ====================================== 14-15. the final file, exactly once =
class TestFinalFile(CrawlCase):
    def test_the_keyword_xlsx_is_written_exactly_once(self):
        page, scraper, collector = self.run_crawl([PAGE1, PAGE2, PAGE3],
                                                  total=22)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        self.assertEqual(
            [name for name in os.listdir(self.context.keyword_download_dir)],
            ["%s.xlsx" % self.slug])
        # the page exports went to staging, never to keyword_downloads
        for path in scraper.exports:
            self.assertIn("crawl_staging", path)

    def test_the_row_count_matches_the_merged_unique_rows(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        from student_scraper import validate_export
        self.assertEqual(validate_export(self.target), rows)
        self.assertEqual(len(rows_of(self.target)), rows)

    def test_no_page_file_leaks_into_the_output_directory(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        for directory in (C.FINAL_OUTPUT_DIR,
                          self.context.keyword_download_dir):
            if not os.path.isdir(directory):
                continue
            leaked = [name for name in os.listdir(directory)
                      if name.startswith("page_")]
            self.assertEqual(leaked, [], directory)

    def test_no_temp_file_survives_the_merge(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        leftovers = [name
                     for name in os.listdir(self.context.keyword_download_dir)
                     if name.endswith(".tmp") or name.startswith(".")]
        self.assertEqual(leftovers, [])

    def test_cleanup_removes_staging_only_after_success(self):
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        self.assertTrue(os.path.isdir(self.staging))
        collector.cleanup()
        self.assertFalse(os.path.isdir(self.staging))
        self.assertTrue(os.path.exists(self.target))


# ====================================== 18-19. rank isolation of staging ====
class TestRankIsolation(CrawlCase):
    def test_staging_lives_inside_the_rank(self):
        self.assertTrue(self.context.owns(self.context.crawl_staging_dir))
        self.assertTrue(self.context.owns(self.staging))

    def test_two_ranks_never_share_staging(self):
        self.assertNotEqual(self.context.crawl_staging_dir,
                            self.other.crawl_staging_dir)
        self.assertFalse(self.context.owns(self.other.crawl_staging_dir))
        self.assertFalse(self.other.owns(self.staging))
        self.assertNotIn(self.context.crawl_staging_dir,
                         set(self.other.paths().values()))
        self.assertFalse(self.other.owns(self.context.crawl_staging_dir))

    def test_another_ranks_staging_is_untouched_by_a_crawl(self):
        theirs = os.path.join(self.other.staging_for(self.slug),
                              "page_001.xlsx")
        write_page_export(theirs, ["theirperson"])
        page, _scraper, collector = self.run_crawl([PAGE1, PAGE2], total=19)
        asyncio.run(collector.collect(page, self.slug))
        collector.merge_into(self.target)
        collector.cleanup()
        self.assertTrue(os.path.exists(theirs),
                        "rank 1 must not touch rank 2's staging")
        names = [row[1] for row in rows_of(self.target)]
        self.assertFalse(any("theirperson" in name for name in names))

    def test_different_searches_stage_separately(self):
        other_slug = "college-student-visakhapatnam"
        write_page_export(
            os.path.join(self.context.staging_for(other_slug),
                         "page_001.xlsx"), ["othersearch"])
        page, _scraper, collector = self.run_crawl([PAGE1], total=10)
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 10)
        self.assertTrue(os.path.exists(
            os.path.join(self.context.staging_for(other_slug),
                         "page_001.xlsx")))


# =============================== 10-11, 16-17, 20. the rest of the pipeline =
class TestPipelineUnchanged(unittest.TestCase):
    """Single page, zero results, --limit, duplicates, and the campaign."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pbp_pipe_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan(
            ["College Student", "University Student"],
            [type("City", (), {"country": "India",
                               "state": "Andhra Pradesh",
                               "canonical": "Visakhapatnam",
                               "search_aliases": ("Visakhapatnam",)})()])
        self.context = sandbox.enter(self.runtime)
        self.assigned = self.plan.assign(1, 1)
        self.stores = []

    def tearDown(self):
        sandbox.restore()
        for store in self.stores:
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def workflow(self, unit, state, verdict, rows=10, outcome=None):
        from test_single_page import (FakeBrowser, FakePage, FakePopup,
                                      Options, StubScraper)
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
        workflow.unit = unit
        workflow.state = StateManager(path=self.context.state_file)
        workflow.browser = FakeBrowser(page)
        workflow.scraper = StubScraper(verdict, FakePopup(), rows=rows,
                                       outcome=outcome)
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

    def test_a_single_page_search_is_unchanged(self):
        unit = self.assigned[0]
        target = os.path.join(self.context.keyword_download_dir,
                              "%s.xlsx" % unit.slug)
        workflow, page = self.workflow(unit, page_state(),
                                       S.NextControl.SINGLE_PAGE, rows=7)
        status, rows = asyncio.run(
            workflow.collect(page, unit.search_query, target))
        self.assertEqual(status, S.CollectResult.RESULTS)
        self.assertEqual(rows, 7)
        self.assertEqual(workflow.scraper.single_pages, 1)
        self.assertFalse(os.path.isdir(self.context.crawl_staging_dir),
                         "a single page needs no staging")

    def test_a_zero_result_search_is_unchanged(self):
        from test_zero_results import zero_state
        unit = self.assigned[0]
        target = os.path.join(self.context.keyword_download_dir,
                              "%s.xlsx" % unit.slug)
        workflow, page = self.workflow(unit, zero_state(),
                                       S.NextControl.SINGLE_PAGE)
        status, rows = asyncio.run(
            workflow.collect(page, unit.search_query, target))
        self.assertEqual(status, S.CollectResult.ZERO_RESULTS)
        self.assertEqual(rows, 0)
        self.assertFalse(os.path.isdir(self.context.crawl_staging_dir))

    def test_the_crawl_state_block_is_recorded_and_reset(self):
        unit = self.assigned[0]
        state = StateManager(path=self.context.state_file)
        state.begin_unit(0, unit)
        state.set_in("crawl", mode="page-by-page", completed_pages=2,
                     expected_total=19)
        self.assertEqual(state.data["crawl"]["completed_pages"], 2)
        state.begin_unit(1, self.assigned[1])
        self.assertEqual(state.data["crawl"], {},
                         "the crawl block belongs to ONE search")

    def test_the_crawl_block_is_cleared_by_a_retry_reset(self):
        unit = self.assigned[0]
        state = StateManager(path=self.context.state_file)
        state.begin_unit(0, unit)
        state.set_in("crawl", mode="page-by-page", completed_pages=1)
        state.reset_for_recollection()
        self.assertEqual(state.data["crawl"], {})

    def test_the_search_identity_survives_a_page_by_page_crawl(self):
        """Staging a crawl must not disturb which search is in hand."""
        unit = self.assigned[0]
        state = StateManager(path=self.context.state_file)
        state.begin_unit(0, unit)
        state.set_in("crawl", mode="page-by-page")
        described = state.describe_current()
        self.assertEqual(described["search_query"], unit.search_query)
        self.assertEqual(described["canonical_city"], unit.canonical_city)
        self.assertEqual(described["slug"], unit.slug)
        reloaded = StateManager(path=self.context.state_file)
        self.assertEqual(reloaded.search_query, unit.search_query)
        self.assertEqual(reloaded.data["crawl"]["mode"], "page-by-page")

    def test_limit_counts_students_not_pages(self):
        """--limit is unchanged: it counts saved students."""
        import inspect
        # run() is a thin wrapper that flushes the workbook on exit; the loop
        # itself lives in _run().  Both are inspected so a future rename
        # cannot hide the check.
        source = (inspect.getsource(S.Workflow.run)
                  + inspect.getsource(S.Workflow._run))
        self.assertIn("self.processed >= self.options.limit", source)
        crawl_source = inspect.getsource(student_scraper.PageByPageCollector)
        # The crawler must not consult the --limit OPTION.  (It legitimately
        # mentions rate-limit handling, so check for the option specifically.)
        self.assertNotIn("options.limit", crawl_source,
                         "page collection must not consult --limit")
        self.assertNotIn("self.options", crawl_source,
                         "the crawler has no options; --limit is the run loop's")

    def test_already_saved_students_are_still_skipped(self):
        unit = self.assigned[0]
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        record = {column: "" for column in C.FINAL_COLUMNS}
        record["GitHub Profile"] = "https://github.com/alpha"
        record["github"] = record["GitHub Profile"]
        store.save_student(record, search_unit=unit)
        target = os.path.join(self.context.keyword_download_dir,
                              "%s.xlsx" % unit.slug)
        write_page_export(target, PAGE1)
        candidates = S.XlsxReader().read(target, unit.search_query)
        pending = [candidate for candidate in candidates
                   if not store.is_known(username=candidate.username,
                                         profile_url=candidate.profile_url)]
        self.assertEqual(len(candidates), 10)
        self.assertEqual(len(pending), 9)


# ======================================== the total/mode parser in isolation =
class TestTotalParsing(unittest.TestCase):
    def test_an_exact_small_total_is_exact(self):
        self.assertEqual(expected_total_users({"reportedTotalUsers": "19"}),
                         (19, ValidationMode.EXACT_TOTAL))

    def test_an_abbreviated_total_is_not_invented(self):
        total, mode = expected_total_users({"reportedTotalUsers": "9.1k"})
        self.assertIsNone(total)
        self.assertEqual(mode, ValidationMode.PAGE_COMPLETENESS)

    def test_a_total_above_the_result_cap_uses_page_completeness(self):
        total, mode = expected_total_users({"reportedTotalUsers": "1001"})
        self.assertEqual(total, 1001)
        self.assertEqual(mode, ValidationMode.PAGE_COMPLETENESS)

    def test_a_missing_total_uses_page_completeness(self):
        for raw in (None, "", "many"):
            _total, mode = expected_total_users({"reportedTotalUsers": raw})
            self.assertEqual(mode, ValidationMode.PAGE_COMPLETENESS, raw)

    def test_page_counts(self):
        self.assertEqual(expected_page_count(19), 2)
        self.assertEqual(expected_page_count(10), 1)
        self.assertEqual(expected_page_count(11), 2)
        self.assertIsNone(expected_page_count(None))


class TestASlowPageIsRetriedNotAbandoned(CrawlCase):
    """A live browser that served a slow page must not lose the search.

    OBSERVED 2026-10-08: a Delhi crawl exported pages 1-24 at a steady pace,
    page 25 missed the 45s stability window once, and the whole search was
    abandoned -- no rate limit, no block, and page 25 served normally when
    checked afterwards.  _stable_page() restarted only a DEAD browser, so a
    live one that was merely slow raised straight away.
    """

    class _Scraper:
        """Fails to stabilise `fail_times` times, then succeeds."""

        def __init__(self, page, fail_times):
            self.page = page
            self.fail_times = fail_times
            self.calls = 0

        async def wait_for_stable_results(self, _page):
            self.calls += 1
            if self.calls <= self.fail_times:
                return False, {}
            return True, {"url": "https://github.com/search?p=25",
                          "resultsListChildren": 10}

    def _stable(self, fail_times):
        page = FakeCrawlPage([["alfa"]])
        scraper = self._Scraper(page, fail_times)
        collector = self.collector(scraper)
        return page, scraper, collector

    def test_a_page_that_settles_on_the_second_try_is_kept(self):
        page, scraper, collector = self._stable(fail_times=1)
        _page, state = asyncio.run(collector._stable_page(page, 25))
        self.assertEqual(int(state["resultsListChildren"]), 10)
        self.assertEqual(scraper.calls, 2, "the page must be re-waited")
        self.assertEqual(getattr(page, "reloads", 0), 1,
                         "the page must be reloaded between attempts")

    def test_a_page_that_never_settles_still_fails(self):
        page, scraper, collector = self._stable(fail_times=99)
        with self.assertRaises(student_scraper.PageCrawlError):
            asyncio.run(collector._stable_page(page, 25))
        self.assertEqual(scraper.calls, C.PAGE_STABLE_ATTEMPTS,
                         "a dead end must not retry forever")

    def test_a_page_that_settles_first_time_is_not_reloaded(self):
        page, scraper, collector = self._stable(fail_times=0)
        asyncio.run(collector._stable_page(page, 25))
        self.assertEqual(scraper.calls, 1)
        self.assertEqual(getattr(page, "reloads", 0), 0,
                         "the healthy path must cost nothing extra")


if __name__ == "__main__":
    unittest.main(verbosity=2)
