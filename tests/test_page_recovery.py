"""Page-level export recovery, and surviving a browser death between pages.

THE BUG THIS FILE EXISTS FOR
----------------------------
    [PAGE 1] stable: 10 result(s)
    [INCOMPLETE] the browser died during page-by-page collection
    [STOP] Download.save_as: Target page, context or browser has been closed

Chromium writes downloads to its configured directory -- keyword_downloads/ --
but a page export's destination is crawl_staging/<slug>/page_001.xlsx.  The
old recovery scanned dirname(destination), i.e. the staging directory, which
is empty, so it never found the file Chromium had actually written, and a
survivable browser hiccup became a fatal BrowserLost.

Recovery now scans the real download directory, snapshotting it before the
click so it can only ever adopt THIS page's export -- never a previous page's,
another rank's, or a half-written one -- and the crawl restarts the browser
between pages rather than requiring one process to survive the whole search.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook                            # noqa: E402

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (BrowserLost, InstantDataScraper)   # noqa: E402
import sandbox                                         # noqa: E402


def write_usernames(path, usernames):
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar src", "prc-Link href", "search-match"])
    for name in usernames:
        sheet.append(["https://github.com/%s.png" % name,
                      "https://github.com/%s" % name, name.title()])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    book.save(path)
    book.close()
    return path


def make_xlsx(path, count, header=("prc-Avatar src", "prc-Link href",
                                   "search-match")):
    book = Workbook()
    sheet = book.active
    sheet.append(list(header))
    for number in range(1, count + 1):
        sheet.append(["https://github.com/person%d.png" % number,
                      "https://github.com/person%d" % number,
                      "Person %d" % number][:len(header)])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    book.save(path)
    book.close()
    return path


# ============================================ 1-6, 10-11. export_xlsx recovery
class FakeDownload:
    """A Playwright download whose save_as / path can be made to fail."""

    def __init__(self, save_ok=True, path_value=None, closed_error=False):
        self.save_ok = save_ok
        self.path_value = path_value
        self.closed_error = closed_error

    async def save_as(self, destination):
        if self.save_ok:
            shutil.copyfile(self._src, destination)
            return
        if self.closed_error:
            raise RuntimeError("Download.save_as: Target page, context or "
                               "browser has been closed")
        raise RuntimeError("save_as failed for some other reason")

    async def path(self):
        if self.path_value is None:
            raise RuntimeError("Download.path: browser has been closed")
        return self.path_value


class FakeExportPopup:
    """The extension popup, just enough for export_xlsx.

    on_click runs when the XLSX button is pressed -- it models what Chromium
    does with the download (write it to the download dir, or not).
    """

    def __init__(self, download, on_click=None):
        self.download = download
        self.on_click = on_click
        self.closed = False

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    async def wait_for_selector(self, selector, timeout=None):
        return True

    def expect_download(self, timeout=None):
        popup = self

        class Ctx:
            async def __aenter__(self_):
                return self_

            async def __aexit__(self_, *exc):
                return False

            @property
            def value(self_):
                async def _get():
                    return popup.download
                return _get()
        return Ctx()

    async def click(self, selector, timeout=None):
        if self.on_click is not None:
            self.on_click()


class ExportRecoveryCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pxr_")
        self.download_dir = os.path.join(self.root, "keyword_downloads")
        self.staging = os.path.join(self.root, "crawl_staging", "slug")
        os.makedirs(self.download_dir)
        os.makedirs(self.staging)
        self.destination = os.path.join(self.staging, "page_001.xlsx")
        self.said = []
        self.scraper = InstantDataScraper(browser=None, state=None,
                                          say=self.said.append)

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def export(self, popup):
        return asyncio.run(self.scraper.export_xlsx(
            popup, self.destination, download_dir=self.download_dir))

    def test_a_normal_save_succeeds(self):
        src = make_xlsx(os.path.join(self.root, "src.xlsx"), 10)
        download = FakeDownload(save_ok=True)
        download._src = src
        popup = FakeExportPopup(download)
        rows = self.export(popup)
        self.assertEqual(rows, 10)
        self.assertTrue(os.path.exists(self.destination))
        self.assertFalse(self.scraper.last_export_recovered)

    def test_save_as_closed_but_valid_file_on_disk_recovers(self):
        """The exact live failure: save_as closed, file already downloaded."""
        def on_click():
            make_xlsx(os.path.join(self.download_dir, "export.xlsx"), 10)
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        rows = self.export(popup)
        self.assertEqual(rows, 10)
        self.assertTrue(os.path.exists(self.destination),
                        "the recovered file must land at the destination")
        self.assertTrue(self.scraper.last_export_recovered)
        self.assertTrue(any("recovered the export from disk" in line
                            for line in self.said), self.said)

    def test_a_genuinely_partial_crdownload_is_incomplete(self):
        """A .crdownload that will NOT open is a real partial -> incomplete."""
        def on_click():
            with open(os.path.join(self.download_dir,
                                   "export.xlsx.crdownload"), "wb") as handle:
                handle.write(b"partial bytes, not a workbook")
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        saved = C.DOWNLOAD_TIMEOUT
        C.DOWNLOAD_TIMEOUT = 600
        try:
            with self.assertRaises(BrowserLost):
                self.export(popup)
        finally:
            C.DOWNLOAD_TIMEOUT = saved
        self.assertFalse(os.path.exists(self.destination))

    def test_a_complete_file_left_as_crdownload_is_recovered(self):
        """THE live scenario: Playwright leaves a finished export unrenamed.

        In this environment a completed download stays as
        <guid>.crdownload in the downloads directory (verified against three
        real 20 KB, 10-row files).  Completeness is decided by opening the
        file, not by its extension, so it must be recovered -- exactly as the
        old single-page flow did.
        """
        def on_click():
            # A COMPLETE, valid workbook that merely has the .crdownload name.
            src_path = make_xlsx(os.path.join(self.root, "complete.xlsx"), 10)
            import shutil as _sh
            _sh.copyfile(src_path, os.path.join(
                self.download_dir,
                "3f2a-guid-9c1d.crdownload"))
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        rows = self.export(popup)
        self.assertEqual(rows, 10, "a complete .crdownload must be recovered")
        self.assertTrue(os.path.exists(self.destination))
        self.assertTrue(self.destination.endswith(".xlsx"))
        self.assertTrue(self.scraper.last_export_recovered)
        # the destination is a real, openable workbook
        from student_scraper import validate_export
        self.assertEqual(validate_export(self.destination), 10)

    def test_a_stale_crdownload_from_a_prior_attempt_is_ignored(self):
        """The 3 leftover .crdownload files a failed run leaves behind.

        They are complete and would validate, but they were present BEFORE
        this click, so the snapshot excludes them; only this attempt's file
        is recovered.
        """
        # three complete leftovers, as the failing live runs produced
        for guid in ("55969217", "7ef72214", "e5322373"):
            make_xlsx(os.path.join(self.download_dir, "%s.crdownload" % guid),
                      10)

        def on_click():
            src_path = make_xlsx(os.path.join(self.root, "fresh.xlsx"), 9)
            import shutil as _sh
            _sh.copyfile(src_path, os.path.join(self.download_dir,
                                                "new-guid.crdownload"))
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        rows = self.export(popup)
        self.assertEqual(rows, 9, "only THIS attempt's file is recovered")
        # the three leftovers are untouched
        for guid in ("55969217", "7ef72214", "e5322373"):
            self.assertTrue(os.path.exists(
                os.path.join(self.download_dir, "%s.crdownload" % guid)))

    def test_no_file_at_all_is_a_browser_lost(self):
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download)          # click writes nothing
        saved = C.DOWNLOAD_TIMEOUT
        C.DOWNLOAD_TIMEOUT = 600
        try:
            with self.assertRaises(BrowserLost):
                self.export(popup)
        finally:
            C.DOWNLOAD_TIMEOUT = saved
        self.assertFalse(os.path.exists(self.destination))

    def test_two_new_exports_refuse_as_ambiguous(self):
        def on_click():
            make_xlsx(os.path.join(self.download_dir, "one.xlsx"), 10)
            make_xlsx(os.path.join(self.download_dir, "two.xlsx"), 9)
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        saved = C.DOWNLOAD_TIMEOUT
        C.DOWNLOAD_TIMEOUT = 600
        try:
            with self.assertRaises(BrowserLost):
                self.export(popup)
        finally:
            C.DOWNLOAD_TIMEOUT = saved
        self.assertTrue(any("cannot tell which" in line
                            for line in self.said), self.said)
        self.assertFalse(os.path.exists(self.destination))

    def test_a_malformed_recovered_file_is_refused(self):
        def on_click():
            with open(os.path.join(self.download_dir, "export.xlsx"),
                      "wb") as handle:
                handle.write(b"not a real workbook")
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        saved = C.DOWNLOAD_TIMEOUT
        C.DOWNLOAD_TIMEOUT = 600
        try:
            with self.assertRaises(BrowserLost):
                self.export(popup)
        finally:
            C.DOWNLOAD_TIMEOUT = saved
        self.assertFalse(os.path.exists(self.destination))

    def test_another_ranks_pre_existing_file_is_never_adopted(self):
        """A file present BEFORE the click is not this export."""
        make_xlsx(os.path.join(self.download_dir, "stale.xlsx"), 4)
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download)          # click writes nothing new
        saved = C.DOWNLOAD_TIMEOUT
        C.DOWNLOAD_TIMEOUT = 600
        try:
            with self.assertRaises(BrowserLost):
                self.export(popup)
        finally:
            C.DOWNLOAD_TIMEOUT = saved
        self.assertFalse(os.path.exists(self.destination),
                         "a pre-existing file must not be recovered")

    def test_a_stale_old_export_before_the_click_is_ignored(self):
        old = make_xlsx(os.path.join(self.download_dir, "old.xlsx"), 3)
        os.utime(old, (time.time() - 3600, time.time() - 3600))

        def on_click():
            make_xlsx(os.path.join(self.download_dir, "fresh.xlsx"), 10)
        download = FakeDownload(save_ok=False, closed_error=True)
        popup = FakeExportPopup(download, on_click=on_click)
        rows = self.export(popup)
        self.assertEqual(rows, 10, "only the fresh export is recovered")
        # the stale one is left untouched
        self.assertTrue(os.path.exists(old))

    def test_download_path_fallback_is_used_when_available(self):
        """If save_as fails but download.path() works, copy from there."""
        src = make_xlsx(os.path.join(self.root, "viapath.xlsx"), 8)
        download = FakeDownload(save_ok=False, closed_error=False,
                               path_value=src)
        popup = FakeExportPopup(download)
        rows = self.export(popup)
        self.assertEqual(rows, 8)
        self.assertTrue(os.path.exists(self.destination))


# ================================ 12. single-page recovery still works =======
class TestSinglePageRecoveryUnchanged(ExportRecoveryCase):
    def setUp(self):
        super().setUp()
        # Single-page callers export straight INTO the download directory.
        self.destination = os.path.join(self.download_dir,
                                        "college-student.xlsx")

    def export(self, popup):
        # No download_dir passed: it must default to the destination's own
        # directory, exactly as single-page collect() relies on.
        return asyncio.run(self.scraper.export_xlsx(popup, self.destination))

    def test_single_page_save_success(self):
        src = make_xlsx(os.path.join(self.root, "src.xlsx"), 10)
        download = FakeDownload(save_ok=True)
        download._src = src
        rows = self.export(FakeExportPopup(download))
        self.assertEqual(rows, 10)

    def test_single_page_disk_recovery(self):
        def on_click():
            make_xlsx(os.path.join(self.download_dir, "export.xlsx"), 10)
        download = FakeDownload(save_ok=False, closed_error=True)
        rows = self.export(FakeExportPopup(download, on_click=on_click))
        self.assertEqual(rows, 10)
        self.assertTrue(os.path.exists(self.destination))
        self.assertTrue(self.scraper.last_export_recovered)


# ============================ 7-9. browser restart between pages ============
PAGE1 = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]
PAGE2 = ["k", "l", "m", "n", "o", "p", "q", "r", "s"]
SEARCH = "https://github.com/search?q=University+Student+Visakhapatnam&type=users"


class RestartPage:
    """A GitHub tab that can 'die' and be replaced by restart()."""

    def __init__(self, world, generation):
        self.world = world
        self.generation = generation

    @property
    def dead(self):
        return self.generation < self.world.generation

    @property
    def url(self):
        n = self.world.index + 1
        return SEARCH if n == 1 else "%s&p=%d" % (SEARCH, n)

    def is_closed(self):
        return self.dead

    async def bring_to_front(self):
        return None

    async def reload(self, **_kwargs):
        """Mirrors a Playwright page: _stable_page() reloads a page that did
        not settle before abandoning the search."""
        self.reloads = getattr(self, "reloads", 0) + 1

    async def goto(self, url, **kwargs):
        if self.dead:
            raise RuntimeError("browser has been closed")
        import urllib.parse
        q = urllib.parse.parse_qs(urllib.parse.urlsplit(url).query)
        self.world.index = int((q.get("p") or ["1"])[0]) - 1

    async def wait_for_function(self, *a, **k):
        return True

    def locator(self, selector):
        from test_single_page import FakeLocator
        world = self.world
        if selector == 'a[rel="next"]' and world.has_next and not self.dead:
            def _click():
                if self.dead:
                    raise RuntimeError("browser has been closed")
                world.index += 1
            return FakeLocator(count=1, visible=True, on_click=_click)
        if self.dead:
            # a dead page raises on any access, as Playwright would
            class Dead:
                async def count(self_inner):
                    raise RuntimeError("browser has been closed")
            return Dead()
        return FakeLocator(count=0)


class RestartWorld:
    """Models a browser that dies at a chosen point and can be restarted."""

    def __init__(self, pages, die_before_export_page=None,
                 die_after_export_page=None, total=19):
        self.pages = pages
        self.index = 0
        self.generation = 0
        self.total = total
        self.die_before = die_before_export_page
        self.die_after = die_after_export_page
        self.restarts = 0
        self.download_dir = None          # set by the test
        self.logins = 0

    @property
    def has_next(self):
        return self.index < len(self.pages) - 1

    def current_page(self):
        return RestartPage(self, self.generation)

    # -- BrowserManager surface -------------------------------------------
    async def is_alive(self):
        return True                        # deaths surface as raised errors

    async def restart(self, with_extension=True):
        self.restarts += 1
        self.generation += 1              # old page handles become dead
        return self.current_page()

    async def github_login_name(self, page):
        self.logins += 1
        return "tester"


class RestartScraper:
    """Scraper whose export writes to the download dir, and can 'die'."""

    def __init__(self, world):
        self.world = world
        self.last_export_recovered = False
        self.exports = []
        self.popups = 0

    async def wait_for_stable_results(self, page, timeout=None, **kwargs):
        if page.dead:
            raise BrowserLost("browser has been closed")
        users = self.world.pages[self.world.index]
        return True, page_state_for(page.url, users, self.world)

    async def open_instant_data_scraper(self, page):
        self.popups += 1
        if page.dead:
            raise BrowserLost("browser has been closed")
        return _RestartPopup(self.world.pages[self.world.index], page.url)

    async def wait_for_table(self, popup, timeout=None):
        return True

    async def establish_results_table(self, popup, page, expected_ids):
        return True

    async def stabilize_results_table(self, popup, page, expected_ids):
        return True

    async def export_page_from_preview(self, popup, destination,
                                       expected_ids):
        return await self.export_xlsx(popup, destination)

    def dom_result_identities(self, state):
        from student_scraper import InstantDataScraper
        return InstantDataScraper.dom_result_identities(state)

    async def read_stats(self, popup):
        return 1, len(popup.usernames)

    async def export_xlsx(self, popup, destination, download_dir=None):
        number = self.world.index + 1
        self.last_export_recovered = False
        # Optionally the browser dies mid-export but the file lands on disk.
        if self.world.die_after == number:
            self.world.generation += 1     # the tab is now dead
            self.last_export_recovered = True
        write_usernames(destination, popup.usernames)
        self.exports.append(destination)
        return len(popup.usernames)


class _RestartPopup:
    def __init__(self, usernames, url):
        self.usernames = usernames
        self.url = "chrome-extension://x/popup.html?url=%s" % url
        self.closed = False

    def is_closed(self):
        return self.closed

    async def close(self):
        self.closed = True


def page_state_for(url, users, world):
    from test_single_page import page_state
    return page_state(
        url=url,
        hasNext=world.has_next,
        nextByShape=1 if world.has_next else 0,
        nextDisabled=None if world.has_next else True,
        resultsListChildren=len(users),
        profileLinkCount=len(users),
        resultProfileLinkCount=len(users),
        reportedTotalUsers=str(world.total),
    )


class RestartCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="pxrs_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.context = sandbox.enter(self.runtime)
        self.slug = "university-student-visakhapatnam"
        self.target = os.path.join(self.context.keyword_download_dir,
                                   "%s.xlsx" % self.slug)
        self.said = []

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def collector(self, world):
        world.download_dir = self.context.keyword_download_dir
        scraper = RestartScraper(world)
        return student_scraper.PageByPageCollector(
            world, scraper, self.said.append), scraper

    def test_page1_recovered_then_restart_then_page2_continues(self):
        """Requirement 7: page1 recovered after death -> restart -> page2."""
        world = RestartWorld([PAGE1, PAGE2], die_after_export_page=1)
        collector, scraper = self.collector(world)
        page = world.current_page()
        asyncio.run(collector.collect(page, self.slug))
        rows, why = collector.merge_into(self.target)
        self.assertEqual(rows, 19)
        self.assertEqual(world.restarts, 1, "exactly one restart was needed")
        self.assertEqual(len(collector.pages), 2)
        self.assertEqual(len(rows_of(self.target)), 19)

    def test_page1_staging_is_not_re_exported_after_restart(self):
        """Requirement 8: a validated page is kept, not collected again."""
        world = RestartWorld([PAGE1, PAGE2], die_after_export_page=1)
        collector, scraper = self.collector(world)
        page = world.current_page()
        asyncio.run(collector.collect(page, self.slug))
        page1_exports = [p for p in scraper.exports if p.endswith(
            "page_001.xlsx")]
        self.assertEqual(len(page1_exports), 1,
                         "page 1 must be exported exactly once")

    def test_browser_dies_after_page2_export_then_merge_is_19(self):
        """Requirement 9: death after the LAST page's export still merges."""
        world = RestartWorld([PAGE1, PAGE2], die_after_export_page=2)
        collector, scraper = self.collector(world)
        page = world.current_page()
        asyncio.run(collector.collect(page, self.slug))
        rows, _why = collector.merge_into(self.target)
        self.assertEqual(rows, 19)
        # page 2 was the last, so no restart is needed after it.
        self.assertEqual(len(collector.pages), 2)

    def test_the_restart_confirms_login_and_navigates_to_the_next_page(self):
        world = RestartWorld([PAGE1, PAGE2], die_after_export_page=1)
        collector, scraper = self.collector(world)
        page = world.current_page()
        asyncio.run(collector.collect(page, self.slug))
        self.assertGreaterEqual(world.logins, 1,
                                "login must be reconfirmed after a restart")
        self.assertTrue(any("browser restarted" in line
                            for line in self.said), self.said)

    def test_the_restart_budget_is_respected(self):
        world = RestartWorld([PAGE1, PAGE2], die_after_export_page=1)
        world.restarts = C.BROWSER_MAX_RESTARTS      # already exhausted
        collector, scraper = self.collector(world)
        page = world.current_page()
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            asyncio.run(collector.collect(page, self.slug))
        self.assertIn("restart budget", str(caught.exception))
        # page 1 staging is still there for a later retry
        self.assertTrue(os.path.exists(
            student_scraper.page_path(collector.staging, 1)))


def rows_of(path):
    from openpyxl import load_workbook
    book = load_workbook(path, read_only=True)
    sheet = book.active
    out = [r for i, r in enumerate(sheet.iter_rows(values_only=True))
           if i and any(v not in (None, "") for v in r)]
    book.close()
    return out


# ================================= with_page_number helper ==================
class TestWithPageNumber(unittest.TestCase):
    def test_page_one_drops_the_parameter(self):
        self.assertEqual(student_scraper.with_page_number(SEARCH + "&p=3", 1),
                         SEARCH)

    def test_a_later_page_sets_it(self):
        self.assertIn("p=2", student_scraper.with_page_number(SEARCH, 2))
        self.assertIn("p=5", student_scraper.with_page_number(SEARCH + "&p=2", 5))

    def test_the_query_is_preserved(self):
        url = student_scraper.with_page_number(SEARCH, 4)
        self.assertIn("type=users", url)
        self.assertIn("q=University", url)


if __name__ == "__main__":
    unittest.main(verbosity=2)
