"""Regression: the diagnostic must REPORT even when export never happens.

THE CRASH THIS FILE EXISTS FOR
------------------------------
student_scraper.py reached the report and died with

    NameError: name 'exported' is not defined

at  say("IDS exported rows : %d (raw)" % max(exported, 0))

because the reporting variables were only ever bound inside the export path.
When the table did not stabilise (or export raised), that path was skipped
and the report referenced an unbound name.  The fix initialises every
reporting variable BEFORE the export phase, so the diagnostic always prints a
report -- FAILED / N/A when the export could not run -- instead of crashing.

This drives main() end-to-end with fakes (no browser, no network, no live
scrape) through to the report and asserts it completes without NameError.
"""

import asyncio
import io
import os
import sys
import tempfile
import types
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                     # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (InstantDataScraper)   # noqa: E402
import student_scraper as diag                            # noqa: E402

ENG_P2 = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]


class FakePage:
    url = "https://github.com/search?q=x&type=users&p=2"

    async def evaluate(self, script, arg=None):
        return {"url": self.url, "resultsListPresent": True,
                "users": [{"username": u, "href": "/" + u, "hasImg": True,
                           "anchorText": u, "anchor": {"tag": "a"},
                           "card": {"tag": "div", "testid": "results"}}
                          for u in ENG_P2]}


class FakeBrowser:
    download_dir = None

    def __init__(self, headless=False, profile_dir=None, download_dir=None):
        self._page = FakePage()

    async def start(self, with_extension=False):
        return None

    async def page(self):
        return self._page

    async def close_extra_pages(self, keep=()):
        return None

    async def close(self):
        return None


def point_config_at(root):
    """Redirect the diagnostic's paths into a temp dir.  Returns a restorer."""
    saved = {name: getattr(C, name) for name in
             ("PROFILE_DIR", "KEYWORD_DOWNLOAD_DIR", "TEMP_DOWNLOAD_DIR")}
    C.PROFILE_DIR = os.path.join(root, "pw_profile")
    C.KEYWORD_DOWNLOAD_DIR = os.path.join(root, "keyword_downloads")
    C.TEMP_DOWNLOAD_DIR = os.path.join(root, "temp_downloads")
    for path in (C.KEYWORD_DOWNLOAD_DIR, C.TEMP_DOWNLOAD_DIR):
        os.makedirs(path, exist_ok=True)

    def restore():
        for name, value in saved.items():
            setattr(C, name, value)
    return restore


class BaseFakeScraper:
    """A scraper whose table always establishes; subclasses vary export."""

    def __init__(self, *a, **k):
        self.last_export_recovered = False

    async def open_instant_data_scraper(self, page):
        return object()

    async def wait_for_table(self, popup, timeout=None):
        return True

    async def read_stats(self, popup):
        return 1, len(ENG_P2)

    def dom_result_identities(self, state):
        return {u.lower() for u in ENG_P2}

    async def _read_preview(self, popup):
        return {u.lower() for u in ENG_P2}, "sig"

    async def read_preview_grids(self, popup):
        return [{"headers": ["prc-Avatar src", "prc-Link href", "search-match"],
                 "rows": [["https://github.com/%s.png" % u,
                           "https://github.com/%s" % u, u] for u in ENG_P2],
                 "identities": {u.lower() for u in ENG_P2},
                 "signature": "sig"}]

    async def establish_results_table(self, popup, page, expected):
        return True

    async def preview_identities(self, popup, expected_ids=None):
        return {u.lower() for u in ENG_P2}

    def _select_grid(self, grids, expected):
        return InstantDataScraper._select_grid(grids, expected)

    def _identities_from_cells(self, cells):
        return InstantDataScraper._identities_from_cells(cells)


def run_main(monkey_scraper, tmp):
    """Drive diag.mode_diagnose_ids_rows() with all live parts faked; return (rc, printed)."""
    saved = {name: getattr(diag, name) for name in
             ("BrowserManager", "InstantDataScraper",
              "classify_search_page", "expected_total_users", "page_number",
              "reach_page")}
    restore_config = point_config_at(tmp)

    async def fake_reach(page, scraper, query, target_page):
        return {"url": page.url, "reportedTotalUsers": "4",
                "resultProfileUsers": list(ENG_P2)}

    async def no_sleep(_seconds):
        return None

    orig_sleep = diag.asyncio.sleep
    try:
        diag.BrowserManager = FakeBrowser
        diag.InstantDataScraper = monkey_scraper
        diag.classify_search_page = lambda state: ("RESULTS_MULTI_PAGE", "ok")
        diag.expected_total_users = lambda state: (4, "reported")
        diag.page_number = lambda url: 2
        diag.reach_page = fake_reach
        diag.asyncio.sleep = no_sleep
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = asyncio.run(diag.mode_diagnose_ids_rows("Engineering Student Visakhapatnam",
                                       2, False))
        return rc, buf.getvalue()
    finally:
        diag.asyncio.sleep = orig_sleep
        restore_config()
        for name, value in saved.items():
            setattr(diag, name, value)


class TestDiagnosticNeverCrashesReporting(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="diagtest_")

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_unstable_table_reports_failed_without_nameerror(self):
        # stabilize_results_table -> False: the export path is SKIPPED, which
        # is exactly the branch that used to hit `NameError: exported`.
        class Unstable(BaseFakeScraper):
            async def stabilize_results_table(self, popup, page, expected):
                return False

            async def export_page_from_preview(self, popup, dest, expected):
                raise AssertionError("export must not run when unstable")

        rc, out = run_main(Unstable, self.tmp)
        self.assertEqual(rc, 0)                       # completed, did not crash
        self.assertIn("IDS export status      : FAILED", out)
        self.assertIn("written XLSX rows      : N/A", out)
        self.assertIn("Exact export match     : NO", out)
        self.assertNotIn("Traceback", out)

    def test_export_raises_is_caught_and_reported(self):
        # the preview export blows up: still no crash, still a FAILED report.
        class Boom(BaseFakeScraper):
            async def stabilize_results_table(self, popup, page, expected):
                return True

            async def export_page_from_preview(self, popup, dest, expected):
                raise RuntimeError("browser exploded mid-export")

        rc, out = run_main(Boom, self.tmp)
        self.assertEqual(rc, 0)
        self.assertIn("IDS export status      : FAILED", out)
        self.assertIn("written XLSX rows      : N/A", out)
        self.assertIn("browser exploded mid-export", out)

    def test_successful_export_reports_exact_match(self):
        # the happy path the live run is chasing: 4 usable / 4 expected,
        # written from the verified preview (no #xlsx click).
        from test_export_identity import write_identities

        class Good(BaseFakeScraper):
            async def stabilize_results_table(inner, popup, page, expected):
                return True

            async def export_page_from_preview(inner, popup, dest, expected):
                write_identities(dest, ENG_P2)
                return len(ENG_P2)

        rc, out = run_main(Good, self.tmp)
        self.assertEqual(rc, 0)
        self.assertIn("IDS export status      : OK", out)
        self.assertIn("written XLSX rows      : 4 (raw)", out)
        self.assertIn("written XLSX usable    : 4", out)
        self.assertIn("Exact export match     : YES", out)
        self.assertIn("MISSING                : 0", out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
