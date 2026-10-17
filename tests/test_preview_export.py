"""Export a GitHub Users page from the VERIFIED PREVIEW, never IDS #xlsx.

WHY THIS FILE EXISTS
--------------------
Final proof on "Engineering Student Visakhapatnam" page 2:

    GitHub DOM        = 4
    IDS established    = 4/4 exact
    IDS stable         = 4/4
    final pre-click    = 4/4
    IDS #xlsx export   = WRONG data (1 user + page chrome)

The extension's #xlsx button can export a table that diverged from the
validated preview, and clicking it can kill Chromium here.  So #xlsx is no
longer used to collect a GitHub Users results page.  Instead the verified
preview grid IS the source: Python reads its headers and rows and writes the
page XLSX directly, with the identity set checked BEFORE the write (preview ==
DOM) and AFTER (written file == DOM).

These tests exercise the REAL InstantDataScraper.read_preview_grids /
write_preview_xlsx / export_page_from_preview, the REAL XlsxReader, and the
REAL merge_staged_pages -- no #xlsx click anywhere.  The identity of a page
comes from ONE structured grid's rows, never a flat union across grids.
"""

import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import asyncio

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                         # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (InstantDataScraper)   # noqa: E402

NODE = shutil.which("node") or shutil.which("nodejs")

# Realistic IDS column names: an avatar-URL column, a profile-URL column, and a
# highlighted search-match (bio) column -- the shape XlsxReader expects.
HEADERS = ["prc-Avatar-Avatar-0xaUi src", "prc-Link-Link-9ZwDx href",
           "search-match"]


def one_grid(usernames, garbage_rows=()):
    """A single structured grid: {headers, rows} for these users (+ chrome)."""
    rows = []
    for name in usernames:
        rows.append(["https://github.com/%s.png?size=40" % name,
                     "https://github.com/%s" % name, name.title()])
    for text in garbage_rows:
        rows.append(["", text, "chrome"])
    return {"headers": list(HEADERS), "rows": rows}


def payload(grids):
    """The dict popup.evaluate(IDS_PREVIEW_JS) returns: a LIST of concrete
    grids plus the LEGACY flat `cells` union across all of them (so a test can
    prove the flat union is ignored)."""
    cells = []
    for grid in grids:
        cells.extend(c for c in grid["headers"] if c)
        for row in grid["rows"]:
            cells.extend(c for c in row if c)
    return {"hasTable": True, "grids": list(grids), "cells": cells}


def grid_for(usernames, garbage_rows=()):
    """Convenience: a payload with exactly one grid of these users."""
    return payload([one_grid(usernames, garbage_rows)])


class GridPopup:
    """A popup whose preview is a fixed payload; records any #xlsx click."""

    def __init__(self, grid):
        self.grid = grid
        self.xlsx_clicks = 0
        self.closed = False

    async def evaluate(self, script, arg=None):
        return self.grid

    async def click(self, selector, timeout=None):
        if selector == C.SEL_XLSX:
            self.xlsx_clicks += 1

    async def bring_to_front(self):
        return None

    def is_closed(self):
        return self.closed

    def locator(self, selector):
        from test_single_page import FakeLocator
        # No "Try another table" control -> establishment can't cycle.
        return FakeLocator(count=0)


class CyclingGridPopup:
    """A popup whose #wrongTable cycles through a list of payloads."""

    def __init__(self, payloads, has_wrong_table=True):
        self.payloads = payloads
        self.index = 0
        self._has_wrong = has_wrong_table
        self.clicks = 0
        self.xlsx_clicks = 0
        self.closed = False

    async def evaluate(self, script, arg=None):
        return self.payloads[self.index]

    async def click(self, selector, timeout=None):
        if selector == C.SEL_XLSX:
            self.xlsx_clicks += 1
        elif selector == C.SEL_WRONG_TABLE:
            self.clicks += 1
            self.index = min(self.index + 1, len(self.payloads) - 1)

    async def bring_to_front(self):
        return None

    def is_closed(self):
        return self.closed

    def locator(self, selector):
        from test_single_page import FakeLocator
        if selector == C.SEL_WRONG_TABLE:
            more = self.index < len(self.payloads) - 1
            return FakeLocator(count=1 if (self._has_wrong and more) else 0)
        return FakeLocator(count=0)


def scraper():
    said = []
    return InstantDataScraper(browser=None, state=None,
                              say=said.append), said


def expected_of(usernames):
    return InstantDataScraper.dom_result_identities(
        {"resultProfileUsers": list(usernames)})


class TestVerifiedPreviewExport(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pvexp_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def export(self, usernames, expected=None, garbage=()):
        instance, said = scraper()
        popup = GridPopup(grid_for(usernames, garbage))
        dest = os.path.join(self.dir, "page_001.xlsx")
        exp = expected_of(usernames if expected is None else expected)
        rows = asyncio.run(instance.export_page_from_preview(popup, dest, exp))
        return rows, dest, popup, said

    # 1. preview 4/4 -> Python XLSX -> 4/4
    def test_preview_four_writes_four(self):
        users = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]
        rows, dest, popup, _said = self.export(users)
        self.assertEqual(rows, 4)
        self.assertTrue(os.path.exists(dest))
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))
        self.assertEqual(len(student_scraper.XlsxReader().read(dest, "")), 4)
        self.assertEqual(popup.xlsx_clicks, 0)            # item 9

    # 2. preview 5/5 -> XLSX -> 5/5
    def test_preview_five_writes_five(self):
        users = ["u1", "u2", "u3", "u4", "u5"]
        rows, dest, popup, _said = self.export(users)
        self.assertEqual(rows, 5)
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))
        self.assertEqual(popup.xlsx_clicks, 0)

    # 3. wrong preview -> no XLSX written
    def test_wrong_preview_writes_nothing(self):
        instance, said = scraper()
        # preview shows only 1 of the 4 expected users (+ page chrome)
        popup = GridPopup(grid_for(["Suresh240618"],
                                   garbage_rows=["Skip to content", "images"]))
        dest = os.path.join(self.dir, "page_001.xlsx")
        expected = expected_of(["Suresh240618", "Dorothy-21", "Sri1234rama",
                                "Nasirkhan6"])
        rows = asyncio.run(instance.export_page_from_preview(popup, dest,
                                                             expected))
        self.assertEqual(rows, -1)
        self.assertFalse(os.path.exists(dest))
        self.assertEqual(popup.xlsx_clicks, 0)
        self.assertTrue(any("no single structured grid" in s for s in said),
                        said)

    # 5. headers/rows preserved and re-readable
    def test_headers_and_rows_preserved(self):
        from openpyxl import load_workbook
        users = ["alpha", "bravo", "charlie"]
        _rows, dest, _popup, _said = self.export(users)
        book = load_workbook(dest, data_only=True)
        sheet = book.active
        values = [["" if c is None else str(c) for c in row]
                  for row in sheet.iter_rows(values_only=True)]
        book.close()
        self.assertEqual(values[0], HEADERS)              # header row preserved
        self.assertEqual(len(values), 1 + len(users))
        # each data row keeps its avatar URL, profile URL and name, in order
        self.assertEqual(values[1], ["https://github.com/alpha.png?size=40",
                                     "https://github.com/alpha", "Alpha"])

    # 9. explicitly: no IDS #xlsx click occurs on export
    def test_no_xlsx_click_on_export(self):
        _rows, _dest, popup, _said = self.export(["a", "b", "c"])
        self.assertEqual(popup.xlsx_clicks, 0)

    # 10. the browser/popup stays alive: export does not close or navigate it
    def test_browser_stays_alive_after_export(self):
        _rows, _dest, popup, _said = self.export(["a", "b", "c", "d"])
        self.assertFalse(popup.closed)
        self.assertFalse(popup.is_closed())
        self.assertEqual(popup.xlsx_clicks, 0)

    # empty expected set is refused (never writes an unverified page)
    def test_empty_expected_is_refused(self):
        instance, _said = scraper()
        popup = GridPopup(grid_for(["a"]))
        dest = os.path.join(self.dir, "p.xlsx")
        rows = asyncio.run(instance.export_page_from_preview(popup, dest,
                                                             set()))
        self.assertEqual(rows, -1)
        self.assertFalse(os.path.exists(dest))


class TestMultiPageMerge(unittest.TestCase):
    """Verified-preview page files merge to the right totals (real merge)."""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pvmerge_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _page(self, number, usernames):
        instance, _said = scraper()
        grid = one_grid(usernames)
        path = student_scraper.page_path(self.dir, number)
        instance.write_preview_xlsx(path, grid["headers"], grid["rows"])
        return path

    # 7. multi-page 10 + 4 -> 14
    def test_ten_plus_four_is_fourteen(self):
        p1 = self._page(1, ["u%02d" % i for i in range(10)])
        p2 = self._page(2, ["v%02d" % i for i in range(4)])
        dest = os.path.join(self.dir, "merged.xlsx")
        unique, dups, _unusable = student_scraper.merge_staged_pages(
            [p1, p2], dest, say=lambda *_a: None)
        self.assertEqual(unique, 14)
        self.assertEqual(dups, 0)
        self.assertEqual(len(student_scraper.XlsxReader().read(dest, "")), 14)

    # 8. multi-page 10 + 10 + 5 -> 25
    def test_ten_ten_five_is_twentyfive(self):
        p1 = self._page(1, ["a%02d" % i for i in range(10)])
        p2 = self._page(2, ["b%02d" % i for i in range(10)])
        p3 = self._page(3, ["c%02d" % i for i in range(5)])
        dest = os.path.join(self.dir, "merged.xlsx")
        unique, _dups, _unusable = student_scraper.merge_staged_pages(
            [p1, p2, p3], dest, say=lambda *_a: None)
        self.assertEqual(unique, 25)
        self.assertEqual(len(student_scraper.XlsxReader().read(dest, "")), 25)


class TestPostWriteRejection(unittest.TestCase):
    """Item 4: a written page whose identity set is wrong -> INCOMPLETE.

    Drives the REAL PageByPageCollector._export_current_page with a scraper
    whose export writes the WRONG users (bypassing the pre-write guard); the
    post-write identity-set check must raise and staging must be preserved.
    """

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pvpost_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_written_wrong_identities_raise(self):
        from test_export_identity import write_identities
        from test_single_page import FakePopup, FakePage, page_state

        dom_users = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]
        staging = self.dir

        class Sc:
            async def open_instant_data_scraper(self, page):
                return FakePopup()

            async def wait_for_table(self, popup, timeout=None):
                return True

            async def read_stats(self, popup):
                return 1, len(dom_users)

            async def stabilize_results_table(self, popup, page, expected):
                return True

            async def preview_identities(self, popup, expected_ids=None):
                import student_scraper as _id
                return {_id.normalize_username(u) for u in dom_users}

            async def export_page_from_preview(self, popup, destination,
                                               expected):
                # writes only 1 real user -> written identities != DOM
                write_identities(destination, ["Suresh240618"],
                                 ("Skip to content", "images"))
                return 3

            def dom_result_identities(self, state):
                return InstantDataScraper.dom_result_identities(state)

            def export_identities(self, path):
                return InstantDataScraper.export_identities(path)

        collector = student_scraper.PageByPageCollector(
            None, Sc(), type("Ctx", (), {"keyword_download_dir": self.dir})(),
            lambda *_a: None)
        collector.staging = staging
        import student_scraper as _id
        expected = {_id.normalize_username(u) for u in dom_users}
        with self.assertRaises(student_scraper.PageCrawlError) as caught:
            asyncio.run(collector._export_current_page(
                FakePage(state=page_state()), 2, expected))
        self.assertIn("do not match", str(caught.exception))
        # staging file preserved for diagnosis (not silently deleted)
        self.assertTrue(os.path.exists(student_scraper.page_path(staging, 2)))


class TestOneSourceOfTruth(unittest.TestCase):
    """The EXACT live bug: flat/global `cells` lied; per-grid rows are truth.

    A flat union of every td/th across Handsontable's overlay clones can
    contain all expected identities while no single concrete grid does -- and
    that is what let a validated "4/4" write a 1-row file.  Establishment,
    stability, verification and the WRITE must all read ONE grid's own rows.
    """

    USERS = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pvtruth_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def expected(self):
        return expected_of(self.USERS)

    # 1. flat cells hold A,B,C,D but the only structured grid is A + chrome
    def test_flat_union_never_establishes(self):
        inst, _said = scraper()
        grid = one_grid(["Suresh240618"], ["Skip to content", "images"])
        pay = payload([grid])
        pay["cells"] += ["https://github.com/Dorothy-21",       # as clones would
                         "https://github.com/Sri1234rama",
                         "https://github.com/Nasirkhan6"]
        # the flat union DOES contain all four -- proving the old source lied
        self.assertEqual(
            InstantDataScraper._identities_from_cells(pay["cells"]),
            self.expected())
        popup = GridPopup(pay)
        self.assertFalse(asyncio.run(
            inst.establish_results_table(popup, None, self.expected())))
        rows = asyncio.run(inst.export_page_from_preview(
            popup, os.path.join(self.dir, "p.xlsx"), self.expected()))
        self.assertEqual(rows, -1)
        self.assertEqual(popup.xlsx_clicks, 0)

    # 2. candidate 1 structured rows wrong; candidate 2 = A,B,C,D -> establish 2
    def test_cycles_to_the_correct_structured_grid(self):
        inst, _said = scraper()
        wrong = payload([one_grid(["Suresh240618"], ["chrome"])])
        right = payload([one_grid(self.USERS)])
        popup = CyclingGridPopup([wrong, right])
        self.assertTrue(asyncio.run(
            inst.establish_results_table(popup, None, self.expected())))
        self.assertEqual(popup.clicks, 1)

    # 3. stability is judged on IDENTITY, not full row text: a non-identity
    #    cell changing every read (e.g. a bio/timing cell) is still STABLE
    def test_stability_ignores_non_identity_changes(self):
        inst, _said = scraper()
        inst.STABILITY_READS = 3
        inst.STABILITY_POLL = 0.001
        inst.STABILITY_TIMEOUT = 0.02
        stable = GridPopup(payload([one_grid(self.USERS)]))
        self.assertTrue(asyncio.run(
            inst._await_stable_preview(stable, self.expected())))

        class NonIdentityChurn(GridPopup):
            def __init__(self, users):
                super().__init__(None)
                self._users = users
                self._n = 0

            async def evaluate(self, script, arg=None):
                self._n += 1
                grid = one_grid(self._users)
                grid["rows"][0][2] = "updated %d ms ago" % self._n  # not an id
                return payload([grid])

        # identities match every read; only non-identity text churns -> STABLE
        self.assertTrue(asyncio.run(
            inst._await_stable_preview(NonIdentityChurn(self.USERS),
                                       self.expected())))

    # 4 & 5. final structured rows exact -> XLSX exact, 4/4
    def test_final_structured_rows_write_exact(self):
        inst, _said = scraper()
        popup = GridPopup(payload([one_grid(self.USERS)]))
        dest = os.path.join(self.dir, "p.xlsx")
        rows = asyncio.run(inst.export_page_from_preview(popup, dest,
                                                         self.expected()))
        self.assertEqual(rows, 4)
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         self.expected())
        self.assertEqual(popup.xlsx_clicks, 0)

    # 6. NO union across overlay clones: two half grids -> reject
    def test_no_union_across_clones(self):
        inst, _said = scraper()
        pay = payload([one_grid(["Suresh240618", "Dorothy-21"]),
                       one_grid(["Sri1234rama", "Nasirkhan6"])])
        popup = GridPopup(pay)
        self.assertFalse(asyncio.run(
            inst.establish_results_table(popup, None, self.expected())))
        rows = asyncio.run(inst.export_page_from_preview(
            popup, os.path.join(self.dir, "p.xlsx"), self.expected()))
        self.assertEqual(rows, -1)

    # 7. split identities never merge into a fake match; but the ONE complete
    #    grid IS selected when it exists among partial grids
    def test_split_does_not_merge_but_whole_grid_is_selected(self):
        inst, _said = scraper()
        split = payload([one_grid(["Suresh240618", "Dorothy-21"]),
                         one_grid(["Sri1234rama", "Nasirkhan6"]),
                         one_grid(["Dorothy-21", "Sri1234rama"])])
        self.assertFalse(asyncio.run(
            inst.establish_results_table(GridPopup(split), None,
                                         self.expected())))
        whole = payload([one_grid(["Suresh240618", "Dorothy-21"]),
                         one_grid(self.USERS),
                         one_grid(["Sri1234rama"])])
        popup = GridPopup(whole)
        self.assertTrue(asyncio.run(
            inst.establish_results_table(popup, None, self.expected())))
        dest = os.path.join(self.dir, "whole.xlsx")
        rows = asyncio.run(inst.export_page_from_preview(popup, dest,
                                                         self.expected()))
        self.assertEqual(rows, 4)
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         self.expected())


def grid_extra(users, extra, order=None):
    """A grid with an extra NON-IDENTITY trailing cell per row (timing/button
    text).  `order` optionally reorders the users."""
    names = list(order) if order is not None else list(users)
    rows = []
    for name in names:
        rows.append(["https://github.com/%s.png" % name,
                     "https://github.com/%s" % name, name.title(), extra])
    return {"headers": HEADERS + ["state"], "rows": rows}


class SequencePopup:
    """A popup that returns a different payload on each read (last repeats)."""

    def __init__(self, payloads):
        self.payloads = payloads
        self.i = 0

    async def evaluate(self, script, arg=None):
        p = self.payloads[min(self.i, len(self.payloads) - 1)]
        self.i += 1
        return p

    async def bring_to_front(self):
        return None

    def is_closed(self):
        return False

    def locator(self, selector):
        from test_single_page import FakeLocator
        return FakeLocator(count=0)


class TestIdentityProjectionStability(unittest.TestCase):
    """Stability is judged on the IDENTITY projection, not full row text.

    Non-identity cells (result timing "(29 ms)", counters, button/state text)
    change while the users stay the same and must NOT read as instability; a
    removed or replaced identity must.
    """

    USERS = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]

    def _scraper(self):
        inst, _said = scraper()
        inst.STABILITY_READS = 3
        inst.STABILITY_POLL = 0.001
        inst.STABILITY_TIMEOUT = 0.05      # budget comfortably > 3 reads
        inst.MAX_ESTABLISH_ATTEMPTS = 2
        return inst

    def expected(self):
        return expected_of(self.USERS)

    def _await(self, payloads):
        inst = self._scraper()
        popup = SequencePopup(payloads)
        return asyncio.run(inst._await_stable_preview(popup, self.expected()))

    # 1. same identities, changing timing text -> stable
    def test_changing_timing_text_is_stable(self):
        self.assertTrue(self._await([
            payload([grid_extra(self.USERS, "(29 ms)")]),
            payload([grid_extra(self.USERS, "(45 ms)")]),
            payload([grid_extra(self.USERS, "(78 ms)")]),
            payload([grid_extra(self.USERS, "(91 ms)")]),
        ]))

    # 2. same identities, changing unrelated button/state text -> stable
    def test_changing_button_text_is_stable(self):
        self.assertTrue(self._await([
            payload([grid_extra(self.USERS, "Follow")]),
            payload([grid_extra(self.USERS, "Following")]),
            payload([grid_extra(self.USERS, "Sponsor")]),
            payload([grid_extra(self.USERS, "Unfollow")]),
        ]))

    # 3. identity removed -> unstable
    def test_identity_removed_is_unstable(self):
        three = self.USERS[:3]
        self.assertFalse(self._await([
            payload([grid_extra(self.USERS, "(29 ms)")]),
            payload([grid_extra(three, "(45 ms)")]),   # one user gone, and stays
        ]))

    # 4. identity replaced -> unstable
    def test_identity_replaced_is_unstable(self):
        swapped = self.USERS[:3] + ["someoneelse"]
        self.assertFalse(self._await([
            payload([grid_extra(self.USERS, "(29 ms)")]),
            payload([grid_extra(swapped, "(45 ms)")]),  # X for D, and stays
        ]))

    # 5. same identities, different row order -> stable
    def test_different_order_is_stable(self):
        rev = list(reversed(self.USERS))
        self.assertTrue(self._await([
            payload([grid_extra(self.USERS, "(29 ms)", order=self.USERS)]),
            payload([grid_extra(self.USERS, "(45 ms)", order=rev)]),
            payload([grid_extra(self.USERS, "(78 ms)", order=self.USERS)]),
            payload([grid_extra(self.USERS, "(91 ms)", order=rev)]),
        ]))

    # 6. no matching grid -> unstable
    def test_no_matching_grid_is_unstable(self):
        self.assertFalse(self._await([
            payload([one_grid(self.USERS[:3])]),   # only 3 of 4, forever
        ]))

    # 7. max establishment/stability attempts stay bounded
    def test_attempts_bounded_when_never_stable(self):
        inst = self._scraper()
        calls = {"n": 0}

        async def counting_establish(popup, page, expected):
            calls["n"] += 1
            return True                              # establishment "succeeds"

        async def never_matches(popup, expected_ids):
            return (None, "sig", [])                 # ...but no grid ever matches

        inst.establish_results_table = counting_establish
        inst._selected_grid = never_matches
        ok = asyncio.run(inst.stabilize_results_table(
            SequencePopup([payload([one_grid(self.USERS)])]), None,
            self.expected()))
        self.assertFalse(ok)
        self.assertEqual(calls["n"], inst.MAX_ESTABLISH_ATTEMPTS)


class TestSerializationRoundTrip(unittest.TestCase):
    """grid rows -> write_preview_xlsx -> XlsxReader must keep every identity.

    The live p2 bug: the grid parser (per CELL) declared 4 usable, but after
    write the per-ROW XlsxReader recovered 1, because the grid packed all four
    users' links into ONE row (the other rows were page chrome).  The writer
    now emits one identity per row, so the round-trip is invariant.
    """

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="pvround_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _matrix(self, path):
        from openpyxl import load_workbook
        book = load_workbook(path, data_only=True)
        sheet = book.active
        values = [["" if c is None else str(c) for c in row]
                  for row in sheet.iter_rows(values_only=True)]
        book.close()
        return values

    # clean grid (one user per row) -> openpyxl matrix == input matrix
    def test_clean_grid_matrix_and_identities_roundtrip(self):
        users = ["Aaa", "Bbb", "Ccc", "Ddd"]
        grid = one_grid(users)
        inst, _said = scraper()
        dest = os.path.join(self.dir, "clean.xlsx")
        n = inst.write_preview_xlsx(dest, grid["headers"], grid["rows"])
        self.assertEqual(n, 4)
        matrix = self._matrix(dest)
        self.assertEqual(matrix[0], grid["headers"])       # headers identical
        self.assertEqual(matrix[1:], grid["rows"])         # rows identical
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))

    # 5 users, same invariant
    def test_five_user_grid_roundtrips(self):
        users = ["u1", "u2", "u3", "u4", "u5"]
        grid = one_grid(users)
        inst, _said = scraper()
        dest = os.path.join(self.dir, "five.xlsx")
        n = inst.write_preview_xlsx(dest, grid["headers"], grid["rows"])
        self.assertEqual(n, 5)
        self.assertEqual(self._matrix(dest)[1:], grid["rows"])
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))
        self.assertEqual(len(student_scraper.XlsxReader().read(dest, "")), 5)

    # THE LIVE BUG: all four users packed into ONE grid row + chrome rows.
    def test_packed_row_still_recovers_all_four(self):
        users = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]
        packed = []
        for name in users:
            packed += ["https://github.com/%s.png?size=40" % name,
                       "https://github.com/%s" % name]
        rows = [packed,
                ["", "Skip to content", ""],
                ["", "images", ""]]
        # the naive per-ROW parse of the packed row yields only ONE identity --
        # this is exactly why writing rows verbatim lost three users
        naive = student_scraper.XlsxReader._username_from(packed)
        self.assertEqual(naive.lower(), "suresh240618")

        inst, _said = scraper()
        dest = os.path.join(self.dir, "packed.xlsx")
        n = inst.write_preview_xlsx(dest, HEADERS, rows)
        self.assertEqual(n, 4)                              # one row per user
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))
        self.assertEqual(len(student_scraper.XlsxReader().read(dest, "")), 4)

    # a chrome-only row contributes no identity and is dropped
    def test_chrome_rows_are_dropped(self):
        users = ["alpha", "bravo"]
        grid = one_grid(users, garbage_rows=["Skip to content", "images"])
        inst, _said = scraper()
        dest = os.path.join(self.dir, "chrome.xlsx")
        n = inst.write_preview_xlsx(dest, grid["headers"], grid["rows"])
        self.assertEqual(n, 2)                              # 2 users, chrome gone
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))

    # end-to-end via export_page_from_preview on a packed grid
    def test_export_page_from_preview_roundtrips_packed_grid(self):
        users = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]
        packed = []
        for name in users:
            packed += ["https://github.com/%s.png" % name,
                       "https://github.com/%s" % name]
        pay = payload([{"headers": HEADERS, "rows": [packed]}])
        inst, _said = scraper()
        popup = GridPopup(pay)
        dest = os.path.join(self.dir, "e2e.xlsx")
        n = asyncio.run(inst.export_page_from_preview(popup, dest,
                                                      expected_of(users)))
        self.assertEqual(n, 4)
        self.assertEqual(InstantDataScraper.export_identities(dest),
                         expected_of(users))
        self.assertEqual(popup.xlsx_clicks, 0)


# ============================ the JS actually extracts structure ============
# Two concrete .htCore grids on the page: a WRONG one (1 user + chrome) and the
# CORRECT one (4 users).  The JS must return them SEPARATELY as grids[0]/[1] --
# never a union -- so the flat `cells` union (which holds all 5 identities) is
# provably not what selection uses.
HARNESS = r"""
function cell(text){ return { textContent: text, querySelectorAll: () => [] }; }
function makeCore(headerTexts, dataRows){
  const headerThs = headerTexts.map(cell);
  const trs = dataRows.map(cells => {
    const tds = cells.map(cell);
    return { querySelectorAll: (sel) => sel === 'td' ? tds : [] };
  });
  let cells = headerThs.slice();
  trs.forEach(tr => { cells = cells.concat(tr.querySelectorAll('td')); });
  const core = {
    _cells: cells,
    querySelectorAll: (sel) => {
      if (sel === 'thead tr th') return headerThs;
      if (sel === 'tbody tr') return trs;
      return [];
    },
  };
  return core;
}
const wrongCore = makeCore(
  ['', 'prc-Avatar src', 'prc-Link href', 'search-match'],
  [['https://github.com/Suresh240618.png', 'https://github.com/Suresh240618',
    'Suresh'],
   ['', 'Skip to content', 'chrome']]);
const rightCore = makeCore(
  ['', 'prc-Avatar src', 'prc-Link href', 'search-match'],
  [['https://github.com/Suresh240618.png', 'https://github.com/Suresh240618',
    'Suresh'],
   ['https://github.com/Dorothy-21.png', 'https://github.com/Dorothy-21',
    'Dorothy'],
   ['https://github.com/Sri1234rama.png', 'https://github.com/Sri1234rama',
    'Sri'],
   ['https://github.com/Nasirkhan6.png', 'https://github.com/Nasirkhan6',
    'Nasir']]);
const cores = [wrongCore, rightCore];
let allCells = [];
cores.forEach(c => { allCells = allCells.concat(c._cells); });
const host = {
  hotInstance: undefined,
  querySelector: () => null,
  querySelectorAll: (sel) => {
    if (sel === 'td, th') return allCells;
    if (sel === '.htCore') return cores;
    if (sel === '.handsontable') return [];
    if (sel === 'table') return cores;
    return [];
  },
};
global.document = {
  querySelector: (sel) => sel === '#hot' ? host : null,
};
const out = PROBE();
process.stdout.write(JSON.stringify({
  ngrids: out.grids.length,
  headers0: out.grids[0].headers, nrows0: out.grids[0].rows.length,
  nrows1: out.grids[1].rows.length,
  row0_1: out.grids[1].rows[0][1],
  ncells: out.cells.length,
}));
"""


@unittest.skipUnless(NODE, "node is not installed; JS execution not checked")
class TestPreviewJsExtractsStructure(unittest.TestCase):
    def test_js_returns_separate_grids_not_a_union(self):
        script = "const PROBE = %s;\n%s" % (student_scraper.IDS_PREVIEW_JS, HARNESS)
        result = subprocess.run([NODE, "-e", script], capture_output=True,
                                text=True, timeout=30)
        self.assertEqual(result.returncode, 0,
                         result.stderr or result.stdout)
        import json
        out = json.loads(result.stdout)
        # two SEPARATE grids, never merged
        self.assertEqual(out["ngrids"], 2)
        # leading corner <th> dropped: headers align to the 3 data columns
        self.assertEqual(out["headers0"],
                         ["prc-Avatar src", "prc-Link href", "search-match"])
        self.assertEqual(out["nrows0"], 2)     # wrong grid: 1 user + chrome row
        self.assertEqual(out["nrows1"], 4)     # correct grid: 4 users
        self.assertEqual(out["row0_1"], "https://github.com/Suresh240618")
        # the flat union spans BOTH grids -- which is exactly why it must not
        # be the identity source
        self.assertGreater(out["ncells"], out["nrows1"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
