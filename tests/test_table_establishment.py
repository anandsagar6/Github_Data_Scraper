"""IDS auto-selects a wrong table on small pages; establish the right one.

THE LIVE BUG THIS FILE EXISTS FOR
---------------------------------
On a small GitHub Users results page IDS sometimes auto-selects a table that
is mostly page chrome:

    College Student Vijayawada — GitHub DOM 5 users, IDS raw 3, usable 1
        Nouluri-Nagababu
        "Skip to content ..."               <- garbage row
        "You can't perform that action ..." <- garbage row

The missing and extracted cards have the SAME structure
(div.Result-module__Result__..., a.prc-Link-Link-...), so this is not a
heterogeneous-card problem: IDS picked the wrong table.

THE FIX
-------
Before export, compare the identities in IDS's DISPLAYED preview to the live
DOM identities and cycle "Try another table" (#wrongTable) until they are
EXACTLY equal.  Correctness is identity-set equality -- never raw row count,
never "table detected".  Done pre-export because export can kill Chromium.
"""

import asyncio
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (InstantDataScraper)   # noqa: E402


def cells_for(usernames, garbage=()):
    """A preview's cell texts: an avatar + profile URL per user, plus chrome."""
    cells = []
    for name in usernames:
        cells.append("https://github.com/%s.png?size=40" % name)
        cells.append("https://github.com/%s" % name)
        cells.append(name.title())
    cells.extend(garbage)
    return cells


class FakeTablePopup:
    """A popup whose #wrongTable cycles through a list of candidate tables.

    Each candidate is a list of preview cell texts.  Clicking #wrongTable
    advances to the next; when exhausted it either stops changing or the
    control disappears, per the scenario.
    """

    def __init__(self, candidates, has_wrong_table=True,
                 wrong_table_after=None):
        self.candidates = candidates
        self.index = 0
        self._has_wrong = has_wrong_table
        # after this many clicks, the control disappears
        self.wrong_table_after = wrong_table_after
        self.clicks = 0
        self.closed = False

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    async def evaluate(self, script, arg=None):
        # IDS_PREVIEW_JS: return the current candidate's cells.
        return {"hasTable": True, "cells": self.candidates[self.index]}

    async def click(self, selector, timeout=None):
        if selector == C.SEL_WRONG_TABLE:
            self.clicks += 1
            # wrap around, so cycling back to a seen table is detected by the
            # signature set rather than by a no-change timeout
            self.index = (self.index + 1) % len(self.candidates)

    def locator(self, selector):
        from test_single_page import FakeLocator
        if selector == C.SEL_WRONG_TABLE:
            available = self._has_wrong and (
                self.wrong_table_after is None
                or self.clicks < self.wrong_table_after)
            return FakeLocator(count=1 if available else 0)
        return FakeLocator(count=0)


def scraper():
    said = []
    return InstantDataScraper(browser=None, state=None,
                              say=said.append), said


class TestEstablishTable(unittest.TestCase):
    def establish(self, candidates, expected_users, **kw):
        instance, said = scraper()
        popup = FakeTablePopup([cells_for(*c) if isinstance(c, tuple)
                                else cells_for(c) for c in candidates], **kw)
        expected = InstantDataScraper.dom_result_identities(
            {"resultProfileUsers": expected_users})
        ok = asyncio.run(instance.establish_results_table(popup, None,
                                                          expected))
        return ok, popup, said

    # 1. initial table already exact -> no wrongTable click
    def test_initial_exact_match_needs_no_cycle(self):
        ok, popup, _said = self.establish(
            [["a", "b", "c", "d", "e"]], ["a", "b", "c", "d", "e"])
        self.assertTrue(ok)
        self.assertEqual(popup.clicks, 0)

    # 2. raw 3 / usable 1 / expected 5 -> rejected (if no better candidate)
    def test_wrong_table_with_no_alternative_is_rejected(self):
        ok, _popup, said = self.establish(
            [(["realuser"], ["Skip to content",
                             "You can't perform that action"])],
            ["a", "b", "c", "d", "e"], has_wrong_table=False)
        self.assertFalse(ok)
        self.assertTrue(any("candidate 1: best structured grid has 1 usable"
                             in s for s in said), said)

    # 3. wrong then exact 5/5 -> accepted, cycled once
    def test_second_candidate_exact_is_accepted(self):
        ok, popup, said = self.establish(
            [(["realuser"], ["Skip to content", "garbage"]),
             ["a", "b", "c", "d", "e"]],
            ["a", "b", "c", "d", "e"])
        self.assertTrue(ok)
        self.assertEqual(popup.clicks, 1)
        self.assertTrue(any("exactly match" in s for s in said), said)

    # 4. the exact live case: 4 DOM, 1 usable + 2 garbage, then 4/4
    def test_the_live_four_user_case(self):
        ok, popup, _said = self.establish(
            [(["Suresh240618"], ["Skip to content",
                                 "You can't perform that action at this time"]),
             ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]],
            ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"])
        self.assertTrue(ok)
        self.assertEqual(popup.clicks, 1)

    # 5. signature repeats -> stop incomplete
    def test_a_repeating_signature_stops(self):
        # two candidates, both wrong; clicking wraps back to the first, whose
        # signature was already seen -> stop as cycled.
        ok, popup, said = self.establish(
            [(["x"], ["chrome"]), (["y"], ["chrome2"])],
            ["a", "b", "c"])
        self.assertFalse(ok)
        self.assertTrue(any("cycled" in s for s in said), said)
        self.assertLessEqual(popup.clicks, 4)

    # 6. wrongTable unavailable -> incomplete
    def test_no_wrong_table_control_is_incomplete(self):
        ok, _popup, said = self.establish(
            [(["realuser"], ["chrome"])], ["a", "b", "c"],
            has_wrong_table=False)
        self.assertFalse(ok)
        self.assertTrue(any("unavailable" in s for s in said), said)

    # 7. bounded attempts -> no infinite loop
    def test_bounded_candidate_attempts(self):
        # every candidate distinct and wrong: must stop at MAX_TABLE_CANDIDATES
        cands = [([("u%d" % i)], ["chrome%d" % i]) for i in range(20)]
        ok, popup, said = self.establish(cands, ["a", "b", "c"])
        self.assertFalse(ok)
        self.assertLessEqual(popup.clicks,
                             InstantDataScraper.MAX_TABLE_CANDIDATES)
        self.assertTrue(any("exhausted" in s or "cycled" in s for s in said),
                        said)

    # 8. same raw count, wrong identities -> reject
    def test_same_count_wrong_identities_is_rejected(self):
        # DOM = A,B,C ; IDS = A,X,Y (X,Y are real-looking but wrong users)
        ok, _popup, _said = self.establish(
            [["A", "X", "Y"]], ["A", "B", "C"], has_wrong_table=False)
        self.assertFalse(ok)     # no better candidate offered

    # 9. correct identities in different row order -> accept
    def test_different_order_is_accepted(self):
        ok, popup, _said = self.establish(
            [["c", "a", "b"]], ["a", "b", "c"])
        self.assertTrue(ok)
        self.assertEqual(popup.clicks, 0)

    # 10. duplicate IDS identity -> reject (a duplicate cannot make up a
    #     missing distinct user)
    def test_duplicate_identity_does_not_satisfy(self):
        # preview has A twice and B; DOM is A,B,C -> missing C -> reject
        ok, _popup, _said = self.establish(
            [(["a", "a", "b"], [])], ["a", "b", "c"], has_wrong_table=False)
        self.assertFalse(ok)

    # 11. page-chrome rows never count toward completeness
    def test_chrome_rows_do_not_count(self):
        instance, _said = scraper()
        ids = instance._identities_from_cells(
            ["Skip to content", "You can't perform that action at this time",
             "Sign in", "https://github.com/realuser"])
        self.assertEqual(ids, {"realuser"})

    def test_dom_identities_are_normalised_and_chrome_free(self):
        ids = InstantDataScraper.dom_result_identities(
            {"resultProfileUsers": ["Alice", "BOB", "alice"]})
        self.assertEqual(ids, {"alice", "bob"})

    def test_empty_expected_is_refused(self):
        instance, said = scraper()
        popup = FakeTablePopup([cells_for(["a"])])
        ok = asyncio.run(instance.establish_results_table(popup, None, set()))
        self.assertFalse(ok)


class TestPreviewReading(unittest.TestCase):
    def test_identities_from_cells_uses_the_export_parser(self):
        instance, _said = scraper()
        ids = instance._identities_from_cells([
            "https://github.com/Dorothy-21",              # profile URL
            "https://github.com/Sri1234rama.png?size=40",  # avatar URL
            "Some Name",                                   # plain text
            "https://github.com/octocat/repo",             # repo -> owner
        ])
        self.assertEqual(ids, {"dorothy-21", "sri1234rama", "octocat"})

    def test_signature_changes_with_content(self):
        instance, _said = scraper()

        class P:
            def __init__(self, cells):
                self.cells = cells
            async def evaluate(self, script, arg=None):
                return {"hasTable": True, "cells": self.cells}
        _ids1, sig1 = asyncio.run(instance._read_preview(P(["a", "b"])))
        _ids2, sig2 = asyncio.run(instance._read_preview(P(["a", "c"])))
        self.assertNotEqual(sig1, sig2)


# ============================ table STABILITY wait ==========================
class ScriptedPopup:
    """A popup whose preview is CORRECT or WRONG per read, decided by index.

    is_correct(read_index) picks which cell set each evaluate() returns, so a
    test can make the preview hold steady, flicker, or settle only after a
    number of reads -- exactly the asynchronous mn.data mutation the stability
    wait defends against.
    """

    def __init__(self, is_correct, correct_cells, wrong_cells,
                 has_wrong_table=False):
        self.is_correct = is_correct
        self.correct = correct_cells
        self.wrong = wrong_cells
        self.reads = 0
        self.clicks = 0
        self._has_wrong = has_wrong_table
        self.closed = False

    def is_closed(self):
        return self.closed

    async def bring_to_front(self):
        return None

    async def evaluate(self, script, arg=None):
        cells = self.correct if self.is_correct(self.reads) else self.wrong
        self.reads += 1
        return {"hasTable": True, "cells": cells}

    async def click(self, selector, timeout=None):
        if selector == C.SEL_WRONG_TABLE:
            self.clicks += 1

    def locator(self, selector):
        from test_single_page import FakeLocator
        if selector == C.SEL_WRONG_TABLE:
            return FakeLocator(count=1 if self._has_wrong else 0)
        return FakeLocator(count=0)


class TestStabilizeTable(unittest.TestCase):
    """stabilize_results_table: establish, THEN require the table to hold.

    Timings are shrunk so the read budget is a known, small number:
        budget = max(int(TIMEOUT / POLL) + 1, READS) = max(5, 3) = 5
    so per attempt: 1 establish read, then up to 5 stability reads.
    """

    USERS = ["Suresh240618", "Dorothy-21", "Sri1234rama", "Nasirkhan6"]

    def stabilizer(self):
        instance, said = scraper()
        instance.STABILITY_READS = 3
        instance.STABILITY_POLL = 0.001
        instance.STABILITY_TIMEOUT = 0.004     # -> budget 5
        instance.MAX_ESTABLISH_ATTEMPTS = 2
        return instance, said

    def expected(self):
        return InstantDataScraper.dom_result_identities(
            {"resultProfileUsers": self.USERS})

    def popup(self, is_correct, **kw):
        return ScriptedPopup(is_correct, cells_for(self.USERS),
                             cells_for(["intruder"], ["Skip to content"]), **kw)

    def stabilize(self, popup):
        instance, said = self.stabilizer()
        ok = asyncio.run(instance.stabilize_results_table(
            popup, None, self.expected()))
        return ok, said

    # 1. correct preview, stable for 3 reads -> allowed to export
    def test_stable_three_reads_allows_export(self):
        popup = self.popup(lambda i: True)          # always correct
        ok, said = self.stabilize(popup)
        self.assertTrue(ok)
        self.assertTrue(any("stable" in s for s in said), said)
        self.assertEqual(popup.clicks, 0)           # already correct

    # 2. preview correct then changes and never recovers -> no export
    def test_preview_changes_then_no_export(self):
        # correct only on the very first read; wrong forever after, and no
        # 'Try another table' to fix it -> both attempts fail.
        popup = self.popup(lambda i: i == 0, has_wrong_table=False)
        ok, said = self.stabilize(popup)
        self.assertFalse(ok)
        self.assertTrue(any("did not remain stable" in s for s in said), said)

    # 3. correct -> changes -> re-establish -> stable -> export
    def test_reestablish_then_stable_succeeds(self):
        # attempt 1: establish ok (read 0) but stability fails (reads 1..5
        # wrong); attempt 2: establish ok (read 6) and stability holds
        # (reads 7,8,9 correct).
        popup = self.popup(lambda i: i == 0 or i >= 6)
        ok, said = self.stabilize(popup)
        self.assertTrue(ok)
        self.assertTrue(any("re-establishing" in s for s in said), said)

    # 4. preview never stable -> INCOMPLETE
    def test_never_stable_is_incomplete(self):
        popup = self.popup(lambda i: i % 2 == 0)    # flickers every read
        ok, said = self.stabilize(popup)
        self.assertFalse(ok)
        self.assertTrue(any("did not remain stable" in s for s in said), said)

    # 5. establishment attempts are bounded to MAX_ESTABLISH_ATTEMPTS
    def test_attempts_are_bounded(self):
        instance, _said = self.stabilizer()
        calls = {"n": 0}

        async def counting_establish(popup, page, expected):
            calls["n"] += 1
            return True                              # establishment "succeeds"

        async def always_wrong(popup):
            return {"nobody"}                        # ...but never stabilises

        instance.establish_results_table = counting_establish
        instance.preview_identities = always_wrong
        ok = asyncio.run(instance.stabilize_results_table(
            object(), None, self.expected()))
        self.assertFalse(ok)
        self.assertEqual(calls["n"], instance.MAX_ESTABLISH_ATTEMPTS)

    # 6. post-export wrong identities are STILL rejected (guard unchanged)
    def test_post_export_wrong_identities_still_rejected(self):
        from openpyxl import Workbook
        path = os.path.join(
            os.environ.get("TMPDIR", "/tmp"), "stab_bad_export.xlsx")
        book = Workbook()
        sheet = book.active
        sheet.append(["prc-Avatar src", "prc-Link href", "name"])
        # 1 real user + page chrome -> export identities != DOM identities
        sheet.append(["https://github.com/Suresh240618.png",
                      "https://github.com/Suresh240618", "Suresh"])
        sheet.append(["", "Skip to content", "chrome"])
        sheet.append(["", "images", "chrome"])
        book.save(path)
        book.close()
        got = InstantDataScraper.export_identities(path)
        os.remove(path)
        self.assertEqual(got, {"suresh240618"})
        self.assertNotEqual(got, self.expected(),
                            "a shifted export must fail the identity-set check")


if __name__ == "__main__":
    unittest.main(verbosity=2)
