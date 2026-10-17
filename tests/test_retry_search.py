"""--retry-current-search: re-collect ONE search, safely.

WHY THIS COMMAND EXISTS
-----------------------
The "University Student Visakhapatnam" loss left a search unit marked
COMPLETE with a 10-row XLSX that should have held 19.  Recovering from that
means re-collecting exactly that one search -- and the obvious manual route
(rm the XLSX, edit scraper_state.json with json.dump) bypasses StateManager's
atomic, validated writes and can leave the campaign inconsistent.

WHAT THE COMMAND GUARANTEES
---------------------------
    report by default        --apply is required to change anything
    one atomic state write   never a half-reset state on disk
    narrow blast radius      only collection/processing of the CURRENT unit
    files moved, not deleted the XLSX goes to retry_backup/
    rank-local               another rank's files are never touched
    saved students kept      they are skipped as duplicates on re-collection
"""

import asyncio
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                      # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (NoCurrentUnit, StateManager)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                             # noqa: E402
from test_single_page import (FakeBrowser, FakePage,      # noqa: E402
                              FakePopup, StubScraper, make_export,
                              page_state)


class FakeCity:
    def __init__(self, state, canonical, aliases):
        self.country = "India"
        self.state = state
        self.canonical = canonical
        self.search_aliases = tuple(aliases)


CITIES = [FakeCity("Andhra Pradesh", "Visakhapatnam", ["Visakhapatnam"])]
KEYWORDS = ["College Student", "University Student", "Fresher"]


class Options:
    """Exactly the attributes mode_retry_current_search() reads."""

    def __init__(self, rank=1, total_ranks=1, apply=False):
        self.rank = rank
        self.total_ranks = total_ranks
        self.apply = apply
        self.verbose = False


class RetryCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="retry_")
        self.saved = {name: getattr(C, name) for name in
                      (
                       "HISTORY_DIR", "HISTORY_DB")}
        C.HISTORY_DIR = os.path.join(self.root, "history")
        C.HISTORY_DB = os.path.join(C.HISTORY_DIR, "identities.db")

        # A small plan, so the tests are fast and the queries predictable.
        self.plan = sandbox.plan(KEYWORDS, CITIES)
        self.assigned = self.plan.assign(1, 1)
        self.unit = self.assigned[1]          # "University Student ..."
        self.context = sandbox.enter(os.path.join(self.root, "copy1"))
        # A SECOND copy -- a different checkout, so a different folder.  It
        # must come through the retry untouched.
        self.other = sandbox.other_folder(os.path.join(self.root, "copy2"))
        os.makedirs(self.other.keyword_download_dir, exist_ok=True)
        self.stores = []
        self._set_up_rank_one()
        self._set_up_rank_two()

    def tearDown(self):
        sandbox.restore()
        for store in self.stores:
            store.close()
        for name, value in self.saved.items():
            setattr(C, name, value)
        shutil.rmtree(self.root, ignore_errors=True)

    # -- fixtures ----------------------------------------------------------
    def _set_up_rank_one(self):
        """Rank 1 mid-way through unit 2, collection wrongly 'complete'."""
        self.xlsx = os.path.join(self.context.keyword_download_dir,
                                 "%s.xlsx" % self.unit.slug)
        make_export(self.xlsx, 10)
        state = StateManager(path=self.context.state_file)
        state.begin_unit(1, self.unit)
        state.set_in("collection", complete=True, xlsx_file=self.xlsx, rows=10)
        state.set_in("processing", next_row_index=2, total_rows=10,
                     complete=False)
        state.set(phase="profiles")

        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        for name in ("person1", "person2"):
            store.save_student(self._record(name), search_unit=self.unit,
                               campaign_id="",
                               rank=1, total_ranks=1)
        # A student from a DIFFERENT search, which must also survive.
        store.save_student(self._record("earlier"),
                           search_unit=self.assigned[0],
                           campaign_id="", rank=1,
                           total_ranks=1)
        store.close()
        self.stores.remove(store)

    def _set_up_rank_two(self):
        """A second rank whose files must not be touched."""
        self.other_xlsx = os.path.join(self.other.keyword_download_dir,
                                       "other.xlsx")
        make_export(self.other_xlsx, 4)
        other_plan_unit = self.assigned[0]
        state = StateManager(path=self.other.state_file)
        state.begin_unit(0, other_plan_unit)
        state.set_in("collection", complete=True, xlsx_file=self.other_xlsx,
                     rows=4)
        store = student_scraper.open_store(self.other.store_db)
        store.save_student(self._record("rank2person"))
        store.close()

    @staticmethod
    def _record(username):
        record = {column: "" for column in C.FINAL_COLUMNS}
        record["Student Name"] = username.title()
        record["GitHub Profile"] = "https://github.com/%s" % username
        record["github"] = record["GitHub Profile"]
        return record

    # -- helpers -----------------------------------------------------------
    def run_command(self, rank=1, total_ranks=1, apply=False):
        buffer = io.StringIO()
        with redirect_stdout(buffer):
            code = S.mode_retry_current_search(
                Options(rank, total_ranks, apply))
        return code, buffer.getvalue()

    def fingerprint_of(self, path):
        import hashlib
        with open(path, "rb") as handle:
            return hashlib.sha256(handle.read()).hexdigest()

    def snapshot(self, root):
        found = {}
        for base, _dirs, files in os.walk(root):
            for name in files:
                if name.endswith((".log", "-wal", "-shm")):
                    continue
                path = os.path.join(base, name)
                found[path] = self.fingerprint_of(path)
        return found

    def state_of(self, context=None):
        where = context or self.context
        return StateManager(path=where.state_file).describe_current()


# ================================================= 1. report changes nothing =
class TestReportOnly(RetryCase):
    def test_the_report_changes_nothing_at_all(self):
        before = self.snapshot(self.context.root)
        code, output = self.run_command()
        self.assertEqual(code, 0)
        self.assertEqual(self.snapshot(self.context.root), before,
                         "a report must not modify one byte")

    def test_the_report_says_it_is_a_report(self):
        _code, output = self.run_command()
        self.assertIn("REPORT ONLY", output)
        self.assertIn("Nothing has been changed", output)
        self.assertIn("--apply", output)

    def test_the_report_states_every_required_field(self):
        _code, output = self.run_command()
        for fragment in ("Rank", "Folder", "Search",
                         "Current collection status", "Current XLSX",
                         "Current collected rows",
                         "Current processing position"):
            self.assertIn(fragment, output, fragment)
        self.assertIn(self.unit.search_query, output)
        self.assertIn("complete", output)
        self.assertIn(os.path.basename(self.xlsx), output)
        self.assertIn("row 2 of 10", output)

    def test_the_report_counts_students_already_saved_here(self):
        _code, output = self.run_command()
        self.assertIn("People already saved from this search : 2", output)
        self.assertIn("People in this folder's store in total : 3", output)

    def test_the_report_names_the_file_it_would_move(self):
        _code, output = self.run_command()
        self.assertIn("retry_backup", output)
        self.assertIn(os.path.basename(self.xlsx), output)

    def test_no_retry_backup_directory_is_created_by_a_report(self):
        self.run_command()
        self.assertFalse(os.path.exists(self.context.retry_backup_dir))


# ============================================ 2. --apply resets only this one =
class TestApplyResetsOnlyThisSearch(RetryCase):
    def test_collection_and_processing_are_cleared(self):
        code, _output = self.run_command(apply=True)
        self.assertEqual(code, 0)
        after = self.state_of()
        self.assertFalse(after["collection_complete"])
        self.assertEqual(after["collection_file"], "")
        self.assertEqual(after["collection_rows"], 0)
        self.assertEqual(after["next_row_index"], 0)
        self.assertEqual(after["total_rows"], 0)
        self.assertFalse(after["processing_complete"])
        self.assertFalse(after["cleanup_complete"])
        self.assertEqual(after["phase"], "search")

    def test_the_search_unit_identity_is_unchanged(self):
        before = self.state_of()
        self.run_command(apply=True)
        after = self.state_of()
        for key in ("search_query", "canonical_city", "search_city",
                    "base_keyword", "slug", "state"):
            self.assertEqual(after[key], before[key], key)

    def test_the_rank_position_is_not_rewound(self):
        """A retry of THIS search, not a rewind of the rank."""
        before = self.state_of()
        self.run_command(apply=True)
        after = self.state_of()
        self.assertEqual(after["unit_index"], before["unit_index"])
        self.assertEqual(after["units_done"], before["units_done"])

    def test_the_resume_position_is_unchanged(self):
        """A retry of the current search is not a rewind of the list."""
        import student_scraper as S
        S.save_current_row(1)                 # the run is on search 2
        row_before = self.fingerprint_of(C.CURRENT_ROW_FILE)
        before = self.state_of()
        self.run_command(apply=True)
        after = self.state_of()
        self.assertEqual(after["unit_index"], before["unit_index"])
        self.assertEqual(after["units_done"], before["units_done"])
        self.assertEqual(after["search_query"], before["search_query"])
        self.assertEqual(self.fingerprint_of(C.CURRENT_ROW_FILE), row_before,
                         "current_row.json must not be rewritten")

    def test_learned_pagination_selectors_survive(self):
        state = StateManager(path=self.context.state_file)
        state.remember_next_selector("github.com", 'a[rel="next"]')
        self.run_command(apply=True)
        again = StateManager(path=self.context.state_file)
        self.assertEqual(again.next_selector_for("github.com"),
                         'a[rel="next"]')

    def test_the_state_file_stays_valid_json_and_loadable(self):
        self.run_command(apply=True)
        with open(self.context.state_file, encoding="utf-8") as handle:
            data = json.load(handle)
        self.assertEqual(data["version"], C.STATE_VERSION)
        StateManager(path=self.context.state_file)          # must not raise

    def test_no_temporary_state_file_is_left_behind(self):
        self.run_command(apply=True)
        self.assertFalse(os.path.exists(self.context.state_file + ".tmp"))


# ================================================== 3. students.db untouched =
class TestStoreIsUntouched(RetryCase):
    def test_every_saved_student_survives(self):
        before = self.fingerprint_of(self.context.store_db)
        self.run_command(apply=True)
        store = student_scraper.open_store(self.context.store_db,
                                        read_only=True)
        self.addCleanup(store.close)
        self.assertEqual(store.count(), 3)
        names = {row["github_username_key"] for row in store.iter_records()}
        self.assertEqual(names, {"person1", "person2", "earlier"})
        self.assertEqual(self.fingerprint_of(self.context.store_db), before,
                         "the store must not be rewritten at all")

    def test_students_from_this_search_are_still_known(self):
        self.run_command(apply=True)
        store = student_scraper.open_store(self.context.store_db,
                                        read_only=True)
        self.addCleanup(store.close)
        self.assertEqual(store.is_known(username="person1"), "username")
        self.assertEqual(store.is_known(username="person2"), "username")

    def test_the_history_index_is_not_created_or_touched(self):
        self.run_command(apply=True)
        self.assertFalse(os.path.exists(C.HISTORY_DB))


# ============================================== 4. the XLSX is archived ======
class TestFileArchiving(RetryCase):
    def test_the_xlsx_is_moved_not_deleted(self):
        original = self.fingerprint_of(self.xlsx)
        self.run_command(apply=True)
        self.assertFalse(os.path.exists(self.xlsx),
                         "the stale XLSX must leave keyword_downloads/")
        archived = os.listdir(self.context.retry_backup_dir)
        self.assertEqual(len(archived), 1, archived)
        self.assertTrue(archived[0].startswith(os.path.basename(self.xlsx)))
        self.assertEqual(
            self.fingerprint_of(os.path.join(self.context.retry_backup_dir,
                                             archived[0])),
            original, "the archived copy must be byte-identical")

    def test_a_partial_xlsx_is_archived_too(self):
        partial = self.xlsx.replace(".xlsx", ".partial.xlsx")
        make_export(partial, 3)
        self.run_command(apply=True)
        self.assertFalse(os.path.exists(partial))
        archived = sorted(os.listdir(self.context.retry_backup_dir))
        self.assertEqual(len(archived), 2, archived)
        self.assertTrue(any("partial" in name for name in archived), archived)

    def test_the_archive_lives_inside_this_rank(self):
        self.run_command(apply=True)
        self.assertTrue(self.context.owns(self.context.retry_backup_dir))
        self.assertFalse(self.other.owns(self.context.retry_backup_dir))

    def test_a_missing_xlsx_is_not_an_error(self):
        os.remove(self.xlsx)
        code, output = self.run_command(apply=True)
        self.assertEqual(code, 0)
        self.assertIn("nothing to move", output)
        self.assertFalse(self.state_of()["collection_complete"])

    def test_two_retries_do_not_overwrite_the_first_archive(self):
        self.run_command(apply=True)
        make_export(self.xlsx, 7)
        state = StateManager(path=self.context.state_file)
        state.set_in("collection", complete=True, xlsx_file=self.xlsx, rows=7)
        self.run_command(apply=True)
        self.assertEqual(len(os.listdir(self.context.retry_backup_dir)), 2,
                         "each retry keeps its own timestamped copy")


# ============================================== 5. other ranks untouched =====
class TestOtherRanksUntouched(RetryCase):
    def test_rank_two_is_not_modified(self):
        before = self.snapshot(self.other.root)
        self.run_command(apply=True)
        self.assertEqual(self.snapshot(self.other.root), before,
                         "rank 2 must not be touched by rank 1's retry")

    def test_rank_two_keeps_its_xlsx_and_its_state(self):
        self.run_command(apply=True)
        self.assertTrue(os.path.exists(self.other_xlsx))
        other = self.state_of(self.other)
        self.assertTrue(other["collection_complete"])
        self.assertEqual(other["collection_rows"], 4)

    def test_no_retry_backup_appears_in_rank_two(self):
        self.run_command(apply=True)
        self.assertFalse(os.path.exists(self.other.retry_backup_dir))

    def test_the_retry_only_ever_touches_its_own_folder(self):
        """There is no --rank: the command acts on the folder it runs in."""
        before = self.snapshot(self.other.root)
        code, _output = self.run_command(apply=True)
        self.assertEqual(code, 0)
        self.assertEqual(self.snapshot(self.other.root), before)


# ================================ 6. the rerun re-collects the same search ===
class TestRerunBehaviour(RetryCase):
    def test_the_same_search_unit_resumes_from_the_collection_phase(self):
        self.run_command(apply=True)
        state = StateManager(path=self.context.state_file)
        self.assertEqual(state.search_query, self.unit.search_query,
                         "the same search, not the next one")
        self.assertEqual(state.unit_index, 1)
        self.assertFalse(state.collection_complete,
                         "so do_search_unit() collects instead of processing")

    def test_do_search_unit_actually_re_collects(self):
        """Drive the real workflow: it must call collect(), not process."""
        self.run_command(apply=True)
        page = FakePage(state=page_state())
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        workflow = S.Workflow.__new__(S.Workflow)
        workflow.options = type("O", (), {
            "show": False, "limit": 0, "process_existing": False,
            "max_pages": None, "verbose": False, "rank": 1,
            "total_ranks": 1})()
        workflow.context = self.context
        workflow.searches = self.assigned
        workflow.store = store
        workflow.existing = {}
        workflow._master_dirty = False
        workflow.unit = self.unit
        workflow.state = StateManager(path=self.context.state_file)
        workflow.browser = FakeBrowser(page)
        workflow.scraper = StubScraper(S.NextControl.SINGLE_PAGE, FakePopup(),
                                       rows=19)
        workflow.profiles = None
        workflow.portfolios = None
        workflow.resumes = None
        workflow.merger = S.StudentMerger()
        workflow.reader = S.XlsxReader()
        workflow.transient = S.TransientFailureLog()
        workflow.processed = 0
        workflow.keyword_number = 2
        workflow.keyword_total = len(self.assigned)

        extracted = []
        skipped = []

        async def fake_candidate(_page, candidate):
            # The same gate the real do_candidate() applies first: a known
            # person costs one query, not a profile fetch.
            if workflow.store.is_known(username=candidate.username,
                                       profile_url=candidate.profile_url):
                skipped.append(candidate.username)
                return
            extracted.append(candidate.username)

        workflow.do_candidate = fake_candidate
        asyncio.run(workflow.do_search_unit(page, self.unit))

        self.assertEqual(workflow.scraper.exports, [self.xlsx],
                         "the search was re-collected from GitHub")
        self.assertEqual(sorted(skipped), ["person1", "person2"],
                         "the two already saved must be skipped")
        self.assertEqual(len(extracted), 17,
                         "only the 17 genuinely new people are extracted")
        self.assertNotIn("person1", extracted)
        self.assertNotIn("person2", extracted)

    def test_already_saved_students_are_skipped_as_duplicates(self):
        self.run_command(apply=True)
        # Re-collection produces a bigger export that still contains them.
        make_export(self.xlsx, 19)
        store = student_scraper.open_store(self.context.store_db,
                                        read_only=True)
        self.addCleanup(store.close)
        reader = S.XlsxReader()
        candidates = reader.read(self.xlsx, self.unit.search_query)
        self.assertEqual(len(candidates), 19)
        pending = [candidate for candidate in candidates
                   if not store.is_known(username=candidate.username,
                                         profile_url=candidate.profile_url)]
        self.assertEqual(len(pending), 17)
        skipped = {candidate.username for candidate in candidates
                   if candidate not in pending}
        self.assertEqual(skipped, {"person1", "person2"})


# ===================================================== 7. guards and edges ===
class TestGuards(RetryCase):
    def test_no_state_file_is_reported_not_an_error(self):
        os.remove(C.STATE_FILE)
        code, output = self.run_command()
        self.assertEqual(code, 0)
        self.assertIn("no state", output)
        self.assertIn("nothing to retry", output.lower())

    def test_a_state_with_no_current_search_is_reported(self):
        os.remove(C.STATE_FILE)
        StateManager().save()                 # an empty, valid state
        code, output = self.run_command()
        self.assertEqual(code, 0)
        self.assertIn("No current search is recorded", output)
        self.assertIn("nothing to retry", output.lower())

    def test_a_report_changes_nothing_at_all(self):
        before = self.snapshot(self.context.root)
        code, _output = self.run_command()
        self.assertEqual(code, 0)
        self.assertEqual(self.snapshot(self.context.root), before,
                         "a report must change nothing")

    def test_a_changed_keyword_list_no_longer_refuses(self):
        """The campaign gate is gone: a retry never refuses on the plan.

        The old version compared plan fingerprints and returned 2 rather than
        retry.  There is no plan object to compare any more, and the retry
        only ever touches the search recorded in this folder's own state.
        """
        changed = sandbox.plan(KEYWORDS + ["DevOps Engineer"], CITIES)
        self.assertTrue(changed.assign(1, 1))
        code, output = self.run_command(apply=True)
        self.assertEqual(code, 0)
        self.assertNotIn("REFUSING", output)


# ============================== 8. the StateManager primitive in isolation ===
class TestResetPrimitive(RetryCase):
    def state(self):
        return StateManager(path=self.context.state_file)

    def test_it_returns_the_unit_being_retried(self):
        unit = self.state().reset_for_recollection()
        self.assertEqual(unit["search_query"], self.unit.search_query)

    def test_it_writes_once_and_atomically(self):
        """One save(), so no half-reset state can be observed on disk."""
        state = self.state()
        writes = []
        original = state.save

        def counting_save():
            writes.append(1)
            original()

        state.save = counting_save
        state.reset_for_recollection()
        self.assertEqual(len(writes), 1,
                         "two writes would expose an inconsistent state")

    def test_it_refuses_when_there_is_no_current_unit(self):
        empty = os.path.join(self.root, "empty_state.json")
        state = StateManager(path=empty)
        with self.assertRaises(NoCurrentUnit):
            state.reset_for_recollection()

    def test_describe_current_reports_without_changing_anything(self):
        before = self.fingerprint_of(self.context.state_file)
        described = self.state().describe_current()
        self.assertEqual(described["search_query"], self.unit.search_query)
        self.assertEqual(described["collection_rows"], 10)
        self.assertEqual(self.fingerprint_of(self.context.state_file), before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
