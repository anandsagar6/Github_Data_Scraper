"""The safety rule: success cleans up, any failure preserves everything.

    search unit fully resolved -> delete that unit's XLSX
    crash / Ctrl+C / limit     -> keep the XLSX, the state and temporary.txt

These tests drive the real Workflow with the browser stages stubbed out, so
the decision logic is exercised rather than re-described.

Everything now lives in a throwaway runtime/rank_001 under a temp directory:
the rule is unchanged, but the files it protects are rank-local.
"""

import asyncio
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook              # noqa: E402

import config as C                         # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (StateManager)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                # noqa: E402


class FakeCity:
    """One city, so a test plan is small and its queries are predictable."""

    def __init__(self, state, canonical, aliases):
        self.country = "India"
        self.state = state
        self.canonical = canonical
        self.search_aliases = tuple(aliases)


def make_plan():
    return sandbox.plan(
        ["lnct"], [FakeCity("Madhya Pradesh", "Bhopal", ["Bhopal"])])


class Options:
    show = False
    limit = 0
    process_existing = True
    max_pages = None
    verbose = False
    rank = 1
    total_ranks = 1


def make_export(path, count):
    """An Instant Data Scraper style user-search export with `count` rows."""
    book = Workbook()
    sheet = book.active
    sheet.append(["prc-Avatar-Avatar-0xaUi src", "prc-Link-Link-9ZwDx href",
                  "search-match", "search-match 2"])
    for number in range(1, count + 1):
        sheet.append(["https://github.com/person%d.png?size=40" % number,
                      "https://github.com/person%d" % number,
                      "Person %d" % number, "student"])
    book.save(path)


class StubWorkflow(S.Workflow):
    """Workflow with GitHub/portfolio/resume replaced by canned answers."""

    def __init__(self, options, context, plan, store, fail_at=None):
        self.options = options
        self.context = context
        self.searches = plan.assign(1, 1)
        self.store = store
        self.existing = {}
        self._master_dirty = False
        self.existing = {}
        self._master_dirty = False
        self.unit = self.searches[0]
        self.state = StateManager(path=context.state_file)
        self.browser = None
        self.scraper = None
        self.profiles = None
        self.portfolios = None
        self.resumes = None
        self.reader = S.XlsxReader()
        self.merger = S.StudentMerger()
        self.processed = 0
        self.fail_at = fail_at
        self.seen = []

    async def do_candidate(self, page, candidate):
        self.seen.append(candidate.username)
        if candidate.username == self.fail_at:
            raise RuntimeError("simulated crash")
        duplicate = self.store.is_known(
            username=candidate.username, profile_url=candidate.profile_url)
        if duplicate:
            return
        bundle = S.SourceBundle(candidate=candidate)
        bundle.github_ok = True
        bundle.github_fields = {"fullname": candidate.xlsx_name,
                                "nickname": candidate.username, "links": []}
        bundle.github_text = "Person page"
        student_scraper.write(bundle)
        saved, _reason = self.store.save_student(
            self.merger.build(bundle), search_unit=self.unit,
            campaign_id="", rank=S.rank, total_ranks=S.num_sys)
        if saved:
            self.processed += 1
        student_scraper.clear()


class TestLifecycle(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="lifecycle_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = make_plan()
        self.context = sandbox.enter(self.runtime)
        self.history = os.path.join(self.root, "identities.db")
        self.unit = self.plan.assign(1, 1)[0]
        self.export = os.path.join(self.context.keyword_download_dir,
                                   "%s.xlsx" % self.unit.slug)
        make_export(self.export, 20)
        self.stores = []

    def tearDown(self):
        sandbox.restore()
        for store in self.stores:
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def _workflow(self, options, fail_at=None):
        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        return StubWorkflow(options, self.context, self.plan, store, fail_at=fail_at)

    def _run(self, workflow):
        return asyncio.run(workflow.do_search_unit(None, self.unit))

    def test_success_deletes_the_keyword_xlsx(self):
        workflow = self._workflow(Options())
        self._run(workflow)
        self.assertFalse(os.path.exists(self.export),
                         "a fully resolved keyword must delete its XLSX")
        self.assertEqual(workflow.processed, 20)
        self.assertTrue(workflow.state.data["cleanup_complete"])

    def test_crash_preserves_everything(self):
        workflow = self._workflow(Options(), fail_at="person8")
        with self.assertRaises(S.Stopped):
            self._run(workflow)
        self.assertTrue(os.path.exists(self.export),
                        "a crash must NEVER delete the keyword XLSX")
        self.assertFalse(workflow.state.data["cleanup_complete"])
        # Rows 1-7 were saved; the resume point is row 8 (index 7).
        self.assertEqual(workflow.processed, 7)
        self.assertEqual(workflow.state.next_row_index, 7)

    def test_restart_after_crash_resumes_without_duplicates(self):
        first = self._workflow(Options(), fail_at="person8")
        with self.assertRaises(S.Stopped):
            self._run(first)

        second = self._workflow(Options())          # no failure this time
        self._run(second)

        # It restarted at row 8, not row 1.
        self.assertEqual(second.seen[0], "person8")
        self.assertFalse(os.path.exists(self.export))

        # Exactly 20 rows overall, each person once.
        store = student_scraper.open_store(self.context.store_db,
                                        read_only=True)
        self.addCleanup(store.close)
        self.assertEqual(store.count(), 20)

    def test_rerunning_a_saved_row_creates_no_duplicate(self):
        """Even if state lags, the store's UNIQUE index makes a replay safe."""
        first = self._workflow(Options(), fail_at="person8")
        with self.assertRaises(S.Stopped):
            self._run(first)
        # Rewind the state as though it had not been flushed.
        first.state.set_in("processing", next_row_index=0)

        second = self._workflow(Options())
        self._run(second)
        self.assertEqual(second.seen[0], "person1")     # really replayed
        store = student_scraper.open_store(self.context.store_db,
                                        read_only=True)
        self.addCleanup(store.close)
        self.assertEqual(store.count(), 20)

    def test_limit_stops_without_cleanup(self):
        options = Options()
        options.limit = 5
        workflow = self._workflow(options)
        with self.assertRaises(S.Stopped):
            self._run(workflow)
        self.assertTrue(os.path.exists(self.export))
        self.assertEqual(workflow.processed, 5)

    def test_no_stale_temp_files_after_a_completed_keyword(self):
        os.makedirs(self.context.temp_download_dir, exist_ok=True)
        leftover = os.path.join(self.context.temp_download_dir, "person3_resume.pdf")
        with open(leftover, "wb") as handle:
            handle.write(b"%PDF-1.4 leftover")
        with open(self.context.temp_student_file, "w", encoding="utf-8") as handle:
            handle.write("stale")

        workflow = self._workflow(Options())
        self._run(workflow)
        self.assertEqual(student_scraper.stale_student_files(), [])
        self.assertFalse(os.path.exists(leftover))
        self.assertFalse(os.path.exists(self.context.temp_student_file))

    def test_crash_keeps_resume_downloads(self):
        """An unresolved student's downloaded resume must survive."""
        os.makedirs(self.context.temp_download_dir, exist_ok=True)
        pending = os.path.join(self.context.temp_download_dir, "person8_resume.pdf")
        with open(pending, "wb") as handle:
            handle.write(b"%PDF-1.4 pending")
        workflow = self._workflow(Options(), fail_at="person8")
        with self.assertRaises(S.Stopped):
            self._run(workflow)
        self.assertTrue(os.path.exists(pending),
                        "resume of an unresolved student must be kept")

    def test_temporary_file_survives_a_crash(self):
        """temporary.txt must still describe the person being processed."""
        class Failing(StubWorkflow):
            async def do_candidate(self, page, candidate):
                self.seen.append(candidate.username)
                bundle = S.SourceBundle(candidate=candidate)
                bundle.github_ok = True
                bundle.github_text = "half-collected"
                student_scraper.write(bundle)
                raise RuntimeError("crash before the Excel write")

        store = student_scraper.open_store(self.context.store_db)
        self.stores.append(store)
        workflow = Failing(Options(), self.context, self.plan, store)
        with self.assertRaises(S.Stopped):
            self._run(workflow)
        self.assertTrue(os.path.exists(self.context.temp_student_file),
                        "temporary.txt must NOT be cleared on failure")
        with open(self.context.temp_student_file, encoding="utf-8") as handle:
            self.assertIn("person1", handle.read())
        self.assertTrue(os.path.exists(self.export))


if __name__ == "__main__":
    unittest.main(verbosity=2)
