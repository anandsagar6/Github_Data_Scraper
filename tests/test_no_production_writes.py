"""No test may write into the real project folder.

WHY THIS FILE EXISTS
--------------------
Production paths are module-level constants in `config.py` pointing next to
`student_scraper.py`.  That is what makes "one folder is one rank" true, and
it is also a trap for tests: anything that builds a real `Workflow`, or calls
`cycle_writer`, or opens the store without redirecting those constants writes
into the developer's own working folder.

It is not hypothetical.  While this refactor was in progress
`test_require_contact.py` quietly appended 60 people to the project's own
`current_data.xlsx`, because `do_candidate` writes the saved person to the
cycle file and only `SKIPPED_FILE` had been redirected.  Nothing failed; the
rows just appeared in a real output file.

`tests/sandbox.py` is the fix -- `enter()` redirects every one of those
constants -- and this file is the alarm for the next time somebody forgets.
It replaces the per-rank path checks that lived in the deleted
`test_isolation.py`, where isolation was a property of path construction.
Isolation is now a property of the test fixture, so that is what gets tested.
"""

import os
import subprocess
import sys
import unittest

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(TESTS_DIR)
sys.path.insert(0, PROJECT)

import student_scraper                                 # noqa: E402

import config as C                                       # noqa: E402

# Files a run creates in its own folder.  None of them may appear because a
# test ran.  Directories are listed separately: several are created eagerly by
# config.ensure_dirs() and their mere existence is not a leak.
FORBIDDEN_FILES = (
    C.DETAIL_OUTPUT_FILE,
    C.CYCLE_STATE_FILE,
    C.CURRENT_ROW_FILE,
    C.KEYWORD_CYCLE_FILE,
    C.STATE_FILE,
    C.STORE_DB,
    C.DEFERRED_FILE,
    C.TRANSIENT_FILE,
    C.THROTTLE_FILE,
    C.RATE_LIMIT_FILE,
    C.TEMP_STUDENT_FILE,
    C.TEMP_STUDENT_JSON,
    C.SKIPPED_FILE,
    C.FAILED_FILE,
)

# The duplicate index is not a config constant -- student_scraper.master_path()
# derives it from C.FINAL_OUTPUT_DIR on every call -- so it is checked
# separately.  It holds every person ever collected in this folder, which makes
# it the single worst file for a test to rewrite.
MASTER_FILE = student_scraper.master_path()


def cycle_files_in_project():
    stem = os.path.basename(os.path.splitext(C.DETAIL_OUTPUT_FILE)[0]) + "_cycle"
    try:
        return sorted(name for name in os.listdir(C.PROJECT_ROOT)
                      if name.startswith(stem) and name.endswith(".xlsx"))
    except OSError:
        return []


class TestFixtureRedirectsEverything(unittest.TestCase):
    """sandbox.enter() must cover every production path, not most of them."""

    def test_every_forbidden_file_is_redirected_by_the_sandbox(self):
        import tempfile, shutil
        import sandbox
        root = tempfile.mkdtemp(prefix="leakcheck_")
        try:
            sandbox.enter(root)
            for name in ("DETAIL_OUTPUT_FILE", "CYCLE_STATE_FILE",
                         "CURRENT_ROW_FILE", "STATE_FILE", "STORE_DB",
                         "DEFERRED_FILE", "TRANSIENT_FILE", "THROTTLE_FILE",
                         "RATE_LIMIT_FILE", "TEMP_STUDENT_FILE",
                         "TEMP_STUDENT_JSON", "SKIPPED_FILE", "FAILED_FILE",
                         "KEYWORD_DOWNLOAD_DIR", "TEMP_DOWNLOAD_DIR",
                         "CRAWL_STAGING_DIR", "PROFILE_DIR", "LOG_DIR",
                         "RETRY_BACKUP_DIR", "FINAL_OUTPUT_DIR"):
                self.assertTrue(
                    getattr(C, name).startswith(root),
                    "sandbox.enter() does not redirect C.%s; a test using it "
                    "would write into the real project folder" % name)
            self.assertTrue(
                student_scraper.master_path().startswith(root),
                "master_path() still points at the real duplicate index; a "
                "test calling load_existing_master() or save_master() without "
                "an explicit path would rewrite every person collected here")
        finally:
            sandbox.restore()
            shutil.rmtree(root, ignore_errors=True)

    def test_restore_puts_every_path_back(self):
        import tempfile, shutil
        import sandbox
        before = {name: getattr(C, name) for name in dir(C)
                  if name.isupper() and isinstance(getattr(C, name), str)}
        root = tempfile.mkdtemp(prefix="leakcheck_")
        try:
            sandbox.enter(root)
            sandbox.restore()
        finally:
            shutil.rmtree(root, ignore_errors=True)
        after = {name: getattr(C, name) for name in dir(C)
                 if name.isupper() and isinstance(getattr(C, name), str)}
        self.assertEqual(after, before,
                         "sandbox.restore() left config pointing elsewhere")


class TestTheSuiteLeavesTheProjectAlone(unittest.TestCase):
    """Run every other test file and prove none of them wrote here.

    Skipped unless IDS_LEAK_CHECK=1, because it runs the whole suite a second
    time and that takes minutes.  Run it before committing a change to the
    fixture:

        IDS_LEAK_CHECK=1 python3 tests/test_no_production_writes.py
    """

    @unittest.skipUnless(os.environ.get("IDS_LEAK_CHECK") == "1",
                         "set IDS_LEAK_CHECK=1 to run the full sweep")
    def test_no_test_file_writes_into_the_project_folder(self):
        watched = FORBIDDEN_FILES + (MASTER_FILE,)
        present_before = {path for path in watched if os.path.exists(path)}
        master_before = (open(MASTER_FILE, "rb").read()
                         if os.path.exists(MASTER_FILE) else None)
        cycles_before = set(cycle_files_in_project())
        offenders = []
        for name in sorted(os.listdir(TESTS_DIR)):
            if not name.startswith("test_") or not name.endswith(".py"):
                continue
            if name == os.path.basename(__file__):
                continue
            subprocess.run([sys.executable, os.path.join(TESTS_DIR, name)],
                           capture_output=True, text=True, timeout=600)
            created = [path for path in watched
                       if os.path.exists(path) and path not in present_before]
            new_cycles = set(cycle_files_in_project()) - cycles_before
            if master_before is not None and os.path.exists(MASTER_FILE) \
               and open(MASTER_FILE, "rb").read() != master_before:
                # Rewritten in place rather than created: openpyxl output is
                # not byte-stable, so this can only be trusted as "something
                # saved the master", which is already a failure.
                created.append(MASTER_FILE)
            if created or new_cycles:
                offenders.append((name, [os.path.basename(p) for p in created]
                                  + sorted(new_cycles)))
                for path in created:
                    if path == MASTER_FILE:
                        # Never delete the duplicate index to tidy up: put the
                        # bytes back.
                        with open(path, "wb") as handle:
                            handle.write(master_before)
                        continue
                    os.remove(path)
                for leaked in new_cycles:
                    os.remove(os.path.join(C.PROJECT_ROOT, leaked))
        self.assertEqual(offenders, [],
                         "these test files wrote into the real project "
                         "folder; they must use tests/sandbox.py")


if __name__ == "__main__":
    unittest.main(verbosity=2)
