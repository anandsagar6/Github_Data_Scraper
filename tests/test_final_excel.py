"""One visible workbook, refreshed when a rank completes -- SQLite stays truth.

The rank-local students.db files remain the internal source of truth (resume,
duplicate protection, crash safety).  When a rank reaches COMPLETE the ONE
visible Student Data original/student_data.xlsx is regenerated from every local
rank store, globally de-duplicated, using the existing 65-column export written
atomically.  A failed export never corrupts the previous workbook and never
affects the rank's saved completion.
"""

import hashlib
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                         # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper as export_students                                     # noqa: E402
from openpyxl import load_workbook                          # noqa: E402


def _sha(path):
    with open(path, "rb") as handle:
        return hashlib.sha256(handle.read()).hexdigest()


def _read_header(path):
    book = load_workbook(path, read_only=True)
    header = [c.value for c in next(book.active.iter_rows(max_row=1))]
    book.close()
    return header


def _data_rows(path):
    book = load_workbook(path, read_only=True, data_only=True)
    rows = list(book.active.iter_rows(min_row=2, values_only=True))
    book.close()
    return [r for r in rows if any(v not in (None, "") for v in r)]


class FinalExcelCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="finalxl_")
        self.runtime = os.path.join(self.tmp, "runtime")
        self.out_dir = os.path.join(self.tmp, "Student Data original")
        os.makedirs(self.runtime)
        os.makedirs(self.out_dir)
        self.out_path = os.path.join(self.out_dir, "student_data.xlsx")
        # Point the export at the isolated temp dirs; restore in tearDown.
        self._saved = {k: getattr(C, k) for k in
                       ("LEGACY_IMPORT_DB",
                        "FINAL_OUTPUT_DIR", "FINAL_WORKBOOK_PATH")}
        C.LEGACY_IMPORT_DB = os.path.join(self.tmp, "nonexistent.db")
        C.FINAL_OUTPUT_DIR = self.out_dir
        C.FINAL_WORKBOOK_PATH = self.out_path
        self.store_paths = []

    def tearDown(self):
        for key, value in self._saved.items():
            setattr(C, key, value)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def build_rank(self, rank, total_ranks, students):
        """One copy's store.  students: [(github_url, email)].

        Its path is remembered so refresh() can name it: discover_stores no
        longer scans for stores, by design.
        """
        directory = os.path.join(self.runtime, "rank_%03d" % rank)
        os.makedirs(directory, exist_ok=True)
        path = os.path.join(directory, "students.db")
        self.store_paths.append(path)
        store = student_scraper.open_store(path)
        for url, email in students:
            record = {col: "" for col in C.FINAL_COLUMNS}
            record["GitHub Profile"] = url
            record["Email Address"] = email
            store.save_student(record, rank=rank, total_ranks=total_ranks)
        store.close()
        return path

    def refresh(self):
        printed = []
        ok = student_scraper.refresh_final_excel(
            say_fn=printed.append,
            store_db=os.path.join(self.runtime, "nonexistent.db"),
            extra=list(self.store_paths))
        return ok, "\n".join(printed)

    # 1. rank databases combine into ONE Excel
    def test_rank_dbs_combine_into_one_excel(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.build_rank(2, 80, [("https://github.com/bob", "b@x.com")])
        ok, out = self.refresh()
        self.assertTrue(ok, out)
        self.assertTrue(os.path.exists(self.out_path))
        self.assertEqual(len(_data_rows(self.out_path)), 2)
        self.assertIn("[EXCEL] Updated:", out)
        self.assertIn("[EXCEL] Students: 2", out)

    # 2. exactly 65 columns
    def test_exactly_65_columns(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        ok, out = self.refresh()
        self.assertTrue(ok, out)
        self.assertEqual(len(_read_header(self.out_path)), 65)
        self.assertIn("[EXCEL] Columns: 65", out)

    # 3. correct column order (the exact FINAL_COLUMNS contract)
    def test_correct_column_order(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.refresh()
        self.assertEqual(_read_header(self.out_path), list(C.FINAL_COLUMNS))

    # 4. duplicate GitHub usernames across ranks collapse to one row
    def test_duplicate_usernames_removed(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.build_rank(2, 80, [("https://github.com/Alice", "")])  # same user
        self.refresh()
        self.assertEqual(len(_data_rows(self.out_path)), 1)

    # 5. duplicate normalized emails collapse unrelated usernames
    def test_duplicate_emails_removed(self):
        self.build_rank(1, 80, [("https://github.com/alice", "shared@x.com")])
        self.build_rank(2, 80, [("https://github.com/bob", "SHARED@x.com")])
        self.refresh()
        self.assertEqual(len(_data_rows(self.out_path)), 1)

    # 6. blank emails must NOT collapse unrelated students
    def test_blank_emails_do_not_collapse(self):
        self.build_rank(1, 80, [("https://github.com/alice", "")])
        self.build_rank(2, 80, [("https://github.com/bob", "")])
        self.refresh()
        self.assertEqual(len(_data_rows(self.out_path)), 2)

    # 7. a completed-rank refresh writes the workbook and reports it
    def test_refresh_writes_and_reports(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com"),
                                ("https://github.com/bob", "b@x.com")])
        ok, out = self.refresh()
        self.assertTrue(ok)
        self.assertIn("[EXCEL] Updated: %s" % self.out_path, out)
        self.assertIn("[EXCEL] Students: 2", out)
        self.assertIn("[EXCEL] Columns: 65", out)

    # 8. a failed refresh does NOT corrupt the existing good workbook
    def test_failed_refresh_preserves_previous(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.assertTrue(self.refresh()[0])
        good = _sha(self.out_path)
        # now make the export blow up mid-way
        original = student_scraper.export

        def boom(*_a, **_k):
            raise RuntimeError("disk full")
        student_scraper.export = boom
        try:
            ok, out = self.refresh()
        finally:
            student_scraper.export = original
        self.assertFalse(ok)
        self.assertIn("previous student_data.xlsx preserved", out)
        self.assertEqual(_sha(self.out_path), good)         # byte-identical

    # 9. a locked workbook is reported and the previous file is preserved
    def test_locked_workbook_preserves_previous(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.assertTrue(self.refresh()[0])
        good = _sha(self.out_path)
        original = student_scraper.export

        def locked(*_a, **_k):
            raise student_scraper.WorkbookLocked("open in LibreOffice")
        student_scraper.export = locked
        try:
            ok, out = self.refresh()
        finally:
            student_scraper.export = original
        self.assertFalse(ok)
        self.assertIn("previous student_data.xlsx preserved", out)
        self.assertEqual(_sha(self.out_path), good)

    # 10. atomic write: an existing file is fully replaced, no temp left behind
    def test_atomic_replacement_no_temp_left(self):
        with open(self.out_path, "w") as handle:
            handle.write("garbage that must be replaced")
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.assertTrue(self.refresh()[0])
        # a valid workbook now, not the garbage
        self.assertEqual(len(_read_header(self.out_path)), 65)
        leftovers = [n for n in os.listdir(self.out_dir)
                     if n.endswith(".tmp")]
        self.assertEqual(leftovers, [], "no temp file should remain")

    # 11. only ONE final workbook -- never per-rank xlsx files
    def test_only_one_final_workbook(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        self.build_rank(2, 80, [("https://github.com/bob", "b@x.com")])
        self.refresh()
        xlsx = [n for n in os.listdir(self.out_dir) if n.endswith(".xlsx")]
        self.assertEqual(xlsx, ["student_data.xlsx"])

    # 12. the SQLite stores are never modified by an export
    def test_sqlite_stores_untouched(self):
        c1 = self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        c2 = self.build_rank(2, 80, [("https://github.com/bob", "b@x.com")])
        before = {p: _sha(p) for p in (c1, c2)}
        self.refresh()
        for path, digest in before.items():
            self.assertEqual(_sha(path), digest, path)

    # 13. the existing export_students CLI still works
    def test_export_students_cli_still_works(self):
        path = self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        # Stores are named explicitly now; the CLI no longer scans for them.
        code = student_scraper.mode_export_final(["--store=%s" % path])
        self.assertEqual(code, 0)
        self.assertTrue(os.path.exists(self.out_path))
        self.assertEqual(len(_read_header(self.out_path)), 65)

    # 14. the master merge across devices (bundle stores) still works
    def test_master_bundle_merge_still_works(self):
        # two "device" stores merged via the bundle path export() supports
        d1 = self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        d2 = self.build_rank(11, 80, [("https://github.com/carol", "c@x.com")])
        report = student_scraper.export(
            self.out_path,
            bundle_stores=[("dev1/rank_001", d1),
                           ("dev2/rank_011", d2)],
            include_imported=False, out=lambda *_a, **_k: None)
        self.assertEqual(report["unique"], 2)
        self.assertEqual(len(_data_rows(self.out_path)), 2)

    # 15. no runtime / state / browser files are modified by an export
    def test_runtime_tree_untouched_by_export(self):
        self.build_rank(1, 80, [("https://github.com/alice", "a@x.com")])
        before = _tree_snapshot(self.runtime)
        self.refresh()
        self.assertEqual(_tree_snapshot(self.runtime), before)


def _tree_snapshot(root):
    # Exclude SQLite's transient journal sidecars (-wal/-shm/-journal): opening
    # a store read-only can create/checkpoint them, which is not a modification
    # of the student DATA or of any state/profile file.  The .db itself is
    # verified byte-for-byte by test_sqlite_stores_untouched.
    snap = {}
    for base, _dirs, files in os.walk(root):
        for name in files:
            if name.endswith(("-wal", "-shm", "-journal")):
                continue
            path = os.path.join(base, name)
            try:
                snap[path] = (os.path.getsize(path), _sha(path))
            except OSError:
                pass
    return snap


if __name__ == "__main__":
    unittest.main(verbosity=2)
