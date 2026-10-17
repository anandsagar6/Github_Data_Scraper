"""The final workbook: ONE file, no rotation, written only by the export.

Replaces tests/test_rotation.py, which asserted the 100-data-rows-per-file
behaviour that has been deliberately removed.  Coverage worth keeping from it
is carried over here: exact header, contiguous rows, no overwrite of an
existing workbook on failure, alias mirroring, validation verdicts and
duplicate-free legacy migration.
"""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook, load_workbook          # noqa: E402

import config as C                                    # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper                                 # noqa: E402
import student_scraper as X                           # noqa: E402


def record(username, name="", email="", city="", extra=0):
    row = {column: "" for column in C.FINAL_COLUMNS}
    row["Student Name"] = name or username
    row["GitHub Profile"] = "https://github.com/%s" % username
    row["Email Address"] = email
    row["Current City"] = city
    for offset in range(extra):
        row[C.FINAL_COLUMNS[20 + offset]] = "filled"
    return student_scraper.alias_fill(row)


def read(path):
    book = load_workbook(path)
    sheet = book.active
    header = [cell.value for cell in sheet[1]]
    rows = [[cell.value for cell in row] for row in sheet.iter_rows(min_row=2)]
    book.close()
    return header, rows


class ExportCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="export_")
        self.runtime = os.path.join(self.root, "runtime")
        self.archive = os.path.join(self.runtime, "archive")
        self.output = os.path.join(self.root, "out")
        os.makedirs(self.runtime)
        os.makedirs(self.output)
        self.workbook = os.path.join(self.output, "student_data.xlsx")
        self.stores = []
        self.store_paths = []

    def tearDown(self):
        for store in self.stores:
            store.close()
        shutil.rmtree(self.root, ignore_errors=True)

    def store_for(self, rank, campaign=None):
        """A store for one copy.  Its path is remembered so export() can name
        it: discover_stores no longer scans, by design -- a merge that
        silently leaves a laptop out looks exactly like one that worked."""
        directory = (os.path.join(self.runtime, "rank_%03d" % rank)
                     if campaign is None
                     else os.path.join(self.archive, campaign,
                                       "rank_%03d" % rank))
        os.makedirs(directory, exist_ok=True)
        path = os.path.join(directory, "students.db")
        store = student_scraper.open_store(path)
        self.stores.append(store)
        self.store_paths.append(path)
        return store

    def export(self):
        for store in self.stores:
            store.commit() if hasattr(store, "commit") else None
        return X.export(self.workbook,
                        store_db=os.path.join(self.runtime, "nonexistent.db"),
                        extra=list(self.store_paths),
                        include_imported=False, out=lambda *_a: None)


# ======================================================== no row rotation ===
class TestNoRotation(ExportCase):
    def test_two_hundred_fifty_records_land_in_one_workbook(self):
        """The old limit was 100 rows per file.  There is no limit now."""
        store = self.store_for(1)
        for number in range(250):
            store.save_student(record("person%d" % number), rank=1)
        report = self.export()
        self.assertEqual(report["written"], 250)
        _header, rows = read(self.workbook)
        self.assertEqual(len(rows), 250)
        self.assertEqual(
            [name for name in os.listdir(self.output)
             if name.endswith(".xlsx")], ["student_data.xlsx"],
            "exactly one workbook, no student_data_2.xlsx")

    def test_rows_are_contiguous_with_no_drift(self):
        store = self.store_for(1)
        for number in range(30):
            store.save_student(record("person%d" % number), rank=1)
        self.export()
        book = load_workbook(self.workbook)
        sheet = book.active
        self.assertEqual(sheet.max_row, 31)          # header + 30
        for position in range(2, 32):
            self.assertTrue(sheet.cell(row=position, column=1).value,
                            "row %d is blank" % position)
        book.close()

    def test_rows_per_file_is_gone_from_the_codebase(self):
        """No half-dead rotation switch left behind."""
        self.assertFalse(hasattr(C, "ROWS_PER_FILE"))
        self.assertFalse(hasattr(student_scraper, "ExcelWriter"))
        self.assertFalse(hasattr(student_scraper, "RotationError"))
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        for name in ("config.py", "student_scraper.py"):
            with open(os.path.join(root, name), encoding="utf-8") as handle:
                body = "".join(line for line in handle
                               if not line.lstrip().startswith("#"))
            self.assertNotIn("ROWS_PER_FILE", body,
                             "%s still references ROWS_PER_FILE" % name)


# ============================================================ the contract ==
class TestSchemaContract(ExportCase):
    def test_header_is_the_exact_sixty_five_columns_in_order(self):
        store = self.store_for(1)
        store.save_student(record("one"), rank=1)
        self.export()
        header, _rows = read(self.workbook)
        self.assertEqual(len(header), 65)
        self.assertEqual(header, C.FINAL_COLUMNS)

    def test_alias_columns_mirror_their_canonical_columns(self):
        store = self.store_for(1)
        row = record("one", name="A Person", email="a@real.com", city="Pune")
        row["Country"] = "India"
        row["Total Experience"] = "1 year"
        row["LinkedIn Profile"] = "https://linkedin.com/in/a"
        store.save_student(row, rank=1)
        self.export()
        header, rows = read(self.workbook)
        index = {name: position for position, name in enumerate(header)}
        for alias, canonical in C.ALIAS_COLUMNS.items():
            self.assertEqual(rows[0][index[alias]],
                             rows[0][index[canonical]],
                             "%s must mirror %s" % (alias, canonical))

    def test_validate_output_passes_a_freshly_written_workbook(self):
        store = self.store_for(1)
        for number in range(5):
            store.save_student(record("person%d" % number), rank=1)
        self.export()
        ok, problems, summary = student_scraper.validate_output(self.workbook)
        self.assertTrue(ok, problems)
        self.assertEqual(summary, [("student_data.xlsx", 65, 5, "PASS")])

    def test_validate_output_reports_a_missing_workbook_without_failing(self):
        ok, problems, summary = student_scraper.validate_output(self.workbook)
        self.assertTrue(ok, problems)
        self.assertEqual(summary, [])

    def test_validate_output_names_a_legacy_workbook(self):
        book = Workbook()
        book.active.append(student_scraper.LEGACY_COLUMNS)
        book.save(os.path.join(self.output, "student_data_1.xlsx"))
        book.close()
        ok, problems, _summary = student_scraper.validate_output(self.workbook)
        self.assertFalse(ok)
        self.assertTrue(any("student_data_1.xlsx" in problem
                            for problem in problems), problems)


# ============================================================ atomic write ==
class TestAtomicWrite(ExportCase):
    def test_existing_workbook_survives_a_failed_write(self):
        store = self.store_for(1)
        store.save_student(record("keeper", name="Keeper"), rank=1)
        self.export()
        before = open(self.workbook, "rb").read()

        class Exploding(dict):
            def get(self, *_a, **_kw):
                raise RuntimeError("boom mid-write")

        with self.assertRaises(RuntimeError):
            student_scraper.write_final([Exploding()], self.workbook)
        self.assertEqual(open(self.workbook, "rb").read(), before,
                         "a failed write must not touch the real workbook")

    def test_no_temp_file_is_left_behind_on_success(self):
        store = self.store_for(1)
        store.save_student(record("one"), rank=1)
        self.export()
        leftovers = [name for name in os.listdir(self.output)
                     if name.endswith(".tmp")]
        self.assertEqual(leftovers, [])

    def test_row_limit_is_refused_not_truncated(self):
        saved = C.XLSX_MAX_DATA_ROWS
        C.XLSX_MAX_DATA_ROWS = 3
        try:
            with self.assertRaises(student_scraper.RowLimitExceeded):
                student_scraper.write_final([record("p%d" % n) for n in range(4)],
                              self.workbook)
            self.assertFalse(os.path.exists(self.workbook),
                             "nothing may be written when the limit is hit")
        finally:
            C.XLSX_MAX_DATA_ROWS = saved


# ================================================== LibreOffice lock safety ==
class TestLocks(ExportCase):
    def _lock(self, name):
        path = os.path.join(self.output, ".~lock.%s#" % name)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(",anand,host,27.09.2026 10:00,file:///x;")
        return path

    def test_active_lock_refuses_the_export(self):
        store = self.store_for(1)
        store.save_student(record("one"), rank=1)
        student_scraper.write_final([record("one")], self.workbook)
        lock = self._lock("student_data.xlsx")
        with self.assertRaises(student_scraper.WorkbookLocked):
            self.export()
        self.assertTrue(os.path.exists(lock),
                        "the lock file must NEVER be removed automatically")

    def test_stale_lock_is_reported_but_does_not_block(self):
        store = self.store_for(1)
        store.save_student(record("one"), rank=1)
        lock = self._lock("student_data.xlsx")     # workbook absent -> stale
        messages = []
        X.export(self.workbook,
                 store_db=os.path.join(self.runtime, "nonexistent.db"),
                 extra=list(self.store_paths),
                 include_imported=False, out=messages.append)
        self.assertTrue(any("stale" in line for line in messages), messages)
        self.assertTrue(os.path.exists(lock), "a stale lock is not deleted")
        self.assertTrue(os.path.exists(self.workbook))


# ======================================================== global de-dupe ====
class TestGlobalDeduplication(ExportCase):
    def test_same_person_in_two_rank_stores_collapses_to_one_row(self):
        self.store_for(1).save_student(record("alice"), rank=1)
        self.store_for(2).save_student(record("ALICE"), rank=2)
        report = self.export()
        self.assertEqual(report["unique"], 1)
        self.assertEqual(report["duplicates"], 1)
        _header, rows = read(self.workbook)
        self.assertEqual(len(rows), 1)

    def test_archived_campaign_stores_are_merged_too(self):
        self.store_for(1).save_student(record("current"), rank=1)
        self.store_for(1, campaign="campaign_20260101_001").save_student(
            record("older"), rank=1)
        report = self.export()
        self.assertEqual(report["unique"], 2)

    def test_the_more_complete_record_wins(self):
        self.store_for(1).save_student(record("alice", name="Sparse"), rank=1)
        self.store_for(2).save_student(
            record("alice", name="Complete", email="alice@real.com",
                   city="Pune", extra=6), rank=2)
        self.export()
        header, rows = read(self.workbook)
        index = {name: position for position, name in enumerate(header)}
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][index["Student Name"]], "Complete")
        self.assertEqual(rows[0][index["Current City"]], "Pune")

    def test_a_shared_real_email_collapses_two_usernames(self):
        self.store_for(1).save_student(
            record("alice", email="same@real.com"), rank=1)
        self.store_for(2).save_student(
            record("alice-alt", email="same@real.com"), rank=2)
        report = self.export()
        self.assertEqual(report["unique"], 1)

    def test_a_placeholder_email_does_not_collapse_anybody(self):
        self.store_for(1).save_student(
            record("one", email="test@test.com"), rank=1)
        self.store_for(2).save_student(
            record("two", email="your_email@example.com"), rank=2)
        report = self.export()
        self.assertEqual(report["unique"], 2)

    def test_the_export_is_deterministic(self):
        self.store_for(1).save_student(record("bravo"), rank=1)
        self.store_for(2).save_student(record("alpha"), rank=2)
        self.export()
        _header, first = read(self.workbook)
        second_path = os.path.join(self.output, "again.xlsx")
        X.export(second_path,
                 store_db=os.path.join(self.runtime, "nonexistent.db"),
                 extra=list(self.store_paths),
                 include_imported=False, out=lambda *_a: None)
        _header, second = read(second_path)
        self.assertEqual(first, second)


# ======================================================= legacy migration ===
class TestLegacyImport(ExportCase):
    def _current_schema_workbook(self, name, usernames):
        book = Workbook()
        sheet = book.active
        sheet.append(C.FINAL_COLUMNS)
        for username in usernames:
            row = record(username, name=username.title())
            sheet.append([row[column] for column in C.FINAL_COLUMNS])
        path = os.path.join(self.output, name)
        book.save(path)
        book.close()
        return path

    def test_current_schema_workbooks_are_read_back_faithfully(self):
        path = self._current_schema_workbook("student_data_1.xlsx",
                                             ["alpha", "bravo"])
        records, problems = student_scraper.read_workbook_records(path)
        self.assertEqual(problems, [])
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Student Name"], "Alpha")
        self.assertEqual(records[0]["GitHub Profile"],
                         "https://github.com/alpha")
        for alias, canonical in C.ALIAS_COLUMNS.items():
            self.assertEqual(records[0][alias], records[0][canonical])

    def test_legacy_sixteen_column_rows_are_converted(self):
        book = Workbook()
        sheet = book.active
        sheet.append(student_scraper.LEGACY_COLUMNS)
        sheet.append([1, "B.Tech Student", "oldperson", "Old Person",
                      "old@real.com", "9876543210", "Pune",
                      "B.Tech CSE student", "", "https://github.com/oldperson"])
        path = os.path.join(self.output, "student_data_1.xlsx")
        book.save(path)
        book.close()
        records, problems = student_scraper.read_workbook_records(path)
        self.assertEqual(problems, [])
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["GitHub Profile"],
                         "https://github.com/oldperson")
        self.assertEqual(records[0]["Student Name"], "Old Person")
        self.assertIn("Migrated", records[0]["Remarks / Notes"])

    def test_importing_the_same_workbook_twice_creates_no_duplicate(self):
        path = self._current_schema_workbook("student_data_1.xlsx",
                                             ["alpha", "bravo"])
        store = student_scraper.open_store(os.path.join(self.root, "import.db"))
        self.stores.append(store)
        for _pass in range(2):
            for row in student_scraper.read_workbook_records(path)[0]:
                store.save_student(row, campaign_id="legacy-import")
        self.assertEqual(store.count(), 2)

    def test_the_source_workbooks_are_never_modified(self):
        path = self._current_schema_workbook("student_data_1.xlsx",
                                             ["alpha"])
        before = open(path, "rb").read()
        student_scraper.read_workbook_records(path)
        student_scraper.legacy_workbooks(self.output)
        self.assertEqual(open(path, "rb").read(), before)
        self.assertTrue(os.path.exists(path))


if __name__ == "__main__":
    unittest.main(verbosity=2)
