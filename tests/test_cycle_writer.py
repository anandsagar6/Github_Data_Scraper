"""current_data.xlsx keeps up with the run, and rolls over every 100 rows.

WHAT THIS REPLACES
------------------
Collected people used to be rendered into ONE student_data.xlsx, rebuilt
every EXPORT_EVERY_N_STUDENTS saves.  A copy now appends each person to
current_data.xlsx and seals it into current_data_cycleN.xlsx every
config.CYCLE_SIZE rows, which is the same output contract the job fetching
scripts use.

The property that mattered then still matters now and is tested here:

    the spreadsheet keeps up with the run        (it is written per person,
                                                  not at the end)
    a failed spreadsheet write never loses a person
                                                 (students.db has already
                                                  accepted the row)
    a sealed cycle file is never rewritten       (a collector may copy it
                                                  away at any moment)
"""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper as S                              # noqa: E402


def person(name):
    return {"Student Name": name, "Email Address": "%s@example.com" % name}


class CycleRolloverCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="cycle_")
        self.saved = (C.DETAIL_OUTPUT_FILE, C.CYCLE_STATE_FILE, C.CYCLE_SIZE)
        C.DETAIL_OUTPUT_FILE = os.path.join(self.root, "current_data.xlsx")
        C.CYCLE_STATE_FILE = os.path.join(self.root, "cycle_state.json")
        C.CYCLE_SIZE = 5

    def tearDown(self):
        (C.DETAIL_OUTPUT_FILE, C.CYCLE_STATE_FILE,
         C.CYCLE_SIZE) = self.saved
        shutil.rmtree(self.root, ignore_errors=True)

    def add(self, count, start=1):
        sealed = []
        for index in range(start, start + count):
            out = student_scraper.append_detail_to_excel(
                person("p%d" % index), say=lambda _m: None)
            if out:
                sealed.append(out)
        return sealed

    # -- the spreadsheet keeps up with the run ---------------------------
    def test_each_person_is_written_immediately(self):
        self.add(1)
        self.assertEqual(student_scraper.row_count(), 1)
        self.add(1, start=2)
        self.assertEqual(student_scraper.row_count(), 2)

    def test_no_rollover_before_the_threshold(self):
        self.assertEqual(self.add(4), [])
        self.assertEqual(student_scraper.row_count(), 4)
        self.assertEqual(student_scraper.cycle_files(self.root), [])

    def test_rollover_seals_and_empties(self):
        sealed = self.add(5)
        self.assertEqual(len(sealed), 1)
        self.assertTrue(sealed[0].endswith("current_data_cycle1.xlsx"))
        self.assertEqual(student_scraper.row_count(), 0,
                         "current_data.xlsx must be emptied after a rollover")
        self.assertEqual(len(student_scraper.read_rows(sealed[0])), 5)

    def test_cycles_are_numbered_in_order(self):
        sealed = self.add(12)
        self.assertEqual([os.path.basename(p) for p in sealed],
                         ["current_data_cycle1.xlsx",
                          "current_data_cycle2.xlsx"])
        self.assertEqual(student_scraper.row_count(), 2)   # 12 - 5 - 5

    def test_every_person_survives_the_rollover(self):
        self.add(12)
        names = []
        for path in student_scraper.cycle_files(self.root):
            names += [row["Student Name"] for row in student_scraper.read_rows(path)]
        names += [row["Student Name"] for row in student_scraper.read_rows()]
        self.assertEqual(names, ["p%d" % n for n in range(1, 13)])

    def test_sealed_files_are_never_rewritten(self):
        self.add(5)
        first = os.path.join(self.root, "current_data_cycle1.xlsx")
        before = open(first, "rb").read()
        self.add(10, start=6)
        self.assertEqual(open(first, "rb").read(), before,
                         "a sealed cycle file must never be touched again")

    # -- the cycle number must not restart after a collector took files --
    def test_number_does_not_restart_when_files_are_collected(self):
        self.add(10)
        for path in student_scraper.cycle_files(self.root):
            os.remove(path)                      # a collector took them
        sealed = self.add(5, start=11)
        self.assertEqual(os.path.basename(sealed[0]),
                         "current_data_cycle3.xlsx",
                         "restarting at cycle1 would overwrite a batch that "
                         "was already handed over")

    def test_an_existing_name_is_never_overwritten(self):
        self.add(5)
        os.remove(C.CYCLE_STATE_FILE)            # state lost, file still there
        sealed = self.add(5, start=6)
        self.assertEqual(os.path.basename(sealed[0]),
                         "current_data_cycle2.xlsx")
        self.assertEqual(
            len(student_scraper.read_rows(
                os.path.join(self.root, "current_data_cycle1.xlsx"))), 5)

    # -- damaged input ----------------------------------------------------
    def test_an_unreadable_workbook_is_preserved_not_deleted(self):
        with open(C.DETAIL_OUTPUT_FILE, "w") as handle:
            handle.write("this is not a workbook")
        student_scraper.append_detail_to_excel(person("after"),
                                            say=lambda _m: None)
        spoiled = [n for n in os.listdir(self.root) if ".corrupt_" in n]
        self.assertEqual(len(spoiled), 1,
                         "the unreadable file must be kept, not thrown away")
        self.assertEqual(student_scraper.row_count(), 1)

    def test_all_65_columns_are_written(self):
        self.add(1)
        self.assertEqual(len(student_scraper.read_rows()[0]),
                         len(C.FINAL_COLUMNS))


class FailureNeverLosesAPersonCase(unittest.TestCase):
    """students.db is the source of truth; the spreadsheet is a rendering."""

    def setUp(self):
        self.said = []
        self._say = S.say
        S.say = self.said.append

    def tearDown(self):
        S.say = self._say

    def test_a_failing_write_is_reported_and_does_not_raise(self):
        workflow = S.Workflow.__new__(S.Workflow)
        broken = lambda *_a, **_k: (_ for _ in ()).throw(OSError("disk full"))
        original = student_scraper.append_detail_to_excel
        student_scraper.append_detail_to_excel = broken
        try:
            self.assertEqual(workflow.write_to_cycle_file(person("x")), "")
        finally:
            student_scraper.append_detail_to_excel = original
        self.assertTrue(any("FAILED" in line for line in self.said))
        self.assertTrue(any("IS saved" in line for line in self.said),
                        "the operator must be told the person is not lost")
        # It must name the file that genuinely holds them.  It used to say
        # students.db, which stopped being written when the master workbook
        # replaced it -- an operator following that message would have gone
        # looking for a file that does not exist.
        self.assertTrue(
            any("master_index.xlsx" in line for line in self.said),
            "the message must name master_index.xlsx, the file that actually "
            "holds the person")
        self.assertFalse(any("students.db" in line for line in self.said),
                         "students.db is never written; naming it sends the "
                         "operator to a file that is not there")


if __name__ == "__main__":
    unittest.main(verbosity=2)
