"""CONTACTABLE-ONLY collection (config.REQUIRE_CONTACT).

A profile is written to the store when ANY contact route was found -- an email
OR a phone number.  Someone with neither is fetched, reported as skipped, and
NOT saved: a row nobody can be reached through is not a lead.
The profile fetch / portfolio / resume / merge collaborators are stubbed so the
test exercises only the save/skip DECISION, with no browser and no network.
"""

import asyncio
import os
import sys
import tempfile
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                         # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (Candidate)   # noqa: E402
import student_scraper as S                                # noqa: E402


class FakeStore:
    def __init__(self):
        self.saved = []
        self.released = []

    def is_known(self, username="", profile_url="", email=""):
        return None

    def claim(self, username, rank):
        return True, rank

    def release_claim(self, username):
        self.released.append(username)

    def save_student(self, record, **kwargs):
        self.saved.append(record)
        return True, "saved"


class FakeProfiles:
    async def fetch(self, page, username):
        # no website, no links, no bio text -> portfolio and resume are skipped
        return {"links": [], "website": "", "email": ""}, ""


class FakeMerger:
    def __init__(self, record):
        self.record = record

    def build(self, bundle):
        return dict(self.record)


class RequireContactCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="reqmail_")
        self._pace = S.pace
        self._tw = student_scraper.write
        self._tc = student_scraper.clear
        S.pace = self._noop
        student_scraper.write = lambda *_a, **_k: None
        student_scraper.clear = lambda *_a, **_k: None
        self._require = C.REQUIRE_CONTACT
        # do_candidate writes the saved person to current_data.xlsx, so the
        # cycle writer must point at the temp dir too -- otherwise this test
        # appends real rows to the project's own output file.
        self._cycle = (C.DETAIL_OUTPUT_FILE, C.CYCLE_STATE_FILE)
        C.DETAIL_OUTPUT_FILE = os.path.join(self.tmp, "current_data.xlsx")
        C.CYCLE_STATE_FILE = os.path.join(self.tmp, "cycle_state.json")

    def tearDown(self):
        C.DETAIL_OUTPUT_FILE, C.CYCLE_STATE_FILE = self._cycle
        if getattr(self, '_skipped', None):
            C.SKIPPED_FILE = self._skipped
        S.pace = self._pace
        student_scraper.write = self._tw
        student_scraper.clear = self._tc
        C.REQUIRE_CONTACT = self._require

    async def _noop(self, *_a, **_k):
        return None

    def workflow(self, record):
        store = FakeStore()
        wf = S.Workflow.__new__(S.Workflow)
        wf.store = store
        wf.existing = {}
        wf._master_dirty = False
        self._skipped = C.SKIPPED_FILE
        C.SKIPPED_FILE = os.path.join(self.tmp, "skipped.txt")
        wf.profiles = FakeProfiles()
        wf.throttle = types.SimpleNamespace(record_success=lambda: None)
        wf.merger = FakeMerger(record)
        wf.transient = types.SimpleNamespace(clear=lambda _u: None)
        wf.resumes = types.SimpleNamespace(discard=lambda _f: None)
        wf.unit = types.SimpleNamespace(search_query="q", slug="q")
        wf.processed = 0
        return wf, store

    def run_candidate(self, wf):
        cand = Candidate(username="someone", profile_url="",
                         source_file="kw.xlsx")
        asyncio.run(wf.do_candidate(object(), cand))

    # neither email nor phone -> not saved, claim released, reported as a skip
    def test_no_contact_at_all_is_not_saved(self):
        C.REQUIRE_CONTACT = True
        wf, store = self.workflow({"Email Address": "", "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(wf.existing, {})              # nothing written
        self.assertEqual(wf.processed, 0)              # not counted
        self.assertEqual(wf.existing, {})              # nothing recorded
        with open(C.SKIPPED_FILE, encoding="utf-8") as handle:
            self.assertIn("no email or phone", handle.read())

    # a valid email -> saved as before
    def test_with_email_saved(self):
        C.REQUIRE_CONTACT = True
        wf, store = self.workflow(
            {"Email Address": "real.person@gmail.com", "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(len(wf.existing), 1)
        self.assertEqual(next(iter(wf.existing.values()))["Email Address"],
                         "real.person@gmail.com")
        self.assertEqual(wf.processed, 1)

    # whitespace-only values count as no contact at all
    def test_blank_contact_not_saved(self):
        C.REQUIRE_CONTACT = True
        wf, store = self.workflow({"Email Address": "   ",
                                   "Phone Number": "  ",
                                   "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(wf.existing, {})

    # PHONE ONLY is still a lead -> saved
    def test_phone_only_is_saved(self):
        C.REQUIRE_CONTACT = True
        wf, store = self.workflow({"Email Address": "",
                                   "Phone Number": "+91 98765 43210",
                                   "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(len(wf.existing), 1)
        self.assertEqual(next(iter(wf.existing.values()))["Phone Number"], "+91 98765 43210")
        self.assertEqual(wf.processed, 1)

    # email only is still a lead -> saved
    def test_email_only_is_saved(self):
        C.REQUIRE_CONTACT = True
        wf, store = self.workflow({"Email Address": "a.person@gmail.com",
                                   "Phone Number": "",
                                   "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(len(wf.existing), 1)

    # both present -> saved
    def test_both_contacts_saved(self):
        C.REQUIRE_CONTACT = True
        wf, store = self.workflow({"Email Address": "a@b.com",
                                   "Phone Number": "9876543210",
                                   "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(len(wf.existing), 1)

    # the helper itself, over the configured contact fields
    def test_contact_fields_are_configurable(self):
        self.assertEqual(C.CONTACT_FIELDS,
                         ("Email Address", "Phone Number"))
        found = S.Workflow._contact_found
        self.assertFalse(found({}))
        self.assertFalse(found({"Email Address": "", "Phone Number": ""}))
        self.assertTrue(found({"Email Address": "x@y.com"}))
        self.assertTrue(found({"Phone Number": "9876543210"}))

    # flag off -> old behaviour, a person with no contact is still saved
    def test_flag_off_saves_without_contact(self):
        C.REQUIRE_CONTACT = False
        wf, store = self.workflow({"Email Address": "", "Phone Number": "",
                                   "Student Name": "X"})
        self.run_candidate(wf)
        self.assertEqual(len(wf.existing), 1)
        self.assertEqual(wf.processed, 1)

    def test_flag_default_on(self):
        self.assertTrue(self._require)                 # ships enabled


if __name__ == "__main__":
    unittest.main(verbosity=2)
