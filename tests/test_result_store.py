"""The store: claims, concurrency, and the limits of a local SQLite backend.

Identity and duplicate rules are covered in tests/test_workflow.py
(TestDuplicateIndex).  This file covers what is specific to the store as a
CONCURRENCY mechanism: the database enforcing uniqueness rather than the code,
claims being leases rather than locks, and the backend being honest about what
it cannot do across machines.
"""

import multiprocessing
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                     # noqa: E402
import student_scraper                                           # noqa: E402


def record(username, email=""):
    row = {column: "" for column in C.FINAL_COLUMNS}
    row["Student Name"] = username.title()
    row["GitHub Profile"] = "https://github.com/%s" % username
    row["github"] = row["GitHub Profile"]
    row["Email Address"] = email
    return row


def _race_worker(path, username, results):
    """A separate PROCESS, so this is a real race, not a simulated one."""
    store = student_scraper.open_store(path)
    saved, reason = store.save_student(record(username))
    store.close()
    results.append((saved, reason))


class StoreCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="store_")
        self.path = os.path.join(self.dir, "students.db")
        self.store = student_scraper.open_store(self.path)

    def tearDown(self):
        self.store.close()
        shutil.rmtree(self.dir, ignore_errors=True)


# ================================================ 1. the database decides ===
class TestUniquenessIsEnforcedByTheDatabase(StoreCase):
    def test_two_processes_racing_produce_exactly_one_row(self):
        """The real race: no SELECT-then-INSERT could survive this."""
        self.store.close()
        with multiprocessing.Manager() as manager:
            results = manager.list()
            workers = [multiprocessing.Process(
                target=_race_worker, args=(self.path, "racer", results))
                for _ in range(6)]
            for worker in workers:
                worker.start()
            for worker in workers:
                worker.join(30)
            outcomes = list(results)
        self.store = student_scraper.open_store(self.path)
        self.assertEqual(self.store.count(), 1)
        self.assertEqual(sum(1 for saved, _reason in outcomes if saved), 1,
                         "exactly one writer may win: %r" % outcomes)

    def test_the_unique_indexes_exist(self):
        names = {row[0] for row in self.store.connection.execute(
            "SELECT name FROM sqlite_master WHERE type='index'")}
        for expected in ("ux_students_username", "ux_students_profile",
                         "ux_students_email"):
            self.assertIn(expected, names)

    def test_a_record_with_no_usable_identity_is_refused(self):
        row = {column: "" for column in C.FINAL_COLUMNS}
        saved, reason = self.store.save_student(row)
        self.assertFalse(saved)
        self.assertIn("identity", reason)
        self.assertEqual(self.store.count(), 0)

    def test_a_read_only_store_refuses_to_write(self):
        self.store.save_student(record("one"))
        read_only = student_scraper.open_store(self.path, read_only=True)
        self.addCleanup(read_only.close)
        self.assertEqual(read_only.count(), 1)
        with self.assertRaises(RuntimeError):
            read_only.save_student(record("two"))
        with self.assertRaises(RuntimeError):
            read_only.claim("two", rank=1)


# ====================================================== 2. claims (leases) ==
class TestClaims(StoreCase):
    def test_one_rank_claims_and_another_sees_it(self):
        self.assertEqual(self.store.claim("student123", rank=4), (True, 4))
        self.assertEqual(self.store.claim("student123", rank=63), (False, 4))

    def test_a_claim_is_case_insensitive_like_the_username(self):
        self.store.claim("Student123", rank=4)
        self.assertEqual(self.store.claim("STUDENT123", rank=63), (False, 4))

    def test_releasing_lets_another_rank_take_over(self):
        self.store.claim("student123", rank=4)
        self.store.release_claim("student123")
        self.assertEqual(self.store.claim("student123", rank=63), (True, 63))

    def test_an_expired_lease_is_taken_over_not_held_forever(self):
        """A crashed rank must not park somebody permanently."""
        self.store.claim("student123", rank=4, lease_seconds=-1)
        self.assertEqual(self.store.claim("student123", rank=63,
                                          lease_seconds=600), (True, 63))

    def test_attempts_accumulate_but_never_blacklist(self):
        for _attempt in range(3):
            self.store.claim("student123", rank=4, lease_seconds=-1)
        self.assertGreaterEqual(self.store.claim_attempts("student123"), 3)
        # Still claimable: attempts are a counter, not a ban.
        self.assertEqual(self.store.claim("student123", rank=9,
                                          lease_seconds=-1)[0], True)

    def test_a_claim_does_not_save_anybody(self):
        self.store.claim("student123", rank=4)
        self.assertEqual(self.store.count(), 0)
        self.assertEqual(self.store.is_known(username="student123"), "")

    def test_an_unusable_username_cannot_be_claimed(self):
        self.assertEqual(self.store.claim("", rank=1), (False, 0))
        self.assertEqual(self.store.claim("settings", rank=1), (False, 0))


# ============================================ 3. honest about the backend ===
class TestBackendHonesty(StoreCase):
    def test_sqlite_does_not_claim_global_deduplication(self):
        self.assertEqual(self.store.backend, "sqlite")
        self.assertFalse(self.store.provides_global_dedupe,
                         "a local file must never claim to be global")

    def test_two_rank_stores_cannot_see_each_other(self):
        """The documented limitation, asserted so it cannot be forgotten."""
        other = student_scraper.open_store(os.path.join(self.dir, "rank63.db"))
        self.addCleanup(other.close)
        self.store.save_student(record("student123"))
        self.assertEqual(self.store.is_known(username="student123"),
                         "username")
        self.assertEqual(other.is_known(username="student123"), "",
                         "rank-local stores are independent by design; "
                         "cross-laptop duplicates are resolved at export")

    def test_a_central_backend_is_refused_clearly_not_half_implemented(self):
        for url in ("postgres://host/db", "postgresql://host/db",
                    "mysql://host/db"):
            with self.assertRaises(NotImplementedError) as caught:
                student_scraper.open_store(url)
            self.assertIn("not implemented", str(caught.exception))

    def test_a_sqlite_url_is_accepted(self):
        store = student_scraper.open_store("sqlite://" + os.path.join(self.dir, "u.db"))
        self.addCleanup(store.close)
        self.assertEqual(store.backend, "sqlite")

    def test_the_interface_is_implemented(self):
        for name in ("is_known", "save_student", "claim", "release_claim",
                     "count", "iter_records", "close"):
            self.assertTrue(callable(getattr(self.store, name)), name)
            self.assertTrue(hasattr(student_scraper.ResultStore, name), name)


# ======================================================== 4. stored shape ===
class TestStoredRecords(StoreCase):
    def test_metadata_is_stored_beside_the_record_not_inside_it(self):
        class Unit:
            country = "India"
            state = "Telangana"
            search_city = "Hyderabad"
            search_query = "Fresher Hyderabad"

        self.store.save_student(record("one"), search_unit=Unit(),
                               campaign_id="campaign_20260927_001", rank=23,
                               total_ranks=80, source_xlsx="fresher.xlsx")
        row = next(self.store.iter_records())
        self.assertEqual(row["campaign_id"], "campaign_20260927_001")
        self.assertEqual(row["rank"], 23)
        self.assertEqual(row["search_city"], "Hyderabad")
        self.assertEqual(row["source_xlsx"], "fresher.xlsx")
        self.assertEqual(sorted(row["record"]), sorted(C.FINAL_COLUMNS))

    def test_the_record_round_trips_all_sixty_five_columns(self):
        row = record("one", email="one@real.com")
        row["Current City"] = "Pune"
        row["Primary Skills"] = "Python, Linux"
        self.store.save_student(row)
        stored = next(self.store.iter_records())["record"]
        self.assertEqual(stored["Current City"], "Pune")
        self.assertEqual(stored["Primary Skills"], "Python, Linux")
        self.assertEqual(len(stored), 65)

    def test_populated_fields_is_counted_for_conflict_resolution(self):
        self.store.save_student(record("sparse"))
        rich = record("rich", email="rich@real.com")
        rich["Current City"] = "Pune"
        rich["State"] = "Maharashtra"
        self.store.save_student(rich)
        rows = {row["github_username_key"]: row["populated_fields"]
                for row in self.store.iter_records()}
        self.assertGreater(rows["rich"], rows["sparse"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
