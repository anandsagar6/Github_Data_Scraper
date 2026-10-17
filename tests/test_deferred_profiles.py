"""A repeatedly-unavailable profile must not block the whole rank.

After MAX_DEFERRED_FAILURES separate transient failures a profile is moved to
the rank's deferred queue (runtime/rank_NNN/student_scraper.json) and the
rank keeps going.  It is retried at end-of-rank and by --retry-deferred, never
blacklisted, never counted as collected, never deleted.  A rank that finishes
all its searches with parked profiles is COMPLETE_WITH_DEFERRED, which lets the
strict runner advance; STOPPED/FAILED still stop the device.
"""

import asyncio
import json
import os
import shutil
import sys
import tempfile
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                         # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (Candidate, DeferredQueue, StateManager, ThrottleCircuit, TransientFailureLog, candidate_from)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                                # noqa: E402
from student_scraper import TransientProfileFailure        # noqa: E402


def cand(user, row=1):
    return Candidate(source_row_number=row, source_file="kw.xlsx",
                     keyword="University Student Guntur", username=user,
                     profile_url="https://github.com/%s" % user)


class FakeReader:
    def __init__(self, candidates):
        self._candidates = candidates

    def read(self, path, keyword):
        return list(self._candidates)


class FakeUnit:
    search_query = "University Student Guntur"
    state = "Andhra Pradesh"
    canonical_city = "Guntur"
    search_city = "Guntur"
    base_keyword = "University Student"


# ============================== the DeferredQueue itself ====================
class DeferredQueueCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="defq_")
        self.ctx = sandbox.enter(self.tmp)

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_add_persists_and_survives_restart(self):          # test 7
        q = DeferredQueue()
        q.add(cand("vu241fa04074-cpu"), "rate limit", 3)
        self.assertTrue(q.has("vu241fa04074-cpu"))
        # a fresh instance re-reads the file
        again = DeferredQueue()
        self.assertTrue(again.has("vu241fa04074-cpu"))
        self.assertEqual(again.count(), 1)
        entry = again.all_entries()[0]
        self.assertEqual(entry["github_url"],
                         "https://github.com/vu241fa04074-cpu")
        self.assertEqual(entry["failures"], 3)

    def test_remove_and_candidate_roundtrip(self):
        q = DeferredQueue()
        q.add(cand("alice", row=6), "rate limit", 4)
        rebuilt = candidate_from(q.all_entries()[0])
        self.assertEqual(rebuilt.username, "alice")
        self.assertEqual(rebuilt.profile_url, "https://github.com/alice")
        self.assertEqual(rebuilt.source_row_number, 6)
        q.remove("alice")
        self.assertEqual(q.count(), 0)
        self.assertFalse(DeferredQueue().has("alice"))


# =================== the profile-phase deferral decision ====================
class ProfilePhaseCase(unittest.TestCase):
    def setUp(self):
        import student_scraper as ratelimit
        self.tmp = tempfile.mkdtemp(prefix="defphase_")
        self.ctx = sandbox.enter(self.tmp)
        self.store = student_scraper.open_store(self.ctx.store_db)
        self._paced = S.pace
        S.pace = self._noop            # no real intra-profile sleeps
        # Capture (and neutralise) the between-profiles pacing delay.
        self.profile_sleeps = []
        self._orig_sleep = ratelimit.sleep_interruptible

        async def capture(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
            self.profile_sleeps.append(seconds)
        ratelimit.sleep_interruptible = capture
        # The profile loop now gates on the battery, so these tests would
        # otherwise pass or fail according to whether this laptop happens to
        # be plugged in.  device_health is injectable for exactly this.
        self._battery = student_scraper.battery_state
        student_scraper.battery_state = lambda *_a, **_k: (100, True)

    def tearDown(self):
        sandbox.restore()
        import student_scraper as ratelimit
        student_scraper.battery_state = self._battery
        ratelimit.sleep_interruptible = self._orig_sleep
        S.pace = self._paced
        self.store.close()
        shutil.rmtree(self.tmp, ignore_errors=True)

    async def _noop(self, *_a, **_k):
        return None

    def workflow(self, candidates, fail_users, succeed_on_first=()):
        wf = S.Workflow.__new__(S.Workflow)
        wf.context = self.ctx
        wf.store = self.store
        wf.state = StateManager(path=self.ctx.state_file)
        wf.transient = TransientFailureLog()
        wf.deferred = DeferredQueue()
        wf.throttle = ThrottleCircuit()
        wf._deferred_this_run = set()
        wf._throttle_min_pos = None
        wf.reader = FakeReader(candidates)
        wf.options = types.SimpleNamespace(limit=0)
        wf.processed = 0
        self.saved = []
        fail_users = set(fail_users)

        async def fake_do(page, candidate):
            # A NON-throttle transient failure (a timeout, not a rate limit).
            # The two are handled differently now -- a rate limit cools down
            # and retries the search, a timeout leaves the row unresolved --
            # so the generic fixture must not accidentally say "rate limit".
            if candidate.username in fail_users:
                raise TransientProfileFailure("navigation timeout")
            self.saved.append(candidate.username)

        async def fake_ensure(page):
            return page

        async def fake_reclaim(page, n):
            return None
        wf.do_candidate = fake_do
        wf.ensure_browser = fake_ensure
        wf._reclaim = fake_reclaim
        return wf

    def run_rows(self, wf):
        return asyncio.run(wf.process_rows(object(), "kw.xlsx", FakeUnit()))

    # profile pacing: one delay BETWEEN two loads, none before the first,
    # none after the last, each within PROFILE_DELAY_MIN..MAX
    def test_profile_delay_between_loads_only(self):
        wf = self.workflow([cand("a"), cand("b"), cand("c")], fail_users=set())
        self.run_rows(wf)
        self.assertEqual(self.saved, ["a", "b", "c"])
        self.assertEqual(len(self.profile_sleeps), 2)   # a->b and b->c only
        for seconds in self.profile_sleeps:
            self.assertTrue(S.PROFILE_DELAY_MIN <= seconds
                            <= S.PROFILE_DELAY_MAX)

    def test_no_profile_delay_for_single_profile(self):
        wf = self.workflow([cand("only")], fail_users=set())
        self.run_rows(wf)
        self.assertEqual(self.profile_sleeps, [])       # never before the first

    # A GitHub RATE LIMIT on a profile cools down and retries the search
    # instead of leaving it unresolved and stopping the rank.
    def test_a_rate_limited_profile_asks_for_a_cooldown(self):
        wf = self.workflow([cand("limited")], fail_users=set())

        async def rate_limited(page, candidate):
            raise TransientProfileFailure(
                "GitHub rate limit / challenge: rate limit")
        wf.do_candidate = rate_limited
        with self.assertRaises(S.RateLimited) as caught:
            self.run_rows(wf)
        self.assertEqual(caught.exception.source, "profile")
        # the retry resumes AT the failed profile, not past it
        self.assertEqual(wf.state.data["processing"]["next_row_index"], 0)
        self.assertEqual(wf._rate_limited_profile, "limited")
        self.assertFalse(wf.deferred.has("limited"),
                         "a cooldown is not a deferral")

    def test_a_profile_at_the_defer_threshold_is_parked_not_cooled_down(self):
        """Once it has failed enough times, parking it beats waiting again."""
        wf = self.workflow([cand("stuck"), cand("ok")], fail_users=set())

        async def rate_limited(page, candidate):
            if candidate.username == "stuck":
                raise TransientProfileFailure(
                    "GitHub rate limit / challenge: rate limit")
            self.saved.append(candidate.username)
        wf.do_candidate = rate_limited
        for _ in range(C.MAX_DEFERRED_FAILURES):
            wf.transient.record(cand("stuck"), "rate limit")
        result = self.run_rows(wf)         # must NOT raise RateLimited
        self.assertTrue(result)
        self.assertTrue(wf.deferred.has("stuck"))
        self.assertIn("ok", self.saved)

    def test_a_non_rate_limit_failure_does_not_cool_down(self):
        """A timeout is not a rate limit, so the old behaviour is kept:
        the row stays unresolved and no cooldown is requested."""
        wf = self.workflow([cand("broken")], fail_users={"broken"})
        result = self.run_rows(wf)         # must NOT raise RateLimited
        self.assertFalse(result)           # search not completed
        self.assertFalse(wf.deferred.has("broken"))

    # test 1: a first-time transient failure stays retryable (not deferred)
    def test_below_threshold_stays_retryable(self):
        wf = self.workflow([cand("x")], fail_users={"x"})
        result = self.run_rows(wf)
        self.assertFalse(result)                     # search not completed
        self.assertTrue(wf.transient.attempts("x") >= 1)
        self.assertFalse(wf.deferred.has("x"))       # NOT deferred yet
        self.assertEqual(self.store.count(), 0)      # not collected

    # test 2/5/6: at threshold -> deferred, next candidate + search continue
    def test_threshold_defers_and_continues(self):
        wf = self.workflow([cand("stuck"), cand("ok")], fail_users={"stuck"})
        # pre-seed so the next failure crosses MAX_DEFERRED_FAILURES
        for _ in range(C.MAX_DEFERRED_FAILURES):
            wf.transient.record(cand("stuck"), "rate limit")
        result = self.run_rows(wf)
        self.assertTrue(result)                      # search COMPLETED
        self.assertTrue(wf.deferred.has("stuck"))    # moved to deferred
        self.assertFalse(wf.transient.attempts("stuck")
                         and "stuck" in wf.transient.data["failures"])
        self.assertIn("ok", self.saved)              # next candidate ran
        self.assertNotIn("stuck", self.saved)        # deferred, not saved

    # test 3/4: deferred candidate not counted, not blacklisted
    def test_deferred_not_counted_or_blacklisted(self):
        wf = self.workflow([cand("stuck")], fail_users={"stuck"})
        for _ in range(C.MAX_DEFERRED_FAILURES):
            wf.transient.record(cand("stuck"), "rate limit")
        self.run_rows(wf)
        self.assertEqual(self.store.count(), 0)      # not collected (test 3/14)
        self.assertFalse(os.path.exists(self.ctx.skipped_file))  # not blacklisted
        self.assertTrue(wf.deferred.has("stuck"))    # kept, not deleted

    # test 13: duplicate protection unaffected by deferral
    def test_dedupe_unaffected(self):
        record = {col: "" for col in C.FINAL_COLUMNS}
        record["GitHub Profile"] = "https://github.com/real"
        self.store.save_student(record, rank=3, total_ranks=80)
        wf = self.workflow([cand("stuck")], fail_users={"stuck"})
        for _ in range(C.MAX_DEFERRED_FAILURES):
            wf.transient.record(cand("stuck"), "rate limit")
        self.run_rows(wf)
        self.assertTrue(self.store.is_known(username="real"))
        self.assertFalse(self.store.is_known(username="stuck"))


# ========================= end-of-rank / --retry-deferred ===================
class RetryDeferredCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="defretry_")
        self.ctx = sandbox.enter(self.tmp)
        self.store = student_scraper.open_store(self.ctx.store_db)
        self._paced = S.pace
        S.pace = self._noop

    def tearDown(self):
        sandbox.restore()
        S.pace = self._paced
        self.store.close()
        shutil.rmtree(self.tmp, ignore_errors=True)

    async def _noop(self, *_a, **_k):
        return None

    def workflow(self, recover_users):
        wf = S.Workflow.__new__(S.Workflow)
        wf.context = self.ctx
        wf.store = self.store
        wf.deferred = DeferredQueue()
        wf.throttle = ThrottleCircuit()
        wf.deferred.add(cand("stuck"), "rate limit", 3)
        wf.deferred.add(cand("recovers"), "rate limit", 3)
        recover = set(recover_users)

        async def fake_do(page, candidate):
            if candidate.username not in recover:
                raise TransientProfileFailure("rate limit")

        async def fake_ensure(page):
            return page
        wf.do_candidate = fake_do
        wf.ensure_browser = fake_ensure
        wf.store.release_claim = lambda *_a, **_k: None
        return wf

    # test 8/9: successful retry removed, failed retry stays deferred
    def test_retry_removes_success_keeps_failure(self):
        wf = self.workflow(recover_users={"recovers"})
        asyncio.run(wf._retry_deferred(object()))
        self.assertFalse(wf.deferred.has("recovers"))   # test 8
        self.assertTrue(wf.deferred.has("stuck"))        # test 9
        # survives restart with the still-deferred one only
        again = DeferredQueue()
        self.assertEqual(again.count(), 1)
        self.assertTrue(again.has("stuck"))


# =============================== distributed runner =========================
class SchemaUnchangedCase(unittest.TestCase):
    # test 15: the 65-column Excel schema is untouched by this feature
    def test_final_columns_still_65(self):
        self.assertEqual(len(C.FINAL_COLUMNS), 65)


if __name__ == "__main__":
    unittest.main(verbosity=2)
