"""Global-throttle circuit breaker: stop the device when the ACCOUNT is blocked.

One unavailable profile is deferred and the rank keeps going.  But when several
DIFFERENT profiles all hit a GitHub rate-limit/challenge inside the window, it
is an account/session block: the rank stops (STOPPED_GLOBAL_THROTTLE), the
device stops, no next rank starts, and those profiles are NOT deferred as if
profile-specific.  Evidence persists so restarting cannot bypass it; a
successful load clears the streak.
"""

import asyncio
import datetime
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
from student_scraper import (Candidate, DeferredQueue, StateManager, ThrottleCircuit, TransientFailureLog, is_throttle_reason)   # noqa: E402
import sandbox                                         # noqa: E402
import student_scraper as S                                # noqa: E402
from student_scraper import GlobalThrottle, TransientProfileFailure  # noqa: E402


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


# ============================== the circuit itself ==========================
class ThrottleCircuitCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="thr_")
        self.path = os.path.join(self.dir, "throttle.json")

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.dir, ignore_errors=True)

    def circuit(self):
        return ThrottleCircuit(path=self.path, distinct=3, window_minutes=10)

    def test_reason_classification(self):
        self.assertTrue(is_throttle_reason(
            "GitHub rate limit / challenge: rate limit"))
        self.assertTrue(is_throttle_reason("GitHub returned HTTP 429"))
        self.assertFalse(is_throttle_reason("navigation timeout"))
        self.assertFalse(is_throttle_reason("404 not found"))

    # test 2: three different profiles within the window -> tripped
    def test_three_distinct_trips(self):
        c = self.circuit()
        self.assertEqual(c.record_throttle("A"), 1)
        self.assertEqual(c.record_throttle("B"), 2)
        self.assertFalse(c.tripped())
        self.assertEqual(c.record_throttle("C"), 3)
        self.assertTrue(c.tripped())

    # same profile repeatedly does NOT trip the global circuit (it is one)
    def test_same_profile_does_not_trip(self):
        c = self.circuit()
        for _ in range(5):
            c.record_throttle("A")
        self.assertFalse(c.tripped())
        self.assertEqual(c.distinct_recent(), 1)

    # test 7: persisted evidence survives restart (cannot be bypassed)
    def test_persists_across_restart(self):
        c = self.circuit()
        for user in ("A", "B", "C"):
            c.record_throttle(user)
        self.assertTrue(self.circuit().tripped())          # fresh instance

    # test 8: a successful load clears the streak
    def test_success_clears_streak(self):
        c = self.circuit()
        for user in ("A", "B"):
            c.record_throttle(user)
        c.record_success()
        self.assertEqual(c.distinct_recent(), 0)
        self.assertFalse(c.tripped())
        # and a later single throttle does not immediately trip
        c.record_throttle("D")
        self.assertFalse(c.tripped())

    # events older than the window are ignored
    def test_window_expiry(self):
        c = self.circuit()
        old = (datetime.datetime.now()
               - datetime.timedelta(minutes=20)).isoformat(timespec="seconds")
        c.data["events"] = [{"username": u, "at": old} for u in ("A", "B")]
        c.save()
        self.assertEqual(c.distinct_recent(), 0)           # both expired
        c.record_throttle("C")
        self.assertFalse(c.tripped())                      # only 1 fresh


# ================== global throttle in the profile phase ====================
class ProfilePhaseThrottleCase(unittest.TestCase):
    def setUp(self):
        import student_scraper as ratelimit
        self.tmp = tempfile.mkdtemp(prefix="thrphase_")
        self.ctx = sandbox.enter(self.tmp)
        self.store = student_scraper.open_store(self.ctx.store_db)
        self._paced = S.pace
        S.pace = self._noop
        # neutralise the between-profiles pacing delay (kept honest elsewhere)
        self._orig_sleep = ratelimit.sleep_interruptible

        async def _fast(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
        ratelimit.sleep_interruptible = _fast
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

    def workflow(self, candidates, fail_users):
        wf = S.Workflow.__new__(S.Workflow)
        wf.context = self.ctx
        wf.store = self.store
        wf.state = StateManager(path=self.ctx.state_file)
        wf.transient = TransientFailureLog()
        wf.deferred = DeferredQueue()
        # small distinct limit is the production default (3)
        wf.throttle = ThrottleCircuit(distinct=3,
                                      window_minutes=10)
        wf._deferred_this_run = set()
        wf._throttle_min_pos = None
        wf.reader = FakeReader(candidates)
        wf.options = types.SimpleNamespace(limit=0)
        wf.processed = 0
        fail = set(fail_users)

        async def fake_do(page, candidate):
            if candidate.username in fail:
                raise TransientProfileFailure("GitHub rate limit / challenge")
            wf.throttle.record_success()          # mirror the real success hook

        async def fake_ensure(page):
            return page

        async def fake_reclaim(page, n):
            return None
        wf.do_candidate = fake_do
        wf.ensure_browser = fake_ensure
        wf._reclaim = fake_reclaim
        wf.store.release_claim = lambda *_a, **_k: None
        return wf

    def run_rows(self, wf):
        return asyncio.run(wf.process_rows(object(), "kw.xlsx", FakeUnit()))

    # test 2/3: three different throttled profiles -> GlobalThrottle raised
    def test_three_distinct_raises_global_throttle(self):
        # pre-seed each so a single failure this run crosses the per-profile
        # deferral threshold, so all three are reached in one pass.
        cands = [cand("A", 1), cand("B", 2), cand("C", 3)]
        wf = self.workflow(cands, fail_users={"A", "B", "C"})
        for name in ("A", "B", "C"):
            for _ in range(C.MAX_DEFERRED_FAILURES):
                wf.transient.record(cand(name), "rate limit")
        with self.assertRaises(GlobalThrottle):
            self.run_rows(wf)

    # test 5: the streak profiles are NOT left deferred by the global throttle
    def test_global_throttle_does_not_leave_profiles_deferred(self):
        cands = [cand("A", 1), cand("B", 2), cand("C", 3)]
        wf = self.workflow(cands, fail_users={"A", "B", "C"})
        for name in ("A", "B", "C"):
            for _ in range(C.MAX_DEFERRED_FAILURES):
                wf.transient.record(cand(name), "rate limit")
        with self.assertRaises(GlobalThrottle):
            self.run_rows(wf)
        # A and B were deferred-this-run then undone; C tripped before deferral
        self.assertFalse(wf.deferred.has("A"))
        self.assertFalse(wf.deferred.has("B"))
        self.assertFalse(wf.deferred.has("C"))
        self.assertEqual(self.store.count(), 0)       # none counted (test 14)

    # test 6: state + XLSX preserved (nothing discarded / deleted)
    def test_state_and_xlsx_preserved(self):
        xlsx = os.path.join(self.ctx.keyword_download_dir, "kw.xlsx")
        with open(xlsx, "w") as handle:
            handle.write("keyword rows")
        cands = [cand("A", 1), cand("B", 2), cand("C", 3)]
        wf = self.workflow(cands, fail_users={"A", "B", "C"})
        for name in ("A", "B", "C"):
            for _ in range(C.MAX_DEFERRED_FAILURES):
                wf.transient.record(cand(name), "rate limit")
        with self.assertRaises(GlobalThrottle):
            self.run_rows(wf)
        self.assertTrue(os.path.exists(xlsx))             # XLSX preserved
        self.assertTrue(os.path.exists(self.ctx.state_file))  # state saved
        # rewound so a resume retries from the first throttled profile
        self.assertEqual(wf.state.next_row_index, 0)

    # test 1 (contrast): one throttled profile among successes -> deferred,
    # NOT a global throttle
    def test_single_profile_defers_not_global(self):
        cands = [cand("bad", 1), cand("good", 2)]
        wf = self.workflow(cands, fail_users={"bad"})
        for _ in range(C.MAX_DEFERRED_FAILURES):
            wf.transient.record(cand("bad"), "rate limit")
        result = self.run_rows(wf)          # no GlobalThrottle
        self.assertTrue(result)             # search completed
        self.assertTrue(wf.deferred.has("bad"))
        self.assertFalse(wf.throttle.tripped())   # good load cleared streak

    # test 7 (restart): persisted trip means a fresh run trips immediately
    def test_restart_cannot_bypass(self):
        # persist 3 distinct throttle events, then a new circuit is tripped
        circuit = ThrottleCircuit(distinct=3,
                                  window_minutes=10)
        for user in ("A", "B", "C"):
            circuit.record_throttle(user)
        fresh = ThrottleCircuit(distinct=3, window_minutes=10)
        self.assertTrue(fresh.tripped())


# =============================== runner side ================================


if __name__ == "__main__":
    unittest.main(verbosity=2)
