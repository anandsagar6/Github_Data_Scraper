"""Battery and network gates.

The rules the operator asked for:
    battery at/below 20% on battery power -> PAUSE until a charger appears,
        then resume by itself at 21% (or the moment mains power is back)
    battery critical                      -> STOP cleanly, state preserved
    no internet / github unreachable      -> PAUSE and resume AUTOMATICALLY
        whenever it comes back, however long that takes -- never give up

And the two things that made 20% look like it did nothing at all:
    the battery is checked per PROFILE, not only between searches -- one
        search can run for tens of minutes, long enough for a 20% battery
        to reach empty while the run carried on
    a run paused for the battery consults BATTERY_RESUME_PERCENT instead of
        continuing the instant battery_verdict stopped saying "pause"

Nothing here touches real hardware, the real clock or the real network: the
sysfs tree is a temp directory, the socket is a stub, and the sleep is faked.
"""

import asyncio
import inspect
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper as S                              # noqa: E402


def fake_sysfs(root, percent=None, ac=None, status=None):
    """Build a /sys/class/power_supply-shaped tree."""
    if percent is not None:
        bat = os.path.join(root, "BAT0")
        os.makedirs(bat, exist_ok=True)
        open(os.path.join(bat, "capacity"), "w").write("%s\n" % percent)
        if status is not None:
            open(os.path.join(bat, "status"), "w").write("%s\n" % status)
    if ac is not None:
        adapter = os.path.join(root, "AC")
        os.makedirs(adapter, exist_ok=True)
        open(os.path.join(adapter, "online"), "w").write("%s\n" % ac)
    return root


class BatteryReadingCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="psy_")

    def test_reads_percent_and_mains(self):
        fake_sysfs(self.root, percent=73, ac=1, status="Charging")
        self.assertEqual(student_scraper.battery_state(self.root), (73, True))

    def test_on_battery_is_not_plugged(self):
        fake_sysfs(self.root, percent=44, ac=0, status="Discharging")
        self.assertEqual(student_scraper.battery_state(self.root), (44, False))

    def test_charging_status_counts_as_plugged_without_an_ac_node(self):
        fake_sysfs(self.root, percent=50, status="Charging")
        self.assertEqual(student_scraper.battery_state(self.root), (50, True))

    def _no_psutil_battery(self):
        """Make the psutil fallback report "this machine has no battery".

        battery_state() falls back to psutil when sysfs shows no BAT* node,
        which is the whole point of the fallback -- a battery that udev has
        not exposed is still a battery.  So an empty FAKE sysfs tree does not
        by itself describe a batteryless machine: without this stub the
        fallback reads the REAL laptop and answers 100.

        That is also why this test used to pass: psutil simply was not
        installed, so the fallback returned (None, None) by accident.  It
        failed the moment requirements.txt was installed into a venv.
        """
        saved = student_scraper._battery_state_psutil
        student_scraper._battery_state_psutil = lambda: (None, None)
        self.addCleanup(setattr, student_scraper,
                        "_battery_state_psutil", saved)

    def test_a_machine_with_no_battery_reports_none(self):
        """A desktop/server has no battery risk, so it is never paused."""
        self._no_psutil_battery()
        percent, plugged = student_scraper.battery_state(self.root)   # empty tree
        self.assertIsNone(percent)
        self.assertIsNone(plugged)
        self.assertEqual(student_scraper.battery_verdict(percent, plugged), "ok")

    def test_an_empty_sysfs_tree_consults_the_psutil_fallback(self):
        """The fallback is REACHED -- nothing was covering this before.

        A laptop whose battery is not under this root must still be found,
        otherwise a real battery at 4% would read as "no battery, never
        pause" and the device would run itself flat.
        """
        saved = student_scraper._battery_state_psutil
        calls = []

        def fake():
            calls.append(True)
            return 4, False

        student_scraper._battery_state_psutil = fake
        self.addCleanup(setattr, student_scraper,
                        "_battery_state_psutil", saved)
        self.assertEqual(student_scraper.battery_state(self.root), (4, False))
        self.assertEqual(len(calls), 1, "the fallback must be consulted")

    def test_unreadable_capacity_is_not_a_crash(self):
        bat = os.path.join(self.root, "BAT0")
        os.makedirs(bat)
        open(os.path.join(bat, "capacity"), "w").write("not-a-number\n")
        self.assertEqual(student_scraper.battery_state(self.root), (None, None))


class BatteryVerdictCase(unittest.TestCase):
    # the operator's rule: 20% on battery -> pause
    def test_twenty_percent_on_battery_pauses(self):
        self.assertEqual(student_scraper.battery_verdict(20, False), "pause")
        self.assertEqual(student_scraper.battery_verdict(19, False), "pause")

    def test_above_the_threshold_is_ok(self):
        self.assertEqual(student_scraper.battery_verdict(21, False), "ok")
        self.assertEqual(student_scraper.battery_verdict(100, False), "ok")

    def test_mains_power_is_always_ok(self):
        """A plugged-in laptop at 5% is charging, not dying."""
        for percent in (5, 20, 90):
            self.assertEqual(student_scraper.battery_verdict(percent, True), "ok")

    def test_critical_stops_rather_than_pausing(self):
        """Waiting cannot help below critical: save state while we can."""
        self.assertEqual(student_scraper.battery_verdict(C.BATTERY_CRITICAL_PERCENT, False),
                         "stop")
        self.assertEqual(student_scraper.battery_verdict(1, False), "stop")

    def test_resume_needs_more_than_the_pause_threshold(self):
        """Hysteresis: the run must not flap around one number."""
        self.assertGreater(C.BATTERY_RESUME_PERCENT, C.BATTERY_PAUSE_PERCENT)
        self.assertFalse(student_scraper.battery_resumed(C.BATTERY_PAUSE_PERCENT, False))
        self.assertTrue(student_scraper.battery_resumed(C.BATTERY_RESUME_PERCENT, False))
        self.assertTrue(student_scraper.battery_resumed(5, True))      # charger back

    def test_the_configured_thresholds_are_the_requested_ones(self):
        self.assertEqual(C.BATTERY_PAUSE_PERCENT, 20)
        self.assertEqual(C.BATTERY_CRITICAL_PERCENT, 10)


class NetworkCase(unittest.TestCase):
    def test_a_successful_connect_is_up(self):
        class Conn:
            def close(self): pass
        seen = {}

        def connect(address, timeout):
            seen["address"] = address
            seen["timeout"] = timeout
            return Conn()
        self.assertTrue(student_scraper.network_ok(connect=connect))
        self.assertEqual(seen["address"], (C.HEALTH_HOST, C.HEALTH_PORT))
        self.assertEqual(seen["timeout"], C.HEALTH_TIMEOUT_SECONDS)

    def test_a_refused_connect_is_down(self):
        def connect(address, timeout):
            raise OSError("connection refused")
        self.assertFalse(student_scraper.network_ok(connect=connect))

    def test_a_timeout_is_down(self):
        def connect(address, timeout):
            raise TimeoutError()
        self.assertFalse(student_scraper.network_ok(connect=connect))


class HealthGateCase(unittest.TestCase):
    """The gate the run loop calls between two searches."""

    def setUp(self):
        self.said = []
        self.slept = []
        self._say = S.say
        self._sleep = student_scraper.sleep_interruptible
        self._battery = student_scraper.battery_state
        self._network = student_scraper.network_ok
        S.say = self.said.append

        async def fake_sleep(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
            self.slept.append(seconds)
        student_scraper.sleep_interruptible = fake_sleep

    def tearDown(self):
        S.say = self._say
        student_scraper.sleep_interruptible = self._sleep
        student_scraper.battery_state = self._battery
        student_scraper.network_ok = self._network

    def gate(self, readings, online=True, check_network=True):
        """readings: list of (percent, plugged) returned in order."""
        series = list(readings)
        student_scraper.battery_state = lambda *a, **k: (series.pop(0) if len(series) > 1
                                           else series[0])
        if callable(online):
            student_scraper.network_ok = online
        else:
            student_scraper.network_ok = lambda *a, **k: online
        wf = S.Workflow.__new__(S.Workflow)
        return asyncio.run(wf._health_gate(check_network=check_network))

    def text(self):
        return "\n".join(self.said)

    def test_healthy_continues_without_waiting(self):
        self.assertTrue(self.gate([(90, True)]))
        self.assertEqual(self.slept, [])

    def test_no_battery_machine_continues(self):
        self.assertTrue(self.gate([(None, None)]))
        self.assertEqual(self.slept, [])

    # 20% on battery -> PAUSE, then continue once the charger appears
    def test_low_battery_pauses_then_resumes_when_plugged_in(self):
        self.assertTrue(self.gate([(18, False), (18, True)]))
        self.assertEqual(self.slept, [C.HEALTH_RECHECK_SECONDS])
        self.assertIn("[PAUSED]", self.text())
        self.assertIn("waiting for a charger", self.text())
        self.assertIn("recovered", self.text())

    # critical battery -> STOP, no waiting at all
    def test_critical_battery_stops_immediately(self):
        self.assertFalse(self.gate([(5, False)]))
        self.assertEqual(self.slept, [])
        self.assertIn("[BATTERY CRITICAL]", self.text())
        self.assertIn("[DEVICE STOPPED]", self.text())

    # server/network down -> PAUSE, then continue when it returns
    def test_network_down_pauses_then_resumes(self):
        calls = {"n": 0}

        def flaky(*_a, **_k):
            calls["n"] += 1
            return calls["n"] > 1
        self.assertTrue(self.gate([(90, True)], online=flaky))
        self.assertEqual(self.slept, [C.HEALTH_RECHECK_SECONDS])
        self.assertIn("unreachable", self.text())

    # ... and WAITS, however long it takes, then resumes by itself.
    # This used to stop the device after HEALTH_MAX_WAIT_MINUTES; the
    # operator wants an overnight outage to find the run still going.
    def test_the_wait_is_unbounded_by_configuration(self):
        self.assertIsNone(C.HEALTH_MAX_WAIT_MINUTES)

    def test_a_long_outage_resumes_automatically_instead_of_stopping(self):
        checks = {"n": 0}

        def down_for_a_long_time(*_a, **_k):
            checks["n"] += 1
            return checks["n"] > 200          # ~100 min at 30s a check
        self.assertTrue(self.gate([(90, True)], online=down_for_a_long_time))
        self.assertIn("recovered", self.text())
        self.assertNotIn("[DEVICE STOPPED]", self.text())
        # Well past the 60 minutes it used to give up after.
        self.assertGreater(sum(self.slept), 60 * 60)

    def test_a_paused_run_says_it_will_come_back_on_its_own(self):
        calls = {"n": 0}

        def flaky(*_a, **_k):
            calls["n"] += 1
            return calls["n"] > 1
        self.gate([(90, True)], online=flaky)
        self.assertIn("resumes by itself", self.text())

    def test_ctrl_c_while_paused_is_still_responsive(self):
        S.STOP_REQUESTED = True
        try:
            with self.assertRaises(S.Stopped):
                self.gate([(90, True)], online=False)
        finally:
            S.STOP_REQUESTED = False

    # ---------------------------------------------------- hysteresis ------
    # BATTERY_RESUME_PERCENT used to be dead config: the gate re-read
    # battery_verdict, which says "ok" one point above the pause threshold,
    # so a run paused at 20% resumed at 21% and could flap straight back.
    def test_a_paused_run_keeps_waiting_below_the_resume_threshold(self):
        """It continues at BATTERY_RESUME_PERCENT and not a point earlier."""
        below = C.BATTERY_RESUME_PERCENT - 1
        readings = ([(C.BATTERY_PAUSE_PERCENT, False)]
                    + [(below, False)] * 5
                    + [(C.BATTERY_RESUME_PERCENT, False)])
        self.assertTrue(self.gate(readings))
        # One wait for the pause itself plus one per reading still too low.
        self.assertEqual(len(self.slept), 6)
        self.assertIn("recovered", self.text())

    def test_a_paused_run_resumes_at_the_resume_threshold(self):
        self.assertTrue(self.gate([(C.BATTERY_PAUSE_PERCENT, False),
                                   (C.BATTERY_RESUME_PERCENT, False)]))
        self.assertIn("recovered", self.text())

    def test_a_healthy_run_is_not_held_to_the_resume_threshold(self):
        """Hysteresis applies only AFTER a pause, never on the first read."""
        between = C.BATTERY_PAUSE_PERCENT + 1
        self.assertTrue(self.gate([(between, False)]))
        self.assertEqual(self.slept, [])

    # --------------------------------------------- the mid-search gate ----
    # The between-search gate alone let a long search run on past 20%: one
    # search can hold the run for tens of minutes.
    def test_the_mid_search_gate_never_touches_the_network(self):
        def explode(*_a, **_k):
            raise AssertionError("network checked during a search")
        self.assertTrue(self.gate([(90, False)], online=explode,
                                  check_network=False))

    def test_the_mid_search_gate_still_stops_on_a_critical_battery(self):
        self.assertFalse(self.gate([(5, False)], check_network=False))
        self.assertIn("[BATTERY CRITICAL]", self.text())

    def test_the_mid_search_gate_pauses_and_waits_for_a_charger(self):
        self.assertTrue(self.gate([(C.BATTERY_PAUSE_PERCENT, False),
                                   (C.BATTERY_PAUSE_PERCENT, True)],
                                  check_network=False))
        self.assertIn("waiting for a charger", self.text())

    def test_the_mid_search_gate_does_not_claim_the_network_is_up(self):
        """Reporting "up" on a reading nobody took would be a lie."""
        self.gate([(C.BATTERY_PAUSE_PERCENT, False),
                   (C.BATTERY_PAUSE_PERCENT, True)], check_network=False)
        self.assertIn("Network    : not checked", self.text())
        self.assertNotIn("Network    : up", self.text())

    def test_the_profile_loop_gates_on_the_battery(self):
        """Wiring: without this call the gate cannot fire inside a search."""
        source = inspect.getsource(S.Workflow.process_rows)
        self.assertIn("_health_gate(check_network=False)", source)

    def test_the_deferred_retry_loop_gates_on_the_battery(self):
        """It can hold hundreds of profiles -- as long a stretch as a search."""
        source = inspect.getsource(S.Workflow._retry_deferred)
        self.assertIn("_health_gate(check_network=False)", source)

    def test_the_page_crawl_gates_between_pages(self):
        """The LAST unguarded stretch: a multi-page crawl runs for minutes."""
        import student_scraper as page_crawler
        source = inspect.getsource(page_crawler.PageByPageCollector.collect)
        self.assertIn("self._health_gate()", source)
        # and the workflow must actually hand one over
        wiring = inspect.getsource(S.Workflow._collect_page_by_page)
        self.assertIn("health_gate=self._crawl_health_gate", wiring)


class CrawlGateCase(unittest.TestCase):
    """The gate as the page crawler uses it: pause in place, or stop."""

    def setUp(self):
        self.said = []
        self._say = S.say
        self._sleep = student_scraper.sleep_interruptible
        self._battery = student_scraper.battery_state
        S.say = self.said.append

        async def fake_sleep(seconds, check_stop=None, **_k):
            if check_stop is not None:
                check_stop()
        student_scraper.sleep_interruptible = fake_sleep

    def tearDown(self):
        S.say = self._say
        student_scraper.sleep_interruptible = self._sleep
        student_scraper.battery_state = self._battery

    def gate(self, readings):
        series = list(readings)
        student_scraper.battery_state = lambda *a, **k: (series.pop(0) if len(series) > 1
                                           else series[0])
        wf = S.Workflow.__new__(S.Workflow)
        return asyncio.run(wf._crawl_health_gate())

    def test_a_low_battery_pauses_the_crawl_then_lets_it_continue(self):
        self.assertIsNone(self.gate([(20, False), (20, True)]))
        text = "\n".join(self.said)
        self.assertIn("[PAUSED]", text)
        self.assertIn("recovered", text)

    def test_a_critical_battery_stops_the_crawl_cleanly(self):
        """Raising is what preserves the staged pages."""
        with self.assertRaises(S.Stopped):
            self.gate([(5, False)])
        self.assertIn("[BATTERY CRITICAL]", "\n".join(self.said))

    def test_a_healthy_battery_does_not_interrupt_the_crawl(self):
        self.assertIsNone(self.gate([(90, False)]))
        self.assertEqual(self.said, [])

    def test_a_collector_built_without_a_gate_is_unaffected(self):
        """Existing callers must see no change at all."""
        import student_scraper as page_crawler
        collector = page_crawler.PageByPageCollector.__new__(
            page_crawler.PageByPageCollector)
        signature = inspect.signature(
            page_crawler.PageByPageCollector.__init__)
        self.assertIsNone(signature.parameters["health_gate"].default)
        del collector


if __name__ == "__main__":
    unittest.main(verbosity=2)
