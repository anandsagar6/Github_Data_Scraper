"""ensure_logged_in: pass through when signed in, pause + auto-resume when not.

Stubs the browser and the sleep so the whole pause/re-login/resume cycle runs
in milliseconds.  Nothing here opens a real browser or touches GitHub.
"""
import asyncio
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                     # noqa: E402
import student_scraper as S           # noqa: E402


class FakeBrowser:
    def __init__(self, logins):
        # `logins` is a list of return values for successive
        # github_login_name() calls (None = signed out, str = signed in).
        self._logins = list(logins)
        self.closed = 0
        self.started = 0
        self.login_calls = 0

    async def github_login_name(self, page):
        self.login_calls += 1
        return self._logins.pop(0) if self._logins else "alice"

    async def close(self):
        self.closed += 1

    async def start(self, with_extension=True):
        self.started += 1

    async def page(self):
        return "PAGE-%d" % self.started

    async def close_extra_pages(self, keep=()):
        return 0


def make_workflow(browser):
    wf = S.Workflow.__new__(S.Workflow)   # skip the heavy __init__
    wf.browser = browser
    wf.state = None
    wf.scraper = None
    wf.profiles = None
    return wf


class LoginWatchCase(unittest.TestCase):
    def setUp(self):
        self._auth = C.AUTH_FILE
        self._recheck = S.LOGIN_RECHECK_SECONDS
        handle, self.auth_path = tempfile.mkstemp(suffix="_auth.json")
        os.close(handle)
        C.AUTH_FILE = self.auth_path
        S.LOGIN_RECHECK_SECONDS = 60      # distinguishes the wait from the 3s guard

        # No real sleeping, no real bell / helpers.
        self._orig = (S.sleep_interruptible, S.bell,
                      S.InstantDataScraper, S.GitHubProfile)
        S.bell = lambda: None
        S.InstantDataScraper = lambda *a, **k: "SCRAPER"
        S.GitHubProfile = lambda *a, **k: "PROFILES"

    def tearDown(self):
        C.AUTH_FILE = self._auth
        S.LOGIN_RECHECK_SECONDS = self._recheck
        (S.sleep_interruptible, S.bell,
         S.InstantDataScraper, S.GitHubProfile) = self._orig
        os.path.exists(self.auth_path) and os.remove(self.auth_path)

    def test_signed_in_passes_straight_through(self):
        browser = FakeBrowser(["alice"])         # first check succeeds
        wf = make_workflow(browser)

        async def fake_sleep(*a, **k):
            raise AssertionError("must not wait when already signed in")
        S.sleep_interruptible = fake_sleep

        page = asyncio.run(wf.ensure_logged_in("P0"))
        self.assertEqual(page, "P0")             # same page, untouched
        self.assertEqual(browser.closed, 0)      # browser never released
        self.assertEqual(browser.started, 0)
        self.assertEqual(browser.login_calls, 1)

    def test_logout_pauses_then_auto_resumes_on_relogin(self):
        # 3 Nones confirm the logout; after reopen the 4th call is signed in.
        browser = FakeBrowser([None, None, None, "bob"])
        wf = make_workflow(browser)

        async def fake_sleep(seconds, *a, **k):
            # The wait tick (60s) is when --login would land: write a fresh
            # github_auth.json so its mtime moves on.
            if int(seconds) == S.LOGIN_RECHECK_SECONDS:
                with open(C.AUTH_FILE, "w") as fh:
                    fh.write('{"cookies":[{"name":"x"}]}')
                    fh.flush()
                os.utime(C.AUTH_FILE, None)
            return
        S.sleep_interruptible = fake_sleep

        page = asyncio.run(wf.ensure_logged_in("P0"))
        self.assertEqual(page, "PAGE-1")         # fresh page from the reopen
        self.assertGreaterEqual(browser.closed, 1)   # released for --login
        self.assertEqual(browser.started, 1)          # reopened once
        self.assertEqual(wf.scraper, "SCRAPER")       # helpers rebuilt
        self.assertEqual(wf.profiles, "PROFILES")

    def test_false_logout_does_not_pause(self):
        # One flaky None, then signed in: the 3x guard must ride it out.
        browser = FakeBrowser([None, "carol"])
        wf = make_workflow(browser)

        async def fake_sleep(seconds, *a, **k):
            # Only the 3s guard sleep may run; a 60s wait means it paused.
            if int(seconds) == S.LOGIN_RECHECK_SECONDS:
                raise AssertionError("a single flaky read must not pause")
            return
        S.sleep_interruptible = fake_sleep

        page = asyncio.run(wf.ensure_logged_in("P0"))
        self.assertEqual(page, "P0")
        self.assertEqual(browser.closed, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
