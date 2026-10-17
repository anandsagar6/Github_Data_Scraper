"""A profile that TALKS about rate limiting is not a rate limit.

THE LIVE BUG THIS FILE EXISTS FOR
---------------------------------
    TRANSIENT PROFILE FAILURE
    Username: ROHITSOMIREDDI
    Reason: GitHub rate limit / challenge: rate limit

GitHub had not limited anything: the run was making ~4 requests per minute,
hundreds of times under any published limit.  A profile page renders the
person's own profile README, and these two real READMEs each contain the
substring "rate limit":

    "JWT auth, bcrypt hashing, rate limiting, env management"
    "and rate limiting (300 GET / 60 write req/min)"

blocked_marker() scanned that text, matched "rate limit" inside "rate
limiting", and the rank stopped -- repeatedly, on the same profiles.

THE RULE NOW
------------
A genuine block or challenge page does not render a user profile.  So the
marker scan is asked LAST and only of a page that did NOT parse as a profile.
HTTP 429/403 remains authoritative and is checked first, because that comes
from GitHub rather than from page text.
"""

import asyncio
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (GitHubProfileError)   # noqa: E402

# Verbatim from the two profiles that stopped the live run.
README_ROHIT = ("Rohit Somireddi\nSkills\n"
                "Security | JWT auth, bcrypt hashing, rate limiting, "
                "env management |\nPinned repositories")
README_SIDDHU = ("Siddhu Pudi\nBackend\n"
                 "and rate limiting (300 GET / 60 write req/min). "
                 "CI/CD via GitHub Actions.")
# What GitHub actually serves when it throttles: no profile at all.
REAL_BLOCK = ("Whoa there!\nYou have triggered an abuse detection "
              "mechanism.\nPlease wait a few minutes before you try again.")


class FakeResponse:
    def __init__(self, status=200):
        self.status = status


class FakeLocator:
    def __init__(self, text):
        self._text = text

    async def inner_text(self, timeout=None):
        return self._text


class FakePage:
    """Just enough Playwright surface for GitHubProfile.fetch()."""

    def __init__(self, body, fields, status=200):
        self._body = body
        self._fields = fields
        self._status = status

    async def goto(self, url, **_kwargs):
        return FakeResponse(self._status)

    def locator(self, _selector):
        return FakeLocator(self._body)

    async def evaluate(self, _script, _arg=None):
        return dict(self._fields)


def profile_fields(**overrides):
    """What PROFILE_FIELDS_JS returns for a healthy user profile."""
    fields = {"isUser": True, "isOrg": False, "notFound": False,
              "login": "someone", "name": "Some One", "email": "",
              "links": [], "website": ""}
    fields.update(overrides)
    return fields


class FalsePositiveCase(unittest.TestCase):
    def fetch(self, body, fields, status=200):
        profile = student_scraper.GitHubProfile(browser=None, say=lambda *_a: None)
        profile.fetch_readme_raw = self._no_readme
        page = FakePage(body, fields, status)
        return asyncio.run(profile.fetch(page, "someone"))

    async def _no_readme(self, _username):
        return ""

    # the two profiles that actually stopped the run
    def test_a_readme_mentioning_rate_limiting_is_collected(self):
        for body in (README_ROHIT, README_SIDDHU):
            fields, text = self.fetch(body, profile_fields())
            self.assertTrue(fields["isUser"])
            self.assertIn("rate limit", text.lower(),
                          "the text is returned as-is")

    def test_every_block_word_in_a_valid_profile_is_ignored(self):
        """Any marker inside a person's own content must not block."""
        for marker in C.BLOCK_MARKERS:
            body = "Jane Dev\nI wrote about %s in my project.\n" % marker
            fields, _text = self.fetch(body, profile_fields())
            self.assertTrue(fields["isUser"], marker)

    # and a REAL block is still caught
    def test_a_real_block_page_is_still_detected(self):
        """A block page carries no profile, so the marker scan runs."""
        with self.assertRaises(GitHubProfileError) as caught:
            self.fetch(REAL_BLOCK, profile_fields(isUser=False, login=""))
        self.assertIn("rate limit / challenge", str(caught.exception))
        self.assertFalse(caught.exception.permanent, "a block is retryable")

    def test_http_429_is_still_authoritative(self):
        """A status code comes from GitHub, not from page text."""
        with self.assertRaises(GitHubProfileError) as caught:
            self.fetch(README_ROHIT, profile_fields(), status=429)
        self.assertIn("429", str(caught.exception))

    def test_http_403_is_still_authoritative(self):
        with self.assertRaises(GitHubProfileError) as caught:
            self.fetch(README_ROHIT, profile_fields(), status=403)
        self.assertIn("403", str(caught.exception))

    def test_404_is_still_permanent(self):
        with self.assertRaises(GitHubProfileError) as caught:
            self.fetch("Not Found", profile_fields(), status=404)
        self.assertTrue(caught.exception.permanent)

    def test_an_unrecognised_page_without_markers_is_retryable(self):
        with self.assertRaises(GitHubProfileError) as caught:
            self.fetch("something else entirely",
                       profile_fields(isUser=False))
        self.assertIn("markup not recognised", str(caught.exception))
        self.assertFalse(caught.exception.permanent)

    def test_an_organisation_is_still_a_permanent_skip(self):
        with self.assertRaises(GitHubProfileError) as caught:
            self.fetch("Acme Inc", profile_fields(isUser=False, isOrg=True))
        self.assertTrue(caught.exception.permanent)

    # blocked_marker() itself is unchanged: it still matches text given to it
    def test_blocked_marker_still_matches_when_asked(self):
        self.assertEqual(student_scraper.blocked_marker("You have triggered an abuse "
                                           "detection mechanism"),
                         "abuse detection")
        self.assertIsNone(student_scraper.blocked_marker("a perfectly normal profile"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
