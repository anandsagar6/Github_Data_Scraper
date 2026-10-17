"""The page-2 diagnostic must diagnose the page it was asked about.

THE BUG THIS FILE EXISTS FOR
----------------------------
    python3 student_scraper.py "University Student Visakhapatnam"
    page classification : unknown-invalid
    (not a GitHub Users search page (https://github.com/))

browser_manager.github_login_name() navigates the tab to https://github.com
to read the user-login meta tag.  The diagnostic called it AFTER navigating
to the search, so the login check threw the search navigation away and every
later phase inspected the GitHub homepage.

The login check now happens first, and the landing is PROVED before anything
is diagnosed: a homepage, a login wall, an error page or any redirect away
from /search?type=users stops the run instead of being described as if it
were the results.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                      # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (PageState)   # noqa: E402
import student_scraper as D                              # noqa: E402

SEARCH = "https://github.com/search?q=University+Student+Visakhapatnam&type=users"
GOOD = (PageState.RESULTS_MULTI_PAGE, "10 result(s) and a next page")


class TestSearchUrl(unittest.TestCase):
    """Requirement: the query must open the Users search URL."""

    def test_the_expected_url_is_built(self):
        self.assertEqual(
            D.search_url_for("University Student Visakhapatnam"), SEARCH)

    def test_it_uses_the_production_template(self):
        """One source for the template, so the two can never diverge."""
        self.assertEqual(
            D.search_url_for("Fresher Pune"),
            C.SEARCH_URL.format(query="Fresher+Pune"))
        self.assertIn("type=users", D.search_url_for("anything"))

    def test_spaces_and_punctuation_are_encoded(self):
        url = D.search_url_for("B.Tech Student Navi Mumbai")
        self.assertIn("B.Tech+Student+Navi+Mumbai", url)
        self.assertNotIn(" ", url)

    def test_the_query_is_not_lost_for_a_single_word(self):
        self.assertEqual(D.search_url_for("Fresher"),
                         "https://github.com/search?q=Fresher&type=users")


class TestLandingVerification(unittest.TestCase):
    """Requirement: reject the homepage, login and error redirects."""

    def verify(self, final_url, page_state=GOOD):
        return D.verify_search_landing(SEARCH, final_url, page_state)

    # -- the good case -----------------------------------------------------
    def test_the_real_search_page_is_accepted(self):
        ok, why = self.verify(SEARCH)
        self.assertTrue(ok, why)

    def test_page_two_of_the_search_is_accepted(self):
        ok, why = self.verify(SEARCH + "&p=2")
        self.assertTrue(ok, why)

    def test_a_single_page_result_set_is_accepted(self):
        ok, _why = self.verify(
            SEARCH, (PageState.RESULTS_SINGLE_PAGE, "10 result(s)"))
        self.assertTrue(ok)

    # -- the homepage, which is the bug ------------------------------------
    def test_the_homepage_is_rejected(self):
        for url in ("https://github.com/", "https://github.com",
                    "http://github.com/"):
            ok, why = self.verify(url)
            self.assertFalse(ok, url)
            self.assertIn("HOMEPAGE", why)

    def test_the_homepage_is_rejected_even_with_a_healthy_classification(self):
        """The classification is not the only check; the URL is checked too."""
        ok, why = self.verify("https://github.com/",
                              (PageState.RESULTS_MULTI_PAGE, "looks fine"))
        self.assertFalse(ok)
        self.assertIn("HOMEPAGE", why)

    # -- login -------------------------------------------------------------
    def test_a_login_redirect_is_rejected(self):
        for url in ("https://github.com/login",
                    "https://github.com/login?return_to=%2Fsearch",
                    "https://github.com/session"):
            ok, why = self.verify(url)
            self.assertFalse(ok, url)
            self.assertIn("LOGIN", why)

    def test_a_login_wall_on_the_search_url_is_rejected(self):
        ok, why = self.verify(
            SEARCH, (PageState.LOGIN_REQUIRED, "GitHub is asking for a sign-in"))
        self.assertFalse(ok)
        self.assertIn("login-required", why)

    # -- errors and throttling --------------------------------------------
    def test_a_rate_limited_page_is_rejected(self):
        ok, why = self.verify(
            SEARCH, (PageState.BLOCKED, "GitHub says: secondary rate limit"))
        self.assertFalse(ok)
        self.assertIn("rate limit", why)

    def test_an_http_error_page_is_rejected(self):
        ok, why = self.verify(SEARCH,
                              (PageState.HTTP_ERROR, "GitHub returned HTTP 429"))
        self.assertFalse(ok)
        self.assertIn("429", why)

    def test_a_github_error_page_is_rejected(self):
        ok, _why = self.verify(
            SEARCH, (PageState.GITHUB_ERROR, "GitHub rendered an error page"))
        self.assertFalse(ok)

    def test_an_unrenderable_page_is_rejected(self):
        ok, why = self.verify(
            SEARCH, (PageState.UNKNOWN_INVALID, "no user results"))
        self.assertFalse(ok)
        self.assertIn("did not render", why)

    # -- other redirects ---------------------------------------------------
    def test_a_redirect_to_another_github_path_is_rejected(self):
        ok, why = self.verify("https://github.com/explore")
        self.assertFalse(ok)
        self.assertIn("redirected away from /search", why)

    def test_a_search_of_the_wrong_type_is_rejected(self):
        ok, why = self.verify(
            "https://github.com/search?q=University+Student&type=repositories")
        self.assertFalse(ok)
        self.assertIn("not type=users", why)

    def test_a_search_with_no_type_is_rejected(self):
        ok, why = self.verify("https://github.com/search?q=University+Student")
        self.assertFalse(ok)
        self.assertIn("not type=users", why)

    def test_leaving_github_is_rejected(self):
        for url in ("https://example.com/search?type=users", "", None):
            ok, why = self.verify(url)
            self.assertFalse(ok, url)
            self.assertIn("left GitHub", why)

    # -- zero results is not a page-2 subject ------------------------------
    def test_a_zero_result_search_stops_with_an_explanation(self):
        """Correct page, but there is no page 2 to diagnose."""
        ok, why = self.verify(
            SEARCH, (PageState.ZERO_RESULTS, "GitHub reports no users"))
        self.assertFalse(ok)
        self.assertIn("ZERO users", why)
        self.assertIn("no page 2", why)


class TestOrderingRegression(unittest.TestCase):
    """The specific mistake must not come back."""

    def test_the_login_check_navigates_and_so_must_come_first(self):
        """Documents WHY the order matters, straight from the source."""
        import inspect

        import student_scraper as browser_manager
        source = inspect.getsource(browser_manager.BrowserManager
                                   .github_login_name)
        self.assertIn("page.goto", source,
                      "github_login_name() navigates; if that ever stops "
                      "being true this test can be retired")
        self.assertIn("https://github.com", source)

    def test_the_diagnostic_checks_login_before_navigating_to_the_search(self):
        import inspect
        source = inspect.getsource(D.Diagnostic.phase_open_search)
        login_at = source.index("github_login_name")
        goto_at = source.index("page.goto")
        self.assertLess(login_at, goto_at,
                        "the login check must precede the search navigation, "
                        "because it navigates the tab itself")

    def test_phase_a_does_not_navigate_at_all(self):
        """Phase A inspects the verified page; it must not re-navigate."""
        import inspect
        source = inspect.getsource(D.Diagnostic.phase_page_one)
        self.assertNotIn("page.goto", source)
        self.assertNotIn("github_login_name", source)

    def test_the_run_stops_when_the_landing_is_wrong(self):
        import inspect
        source = inspect.getsource(D.mode_diagnose_page2)
        self.assertIn("phase_open_search", source)
        self.assertIn("if landed is None", source)
        self.assertLess(source.index("phase_open_search"),
                        source.index("phase_page_one"),
                        "the search must be confirmed before Phase A")

    def test_the_four_required_lines_are_printed(self):
        import inspect
        source = inspect.getsource(D.Diagnostic.phase_open_search)
        for label in ("Search query", "Search URL", "Final browser URL",
                      "Page classification"):
            self.assertIn(label, source, label)


if __name__ == "__main__":
    unittest.main(verbosity=2)
