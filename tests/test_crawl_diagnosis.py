"""'Table not found' is a symptom with four different causes.

The fixtures below are real captures, not invented shapes.  RATE_LIMIT_429
is exactly what was logged when a 28-page crawl of "College Student" was
throttled at page 29 and 290 collected rows were discarded.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                       # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper                                 # noqa: E402
from student_scraper import (CrawlOutcome, blocked_marker, classify_page_state)   # noqa: E402


def state(**overrides):
    base = {
        "title": "User search results",
        "bodyHead": "College Student users Search Results",
        "bodyTail": "",
        "hasNext": True,
        "nextDisabled": False,
        "nextHref": "https://github.com/search?q=x&type=users&p=2",
        "profileLinkCount": 12,
        "resultsListPresent": True,
        "resultsListChildren": 10,
        "noResults": False,
        "saysError": False,
        "blocked_marker": None,
        "ext_pages": 28,
        "ext_rows": 271,
        "ext_error": "Table not found",
    }
    base.update(overrides)
    return base


# The exact capture from logs/student_scraper.log at 13:20:35.
RATE_LIMIT_429 = state(
    title="Error 429",
    bodyHead="Skip to content College Student Unable to load page. "
             "Status: 429 Message: Please reload page and try again",
    bodyTail="Skip to content College Student Unable to load page. "
             "Status: 429 Message: Please reload page and try again",
    hasNext=False, nextDisabled=None, nextHref="",
    profileLinkCount=2, resultsListPresent=False, resultsListChildren=0,
    saysError=True, blocked_marker="unable to load page",
    ext_pages=29, ext_rows=290)


class TestClassification(unittest.TestCase):
    # ---------------------------------------------- 2. rate limit / error
    def test_real_429_capture_is_a_rate_limit(self):
        outcome, why = classify_page_state(RATE_LIMIT_429)
        self.assertEqual(outcome, CrawlOutcome.RATE_LIMITED)
        self.assertIn("429", why)

    def test_429_detected_even_without_the_marker(self):
        """The title alone must be enough."""
        capture = dict(RATE_LIMIT_429)
        capture["blocked_marker"] = None
        capture["bodyHead"] = capture["bodyTail"] = "Skip to content"
        outcome, _why = classify_page_state(capture)
        self.assertEqual(outcome, CrawlOutcome.RATE_LIMITED)

    def test_body_status_line_alone_is_enough(self):
        capture = state(title="", hasNext=False, resultsListPresent=False,
                        resultsListChildren=0, profileLinkCount=0,
                        bodyHead="Unable to load page. Status: 503 "
                                 "Message: Please reload page")
        outcome, why = classify_page_state(capture)
        self.assertEqual(outcome, CrawlOutcome.RATE_LIMITED)
        self.assertIn("503", why)

    def test_classic_rate_limit_wording_still_works(self):
        outcome, _why = classify_page_state(
            state(blocked_marker="secondary rate limit"))
        self.assertEqual(outcome, CrawlOutcome.RATE_LIMITED)

    def test_error_page_is_not_mistaken_for_the_last_page(self):
        """An error page also has no results and no Next link."""
        outcome, _why = classify_page_state(RATE_LIMIT_429)
        self.assertNotEqual(outcome, CrawlOutcome.END_OF_RESULTS)

    def test_new_markers_are_recognised_by_blocked_marker(self):
        body = ("Unable to load page. Status: 429 "
                "Message: Please reload page and try again")
        self.assertTrue(blocked_marker(body))

    # -------------------------------------------------- 1. end of results
    def test_no_users_message(self):
        outcome, _why = classify_page_state(
            state(noResults=True, hasNext=False, resultsListChildren=0,
                  profileLinkCount=0, resultsListPresent=False))
        self.assertEqual(outcome, CrawlOutcome.END_OF_RESULTS)

    def test_last_page_with_results_but_no_next(self):
        outcome, _why = classify_page_state(
            state(hasNext=False, nextDisabled=None, resultsListChildren=7))
        self.assertEqual(outcome, CrawlOutcome.END_OF_RESULTS)

    def test_disabled_next_is_the_last_page(self):
        outcome, _why = classify_page_state(
            state(hasNext=True, nextDisabled=True, resultsListChildren=4))
        self.assertEqual(outcome, CrawlOutcome.END_OF_RESULTS)

    def test_end_of_results_is_exportable(self):
        self.assertIn(CrawlOutcome.END_OF_RESULTS, CrawlOutcome.EXPORTABLE)
        self.assertIn(CrawlOutcome.FINISHED, CrawlOutcome.EXPORTABLE)
        self.assertNotIn(CrawlOutcome.RATE_LIMITED, CrawlOutcome.EXPORTABLE)
        self.assertNotIn(CrawlOutcome.TABLE_LOST, CrawlOutcome.EXPORTABLE)

    # ------------------------------------------------ 3. lost table state
    def test_valid_users_page_is_a_lost_table(self):
        outcome, why = classify_page_state(state())
        self.assertEqual(outcome, CrawlOutcome.TABLE_LOST)
        self.assertIn("lost table detection", why)

    def test_lost_table_recognised_from_profile_links_alone(self):
        outcome, _why = classify_page_state(
            state(resultsListPresent=False, resultsListChildren=0,
                  profileLinkCount=11))
        self.assertEqual(outcome, CrawlOutcome.TABLE_LOST)

    # -------------------------------------------- 4. extension's own fault
    def test_extension_error_when_the_page_has_nothing(self):
        outcome, why = classify_page_state(
            state(resultsListPresent=False, resultsListChildren=0,
                  profileLinkCount=1, ext_error="Something broke"))
        self.assertEqual(outcome, CrawlOutcome.ERROR)
        self.assertIn("Something broke", why)


class TestPacingAndMarkers(unittest.TestCase):
    def test_crawl_delay_raised_above_the_extension_default(self):
        """1 s over 28 pages drew a 429; the floor is now higher."""
        self.assertGreaterEqual(C.CRAWL_MIN_DELAY_SEC, 3)

    def test_recovery_budget_is_bounded(self):
        self.assertGreaterEqual(C.CRAWL_TABLE_RECOVERIES, 1)
        self.assertLessEqual(C.CRAWL_TABLE_RECOVERIES, 5)

    def test_429_is_a_blocking_status(self):
        self.assertIn(429, C.BLOCK_STATUS)


if __name__ == "__main__":
    unittest.main(verbosity=2)
