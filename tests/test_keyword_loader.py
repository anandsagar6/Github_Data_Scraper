"""Keyword loading: only real search phrases get through.

The bug this guards against: the title line of the keyword file,
"INDIA COLLEGE STUDENT & FRESHER TARGET KEYWORDS", was being searched on
GitHub as though it were a keyword.
"""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                   # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (keyword_rejections, load_base_keywords, load_keywords, rejection_reason, slugify)   # noqa: E402

TITLE = "INDIA COLLEGE STUDENT & FRESHER TARGET KEYWORDS"


class TestRejectionRules(unittest.TestCase):
    def test_the_reported_title_is_rejected(self):
        self.assertEqual(rejection_reason(TITLE), "all-caps title")

    def test_empty_lines(self):
        for line in ("", "   ", "\t", "\n"):
            self.assertEqual(rejection_reason(line), "empty line", repr(line))

    def test_hash_headings(self):
        for line in ("# Heading", "## Heading", "### Telangana",
                     "#### Deep", "   ### Indented"):
            self.assertEqual(rejection_reason(line), "markdown heading", line)

    def test_separator_rules(self):
        for line in ("=====", "-----", "_____", "*****", "~~~~~",
                     "=" * 55, "   ------   "):
            self.assertEqual(rejection_reason(line), "separator rule", line)

    def test_metadata_lines(self):
        for line in ("Format: <Base Keyword> <City>", "Purpose: testing",
                     "Country: India", "State: Telangana", "City: Pune",
                     "Notes: none", "Note: none", "Source: somewhere",
                     "Version: 2", "Updated: today", "Author: me",
                     "Total: 3000", "Category: students"):
            self.assertIn(rejection_reason(line),
                          ("metadata line", "format placeholder"), line)

    def test_placeholder_lines(self):
        self.assertEqual(rejection_reason("<Base Keyword> <City>"),
                         "format placeholder")

    def test_section_headings(self):
        for line in ("Education:", "Skills:", "For IT Support / Help Desk",
                     "Our Target Audience For This Job Portal"):
            self.assertEqual(rejection_reason(line), "section heading", line)

    def test_comment_and_aside_lines(self):
        for line in ("*2nd level from azure admin course *", "// comment",
                     "/* block */"):
            self.assertEqual(rejection_reason(line), "comment or aside", line)

    def test_lines_without_letters(self):
        for line in ("12345", "...", "!!!"):
            self.assertIn(rejection_reason(line),
                          ("no letters", "separator rule", "decoration only"),
                          line)

    def test_overlong_lines_are_not_search_phrases(self):
        long_line = " ".join(["Word"] * 15)
        self.assertEqual(rejection_reason(long_line),
                         "too long to be a search phrase")


class TestValidKeywords(unittest.TestCase):
    VALID = [
        "College Student Hyderabad",
        "B.Tech Student Hyderabad",
        "Computer Science Student Hyderabad",
        "Fresher Hyderabad",
        "Software Engineer Intern Hyderabad",
        "Pre-Final Year Student Hyderabad",
        "BTech CSE Hyderabad",
        "Graduate Trainee Gangtok",
        "Desktop Support Engineer",
    ]

    def test_real_keywords_pass(self):
        for line in self.VALID:
            self.assertEqual(rejection_reason(line), "", line)

    def test_single_all_caps_tokens_survive(self):
        """BCA/MCA/MBA/BBA are degrees, not titles."""
        for line in ("BCA", "MCA", "MBA", "BBA", "B.Tech.", "M.Tech.",
                     "Bsc.", "Diploma"):
            self.assertEqual(rejection_reason(line), "", line)

    def test_all_caps_rule_needs_two_words(self):
        self.assertEqual(rejection_reason("MCA"), "")
        self.assertEqual(rejection_reason("MCA STUDENT PUNE"),
                         "all-caps title")


class TestLoading(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="keywords_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _write(self, text, name="kw.txt"):
        path = os.path.join(self.dir, name)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(text)
        return path

    def test_the_real_file_shape(self):
        path = self._write(
            TITLE + "\n"
            "=======================================================\n"
            "\n"
            "Format: <Base Keyword> <City>\n"
            "\n"
            "\n"
            "### Telangana\n"
            "\n"
            "College Student Hyderabad\n"
            "University Student Hyderabad\n"
            "Engineering Student Hyderabad\n"
            "Computer Science Student Hyderabad\n"
            "\n"
            "### Karnataka\n"
            "\n"
            "College Student Bengaluru\n")
        self.assertEqual(load_keywords([path]), [
            "College Student Hyderabad",
            "University Student Hyderabad",
            "Engineering Student Hyderabad",
            "Computer Science Student Hyderabad",
            "College Student Bengaluru",
        ])

    def test_order_is_preserved(self):
        path = self._write("Zebra Student Pune\nAlpha Student Pune\n"
                           "Middle Student Pune\n")
        self.assertEqual(load_keywords([path]),
                         ["Zebra Student Pune", "Alpha Student Pune",
                          "Middle Student Pune"])

    def test_duplicates_are_removed_keeping_first_position(self):
        path = self._write("College Student Pune\nFresher Pune\n"
                           "college student pune\nCollege Student Pune\n")
        self.assertEqual(load_keywords([path]),
                         ["College Student Pune", "Fresher Pune"])

    def test_multiple_files_are_concatenated_and_deduped(self):
        first = self._write("College Student Pune\nFresher Pune\n", "a.txt")
        second = self._write("Fresher Pune\nMCA\n", "b.txt")
        self.assertEqual(load_keywords([first, second]),
                         ["College Student Pune", "Fresher Pune", "MCA"])

    def test_rejections_are_reported_with_reasons(self):
        path = self._write(TITLE + "\n=====\nFormat: <A> <B>\n"
                           "### Telangana\nCollege Student Pune\n")
        reasons = {reason for _f, _n, _t, reason
                   in keyword_rejections([path])}
        # "Format: <A> <B>" matches the metadata rule first; the
        # placeholder rule catches a bare "<A> <B>" line.
        self.assertEqual(reasons, {"all-caps title", "separator rule",
                                   "metadata line", "markdown heading"})
        bare = self._write("<Base Keyword> <City>\n", "bare.txt")
        self.assertEqual(
            [reason for _f, _n, _t, reason in keyword_rejections([bare])],
            ["format placeholder"])

    def test_missing_file_is_ignored(self):
        self.assertEqual(load_keywords(["/no/such/file.txt"]), [])


class TestProductionFile(unittest.TestCase):
    """The live keyword list must contain no metadata keywords.

    It is loaded from new_mylib.py's keyword_by_category -- the same file the
    cities come from.  Cities are NOT part of a keyword: Python crosses each
    keyword with locations.py at plan time, so no entry here holds a city.
    """

    def setUp(self):
        self.keywords = load_base_keywords()

    def test_the_source_is_new_mylib(self):
        import student_scraper as KL
        self.assertEqual(KL.KEYWORD_SOURCE_FILE, "new_mylib.py")
        self.assertEqual(KL.KEYWORD_SOURCE_DICT, "keyword_by_category")
        groups = KL.load_keyword_groups()
        self.assertEqual(len(groups), 19)
        self.assertEqual(len(self.keywords), 3450)

    def test_groups_keep_their_priority_order(self):
        """IT1 first: a keyword in two groups keeps the higher priority."""
        import student_scraper as KL
        names = list(KL.load_keyword_groups())
        self.assertTrue(names[0].startswith("IT1"), names[0])
        self.assertTrue(names[-1].startswith("Shared"), names[-1])

    def test_a_missing_source_is_a_clear_error(self):
        import student_scraper as KL
        with self.assertRaises(KL.BaseKeywordError):
            KL.load_keyword_groups("/no/such/new_mylib.py")

    def test_no_noise_survives(self):
        self.assertTrue(self.keywords)
        for keyword in self.keywords:
            self.assertNotEqual(keyword, TITLE)
            self.assertFalse(keyword.startswith("#"), keyword)
            self.assertFalse(keyword.startswith("Format:"), keyword)
            self.assertNotIn("<", keyword)
            self.assertTrue(keyword.strip(), repr(keyword))
            self.assertNotEqual(keyword.strip("=-_* "), "", keyword)

    def test_expected_beginning(self):
        # Pure keywords, no city appended; IT1 courses come first.
        self.assertEqual(self.keywords[:4], [
            "B.Tech Computer Science",
            "B.Tech Computer Science and Engineering",
            "B.Tech Information Technology",
            "B.Tech Computer Engineering",
        ])

    def test_every_keyword_slugifies_to_something_usable(self):
        for keyword in self.keywords:
            slug = slugify(keyword)
            self.assertTrue(slug)
            self.assertNotEqual(slug, "keyword", keyword)
            self.assertNotIn("/", slug)

    def test_no_duplicates(self):
        lowered = [keyword.lower() for keyword in self.keywords]
        self.assertEqual(len(lowered), len(set(lowered)))


if __name__ == "__main__":
    unittest.main(verbosity=2)
