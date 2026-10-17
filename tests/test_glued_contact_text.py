"""Contact details that a page runs together must still come out clean.

THE LIVE CASE
-------------
A real portfolio renders the contact block as icons and text in adjacent
inline elements, so the rendered text has no separator between the fields:

    82830-89459eshanvohra@gmail.comFirozpur, Punjab, India

inner_text() is correct to return that -- it is what the page shows -- but
EMAIL_RE matched the WHOLE run, because "82830-89459eshanvohra" is a legal
local part and `[A-Za-z]{2,}` is happy to read "comFirozpur" as a TLD.  The
address reached the spreadsheet as

    82830-89459eshanvohra@gmail.comfirozpur

which is not an address at all.  Nothing failed; the row simply had an email
nobody could write to, and the only way to notice was to read it.

THE RULE
--------
Two trims, each safe on its own:

    the TLD        cut at the longest spelling that really is a TLD, and only
                   when the surplus is plain letters
    the local part drop a leading phone number, and only when it is
                   phone-shaped and letters follow it immediately

A tail that already IS a TLD is never touched.  That is the case this file
guards hardest, because getting it wrong corrupts every ordinary address:
"com" starts with "co", and an over-eager trim turns gmail.com into gmail.co
for everyone.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper as S                               # noqa: E402

# Exactly what the page serves, verified by loading it.
LIVE_TEXT = "82830-89459eshanvohra@gmail.comFirozpur, Punjab, India"


class GluedEmailCase(unittest.TestCase):
    def test_the_live_case(self):
        self.assertEqual(S.extract_emails(LIVE_TEXT),
                         ["eshanvohra@gmail.com"])

    def test_a_phone_glued_to_the_front_is_dropped(self):
        self.assertEqual(S.extract_emails("9876543210rahul@gmail.com"),
                         ["rahul@gmail.com"])
        self.assertEqual(S.extract_emails("+91 98765-43210priya@yahoo.in"),
                         ["priya@yahoo.in"])

    def test_prose_glued_to_the_tld_is_cut(self):
        for text, want in (
                ("a@b.comDelhi", "a@b.com"),
                ("dev@site.ioBangalore", "dev@site.io"),
                ("me@x.devPune", "me@x.dev"),
                ("name@dept.edu.inDelhi", "name@dept.edu.in"),
                ("x@y.co.ukLondon", "x@y.co.uk")):
            self.assertEqual(S.extract_emails(text), [want], text)

    # -- the trims must never damage an ordinary address ------------------
    def test_a_clean_address_is_untouched(self):
        for value in ("eshanvohra@gmail.com", "a@b.com",
                      "first.last+tag@sub.domain.org",
                      "student@college.ac.in", "dev@startup.io",
                      "me@mysite.dev", "x@y.co.uk", "a@b.museum"):
            self.assertEqual(S.extract_emails(value), [value], value)

    def test_com_is_not_shortened_to_co(self):
        """The regression this guard exists for."""
        self.assertEqual(S.extract_emails("someone@gmail.com"),
                         ["someone@gmail.com"])
        self.assertNotIn("someone@gmail.co",
                         S.extract_emails("someone@gmail.com"))

    def test_a_local_part_that_really_starts_with_digits_survives(self):
        """A batch year is not a phone number."""
        for value in ("2024batch@college.ac.in", "123@x.com",
                      "4you@gmail.com"):
            self.assertEqual(S.extract_emails(value), [value], value)

    def test_separated_text_is_unaffected(self):
        self.assertEqual(
            S.extract_emails("call 9876543210 or mail a@b.com today"),
            ["a@b.com"])

    def test_the_email_trim_leaves_phone_extraction_alone(self):
        """The trim must not change what extract_phones() sees.

        The phone on that row (8283089459) did NOT come from this glued run --
        extract_phones() rejects it, because a number welded to an email has
        no phone context around it.  That is deliberate and unchanged here;
        the row's phone came from another source.  What matters is that the
        email trim did not alter phone extraction either way.
        """
        before = S.extract_phones(LIVE_TEXT)
        S.extract_emails(LIVE_TEXT)
        self.assertEqual(S.extract_phones(LIVE_TEXT), before)

    def test_trim_is_a_no_op_on_something_that_is_not_an_email(self):
        self.assertEqual(S.trim_glued_email("not-an-email"), "not-an-email")
        self.assertEqual(S.trim_glued_email(""), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
