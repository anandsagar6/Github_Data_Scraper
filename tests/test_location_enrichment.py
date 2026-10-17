"""City -> State -> Country enrichment, and refusing to guess.

Every "must stay blank" case here is a place name that genuinely exists in
more than one country.  Blank beats a coin flip.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                   # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (Candidate, SourceBundle, StudentMerger)   # noqa: E402


def record_for(location="", portfolio_text=""):
    candidate = Candidate(username="u",
                          profile_url="https://github.com/u",
                          keyword="k", source_file="f.xlsx")
    bundle = SourceBundle(candidate=candidate)
    bundle.github_ok = True
    bundle.github_fields = {"location": location, "links": [], "readme": "",
                            "bio": "", "fullname": "U", "email": ""}
    bundle.github_text = ""
    bundle.portfolio_text = portfolio_text
    return StudentMerger().build(bundle)


class TestEnrichment(unittest.TestCase):
    def test_the_stated_examples(self):
        cases = {
            "Bantul": ("Bantul", "Special Region of Yogyakarta",
                       "Indonesia"),
            "Bhopal": ("Bhopal", "Madhya Pradesh", "India"),
            "Pune": ("Pune", "Maharashtra", "India"),
        }
        for raw, expected in cases.items():
            self.assertEqual(student_scraper.parse_location(raw).as_tuple(), expected, raw)

    def test_city_supplies_state_and_country(self):
        for raw in ("Thrissur", "Bengaluru", "Surabaya", "Toronto",
                    "Houston"):
            parsed = student_scraper.parse_location(raw)
            self.assertTrue(parsed.city, raw)
            self.assertTrue(parsed.country, raw)

    def test_state_supplies_country(self):
        for raw, country in (("Madhya Pradesh", "India"),
                             ("Special Region of Yogyakarta", "Indonesia"),
                             ("Ontario", "Canada"), ("Bavaria", "Germany")):
            self.assertEqual(student_scraper.parse_location(raw).country, country, raw)

    def test_space_separated_locations(self):
        """Real GitHub locations are not reliably comma-separated."""
        cases = {
            "Nairobi Kenya": ("Nairobi", "", "Kenya"),
            "Aurangabad Bihar": ("Aurangabad", "Bihar", "India"),
            "Bahawalnagar Punjab Pakistan": ("Bahawalnagar", "Punjab (PK)",
                                             "Pakistan"),
        }
        for raw, expected in cases.items():
            self.assertEqual(student_scraper.parse_location(raw).as_tuple(), expected, raw)

    def test_accented_names_resolve(self):
        self.assertEqual(student_scraper.parse_location("Maceió- Al").country, "Brazil")
        self.assertEqual(
            student_scraper.parse_location("Balneário Camboriú").country, "Brazil")

    def test_country_only_string(self):
        parsed = student_scraper.parse_location("Belgium")
        self.assertEqual(parsed.country, "Belgium")
        self.assertEqual(parsed.city, "")


class TestAmbiguity(unittest.TestCase):
    AMBIGUOUS = ["Victoria", "Cambridge", "Salem", "Hyderabad",
                 "Birmingham", "London", "San Jose"]

    def test_ambiguous_names_get_no_country(self):
        for raw in self.AMBIGUOUS:
            parsed = student_scraper.parse_location(raw)
            self.assertEqual(parsed.country, "", raw)
            self.assertEqual(parsed.state, "", raw)
            self.assertTrue(any(note.startswith("ambiguous")
                                for note in parsed.notes), raw)

    def test_ambiguous_names_still_keep_the_city(self):
        for raw in ("Victoria", "Cambridge", "Salem"):
            self.assertTrue(student_scraper.parse_location(raw).city, raw)

    def test_context_resolves_ambiguity(self):
        cases = {
            "Cambridge, UK": "United Kingdom",
            "Cambridge, Massachusetts": "United States",
            "Victoria, Canada": "Canada",
            "Victoria, Australia": "Australia",
            "London, Ontario": "Canada",
            "Salem, Oregon": "United States",
            "Salem, Tamil Nadu": "India",
            "Hyderabad, India": "India",
            "Hyderabad, Pakistan": "Pakistan",
        }
        for raw, country in cases.items():
            self.assertEqual(student_scraper.parse_location(raw).country, country, raw)

    def test_agreeing_readings_are_not_ambiguous(self):
        """Every Springfield in the table is in the United States."""
        self.assertEqual(student_scraper.parse_location("Springfield").country,
                         "United States")

    def test_region_context_beats_default_reading(self):
        self.assertEqual(student_scraper.parse_location("Aurangabad Bihar").state, "Bihar")
        self.assertEqual(student_scraper.parse_location("Aurangabad Maharashtra").state,
                         "Maharashtra")


class TestNeverGuesses(unittest.TestCase):
    def test_non_places(self):
        for raw in ("Remote", "Earth", "Worldwide", "localhost", "",
                    "somewhere over the rainbow"):
            self.assertEqual(student_scraper.parse_location(raw).country, "", raw)

    def test_unknown_city_keeps_country_blank(self):
        parsed = student_scraper.parse_location("Nowhereville")
        self.assertEqual(parsed.country, "")
        self.assertIn("city not in gazetteer", parsed.notes)

    def test_india_is_never_a_default(self):
        for raw in ("Trikala, Greece", "Lagos", "Dhaka", "Nowhereville"):
            self.assertNotEqual(student_scraper.parse_location(raw).country, "India", raw)


class TestPriority(unittest.TestCase):
    def test_github_wins_and_gaps_are_filled(self):
        best = student_scraper.best_location([("Bantul, Special Region of Yogyakarta",
                                 1.00),
                                ("Mumbai, Maharashtra", 0.80)])
        self.assertEqual(best.as_tuple(),
                         ("Bantul", "Special Region of Yogyakarta",
                          "Indonesia"))

    def test_lower_priority_fills_only_what_is_missing(self):
        # A bare ambiguous city first, an explicit country second.
        best = student_scraper.best_location([("Hyderabad", 1.00), ("India", 0.95)])
        self.assertEqual(best.city, "Hyderabad")
        self.assertEqual(best.country, "India")

    def test_merged_result_is_enriched(self):
        best = student_scraper.best_location([("Pune", 1.00)])
        self.assertEqual(best.state, "Maharashtra")
        self.assertEqual(best.country, "India")


class TestFinalColumns(unittest.TestCase):
    def test_columns_and_aliases(self):
        record = record_for("Bantul, Special Region of Yogyakarta")
        self.assertEqual(record["Current City"], "Bantul")
        self.assertEqual(record["State"], "Special Region of Yogyakarta")
        self.assertEqual(record["Country"], "Indonesia")
        self.assertEqual(record["country"], record["Country"])

    def test_pin_code_mirrors(self):
        record = record_for(
            "Bhopal", "Address: MG Road, Bhopal, PIN 462001")
        self.assertEqual(record["PIN Code"], "462001")
        self.assertEqual(record["pin_code"], record["PIN Code"])

    def test_ambiguity_is_reported_in_remarks(self):
        record = record_for("Hyderabad")
        self.assertEqual(record["Country"], "")
        self.assertIn("ambiguous", record["Remarks / Notes"].lower())
        self.assertIn("Hyderabad", record["Current City"])

    def test_schema_is_unchanged(self):
        record = record_for("Pune")
        self.assertEqual(list(record.keys()), C.FINAL_COLUMNS)
        self.assertEqual(len(record), 65)


if __name__ == "__main__":
    unittest.main(verbosity=2)
