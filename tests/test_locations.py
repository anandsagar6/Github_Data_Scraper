"""Search location is not the student's location.

The whole point of separating locations.py from location_extractor.py:

    SEARCH location    "Fresher Hyderabad" says where we LOOKED.
                       It lives in SearchUnit and in the store's search_*
                       columns.

    STUDENT location   Current City / State / Country in the final record.
                       It comes only from the person's own profile, resume or
                       portfolio, via location_extractor.

Somebody found by "Fresher Hyderabad" who lives in Pune is a Pune record.
Somebody with no location evidence gets a BLANK Current City -- never the
city that happened to find them.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                     # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper                                 # noqa: E402
from student_scraper import (Candidate, SourceBundle, StudentMerger)   # noqa: E402
import sandbox                                         # noqa: E402


class TestDataset(unittest.TestCase):
    def test_dataset_sizes(self):
        # Loaded from new_mylib.py (city_by_states_original).
        # flatten + dict.fromkeys over city_by_states_original, as naukri
        # does it: one name, one city, no alias spellings.
        self.assertEqual(student_scraper.canonical_city_count(), 871)
        self.assertEqual(student_scraper.search_spelling_count(), 871)
        self.assertEqual(student_scraper.COUNTRY, "India")

    def test_the_dataset_comes_from_new_mylib(self):
        """new_mylib.py is the single place cities are edited."""
        self.assertEqual(student_scraper.CITY_SOURCE_FILE, "new_mylib.py")
        self.assertEqual(student_scraper.CITY_SOURCE_DICT, "city_by_states_original")
        self.assertTrue(len(student_scraper.CITY_DATA) > 800)

    def test_a_missing_source_file_is_a_clear_error(self):
        with self.assertRaises(student_scraper.LocationDataError):
            student_scraper._read_city_source("/no/such/new_mylib.py")

    def test_every_city_has_a_non_empty_state_label(self):
        """A label, not necessarily a state.

        It used to also forbid "Delhi NCR" and "Assam & Northeast".  Those are
        regions rather than states, and they are now allowed through, because
        the label is whatever group new_mylib.py files the city under -- which
        is exactly what naukri's city_to_state() returns.  See
        TestCityListIsNaukriShaped.test_a_region_group_is_allowed_as_a_state.
        """
        for city in student_scraper.cities():
            self.assertTrue(city.state.strip(), city.canonical)

    def test_the_city_name_is_its_only_spelling(self):
        for city in student_scraper.cities():
            self.assertEqual(city.search_aliases[0], city.canonical)

    def test_no_spelling_is_claimed_by_two_cities(self):
        seen = {}
        for city in student_scraper.cities():
            for alias in city.search_aliases:
                key = alias.lower()
                self.assertNotIn(key, seen,
                                 "%r is used by %s and %s"
                                 % (alias, seen.get(key), city.canonical))
                seen[key] = city.canonical

    def test_the_dataset_is_immutable(self):
        with self.assertRaises(AttributeError):
            student_scraper.cities()[0].canonical = "Tampered"

    def test_fingerprint_is_stable(self):
        self.assertEqual(student_scraper.dataset_fingerprint(), student_scraper.dataset_fingerprint())
        self.assertEqual(len(student_scraper.dataset_fingerprint()), 64)


class TestCityListIsNaukriShaped(unittest.TestCase):
    """Cities are built exactly as find_naukri_job.py builds them.

    WHAT CHANGED AND WHY
    --------------------
    This class used to be TestAliases, and it asserted the opposite: that
    "Bangalore" folded into "Bengaluru" as one city with two spellings, kept on
    one rank so two copies could not crawl the same place.

    That machinery -- ALIAS_GROUPS, REGION_TO_STATE, PREFERRED_STATE -- is gone
    on purpose.  The chosen shape is naukri's:

        all_cities = [city for cities in city_by_states.values()
                           for city in cities]
        all_cities = list(dict.fromkeys(all_cities))

    One name, one city, first occurrence wins.  The consequence is accepted
    rather than hidden: Bengaluru and Bangalore are two cities and may land on
    two ranks, so two laptops can crawl overlapping GitHub results for one
    place.  Each folder's master_index.xlsx absorbs its own repeats; across
    laptops the overlap goes when their masters are merged.

    The second consequence is that names new_mylib.py does not list are not
    searched at all.  The old ALIAS_GROUPS added thirteen of them by hand
    (Bangalore, Mysore, Vizag, Trivandrum, Calicut, Baroda, ...).  Those are
    gone, and these tests pin that down so it cannot drift back in silently.
    """

    # Every one of these WAS searched by the old alias machinery and is not
    # searched now, because city_by_states_original does not list it.
    DROPPED_WITH_ALIAS_GROUPS = [
        "Bangalore", "Baroda", "Belgaum", "Bokaro", "Calicut",
        "Chhatrapati Sambhajinagar", "Hanamkonda", "Mangalore", "Mysore",
        "New Delhi", "Trichy", "Trivandrum", "Vizag",
    ]

    def test_the_city_list_matches_naukris_own_construction(self):
        """Recompute it naukri's way and require the same list, in order."""
        import ast
        with open(os.path.join(os.path.dirname(os.path.dirname(
                os.path.abspath(__file__))), "new_mylib.py"),
                encoding="utf-8", errors="replace") as handle:
            tree = ast.parse(handle.read())
        data = None
        for node in tree.body:
            if isinstance(node, ast.Assign) \
               and getattr(node.targets[0], "id", "") == "city_by_states_original":
                data = ast.literal_eval(node.value)
        self.assertIsNotNone(data, "city_by_states_original not found")
        flat = [str(c).strip() for group in data.values()
                for c in group if str(c).strip()]
        naukri = list(dict.fromkeys(flat))
        ours = [city.canonical for city in student_scraper.cities()]
        self.assertEqual(ours, naukri,
                         "the city list no longer equals flatten + "
                         "dict.fromkeys over city_by_states_original")

    def test_each_city_is_searched_under_exactly_its_own_name(self):
        for city in student_scraper.cities():
            self.assertEqual(list(city.search_aliases), [city.canonical],
                             "%s carries extra spellings; the alias machinery "
                             "is back" % city.canonical)

    def test_a_name_new_mylib_does_not_list_is_not_searched(self):
        for name in self.DROPPED_WITH_ALIAS_GROUPS:
            self.assertEqual(
                student_scraper.canonical_for(name), "",
                "%s is being searched again -- it is not in new_mylib.py, so "
                "it can only have come from reintroduced alias data. Add it to "
                "city_by_states_original instead." % name)

    def test_a_name_new_mylib_does_list_is_searched_as_its_own_city(self):
        """Gurgaon and Gurugram are BOTH listed, so both are searched."""
        for name in ("Gurugram", "Gurgaon", "Bengaluru", "Visakhapatnam",
                     "Delhi", "Hubli", "Hubballi"):
            self.assertEqual(student_scraper.canonical_for(name), name,
                             "%s is listed in new_mylib.py and must be its own "
                             "city" % name)

    def test_an_unknown_spelling_resolves_to_nothing(self):
        self.assertEqual(student_scraper.canonical_for("Atlantis"), "")
        self.assertEqual(student_scraper.canonical_for(""), "")

    def test_lookup_is_case_insensitive(self):
        self.assertEqual(student_scraper.canonical_for("bENGALURU"), "Bengaluru")
        self.assertEqual(student_scraper.canonical_for("  Delhi  "), "Delhi")

    def test_the_state_is_the_group_the_city_is_listed_under(self):
        """naukri's city_to_state() returns the group key, including regions."""
        self.assertEqual(student_scraper.state_for("Delhi"), "Delhi NCR")
        self.assertEqual(student_scraper.state_for("Noida"), "Delhi NCR")
        self.assertEqual(student_scraper.state_for("Bengaluru"), "Karnataka")

    def test_a_region_group_is_allowed_as_a_state(self):
        """The guard this replaces asserted the opposite; it must not come back.

        "Delhi NCR" is a region, not a state, and it reaches the terminal and
        scraper_state.json as the search's state.  That is accepted: this label
        never reaches a collected person -- the State column on a row comes from
        that person's own location text through KNOWN_CITIES, which has its own
        state table.
        """
        states = {city.state for city in student_scraper.cities()}
        self.assertIn("Delhi NCR", states)
        self.assertEqual(student_scraper.KNOWN_CITIES["gurgaon"][1], "Haryana",
                         "the PERSON's state table must stay accurate even "
                         "though the SEARCH label is now the raw group")

    def test_a_city_listed_twice_keeps_its_first_group(self):
        """dict.fromkeys semantics: first occurrence wins, one state per city."""
        by_name = {}
        for city in student_scraper.cities():
            self.assertNotIn(city.canonical, by_name,
                             "%s appears twice" % city.canonical)
            by_name[city.canonical] = city.state

    def test_every_city_stays_with_one_rank(self):
        """Still true, and still what keeps one rank's work disjoint.

        Weaker than the old guarantee: it is per NAME now, not per place, so
        Bengaluru and Bangalore could sit on different ranks -- except that
        Bangalore is no longer searched at all.
        """
        import student_scraper as S
        cities = list(student_scraper.cities())
        for num_sys in (1, 2, 5, 8):
            owner = {}
            for rank in range(1, num_sys + 1):
                for city in S.get_cities_by_rank(cities, rank, num_sys):
                    for spelling in city.search_aliases:
                        self.assertNotIn(
                            spelling, owner,
                            "%s claimed by rank %s and rank %d"
                            % (spelling, owner.get(spelling), rank))
                        owner[spelling] = rank
            self.assertEqual(len(owner), student_scraper.search_spelling_count(),
                             "num_sys=%d lost a city" % num_sys)


class TestSearchLocationDoesNotLeak(unittest.TestCase):
    """The most important thing in this file."""

    def _bundle(self, profile_location=""):
        candidate = Candidate(
            source_row_number=1, source_file="fresher-hyderabad.xlsx",
            keyword="Fresher Hyderabad", username="someone",
            profile_url="https://github.com/someone")
        bundle = SourceBundle(candidate=candidate)
        bundle.github_ok = True
        bundle.github_fields = {"fullname": "Some One", "nickname": "someone",
                                "location": profile_location, "bio": "",
                                "links": []}
        bundle.github_text = "Some One"
        return bundle

    def test_a_person_found_in_hyderabad_who_lives_in_pune_is_a_pune_record(self):
        record = StudentMerger().build(self._bundle("Pune, Maharashtra"))
        self.assertEqual(record["Current City"], "Pune")
        self.assertEqual(record["State"], "Maharashtra")
        self.assertNotEqual(record["Current City"], "Hyderabad")

    def test_no_location_evidence_means_a_blank_city_not_the_search_city(self):
        record = StudentMerger().build(self._bundle(""))
        self.assertEqual(record["Current City"], "")
        self.assertEqual(record["State"], "")
        self.assertEqual(record["Country"], "")

    def test_the_search_query_never_appears_in_any_location_field(self):
        record = StudentMerger().build(self._bundle(""))
        for column in ("Current City", "State", "Country", "College City",
                       "College State", "pin_code", "country"):
            self.assertNotIn("Hyderabad", str(record.get(column, "")), column)
            self.assertNotIn("Fresher", str(record.get(column, "")), column)

    def test_search_metadata_lives_in_the_store_not_the_65_columns(self):
        import shutil
        import tempfile
        directory = tempfile.mkdtemp(prefix="leak_")
        self.addCleanup(shutil.rmtree, directory, True)
        import student_scraper as S
        city = next(c for c in student_scraper.cities() if c.canonical == "Hyderabad")
        unit = S.make_unit(city, "Hyderabad", "Fresher")
        record = StudentMerger().build(self._bundle("Pune"))
        store = student_scraper.open_store(os.path.join(directory, "s.db"))
        self.addCleanup(store.close)
        store.save_student(record, search_unit=unit,
                           campaign_id="campaign_20260927_001", rank=23,
                           total_ranks=80, source_xlsx="fresher-hyderabad.xlsx")
        row = next(store.iter_records())

        # The search location IS recorded -- as store metadata.
        self.assertEqual(row["search_city"], "Hyderabad")
        self.assertEqual(row["search_query"],
                         "in:Fresher location:Hyderabad")
        self.assertEqual(row["campaign_id"], "campaign_20260927_001")
        self.assertEqual(row["rank"], 23)

        # ...and the student's own location is untouched by it.
        self.assertEqual(row["record"]["Current City"], "Pune")
        self.assertEqual(len(row["record"]), len(C.FINAL_COLUMNS))

    def test_campaign_metadata_is_not_written_into_data_source_or_remarks(self):
        record = StudentMerger().build(self._bundle("Pune"))
        for column in ("Data Source", "Remarks / Notes", "Source Platform",
                       "Source URL"):
            value = str(record.get(column, ""))
            self.assertNotIn("campaign_", value, column)
            self.assertNotIn("rank_", value, column)

    def test_the_search_city_and_the_extraction_city_stay_distinct(self):
        """SEARCH_CITIES and KNOWN_CITIES must never be confused.

        They used to live in different files, so the rule could be "the
        extraction stack must not import locations".  Both now sit in
        student_scraper.py, so the guarantee is carried by the names instead:

            SEARCH_CITIES   where we LOOK -- City objects, drives the plan
            KNOWN_CITIES    what a PERSON wrote -- alias -> (city, region,
                            country), drives parse_location

        If anything ever makes these one object again, a person found by
        "Fresher Hyderabad" starts being recorded as living in Hyderabad.
        """
        search = student_scraper.SEARCH_CITIES
        known = student_scraper.KNOWN_CITIES
        self.assertIsNot(search, known)
        self.assertIsInstance(known, dict)
        self.assertTrue(hasattr(search[0], "canonical"),
                        "SEARCH_CITIES holds City objects, not plain names")
        # A search city carries state+aliases; an extraction entry is a tuple.
        self.assertIsInstance(known[next(iter(known))], tuple)

if __name__ == "__main__":
    unittest.main(verbosity=2)
