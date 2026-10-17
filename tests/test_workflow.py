"""The duplicate index, XLSX reading and state resume.

Rotation tests moved to tests/test_final_export.py when the 100-rows-per-file
behaviour was removed.  The duplicate index is no longer three in-memory sets
rebuilt from workbooks -- it is the store, and the database enforces it -- so
those tests exercise result_store here.

TestXlsxReader is unchanged: the Instant Data Scraper export format did not
change and neither did the parser.

Everything here uses a throwaway directory -- never the real output folder.
"""

import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook                     # noqa: E402

import config as C                                # noqa: E402
import student_scraper                                           # noqa: E402
from student_scraper import (StateManager, XlsxReader)   # noqa: E402
import sandbox                                         # noqa: E402


class FakeCity:
    def __init__(self, state, canonical, aliases):
        self.country = "India"
        self.state = state
        self.canonical = canonical
        self.search_aliases = tuple(aliases)


CITIES = [
    FakeCity("Telangana", "Hyderabad", ["Hyderabad"]),
    FakeCity("Telangana", "Warangal", ["Warangal", "Hanamkonda"]),
    FakeCity("Karnataka", "Bengaluru", ["Bengaluru", "Bangalore"]),
    FakeCity("Kerala", "Kochi", ["Kochi"]),
]
KEYWORDS = ["B.Tech Student", "Fresher"]


def make_record(number, username=None):
    username = username or "person%d" % number
    record = {column: "" for column in C.FINAL_COLUMNS}
    record["Student Name"] = "Person %d" % number
    record["GitHub Profile"] = "https://github.com/%s" % username
    record["github"] = record["GitHub Profile"]
    # A real address, not example.org: placeholder domains are deliberately
    # excluded from identity, so an example.org address would prove nothing.
    record["Email Address"] = "person%d@gmail.com" % number
    record["Source Platform"] = "GitHub"
    return record


class TestDuplicateIndex(unittest.TestCase):
    """The identity rules, now enforced by the store's UNIQUE indexes."""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="store_test_")
        self.store = student_scraper.open_store(
            os.path.join(self.dir, "students.db"))

    def tearDown(self):
        self.store.close()
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_duplicate_by_username_profile_url_and_email(self):
        self.store.save_student(make_record(1))
        self.assertEqual(self.store.is_known(username="person1"), "username")
        self.assertEqual(
            self.store.is_known(profile_url="https://github.com/person1"),
            "username")
        self.assertEqual(
            self.store.is_known(email="person1@gmail.com"), "email")
        self.assertEqual(self.store.is_known(username="nobody"), "")

    def test_username_comparison_is_case_insensitive(self):
        self.store.save_student(make_record(1, username="AgouthamRaju"))
        for spelling in ("agouthamraju", "AGOUTHAMRAJU", "AgouthamRaju"):
            self.assertEqual(self.store.is_known(username=spelling),
                             "username", spelling)

    def test_profile_url_spellings_resolve_to_one_account(self):
        self.store.save_student(make_record(1, username="user"))
        for url in ("https://github.com/User",
                    "https://github.com/user/",
                    "http://github.com/user",
                    "https://www.github.com/USER"):
            self.assertTrue(self.store.is_known(profile_url=url), url)

    def test_the_database_refuses_the_second_insert(self):
        """Not a SELECT-then-INSERT: the UNIQUE index decides."""
        saved, _reason = self.store.save_student(make_record(1))
        self.assertTrue(saved)
        saved, reason = self.store.save_student(
            make_record(1, username="PERSON1"))
        self.assertFalse(saved)
        self.assertEqual(reason, "username")
        self.assertEqual(self.store.count(), 1)

    def test_identity_is_never_established_by_name_city_or_college(self):
        first = make_record(1, username="one")
        first["Current City"] = "Pune"
        first["College Name"] = "COEP"
        first["Primary Skills"] = "Python"
        self.store.save_student(first)
        second = make_record(2, username="two")
        second["Student Name"] = first["Student Name"]
        second["Current City"] = "Pune"
        second["College Name"] = "COEP"
        second["Primary Skills"] = "Python"
        saved, _reason = self.store.save_student(second)
        self.assertTrue(saved, "same name/city/college is NOT a duplicate")
        self.assertEqual(self.store.count(), 2)

    def test_placeholder_emails_are_not_identities(self):
        for number, address in enumerate(
                ("example@example.com", "test@test.com",
                 "your_email@example.com"), start=1):
            record = make_record(number, username="person%d" % number)
            record["Email Address"] = address
            saved, _reason = self.store.save_student(record)
            self.assertTrue(saved, address)
        self.assertEqual(self.store.count(), 3)

    def test_closed_campaign_history_is_consulted_read_only(self):
        self.store.save_student(make_record(1))
        history_path = os.path.join(self.dir, "identities.db")
        history = student_scraper.open_history(history_path, writable=True)
        for keys in self.store.known_keys():
            history.remember(keys, campaign_id="campaign_20260101_001")
        history.commit()
        history.close()

        read_only = student_scraper.open_history(history_path)
        fresh = student_scraper.open_store(os.path.join(self.dir, "next.db"),
                                        history=read_only)
        self.addCleanup(fresh.close)
        self.assertIn("earlier campaign", fresh.is_known(username="person1"))
        self.assertEqual(fresh.is_known(username="nobody"), "")
        with self.assertRaises(RuntimeError):
            read_only.remember({"github_username_key": "x"})


class TestXlsxReader(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="reader_test_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _write(self, header, rows, name="export.xlsx"):
        path = os.path.join(self.dir, name)
        book = Workbook()
        sheet = book.active
        sheet.append(header)
        for row in rows:
            sheet.append(row)
        book.save(path)
        return path

    def test_user_search_layout(self):
        """Highlighted terms arrive as separate cells and must rejoin."""
        path = self._write(
            ["prc-Avatar-Avatar-0xaUi src", "prc-Link-Link-9ZwDx href",
             "search-match", "search-match 2", "prc-Text", "search-match 3",
             "user-location"],
            [["https://github.com/letscodedanish.png?size=40",
              "https://github.com/letscodedanish",
              "Danish Khan", "letscodedanish", "3rd year student at",
              "LNCT", "Bhopal"]])
        rows = XlsxReader().read(path, "LNCT")
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row.username, "letscodedanish")
        self.assertEqual(row.profile_url,
                         "https://github.com/letscodedanish")
        self.assertEqual(row.xlsx_name, "Danish Khan")
        self.assertEqual(row.xlsx_location, "Bhopal")
        self.assertIn("3rd year student at", row.xlsx_bio)
        self.assertIn("LNCT", row.xlsx_bio)

    def test_column_order_is_not_assumed(self):
        """The same data in a different order must parse identically."""
        path = self._write(
            ["a", "b", "c", "d"],
            [["Danish Khan", "https://github.com/letscodedanish",
              "3rd year student", "https://github.com/letscodedanish.png"]])
        rows = XlsxReader().read(path, "LNCT")
        self.assertEqual(rows[0].username, "letscodedanish")

    def test_repository_search_invents_nothing(self):
        path = self._write(
            ["prc-Avatar-Avatar-0xaUi src", "prc-Link-Link-9ZwDx href",
             "search-match", "search-match 2",
             "Footer-module__footerItem__H6aUw"],
            [["https://github.com/haidev2801.png?size=40",
              "https://github.com/haidev2801/FresherMisa2026",
              "haidev2801/Misa2026", "Fresher", "C#"]])
        row = XlsxReader().read(path, "fresher")[0]
        self.assertEqual(row.username, "haidev2801")
        self.assertEqual(row.xlsx_name, "")        # no person name exists
        self.assertEqual(row.xlsx_location, "")    # "C#" is a language

    def test_reserved_paths_are_not_users(self):
        path = self._write(
            ["a"], [["https://github.com/search?q=x"]])
        self.assertEqual(XlsxReader().read(path, "kw"), [])

    def test_duplicate_rows_merge(self):
        path = self._write(
            ["a", "b"],
            [["https://github.com/octocat", "First bio"],
             ["https://github.com/octocat", "Second bio"]])
        rows = XlsxReader().read(path, "kw")
        self.assertEqual(len(rows), 1)


class TestState(unittest.TestCase):
    """Schema v4: identity-based resume, and four refusals."""

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="state_test_")
        self.runtime = os.path.join(self.root, "runtime")
        os.makedirs(self.runtime)
        self.plan = sandbox.plan(KEYWORDS, CITIES)

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def context(self, rank=1, total_ranks=2):
        return sandbox.enter(self.runtime)

    def state(self, rank=1, total_ranks=2, manifest=None):
        return StateManager(path=self.context(rank, total_ranks).state_file)

    def test_resume_point_survives_reload(self):
        assigned = self.plan.assign(1, 2)
        state = self.state()
        state.begin_unit(3, assigned[3])
        state.set_in("processing", next_row_index=137, total_rows=200)

        reloaded = self.state()
        self.assertEqual(reloaded.unit_index, 3)
        self.assertEqual(reloaded.search_query, assigned[3].search_query)
        self.assertEqual(reloaded.next_row_index, 137)

    def test_same_unit_resume_does_not_reset_progress(self):
        assigned = self.plan.assign(1, 2)
        state = self.state()
        state.begin_unit(3, assigned[3])
        state.set_in("processing", next_row_index=137)

        again = self.state()
        again.begin_unit(3, assigned[3])
        self.assertEqual(again.next_row_index, 137)
        # A different search resets it.
        again.begin_unit(4, assigned[4])
        self.assertEqual(again.next_row_index, 0)

    def test_the_search_city_alias_is_part_of_the_identity(self):
        """Two aliases of one city are two different searches."""
        assigned = self.plan.assign(1, 2)
        pairs = {}
        for unit in assigned:
            pairs.setdefault(unit.canonical_city, set()).add(unit.search_city)
        multi = [city for city, spellings in pairs.items()
                 if len(spellings) > 1]
        self.assertTrue(multi, "the fixture needs a multi-alias city")
        city = multi[0]
        units = [unit for unit in assigned if unit.canonical_city == city]
        first = units[0]
        other = next(unit for unit in units
                     if unit.search_city != first.search_city
                     and unit.base_keyword == first.base_keyword)
        state = self.state()
        state.begin_unit(assigned.index(first), first)
        state.set_in("processing", next_row_index=9)
        state.begin_unit(assigned.index(other), other)
        self.assertEqual(state.next_row_index, 0,
                         "a different spelling is a different search")

    def test_next_selector_is_remembered(self):
        state = self.state()
        state.remember_next_selector("github.com", "a[rel='next']")
        self.assertEqual(self.state().next_selector_for("github.com"),
                         "a[rel='next']")

    def test_advance_counts_completed_units(self):
        assigned = self.plan.assign(1, 2)
        state = self.state()
        state.begin_unit(0, assigned[0])
        state.advance_unit()
        self.assertEqual(state.unit_index, 1)
        self.assertEqual(state.units_done, 1)

    # -- nothing can refuse to resume ---------------------------------------
    # The old version compared campaign fingerprints, the rank, the rank count
    # and an assignment-algorithm version, and raised rather than resume when
    # any of them had moved.  All four are gone with the campaign they belonged
    # to.  These tests pin the replacement behaviour: a state is read for what
    # still means something, and a run always continues.

    def test_a_changed_keyword_list_still_resumes(self):
        assigned = self.plan.assign(1, 2)
        self.state().begin_unit(0, assigned[0])
        # The keyword file grew.  Under the old rules this refused to start.
        changed = sandbox.plan(KEYWORDS + ["DevOps Engineer"], CITIES)
        self.assertTrue(changed.assign(1, 2))
        reloaded = self.state()
        self.assertEqual(reloaded.search_query, assigned[0].search_query)

    def test_a_state_written_by_an_older_version_still_loads(self):
        import json
        context = self.context(1, 2)
        with open(context.state_file, "w", encoding="utf-8") as handle:
            json.dump({"version": 3, "keyword_index": 4,
                       "keyword": "Engineering Student Hyderabad",
                       "phase": "profiles"}, handle)
        state = self.state()
        # Keys this version does not know about are dropped; the ones it
        # expects but the file lacks get their defaults.
        self.assertEqual(state.data["version"], C.STATE_VERSION)
        self.assertEqual(state.search_query, "")
        self.assertEqual(state.next_row_index, 0)

    def test_a_state_written_with_campaign_blocks_still_loads(self):
        import json
        context = self.context(1, 2)
        with open(context.state_file, "w", encoding="utf-8") as handle:
            json.dump({"version": 4,
                       "campaign": {"campaign_id": "campaign_20260927_001",
                                    "plan_fingerprint": "deadbeef"},
                       "assignment": {"rank": 9, "total_ranks": 80},
                       "search_unit": {"search_query": "in:x location:Y",
                                       "canonical_city": "Y"},
                       "processing": {"next_row_index": 12}}, handle)
        state = self.state()
        self.assertEqual(state.search_query, "in:x location:Y")
        self.assertEqual(state.next_row_index, 12,
                         "the part that still means something is kept")

    def test_an_unreadable_state_starts_fresh_rather_than_raising(self):
        context = self.context(1, 2)
        with open(context.state_file, "w", encoding="utf-8") as handle:
            handle.write("{not json")
        state = self.state()
        self.assertEqual(state.search_query, "")
        self.assertEqual(state.unit_index, 0)


class TestTheStartupBanner(unittest.TestCase):
    """banner() runs on the way into every real run, so it must not raise.

    WHY THIS EXISTS
    ---------------
    When students.db stopped being the duplicate index, banner() kept its
    `store` parameter and `_run()` kept calling `banner(searches, self.store)`
    -- but nothing assigned self.store any more.  Every real run raised
    AttributeError on its first line of output.  The whole 908-test suite
    stayed green, because no test had ever called banner(): the Workflow tests
    all build stubs with __new__ and call individual methods.

    These tests call it the way _run() does, with the attributes a real
    Workflow actually has.
    """

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="banner_")
        self.sandbox = sandbox.enter(self.root)
        self.lines = []
        self.restore_say = student_scraper.say
        student_scraper.say = self.lines.append

    def tearDown(self):
        student_scraper.say = self.restore_say
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def searches(self):
        """The shape build_search_list() returns: (city, spelling, keyword).

        NOT make_unit() objects.  Production keeps tuples and only builds a
        SearchUnit when that search's turn comes, so a test that handed banner
        SearchUnits would pass while production crashed -- which is how the
        self.store break survived in the first place.
        """
        return [(city, spelling, keyword)
                for city in CITIES
                for spelling in city.search_aliases
                for keyword in KEYWORDS]

    def test_the_fixture_matches_what_build_search_list_returns(self):
        """Guard the guard: if the production shape changes, this fails."""
        real = student_scraper.build_search_list(1, 400, say_total=False)
        if not real:
            self.skipTest("rank 1 of 400 has no cities on this dataset")
        mine = self.searches()[0]
        self.assertEqual(type(real[0]), type(mine),
                         "build_search_list no longer returns what this "
                         "fixture builds; banner's unit[0] indexing and these "
                         "tests must both be updated")
        self.assertEqual(len(real[0]), len(mine))
        # banner() reads unit[0].canonical, so element 0 must carry it in
        # production and in the fixture alike.  FakeCity stands in for City.
        self.assertTrue(hasattr(real[0][0], "canonical"))
        self.assertTrue(hasattr(mine[0], "canonical"))

    def text(self):
        return "\n".join(self.lines)

    def test_banner_runs_with_a_workflow_built_the_way_run_builds_it(self):
        workflow = student_scraper.Workflow.__new__(student_scraper.Workflow)
        searches = self.searches()
        # Exactly the call _run() makes.  An attribute that no longer exists
        # fails here instead of on the user's terminal.
        student_scraper.banner(searches, workflow.existing)
        self.assertIn("GitHub Student Scraper", self.text())

    def test_the_banner_does_not_mention_the_duplicate_index(self):
        """The index is deliberately invisible at startup.

        It is not part of the spec this script follows (append to
        current_data.xlsx, roll at CYCLE_SIZE, save current_row.json, loop),
        and a line naming an index file that does not exist yet only invited
        "what is this file and why does it say 0 people?".  It still works --
        it just does not announce itself here.  The count is reported in the
        per-search rank card and at the end of each lap, where it means
        something.
        """
        workflow = student_scraper.Workflow.__new__(student_scraper.Workflow)
        existing = {"github:someone": make_record(1, "someone"),
                    "github:other": make_record(2, "other")}
        student_scraper.save_master(existing)
        workflow.existing = existing
        student_scraper.banner(self.searches(), workflow.existing)
        text = self.text()
        self.assertNotIn("Duplicate index", text)
        self.assertNotIn("master_index.xlsx", text)
        self.assertNotIn("person/people", text)
        self.assertNotIn("students.db", text)
        # the rest of the banner is unaffected
        self.assertIn("GitHub Student Scraper", text)
        self.assertIn("Cycle size", text)

    def test_the_banner_is_silent_about_the_index_on_a_fresh_folder_too(self):
        self.assertFalse(os.path.exists(student_scraper.master_path()),
                         "this test needs a folder with no index yet")
        student_scraper.banner(self.searches(), {})
        text = self.text()
        self.assertNotIn("Duplicate index", text)
        self.assertNotIn("not created yet", text)
        self.assertIn("GitHub Student Scraper", text)

    def test_banner_runs_on_the_real_production_search_list(self):
        """The strongest form: no fixture at all.

        banner() is handed exactly what main() hands it -- build_search_list()
        output over the real city and keyword data -- and the master index a
        real run loads.  If this passes, the first screen of a real run works.
        """
        searches = student_scraper.build_search_list(say_total=False)
        self.assertTrue(searches, "rank 1 must own some searches")
        workflow = student_scraper.Workflow.__new__(student_scraper.Workflow)
        student_scraper.banner(searches, workflow.existing)
        text = self.text()
        self.assertIn("Rank                : %d of %d"
                      % (student_scraper.rank, student_scraper.num_sys), text)
        self.assertIn("Cities this rank", text)
        self.assertIn("Searches this rank", text)

    def test_banner_accepts_none_for_existing(self):
        """The parameter stays even though nothing is printed from it."""
        student_scraper.banner(self.searches(), None)
        self.assertNotIn("Duplicate index", self.text())
        self.assertIn("GitHub Student Scraper", self.text())

    def test_every_attribute_run_reads_off_self_before_the_banner_exists(self):
        """_run()'s prologue must not read an attribute nothing assigns.

        Generalises the self.store break: scan _run's source for `self.X`
        reads and require each one to be a real member of Workflow.
        """
        import ast
        import inspect
        source = inspect.getsource(student_scraper.Workflow._run)
        tree = ast.parse(source.lstrip() if source.startswith(" ")
                         else source) if False else ast.parse(
                             "\n".join(line[4:] if line.startswith("    ")
                                        else line
                                        for line in source.splitlines()))
        reads = {node.attr for node in ast.walk(tree)
                 if isinstance(node, ast.Attribute)
                 and isinstance(node.ctx, ast.Load)
                 and isinstance(node.value, ast.Name)
                 and node.value.id == "self"}
        missing = sorted(name for name in reads
                         if not hasattr(student_scraper.Workflow, name)
                         and name not in student_scraper.Workflow
                         .__init__.__code__.co_names)
        self.assertEqual(missing, [],
                         "_run() reads self.%s, which Workflow never assigns "
                         "and does not declare" % ", self.".join(missing))


class TestTheMasterPathFollowsConfig(unittest.TestCase):
    """master_path() must resolve when called, not when imported.

    A module-level MASTER_FILE would keep pointing at the real "Student Data
    original" after sandbox.enter() redirected FINAL_OUTPUT_DIR, so any test
    calling load_existing_master()/save_master() without an explicit path
    would read and rewrite the developer's own collected people.
    """

    def test_master_path_moves_with_the_sandbox(self):
        root = tempfile.mkdtemp(prefix="masterpath_")
        try:
            real = student_scraper.master_path()
            sandbox.enter(root)
            self.assertTrue(student_scraper.master_path().startswith(root),
                            "master_path() ignored the redirected output dir")
            sandbox.restore()
            self.assertEqual(student_scraper.master_path(), real,
                             "restore() left master_path() redirected")
        finally:
            sandbox.restore()
            shutil.rmtree(root, ignore_errors=True)

    def test_the_default_master_is_inside_the_output_folder(self):
        self.assertEqual(os.path.dirname(student_scraper.master_path()),
                         C.FINAL_OUTPUT_DIR)


class TestDuplicatesFollowNaukri(unittest.TestCase):
    """The master is a TALLY, not a gate -- find_naukri_job.py's behaviour.

    WHY THIS EXISTS
    ---------------
    This script used to skip a known person twice: once before any fetching
    (saving the profile + README + portfolio + resume -- 111 profiles on one
    real search) and again before writing the row, so current_data.xlsx could
    never hold anyone twice.

    find_naukri_job.py does neither.  Its duplicate branch only bumps the
    counters:

        if key in existing_data:
            existing_data[key]["seen_count"] = old + 1
            existing_data[key]["last_seen_at"] = now_time()
        else:
            existing_data[key] = {...}

        input_row  = {...}                              # line 1033, col 24
        detail_row = await scrape_job_detail_async(...)  # line 1047, col 24
        append_detail_to_excel(detail_row, ...)          # line 1053, col 24

    Those last three sit at the SAME indentation as the if/else, so they run
    either way.  Matching that was asked for explicitly, so these tests hold
    the behaviour in place -- and say out loud what it costs.
    """

    def setUp(self):
        self.root = tempfile.mkdtemp(prefix="naukridupe_")
        self.sandbox = sandbox.enter(self.root)

    def tearDown(self):
        sandbox.restore()
        shutil.rmtree(self.root, ignore_errors=True)

    def test_a_second_sighting_bumps_seen_count_and_writes_another_row(self):
        existing = {}
        first = make_record(1, "repeat")
        key = student_scraper.make_unique_key(first)
        self.assertEqual(key, "github:repeat")

        # first sighting
        first["first_seen_at"] = first["last_seen_at"] = "2026-10-07T10:00:00"
        first["seen_count"] = 1
        existing[key] = first
        student_scraper.append_detail_to_excel(first, say=lambda _m: None)

        # second sighting, as the loop now handles it
        seen = existing[key]
        seen["seen_count"] = int(seen.get("seen_count", 0)) + 1
        seen["last_seen_at"] = "2026-10-07T11:00:00"
        student_scraper.append_detail_to_excel(make_record(1, "repeat"),
                                               say=lambda _m: None)

        self.assertEqual(existing[key]["seen_count"], 2,
                         "the master must count the second sighting")
        self.assertEqual(len(existing), 1,
                         "the master stays ONE row per person")
        self.assertEqual(student_scraper.row_count(), 2,
                         "current_data.xlsx gets a row for EACH sighting -- "
                         "this is the naukri behaviour and is intended")

    def test_nothing_filters_candidates_before_fetching(self):
        """The pre-fetch skip is gone and must not come back quietly.

        _already_known() was the function that did it.  If something
        reintroduces that filter, this fails and whoever did it has to say so.
        """
        self.assertFalse(
            hasattr(student_scraper.Workflow, "_already_known"),
            "Workflow._already_known is back; it skipped known people before "
            "fetching, which find_naukri_job.py does not do. Removing the "
            "skip was an explicit decision -- see this class's docstring.")

    def test_the_progress_denominator_counts_every_candidate(self):
        """'Profiles to process' must equal the candidate count.

        It used to be candidates-minus-known.  With no pre-fetch skip the two
        are the same number, and a difference would mean a filter crept back.
        """
        import inspect
        source = inspect.getsource(student_scraper.Workflow.process_rows)
        self.assertIn("pending = list(range(total))", source)
        self.assertIn("to_process = total", source)
        self.assertNotIn("_already_known", source)


if __name__ == "__main__":
    unittest.main(verbosity=2)
