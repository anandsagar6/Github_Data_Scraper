"""Point config at a throwaway folder, and build SearchUnits, for one test.

WHY THIS EXISTS
---------------
Production is now "one folder is one rank": every path a run writes to is a
module-level constant in config.py pointing next to student_scraper.py.  That
is what makes two copies unable to share a Playwright profile or a state file
-- they are different checkouts -- and it is why RankContext is gone.

A test must not write into the real project folder, so it needs those
constants pointed somewhere temporary for its duration.  `enter()` does that
and `restore()` puts them back.

The object `enter()` returns exposes the same attribute names the old
RankContext did (profile_dir, keyword_download_dir, state_file, store_db ...)
so a test can keep saying `self.context.store_db`.  It is a fixture, not a
production type: nothing outside tests/ imports it, and production code reads
the config constants directly.

    class MyTest(unittest.TestCase):
        def setUp(self):
            self.root = tempfile.mkdtemp()
            self.context = sandbox.enter(self.root)

        def tearDown(self):
            sandbox.restore()
            shutil.rmtree(self.root, ignore_errors=True)
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config as C                                        # noqa: E402
import student_scraper                                           # noqa: E402
import student_scraper                                 # noqa: E402

# Every config name enter() redirects, and the path under the sandbox root it
# is redirected to.  Listed explicitly so a new production path that a test
# should isolate shows up here as a deliberate addition rather than silently
# continuing to point at the real project folder.
_FILES = {
    "PROFILE_DIR": "pw_profile",
    "STATE_FILE": "scraper_state.json",
    "CURRENT_ROW_FILE": "current_row.json",
    "DEFERRED_FILE": "deferred_profiles.json",
    "TRANSIENT_FILE": "transient_failures.json",
    "THROTTLE_FILE": "throttle_state.json",
    "RATE_LIMIT_FILE": "rate_limit_state.json",
    "TEMP_STUDENT_FILE": "temporary.txt",
    "TEMP_STUDENT_JSON": "temporary.json",
    "TEMP_DOWNLOAD_DIR": "temp_downloads",
    "KEYWORD_DOWNLOAD_DIR": "keyword_downloads",
    "CRAWL_STAGING_DIR": "crawl_staging",
    "RETRY_BACKUP_DIR": "retry_backup",
    "LOG_DIR": "logs",
    "STORE_DB": "students.db",
    "SKIPPED_FILE": "skipped_accounts.txt",
    "FAILED_FILE": "failed_profiles.txt",
    "DETAIL_OUTPUT_FILE": "current_data.xlsx",
    "CYCLE_STATE_FILE": "cycle_state.json",
    "KEYWORD_CYCLE_FILE": "keyword_cycle.json",
    "HISTORY_DIR": "history",
    "FINAL_OUTPUT_DIR": "output",
}

# Pacing and schedule constants a sandbox COLLAPSES, and the value it
# collapses them to.  They are module globals at the TOP of
# student_scraper.py (next to rank/num_sys), not config entries, so they are
# set on that module.
#
# WHY.  Several test files drive the real Workflow with only the browser
# stubbed, so they execute the real _search_delay(), _page_delay() and
# _profile_delay() -- and those are deliberately slow in production: a
# results page every 1-3 minutes, a person every 5-25 seconds, taken from
# find_naukri_job.py.  At production values test_lifecycle.py alone sleeps
# for the better part of an hour, and a suite nobody will sit through is a
# suite nobody runs.
#
# A test that cares about the PRODUCTION numbers reads them through
# pristine() below, not through C: see test_ratelimit_pacing.test_config_values.
#
# The scheduled stops go to zero for a harder reason than speed: a lap
# boundary in production cools down for 45 minutes, and a test that happened
# to wrap its tiny search list would hang for 45 minutes with no indication
# of why.
# The ORDERING is preserved, not just the smallness: an empty search must
# still be cheaper than a productive one, because that is a rule the pacing
# tests check rather than a number they happen to read.
_SCRAPER_PACING = {
    "SEARCH_DELAY_MIN": 0.02,
    "SEARCH_DELAY_MAX": 0.02,
    "ZERO_RESULT_DELAY_MIN": 0.01,
    "ZERO_RESULT_DELAY_MAX": 0.01,
    "PAGE_DELAY_MIN": 0.01,
    "PAGE_DELAY_MAX": 0.01,
    "PROFILE_DELAY_MIN": 0.01,
    "PROFILE_DELAY_MAX": 0.01,
    "LAP_COOLDOWN_MINUTES": 0,
    "WORK_PERIOD_MIN_HOURS": 0,
    "WORK_PERIOD_MAX_HOURS": 0,
    # The night stop, for a harder reason than speed: a test that happened to
    # run inside the window would sit until morning.
    "NIGHT_PAUSE_ENABLED": False,
}

# CRAWL_STAGING_DIR is absent on purpose, matching production: it is created
# only when a search actually stages pages, and several tests read "the
# directory exists" as "a crawl was staged".
_DIRECTORIES = ("PROFILE_DIR", "TEMP_DOWNLOAD_DIR", "KEYWORD_DOWNLOAD_DIR",
                "LOG_DIR", "HISTORY_DIR", "FINAL_OUTPUT_DIR")

_saved = {}
_saved_scraper_flags = {}
_saved_temp_paths = {}


class Sandbox:
    """The redirected paths, under RankContext's old attribute names."""

    def __init__(self, root):
        self.root = root
        self.rank = 1
        self.total_ranks = 1
        self.profile_dir = C.PROFILE_DIR
        self.state_file = C.STATE_FILE
        self.current_row_file = C.CURRENT_ROW_FILE
        self.deferred_file = C.DEFERRED_FILE
        self.transient_file = C.TRANSIENT_FILE
        self.throttle_file = C.THROTTLE_FILE
        self.rate_limit_file = C.RATE_LIMIT_FILE
        self.temp_student_file = C.TEMP_STUDENT_FILE
        self.temp_student_json = C.TEMP_STUDENT_JSON
        self.temp_download_dir = C.TEMP_DOWNLOAD_DIR
        self.keyword_download_dir = C.KEYWORD_DOWNLOAD_DIR
        self.crawl_staging_dir = C.CRAWL_STAGING_DIR
        self.retry_backup_dir = C.RETRY_BACKUP_DIR
        self.log_dir = C.LOG_DIR
        self.log_file = os.path.join(C.LOG_DIR, "student_scraper.log")
        self.store_db = C.STORE_DB
        self.skipped_file = C.SKIPPED_FILE
        self.failed_file = C.FAILED_FILE

    def staging_for(self, slug):
        return C.staging_for(slug)

    def ensure(self):
        for name in _DIRECTORIES:
            os.makedirs(getattr(C, name), exist_ok=True)
        return self

    def owns(self, path):
        root = os.path.abspath(self.root) + os.sep
        return os.path.abspath(path).startswith(root)


def enter(root):
    """Redirect every per-folder config path under `root`.  Returns a Sandbox.

    Call restore() in tearDown.  Calling enter() twice restores the first
    sandbox before opening the second rather than raising: a test class that
    forgets its tearDown should fail on its own assertion, not poison every
    test that runs after it.
    """
    if _saved:
        restore()
    for name, leaf in _FILES.items():
        _saved[name] = getattr(C, name)
        setattr(C, name, os.path.join(root, leaf))
    _saved["FINAL_WORKBOOK_PATH"] = C.FINAL_WORKBOOK_PATH
    C.FINAL_WORKBOOK_PATH = os.path.join(C.FINAL_OUTPUT_DIR, C.FINAL_WORKBOOK)
    _saved["HISTORY_DB"] = C.HISTORY_DB
    C.HISTORY_DB = os.path.join(C.HISTORY_DIR, "identities.db")
    _saved["LEGACY_IMPORT_DB"] = C.LEGACY_IMPORT_DB
    C.LEGACY_IMPORT_DB = os.path.join(root, "imported", "legacy_students.db")

    for name, value in _SCRAPER_PACING.items():
        _saved_scraper_flags[name] = getattr(student_scraper, name)
        setattr(student_scraper, name, value)

    _saved_temp_paths.update(student_scraper._PATHS)
    student_scraper.configure(temp_student_file=C.TEMP_STUDENT_FILE,
                         temp_student_json=C.TEMP_STUDENT_JSON,
                         temp_download_dir=C.TEMP_DOWNLOAD_DIR)
    return Sandbox(root).ensure()


def pristine(name):
    """The real shipped value of `name`, even inside a sandbox.

    For the one kind of test that is about the production constants
    themselves rather than about behaviour -- "is the pacing still the
    numbers the operator asked for?".  Looks in both saves (the pacing knobs
    live on student_scraper, the paths on config) and falls back to whichever
    module actually defines it.
    """
    if name in _saved_scraper_flags:
        return _saved_scraper_flags[name]
    if name in _saved:
        return _saved[name]
    if hasattr(student_scraper, name):
        return getattr(student_scraper, name)
    return getattr(C, name)


def restore():
    """Put every redirected config path back.  Safe to call twice."""
    for name, value in _saved.items():
        setattr(C, name, value)
    _saved.clear()
    for name, value in _saved_scraper_flags.items():
        setattr(student_scraper, name, value)
    _saved_scraper_flags.clear()
    if _saved_temp_paths:
        student_scraper._PATHS.update(_saved_temp_paths)
        _saved_temp_paths.clear()


# ------------------------------------------------------------ search units --
def units(count=1, city="Pune", state="Maharashtra", keywords=None,
          spelling=None):
    """A short, deterministic list of SearchUnits for a test.

    Production crosses this rank's cities with keyword_by_category from
    new_mylib.py, which is 603,750 units for one rank at num_sys = 5 -- far
    more than a test wants.  This builds exactly as many as asked for, with
    the same shape.
    """
    import student_scraper as S

    class _City:
        def __init__(self, canonical, state_name, aliases):
            self.country = "India"
            self.state = state_name
            self.canonical = canonical
            self.search_aliases = tuple(aliases)

    keywords = keywords or ["college student", "university student",
                            "fresher", "b.tech student", "engineering student"]
    while len(keywords) < count:
        keywords = keywords + ["%s %d" % (keywords[0], len(keywords))]
    place = _City(city, state, [spelling or city])
    return [S.make_unit(place, place.search_aliases[0], keywords[index])
            for index in range(count)]


def unit(city="Pune", state="Maharashtra", keyword="college student",
         spelling=None):
    """One SearchUnit."""
    return units(1, city=city, state=state, keywords=[keyword],
                 spelling=spelling)[0]


class Plan:
    """A tiny stand-in for the deleted search_plan.SearchPlan.

    Production has no plan object any more: student_scraper crosses this
    rank's cities with every keyword straight into a list.  Several tests
    were written against the old object, and what they actually need from it
    is "give me the units for rank R of T", so that is what this provides --
    over the SAME make_unit() and get_cities_by_rank() production uses, so a
    test cannot pass against an ordering production does not produce.
    """

    def __init__(self, keywords=None, cities=None):
        import student_scraper as locations
        self.base_keywords = tuple(keywords or ["college student"])
        self.cities = tuple(cities if cities is not None else locations.cities())

    def assign(self, rank=1, total_ranks=1):
        import student_scraper as S
        mine = S.get_cities_by_rank(list(self.cities), rank, total_ranks)
        return [S.make_unit(city, spelling, keyword)
                for city in mine
                for spelling in city.search_aliases
                for keyword in self.base_keywords]

    @property
    def units(self):
        return self.assign(1, 1)

    def __len__(self):
        return len(self.units)

    @property
    def canonical_city_count(self):
        return len(self.cities)

    @property
    def search_spelling_count(self):
        return sum(len(city.search_aliases) for city in self.cities)

    @property
    def base_keyword_count(self):
        return len(self.base_keywords)


def plan(keywords=None, cities=None):
    """Build a Plan.  Mirrors the old search_plan.build_plan() signature."""
    return Plan(keywords, cities)


class OtherFolder:
    """A SECOND copy's folder, for tests about two copies not colliding.

    `enter()` can only point config at one folder at a time, because
    production has exactly one set of paths -- that is the whole design. A
    test that needs to talk about another copy's files needs their paths
    without redirecting config at them, which is what this builds.

    The layout is derived from the same `_FILES` table `enter()` uses, so if a
    production path moves, this moves with it.
    """

    def __init__(self, root):
        self.root = root
        for name, leaf in _FILES.items():
            setattr(self, name.lower(), os.path.join(root, leaf))
        self.crawl_staging_dir = os.path.join(root, _FILES["CRAWL_STAGING_DIR"])
        self.store_db = os.path.join(root, _FILES["STORE_DB"])
        self.state_file = os.path.join(root, _FILES["STATE_FILE"])
        self.keyword_download_dir = os.path.join(
            root, _FILES["KEYWORD_DOWNLOAD_DIR"])

    def staging_for(self, slug):
        return os.path.join(self.crawl_staging_dir, slug or "search")

    def owns(self, path):
        root = os.path.abspath(self.root) + os.sep
        return os.path.abspath(path).startswith(root)

    def paths(self):
        return {name.lower(): os.path.join(self.root, leaf)
                for name, leaf in _FILES.items()}


def other_folder(root):
    """Another copy's folder.  Does NOT redirect config."""
    os.makedirs(root, exist_ok=True)
    return OtherFolder(root)
