#!/usr/bin/env python3
"""GitHub student/person data extractor.

    new_mylib.py  (cities x keywords)
              |
              v
      this rank's cities x every keyword  ->  for each search:
              GitHub Users search -> Instant Data Scraper crawls every result
              page -> ONE XLSX -> row by row: GitHub profile, raw README,
              portfolio, resume -> temporary.txt -> extract/normalise ->
              current_data.xlsx -> clear temporary.txt -> ... -> all rows
              resolved -> delete that search's XLSX

The one safety rule that governs everything:

    SUCCESS                 -> clean up
    ANY UNRESOLVED FAILURE  -> preserve the XLSX, the state and the temp
                               files, and resume later

Cleanup is therefore never placed in a finally: block.

HOW THE WORK IS DIVIDED
-----------------------
Two variables at the top of this file, and nothing else:

    rank      = 1      which copy of this script this is
    num_sys   = 5      how many copies are running in total

Each copy takes its own slice of the city list and crosses it with every
keyword.  Copy 1 takes the first slice, copy 2 the next, and so on, so no two
copies ever search the same city.  To deploy, clone this folder onto each
laptop and change `rank` -- exactly the way the job fetching scripts are
deployed, where every rank is its own clone with its own committed rank line.

ONE FOLDER IS ONE RANK
----------------------
Everything a run writes sits next to this file: current_row.json,
scraper_state.json, students.db, current_data.xlsx.  The Chromium profile is
the exception -- it is disposable and lives in the system temp directory, so
the login travels in github_auth.json alone, as naukri_auth.json does.
There is no runtime/rank_NNN/ layering, because a rank is not a subdirectory
-- it is a whole clone of the project.

WHERE IT RESUMES FROM
---------------------
current_row.json holds the index of the next search to run, written
immediately AFTER a search finishes.  A crash therefore resumes on the search
after the one that completed, never repeating it.  When the last search is
done the index wraps back to 0 and the whole list runs again, forever.

WHAT IT PRODUCES
----------------
current_data.xlsx, rolled over into current_data_cycle1.xlsx,
current_data_cycle2.xlsx ... every 100 people.  A cycle file is sealed and
never reopened, so a collector can copy it away while this keeps running.
students.db stays as well, but only as the duplicate index -- it is what makes
"have I already collected this person?" a database question rather than a
guess.
"""

import argparse
import asyncio
import atexit
import datetime
import json
import logging
import logging.handlers
import hashlib
import os
import random
import re
import shutil
import signal
import sys
import urllib.parse

import config as C


import glob
import html
import socket
import sqlite3
import subprocess
import time
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from openpyxl import Workbook, load_workbook
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from playwright.async_api import async_playwright

import tempfile

log = logging.getLogger("student_scraper")

# ==========================================================================
# WORK DIVISION -- the only two numbers that differ between laptops
# ==========================================================================
# rank      which copy of this script this folder is, counting from 1
# num_sys   how many copies are running altogether
#
# These are the DEFAULTS, and they are what a fresh clone runs with.  Set
# them per laptop in rank.json (below) rather than by editing this file.
rank = 1
num_sys = 25

# -- night stop: no collecting between these hours (local, 24-hour clock) --
NIGHT_PAUSE_ENABLED = True     # True = ON, False = OFF
NIGHT_START_HOUR = 23           # collecting stops at this hour
NIGHT_END_HOUR = 5              # collecting resumes at this hour

# -- keyword order ---------------------------------------------------------
KEYWORD_SHUFFLE_ENABLED = True  # False = new_mylib.py's file order

# -- request pacing: every wait is a fresh random draw MIN..MAX, in seconds
SEARCH_DELAY_MIN = 5           # between two completed searches
SEARCH_DELAY_MAX = 20
ZERO_RESULT_DELAY_MIN = 5       # after a search that matched nobody
ZERO_RESULT_DELAY_MAX = 10
PAGE_DELAY_MIN = 20             # between two results pages of one search
PAGE_DELAY_MAX = 180
PROFILE_DELAY_MIN = 5           # between two people
PROFILE_DELAY_MAX = 25
JITTER_FRACTION = 0.35          # retry backoffs are spread up by 0-35%
LAP_COOLDOWN_MIN = 10  
LAP_COOLDOWN_MAX = 60           
LAP_COOLDOWN_MINUTES = random.randrange(LAP_COOLDOWN_MIN, LAP_COOLDOWN_MAX + 1)      # cooldown when the whole list has been walked
RATE_LIMIT_COOLDOWN_MIN = 5     # a GitHub rate limit waits a fresh random
RATE_LIMIT_COOLDOWN_MAX = 30    # MIN..MAX minutes, then retries -- never stops

# -- work / break: both drawn fresh every time, counting WORKING time only -
WORK_PERIOD_MIN_HOURS = 2       # work this long...
WORK_PERIOD_MAX_HOURS = 6
BREAK_MIN_MINUTES = 30          # ...then stop this long
BREAK_MAX_MINUTES = 60
BROWSER_RESTART_PAUSE_SECONDS = 600   # a pause this long closes Chromium

# -- per-laptop override ---------------------------------------------------
# WHY THIS EXISTS.  This project is deployed by cloning it onto each laptop
# and pulling updates into the clone.  Editing `rank` above works until the
# first time somebody changes anything else in this file -- then every
# laptop whose rank is not 1 has a locally modified tracked file, and the
# pull does not merge, it ABORTS:
#
#     error: Your local changes to the following files would be
#     overwritten by merge:  student_scraper.py
#     Aborting
#
# The fix is to keep the one thing that differs per machine OUT of the
# tracked file.  rank.json is gitignored, so a pull can never collide with
# it, and the numbers above stay as the committed defaults.
#
#     rank.json           {"rank": 2, "num_sys": 3}
#
# Precedence, lowest to highest:
#     1. the defaults above
#     2. rank.json, if present
#     3. GPE_RANK / GPE_NUM_SYS in the environment  (for systemd, cron,
#        docker -- anything that sets config by env rather than by file)
#
# Nothing is required: a clone with no rank.json and no environment runs as
# rank 1 of num_sys 3, exactly as this file reads.  See rank.example.json.
RANK_FILE = "rank.json"
RANK_SOURCE = "default"          # shown by --dry-run so the value is traceable


def _load_rank_override():
    """Apply rank.json and then the environment, over the defaults above.

    Deliberately forgiving about a MISSING file -- that is the normal case
    for rank 1 -- and deliberately loud about a file that exists but cannot
    be used, because silently crawling the wrong cities is far worse than
    refusing to start.
    """
    global rank, num_sys, RANK_SOURCE
    import json as _json
    import os as _os

    found = []
    path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)),
                         RANK_FILE)
    if _os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as handle:
                data = _json.load(handle)
            if not isinstance(data, dict):
                raise ValueError("not a JSON object")
            if "rank" in data:
                rank = int(data["rank"])
            if "num_sys" in data:
                num_sys = int(data["num_sys"])
            found.append(RANK_FILE)
        except Exception as exc:
            raise SystemExit(
                "%s is unreadable (%s).\n"
                "It must be a JSON object, for example:\n"
                '    {"rank": 2, "num_sys": 3}\n'
                "Fix it or delete it; deleting it falls back to rank %d of "
                "%d." % (path, exc, rank, num_sys))

    for name, key in (("GPE_RANK", "rank"), ("GPE_NUM_SYS", "num_sys")):
        raw = _os.environ.get(name)
        if raw is None or raw == "":
            continue
        try:
            value = int(raw)
        except ValueError:
            raise SystemExit("%s=%r is not a whole number." % (name, raw))
        if key == "rank":
            rank = value
        else:
            num_sys = value
        found.append(name)

    # Caught HERE rather than at the first search: a bad pair means this
    # laptop crawls the wrong cities or none at all, and the operator should
    # hear about it before the browser opens.
    if num_sys < 1:
        raise SystemExit("num_sys must be at least 1, not %d." % num_sys)
    if rank < 1 or rank > num_sys:
        raise SystemExit(
            "rank %d is outside 1..%d.\n"
            "Each laptop takes a different rank in that range, and num_sys "
            "is the same on all of them." % (rank, num_sys))

    if found:
        RANK_SOURCE = " then ".join(found)


_load_rank_override()


# --------------------------------------------------------------------------
# HOW THE CITIES ARE DIVIDED
# --------------------------------------------------------------------------
# This sits here, next to the two numbers it uses, because it is the whole
# of the work division and it should be readable without searching the file.
# It used to live 11,900 lines down, between make_unit() and
# build_search_list().
#
# The city list is cut into num_sys CONTIGUOUS blocks and this rank takes the
# one at its own position: rank 1 gets the first block, rank 2 the second,
# and so on.  Whole cities only, so every spelling and every keyword of a
# city stays with one rank -- two ranks never crawl the overlapping result
# sets of "Bengaluru" and "Bangalore", and this rank's own duplicate index
# absorbs the overlap instead.
#
# The blocks are EQUAL: with n cities over k ranks every rank gets either
# n//k or n//k + 1 cities and no two ranks differ by more than one.  The
# remainder goes to the lowest ranks, one city each.
#
#     871 cities, num_sys = 3   ->   291, 290, 290
#     871 cities, num_sys = 4   ->   218, 218, 218, 217
#       8 cities, num_sys = 4   ->     2,   2,   2,   2
#
# WHAT THIS REPLACED, AND WHY.  The previous form was
#
#     one_unit  = int(len_city / num_sys) + 1
#     first_idx = one_unit * (rank - 1)
#     last_idx  = one_unit * rank
#
# which rounds the block size UP for everyone and lets the last rank absorb
# whatever is left.  That is not an equal division:
#
#     871 cities, num_sys = 3   ->   291, 291, 289   (a 2-city spread)
#       8 cities, num_sys = 4   ->     3,   3,   2,   0   <-- rank 4 IDLE
#
# A whole laptop sitting idle is the failure that matters -- one_unit is
# large enough that num_sys - 1 ranks can swallow the entire list.  divmod
# cannot do that: it hands out the floor to everybody first, then spreads
# the remainder one city at a time, so no rank is ever starved while another
# holds two extra.
#
# The slice stays a pure function of (len, rank, num_sys), which is what
# lets every laptop work out its own share without talking to any other.
#
# CHANGING THIS CHANGES WHICH CITIES A RANK OWNS.  current_row.json is a
# plain index into the search list this builds, so a rank that is already
# mid-crawl must be reset (./reset.sh) when the division changes -- its old
# index would be an index into a list that no longer exists.
def get_cities_by_rank(all_cities, rank, num_sys):
    """This rank's equal, contiguous slice of the city list.

    Returns [] for a rank outside 1..num_sys, which the callers already
    report as "this rank has no cities to search" rather than crashing.
    """
    if num_sys < 1 or rank < 1 or rank > num_sys:
        return []
    len_city = len(all_cities)
    # base cities for every rank, one extra for the first `extra` ranks
    base, extra = divmod(len_city, num_sys)
    first_idx = base * (rank - 1) + min(rank - 1, extra)
    last_idx = first_idx + base + (1 if rank <= extra else 0)
    return all_cities[first_idx:last_idx]


STOP_REQUESTED = False

# ==========================================================================
# LOCATIONS
# ==========================================================================
# India search-location dataset for the GitHub student/fresher campaign.
#
# THE ONE DISTINCTION THIS MODULE EXISTS TO ENFORCE
# -------------------------------------------------
# A SEARCH location and a STUDENT'S location are different things and must
# never be confused:
#
#     SEARCH location   "Desktop Support Engineer Hyderabad" is a GitHub query.
#                       It says where we LOOKED, nothing about where the person
#                       lives.  It lives in SearchUnit and in the store's
#                       search_* columns.
#
#     STUDENT location  Current City / State / Country in the final record.
#                       Comes only from profile, resume, portfolio or export
#                       evidence, via parse_location().
#
# Somebody found by "Fresher Hyderabad" who lives in Pune is a Pune record.
# Somebody with no location evidence at all gets an empty Current City -- never
# the search city.  This module is therefore deliberately NOT imported by
# location_extractor or student_merger; it feeds the search plan only.
#
# SEARCH ALIASES vs CANONICAL CITY
# --------------------------------
# GitHub users write whichever spelling they grew up with, so both must be
# searched:
#
#     canonical "Bengaluru"      aliases  Bengaluru, Bangalore
#     canonical "Visakhapatnam"  aliases  Visakhapatnam, Vizag
#
# The canonical name is the rank-assignment key, so every alias of one city --
# and every keyword for it -- lands on the same rank and one rank's own
# duplicate index absorbs the overlap between "Bengaluru" and "Bangalore"
# result sets.  Aliases are search inputs only; they never reach a final
# record.
#
# PROVENANCE
# ----------
# Every city here is one of the 153 spellings the existing campaign already
# proved, read out of target_audience.txt's own "### <State>" sections.  The
# file grouped 19 sections; two of them are not states and are split here:
#
#     "Delhi NCR"         -> Delhi (Delhi, New Delhi)
#                            Haryana (Gurugram, Gurgaon, Faridabad)
#     "Assam & Northeast" -> Assam, Meghalaya, Manipur, Tripura, Sikkim
#
# Chandigarh is a union territory, not part of Punjab.  Noida and Greater
# Noida stay in Uttar Pradesh, which is where target_audience.txt already put
# them and what location_extractor agrees with.
#
# RESOLVED CONFLICTS
# ------------------
# Cities that new_mylib.city_by_states_original and location_extractor
# disagreed about, or that neither could place, are settled here explicitly:
#
#     Hanamkonda     Telangana   -- alias of Warangal (tri-city)
#     Nalgonda       Telangana   -- absent from both datasets
#     Dharwad        Karnataka   -- kept SEPARATE from Hubballi; the datasets
#                                   only carried the twin "Hubli-Dharwad"
#     Manipal        Karnataka   -- Udupi district; absent from both
#     Aurangabad     Maharashtra -- alias of Chhatrapati Sambhajinagar, NOT
#                                   the Bihar city of the same name
#     Sagar          Madhya Pradesh  -- not the Karnataka Sagar
#     Udaipur        Rajasthan       -- not the Tripura Udaipur
#     Chandigarh     Chandigarh (UT) -- not Punjab
#     Gurugram       Haryana         -- not "Delhi NCR"
#     Noida          Uttar Pradesh   -- not "Delhi NCR"
#
# Expanding to the larger state-wise dataset later means adding rows to
# CITY_DATA.  Nothing else in the project changes -- but note that it changes
# DATASET_FINGERPRINT, which starts a new campaign by design.
# --------------------------------------------------------------------------
COUNTRY = "India"

# (state, canonical city, [search spellings, canonical first])
#
# Declaration order within a state is the order target_audience.txt used, so
# the generated plan stays recognisable against the campaign it replaces.
# ---------------------------------------------------------------- city data
# The city dataset is LOADED from new_mylib.py (city_by_states_original):
# 29 groups, 866 city entries.  That file is the single place to edit cities,
# the same way keyword_by_category is the single place to edit keywords.
#
# It is PARSED, never imported: new_mylib.py also carries job-role data and
# imports subprocess/sys, and a city list must never execute anything.
# ast.literal_eval sees only the literal dict.
CITY_SOURCE_FILE = "new_mylib.py"
CITY_SOURCE_DICT = "city_by_states_original"






class LocationDataError(RuntimeError):
    """The city dataset is missing or unusable."""


def _read_city_source(path=None):
    """(state, city) pairs from new_mylib.py, in file order.

    Parsed with ast.literal_eval so nothing in that module runs.
    """
    import ast
    import os
    path = path or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                CITY_SOURCE_FILE)
    if not os.path.exists(path):
        raise LocationDataError(
            "%s is missing.  It holds the city dataset; the scraper cannot "
            "build a search plan without it." % path)
    with open(path, encoding="utf-8", errors="replace") as handle:
        tree = ast.parse(handle.read())
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if getattr(node.targets[0], "id", "") != CITY_SOURCE_DICT:
            continue
        try:
            data = ast.literal_eval(node.value)
        except Exception as exc:
            raise LocationDataError("%s in %s is not a literal dict (%s)"
                                    % (CITY_SOURCE_DICT, path, exc))
        pairs = []
        for group, names in data.items():
            group = str(group).strip()
            for name in names:
                name = str(name).strip()
                if not name:
                    continue
                # The group a city is listed under IS its state, which is
                # what naukri's city_to_state() returns: that function walks
                # city_by_states_original and hands back the group key.  So
                # new_mylib.py's groupings come through exactly as written,
                # "Delhi NCR" included.
                pairs.append((group, name))
        if not pairs:
            raise LocationDataError("%s in %s holds no cities"
                                    % (CITY_SOURCE_DICT, path))
        return pairs
    raise LocationDataError("%s not found in %s"
                            % (CITY_SOURCE_DICT, path))


def _city_data(pairs=None):
    """The (state, city, [spellings]) rows the search plan is built from.

    This is find_naukri_job.py's build_search_list(), on this dataset:

        all_cities = [city for cities in city_by_states.values()
                           for city in cities]
        all_cities = list(dict.fromkeys(all_cities))

    Flatten every group, then drop exact duplicates keeping the first -- which
    is what dict.fromkeys does.  A name listed under two groups therefore keeps
    the group it appears under first, so one city never belongs to two states.

    Every city is searched under its own name and nothing else.  There is no
    canonical/alias distinction any more: each row's spelling list is just
    [name], kept as a list so the rest of the file does not care which shape
    the data came from.

    WHAT THIS DELIBERATELY DOES NOT DO
    ----------------------------------
    "Bengaluru" and "Bangalore" are two separate cities here, as they are in
    naukri.  They can land on different ranks, so two copies may crawl
    overlapping GitHub result sets for what is really one place.  Each folder's
    master_index.xlsx absorbs its own repeats; between two laptops the overlap
    is resolved when their masters are merged.  Same for Mysuru/Mysore,
    Gurugram/Gurgaon and the rest.  That is the naukri behaviour and it is
    chosen, not overlooked.
    """
    pairs = pairs if pairs is not None else _read_city_source()
    rows, seen = [], set()
    for state, name in pairs:
        if name in seen:
            continue            # dict.fromkeys: first occurrence wins
        seen.add(name)
        rows.append((state, name, [name]))
    return rows


CITY_DATA = _city_data()


class City:
    """One canonical search city and the spellings to search it by.

    Immutable on purpose: the plan, the fingerprint and every rank's
    assignment are derived from this data, so nothing may mutate it at
    runtime.
    """

    __slots__ = ("country", "state", "canonical", "search_aliases")

    def __init__(self, country, state, canonical, search_aliases):
        object.__setattr__(self, "country", country)
        object.__setattr__(self, "state", state)
        object.__setattr__(self, "canonical", canonical)
        object.__setattr__(self, "search_aliases", tuple(search_aliases))

    def __setattr__(self, name, value):
        raise AttributeError("City is immutable")

    def __repr__(self):
        return "City(%r, %r, aliases=%r)" % (self.state, self.canonical,
                                             self.search_aliases)

    def __eq__(self, other):
        return (isinstance(other, City)
                and (self.country, self.state, self.canonical,
                     self.search_aliases)
                == (other.country, other.state, other.canonical,
                    other.search_aliases))

    def __hash__(self):
        return hash((self.country, self.state, self.canonical,
                     self.search_aliases))


def _build():
    cities, seen_canonical, seen_alias = [], {}, {}
    for state, canonical, aliases in CITY_DATA:
        if canonical in seen_canonical:
            raise ValueError("duplicate canonical city %r (states %r and %r)"
                             % (canonical, seen_canonical[canonical], state))
        seen_canonical[canonical] = state
        if not aliases or aliases[0] != canonical:
            raise ValueError("city %r must list its canonical spelling first"
                             % canonical)
        for alias in aliases:
            key = alias.strip().lower()
            if key in seen_alias:
                raise ValueError("spelling %r used by both %r and %r"
                                 % (alias, seen_alias[key], canonical))
            seen_alias[key] = canonical
        cities.append(City(COUNTRY, state, canonical, aliases))
    return tuple(cities)


SEARCH_CITIES = _build()

# alias (lower-cased) -> canonical city name.  Search inputs only.
_ALIAS_TO_CANONICAL = {alias.lower(): city.canonical
                       for city in SEARCH_CITIES for alias in city.search_aliases}
_CANONICAL_TO_CITY = {city.canonical: city for city in SEARCH_CITIES}


def cities():
    """Every configured search city, in declaration order."""
    return SEARCH_CITIES


def states():
    """Sorted state names.  Sorted, because the plan's canonical ordering is."""
    return sorted({city.state for city in SEARCH_CITIES})




def canonical_city_count():
    return len(SEARCH_CITIES)


def search_spelling_count():
    return sum(len(city.search_aliases) for city in SEARCH_CITIES)


def canonical_for(spelling):
    """'Bangalore' -> 'Bengaluru'.  '' when the spelling is not configured."""
    return _ALIAS_TO_CANONICAL.get((spelling or "").strip().lower(), "")


def state_for(canonical_city):
    city = _CANONICAL_TO_CITY.get(canonical_city)
    return city.state if city else ""


def dataset_fingerprint():
    """Stable hash of the ordered dataset.

    Over the resolved structure, not the file bytes: editing a comment or a
    docstring here does not invalidate a running campaign, while adding,
    removing, renaming or reordering a city or an alias does.
    """
    payload = "\n".join(
        "%s|%s|%s|%s" % (city.country, city.state, city.canonical,
                         ",".join(city.search_aliases))
        for city in SEARCH_CITIES)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


# ==========================================================================
# KEYWORD LOADER
# ==========================================================================
# Reads search keywords out of the plain-text target-audience files.
#
# The files are human documents: they carry a title line, emoji section
# headings, italic asides and blank lines.  Only real role/degree phrases are
# kept, order is preserved and duplicates are dropped.  The files themselves
# are never modified.
# --------------------------------------------------------------------------
# Lines that are decoration rather than a keyword.  Each rule below is
# named so a rejection can be explained, and so a future keyword file can
# be checked against them with keyword_rejections().

# 2. Markdown headings: "#", "##", "### Telangana".
_HASH_HEADING = re.compile(r"^\s*#")

# 3. Separator rules: "=====", "-----", "_____", "*****".
_SEPARATOR = re.compile(r"^\s*[=_*~\-]{3,}\s*$")

# 4. Metadata lines: "Format: <Base Keyword> <City>", "Notes: ...".
_METADATA = re.compile(
    r"^\s*(format|purpose|country|state|city|notes?|source|version|"
    r"updated|author|description|usage|examples?|total|count|category|"
    r"section|region|generated|file|list|target|scope)\s*[:=]",
    re.I)

# A placeholder line such as "<Base Keyword> <City>" describes the format,
# it is not something to search for.
_PLACEHOLDER = re.compile(r"<[^>]+>")

# Comment / aside lines the original files used: "*2nd level from ...*".
_ASIDE_RE = re.compile(r"^\s*[*/]")

# 6. Emoji section headings from the original file layout, e.g.
#    "🖥️ For IT Support / Help Desk", "Our Target Audience For ...".
_LEADING_LABEL = re.compile(r"^(for\b|our\b|list of\b|section\b)", re.I)

_HAS_LETTERS = re.compile(r"[A-Za-z]")
_HAS_LOWERCASE = re.compile(r"[a-z]")

# Decoration stripped before a line is judged.
_DECORATION = "🖥️🌐☁️📱🔧💼📊🎯📌📎🔹🔸▪️▫️●○•-–—*#> \t"

MAX_KEYWORD_WORDS = 10


def rejection_reason(line):
    """Why this line is not a keyword, or "" when it is one.

    Returning the reason rather than a bool keeps the rules auditable and
    lets --dry-run explain a surprising exclusion.
    """
    raw = line.rstrip("\n")
    if not raw.strip():
        return "empty line"
    if _HASH_HEADING.match(raw):
        return "markdown heading"
    if _SEPARATOR.match(raw):
        return "separator rule"
    if _METADATA.match(raw):
        return "metadata line"
    if _PLACEHOLDER.search(raw):
        return "format placeholder"
    if _ASIDE_RE.match(raw):
        return "comment or aside"

    text = raw.strip().lstrip(_DECORATION).strip()
    if not text:
        return "decoration only"
    if not _HAS_LETTERS.search(text):
        return "no letters"
    if text.endswith(":"):
        return "section heading"
    if _LEADING_LABEL.match(text):
        return "section heading"

    words = text.split()
    # 5/6. A title is written in capitals: "INDIA COLLEGE STUDENT & FRESHER
    # TARGET KEYWORDS".  Single all-caps tokens are real keywords (BCA,
    # MCA, MBA), so the rule needs at least two words to fire.
    if len(words) >= 2 and not _HAS_LOWERCASE.search(text):
        return "all-caps title"
    if len(words) > MAX_KEYWORD_WORDS:
        return "too long to be a search phrase"
    return ""


def _is_keyword(line):
    return not rejection_reason(line)


def keyword_rejections(files=None):
    """[(file, line number, text, reason)] for everything skipped."""
    names = files if files is not None else C.KEYWORD_FILES
    rejected = []
    for name in names:
        path = name if os.path.isabs(name) else os.path.join(C.PROJECT_ROOT,
                                                             name)
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8", errors="replace") as handle:
            for number, line in enumerate(handle, start=1):
                reason = rejection_reason(line)
                if reason and reason != "empty line":
                    rejected.append((os.path.basename(path), number,
                                     line.strip(), reason))
    return rejected


def _clean(line):
    """Strip decoration and trailing punctuation from an accepted line."""
    text = line.strip().lstrip(_DECORATION).strip()
    text = re.sub(r"\s+", " ", text).strip()
    # "Site Reliability Engineer (SRE)" -- the parenthetical is an
    # abbreviation of the phrase, not part of what to search for.
    text = re.sub(r"\s*\([^)]*\)\s*$", "", text).strip()
    return text.strip(" ,;")


def load_keywords(files=None):
    """Ordered, de-duplicated keyword list from the configured files."""
    names = files if files is not None else C.KEYWORD_FILES
    keywords = []
    seen = set()
    for name in names:
        path = name if os.path.isabs(name) else os.path.join(C.PROJECT_ROOT,
                                                             name)
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8", errors="replace") as handle:
            for line in handle:
                if not _is_keyword(line):
                    continue
                keyword = _clean(line)
                if not keyword:
                    continue
                key = keyword.lower()
                if key in seen:
                    continue
                seen.add(key)
                keywords.append(keyword)
    return keywords


def slugify(text):
    """'B.Tech Student' -> 'btech-student' (deterministic, filesystem safe)."""
    # Dots inside abbreviations vanish rather than becoming separators, so
    # "B.Tech Student" gives "btech-student" and not "b-tech-student".
    text = re.sub(r"\.(?=[A-Za-z0-9])", "", (text or "")).replace(".", "")
    slug = re.sub(r"[^A-Za-z0-9]+", "-", text.lower()).strip("-")
    return slug or "keyword"


# ==================================================== production base keywords
# new_mylib.py is the ONE file to edit when the target audience
# changes.  It is deliberately a plain text list rather than Python, so a
# keyword edit can never break the program with a syntax error, and so no
# Python knowledge is needed to run a new campaign.
#
# The old target_audience.txt / target_audience_degree.txt pair held the
# manually expanded "<Base Keyword> <City>" lines -- 20 base keywords times
# 153 cities, written out by hand.  load_keywords() still reads them (nothing
# that depended on it changed), but the search plan is built from the base
# keywords below crossed with py instead.



# The live target audience now lives beside the cities, in new_mylib.py.
KEYWORD_SOURCE_FILE = "new_mylib.py"
KEYWORD_SOURCE_DICT = "keyword_by_category"

# A comment.  Unlike target_audience.txt this file is machine-oriented, so
# the only decoration it may carry is a '#' comment and blank space.
_BASE_COMMENT = re.compile(r"^\s*#")


def base_keyword_rejection(line):
    """Why this line of the base file is not a keyword, or "".

    Deliberately stricter and simpler than rejection_reason(): this file is
    ours, so anything odd in it is a mistake worth reporting rather than
    decoration worth tolerating.
    """
    raw = line.rstrip("\n")
    if not raw.strip():
        return "blank line"
    if _BASE_COMMENT.match(raw):
        return "comment"
    text = _clean(raw)
    if not text:
        return "nothing left after cleaning"
    if not _HAS_LETTERS.search(text):
        return "no letters"
    if _PLACEHOLDER.search(raw):
        return "format placeholder"
    if len(text.split()) > MAX_KEYWORD_WORDS:
        return "too long to be a search phrase (%d words)" % len(text.split())
    return ""


def load_base_keywords(path=None):
    """The resolved production base keyword list.

    Loaded from new_mylib.py's keyword_by_category -- the same file the cities
    come from, so one file is the single place the target audience is edited.
    Groups are kept in the file's own PRIORITY order (IT1 -> IT2 -> Cat2 ->
    Cat3 -> Shared) and flattened in that order; the order is part of the
    plan's identity, so it is never sorted.

    De-duplicated case-insensitively, keeping the first (highest-priority)
    spelling, because the same keyword may legitimately appear under more
    than one group.

    PARSED, never imported: new_mylib.py also carries job-role data and
    imports subprocess/sys, and a keyword list must not execute anything.
    """
    if path is not None:
        return _load_base_keywords_from_text_file(path)
    groups = load_keyword_groups()
    keywords, seen = [], set()
    for names in groups.values():
        for name in names:
            name = _clean(name)
            if not name or base_keyword_rejection(name):
                continue
            key = name.lower()
            if key in seen:
                continue
            seen.add(key)
            keywords.append(name)
    if not keywords:
        raise BaseKeywordError(
            "%s holds no usable keywords in %s."
            % (KEYWORD_SOURCE_FILE, KEYWORD_SOURCE_DICT))
    return keywords


def load_keyword_groups(path=None):
    """{group name: [keyword, ...]} straight from new_mylib.py, in file order."""
    import ast
    path = path or os.path.join(C.PROJECT_ROOT, KEYWORD_SOURCE_FILE)
    if not os.path.exists(path):
        raise BaseKeywordError(
            "%s is missing.  It holds the target audience; the scraper "
            "cannot build a search plan without it." % path)
    with open(path, encoding="utf-8", errors="replace") as handle:
        tree = ast.parse(handle.read())
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if getattr(node.targets[0], "id", "") != KEYWORD_SOURCE_DICT:
            continue
        try:
            data = ast.literal_eval(node.value)
        except Exception as exc:
            raise BaseKeywordError("%s in %s is not a literal dict (%s)"
                                   % (KEYWORD_SOURCE_DICT, path, exc))
        return {str(k): list(v) for k, v in data.items()}
    raise BaseKeywordError("%s not found in %s"
                           % (KEYWORD_SOURCE_DICT, path))


def _load_base_keywords_from_text_file(path):
    """The old one-keyword-per-line reader, kept for explicit paths/tests."""
    if not os.path.exists(path):
        raise BaseKeywordError("%s is missing." % path)
    keywords, seen = [], set()
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            if base_keyword_rejection(line):
                continue
            keyword = _clean(line)
            key = keyword.lower()
            if key in seen:
                continue
            seen.add(key)
            keywords.append(keyword)
    if not keywords:
        raise BaseKeywordError(
            "%s contains no usable base keywords." % path)
    return keywords


class BaseKeywordError(RuntimeError):
    """The production keyword file is missing, empty or unusable."""


def base_keyword_problems(path=None):
    """[(where, text, problem)] for --validate-keywords.

    The keywords come from keyword_by_category in new_mylib.py, so the things
    that can be wrong are different from when they lived in a plain text file:
    there are no blank lines or comments to skip, but a keyword can still be
    empty, unusable, or repeated between groups.

    A repeat BETWEEN groups is normal and not reported -- load_base_keywords()
    keeps the first, highest-priority spelling on purpose.  A repeat WITHIN one
    group is a mistake, and so is anything base_keyword_rejection() refuses.
    """
    try:
        groups = load_keyword_groups(path)
    except BaseKeywordError as exc:
        return [(KEYWORD_SOURCE_FILE, str(exc), "cannot be read")]

    problems = []
    for group, keywords in groups.items():
        seen = {}
        for position, keyword in enumerate(keywords, start=1):
            where = "%s[%d]" % (group, position)
            reason = base_keyword_rejection(keyword)
            if reason in ("blank line", "comment"):
                reason = "empty entry" if reason == "blank line" else None
            if reason:
                problems.append((where, str(keyword).strip(), reason))
                continue
            key = _clean(keyword).lower()
            if key in seen:
                problems.append((where, _clean(keyword),
                                 "duplicate of %s[%d]" % (group, seen[key])))
            else:
                seen[key] = position
    return problems


def keywords_fingerprint(keywords=None, path=None):
    """Stable hash of the RESOLVED, ORDERED production keyword list.

    Over the resolved list, so comments, blank lines and stray whitespace in
    the file are invisible to it.  Over the ORDER as well as the contents,
    because SearchUnit ordering -- and therefore every rank's resume point --
    is derived from it: reordering the active keywords really is a different
    plan and must not be mistaken for the same one.
    """
    resolved = keywords if keywords is not None else load_base_keywords(path)
    payload = "\n".join(resolved)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


# ==========================================================================
# THE LIBRARY
# ==========================================================================
# Everything below was thirty separate modules.  They are kept in
# dependency order with their original headings as banners, so the
# boundaries are still visible.  Nothing in them was rewritten.
# --------------------------------------------------------------------------
# ==========================================================================
# BADGES
# ==========================================================================
# Technologies declared as Markdown badges in a profile README.
#
# GitHub profile READMEs almost never write their tech stack as text.  They
# write it as images:
#
#     ![Java](https://img.shields.io/badge/java-%23ED8B00.svg?...&logo=openjdk)
#
# An <img> contributes nothing to innerText, so reading the rendered page
# gives an EMPTY Tech Stack section.  That is why a profile listing C, Java,
# HTML5 and CSS3 produced only "Java" -- picked up from a sentence, not from
# the stack.  The raw Markdown is the only place those names exist.
#
# Two things must not become skills:
#
#   * social badges.  "Socials" sections use the same shields.io style for
#     Instagram/LinkedIn/X, so the badge syntax alone proves nothing.
#   * stats and decoration.  GitHub-stats cards, streak cards, trophies and
#     visitor counters are images with empty or meaningless alt text.
# --------------------------------------------------------------------------
# ![alt](url) and the linked form [![alt](url)](href)
MD_IMAGE = re.compile(r"!\[([^\]]*)\]\(\s*([^)\s]+)")
# A badge that is also a link: [![LinkedIn](badge-url)](profile-url).
# The inner "]" ends the alt text, so a naive [label](url) pattern returns
# the BADGE url instead of the profile url this form exists to carry.
MD_IMAGE_LINK = re.compile(
    r"\[\s*!\[([^\]]*)\]\([^)]*\)\s*\]\(\s*([^)\s]+)")
MD_LINK = re.compile(r"(?<!!)\[([^\]!]*)\]\(\s*([^)\s]+)")
HTML_IMG = re.compile(
    r"<img[^>]*\balt\s*=\s*[\"']([^\"']*)[\"'][^>]*\bsrc\s*=\s*[\"']"
    r"([^\"']+)[\"']", re.I)

# Hosts that serve dashboards and counters rather than technology badges.
DECORATION_HOSTS = (
    "github-readme-stats", "github-readme-streak", "github-readme-activity",
    "github-contributor-stats", "github-profile-trophy",
    "github-profile-summary", "activity-graph", "visitcount",
    "komarev.com", "profile-counter", "hits.seeyoufarm", "hits.dwyl",
    "wakatime.com", "spotify-github-profile", "readme-typing-svg",
    "capsule-render", "github-widgetbox", "streak-stats",
    "leetcode-stats", "codersrank", "holopin", "forthebadge.com/images",
)

SHIELDS_HOSTS = ("img.shields.io", "shields.io", "badgen.net",
                 "badge.fury.io", "custom-icon-badges")


def _is_decoration(url):
    low = (url or "").lower()
    return any(host in low for host in DECORATION_HOSTS)


def is_shields(url):
    low = (url or "").lower()
    return any(host in low for host in SHIELDS_HOSTS)


def shields_label(url):
    """The label a shields.io badge renders, or "".

    https://img.shields.io/badge/java-%23ED8B00.svg?...  ->  "java"
    """
    if not is_shields(url):
        return ""
    match = re.search(r"/badge/([^/?]+)", url)
    if not match:
        # /static/v1?label=Java&...
        query = urllib.parse.urlparse(url).query
        label = urllib.parse.parse_qs(query).get("label", [""])[0]
        return _clean_label(label)
    raw = match.group(1)
    # "java-%23ED8B00.svg" -> "java";  "-" is the field separator, and a
    # literal dash in a label is written "--".
    raw = raw.split(".svg")[0].split(".png")[0]
    raw = re.sub(r"(?<!-)-(?!-)", "\x00", raw).split("\x00")[0]
    return _clean_label(raw)


def logo_name(url):
    """The ?logo= slug, which often names the technology precisely."""
    if not is_shields(url):
        return ""
    query = urllib.parse.urlparse(url).query
    return _clean_label(urllib.parse.parse_qs(query).get("logo", [""])[0])


def _clean_label(value):
    value = urllib.parse.unquote(value or "")
    value = value.replace("--", "-").replace("_", " ")
    value = re.sub(r"%[0-9A-Fa-f]{2}", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def extract_badges(markdown):
    """Every image in the Markdown, as {alt, url, label, logo, shields}."""
    found = []
    seen = set()
    text = markdown or ""
    for pattern in (MD_IMAGE, HTML_IMG):
        for match in pattern.finditer(text):
            alt, url = match.group(1).strip(), match.group(2).strip()
            key = (alt.lower(), url)
            if key in seen:
                continue
            seen.add(key)
            found.append({
                "alt": alt,
                "url": url,
                "label": shields_label(url),
                "logo": logo_name(url),
                "shields": is_shields(url),
                "decoration": _is_decoration(url),
            })
    return found


def badge_names(markdown, include_decoration=False):
    """Candidate technology names from badges, best source first.

    Alt text wins because the author chose it; the shields label and the
    ?logo= slug are fallbacks for badges written without alt text.
    """
    names = []
    for badge in extract_badges(markdown):
        if badge["decoration"] and not include_decoration:
            continue
        for value in (badge["alt"], badge["label"], badge["logo"]):
            value = (value or "").strip()
            if value and value not in names:
                names.append(value)
                break
    return names


def markdown_links(markdown):
    """[(label, destination)] for Markdown links, badge images resolved."""
    text = markdown or ""
    links = []
    seen = set()
    # Linked badges first, so their destination wins over the badge image.
    for match in MD_IMAGE_LINK.finditer(text):
        label, url = match.group(1).strip(), match.group(2).strip()
        if (label, url) not in seen:
            seen.add((label, url))
            links.append((label, url))
    for match in MD_LINK.finditer(text):
        label, url = match.group(1).strip(), match.group(2).strip()
        if not label or is_shields(url):
            continue
        if (label, url) not in seen:
            seen.add((label, url))
            links.append((label, url))
    return links


# ==========================================================================
# BROWSER MANAGER
# ==========================================================================
# Persistent Chromium with Instant Data Scraper loaded.
#
# Extensions only load in the full Chromium build, never in Playwright's
# headless shell, so channel="chromium" is always requested.  The persistent
# profile carries the GitHub login, the cookies and the extension's own stored
# state.
#
# RANK-LOCAL PROFILE -- THE HARD REQUIREMENT
# ------------------------------------------
# The profile directory and the download directory come from a RankContext, not
# from a module-level constant.  Two Playwright persistent contexts must NEVER
# share one user-data-dir: Chromium holds a SingletonLock on it, so the second
# launch fails outright, and if it somehow succeeded both browsers would fight
# over the same cookie jar and the same extension storage.  Deriving the path
# from the rank number makes that collision impossible rather than merely
# discouraged.
#
# Downloads are rank-local for a subtler reason: InstantDataScraper rescues a
# finished-but-unrenamed export by scanning the download directory for the
# newest .xlsx.  Shared, rank 4 would adopt rank 63's export and process the
# wrong city's candidates.
# --------------------------------------------------------------------------


# ---------------------------------------------------------- github session --
# naukri keeps its login in naukri_auth.json and passes it to
# new_context(storage_state=...).  launch_persistent_context() -- which an
# unpacked extension requires -- accepts no storage_state, so the same idea is
# split in two: save_auth() writes the cookies out, load_auth() puts them back.


def github_cookies(cookies):
    """Just the github.com cookies -- the login, and nothing else.

    A crawl visits every student's personal site, so an unfiltered
    storage_state() collects their ad-network and analytics cookies too: 209
    cookies across 73 domains, of which 25 were GitHub's.  Only GitHub's are
    the session, and this file is written on every shutdown now, so the rest
    are dropped rather than accumulated forever.

    Matched on the domain SUFFIX, not a substring: students' portfolio sites
    live on *.github.io, which contains "github" but is not GitHub.
    """
    keep = []
    for cookie in cookies:
        host = (cookie.get("domain") or "").lstrip(".").lower()
        if host == "github.com" or host.endswith(".github.com"):
            keep.append(cookie)
    return keep


async def save_github_auth(context, path=None):
    """Write this context's cookies to github_auth.json.  Returns the count."""
    path = path or C.AUTH_FILE
    state = await context.storage_state()
    cookies = github_cookies(state.get("cookies", []))
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump({"cookies": cookies, "origins": []}, handle, indent=2)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(tmp, path)
    try:
        os.chmod(path, 0o600)             # it IS the session; not world-readable
    except OSError:
        pass
    return len(cookies)


async def load_github_auth(context, path=None):
    """Inject github_auth.json into a context.  Returns the count, or 0.

    Never fatal: a missing or unreadable file just means this copy falls back
    on whatever the profile directory already holds, which on a machine that
    has run --login is the same session.
    """
    path = path or C.AUTH_FILE
    if not os.path.exists(path):
        return 0
    try:
        with open(path, encoding="utf-8") as handle:
            cookies = json.load(handle).get("cookies", [])
    except Exception as exc:
        log.warning("%s could not be read (%s); using the profile as-is",
                    os.path.basename(path), exc)
        return 0
    if not cookies:
        return 0
    try:
        await context.add_cookies(cookies)
    except Exception as exc:
        log.warning("cookies from %s were rejected (%s); using the profile "
                    "as-is", os.path.basename(path), exc)
        return 0
    return len(cookies)


def discard_profile(path=None):
    """Delete the disposable Chromium profile.  Never raises.

    Called at shutdown AND before the first launch of the next run.  The
    second call is the one that actually guarantees the bound: if a process
    is killed outright its shutdown never runs, and without this the leftover
    cache would be inherited by the next run and grow again.
    """
    path = path or C.PROFILE_DIR
    if not C.PROFILE_DISPOSABLE or not os.path.isdir(path):
        return False
    shutil.rmtree(path, ignore_errors=True)
    return not os.path.isdir(path)


# A killed process runs no shutdown code, so this is a courtesy, not the
# guarantee -- start() re-checks before the first launch.
atexit.register(discard_profile)


class BrowserLost(RuntimeError):
    """The Chromium process died; nothing may be cleaned up."""


def browser_died(exc):
    text = str(exc).lower()
    return any(mark in text for mark in (
        "target page, context or browser has been closed",
        "browser has been closed", "browser closed",
        "connection closed", "target closed", "crashed",
    ))


class BrowserManager:
    def __init__(self, headless=True, profile_dir=None, download_dir=None):
        # One folder is one rank, so the profile directory is this folder's
        # own disposable profile in the system temp directory -- named after
        # this folder's path, so two copies can no more share it than two
        # separate checkouts can.  The arguments exist for the test suite.
        self.profile_dir = profile_dir or C.PROFILE_DIR
        self.download_dir = download_dir or C.KEYWORD_DOWNLOAD_DIR
        self.headless = headless
        self.playwright = None
        self.context = None
        self.extension_dir = ""
        self.extension_id = C.EXTENSION_ID
        self._context_closed = False
        self.restarts = 0
        # Students finished since the last restart; see note_healthy().
        self._healthy_since_restart = 0
        # The one page the workflow should ever work through.
        self._page = None

    # -- extension ----------------------------------------------------------
    def validate_extension(self):
        self.extension_dir = C.find_extension_dir()
        if not self.extension_dir:
            raise SystemExit(
                "Instant Data Scraper was not found on this machine.\n"
                "Searched: %s\n"
                "Install it in Chromium, or set IDS_EXTENSION_DIR to the\n"
                "unpacked extension directory, then run:\n"
                "  python3 student_scraper.py --check-extension"
                % ", ".join(C.EXTENSION_SEARCH_ROOTS))
        return self.extension_dir

    # -- lifecycle ----------------------------------------------------------
    async def start(self, with_extension=True):
        C.ensure_dirs()
        # A restart deliberately reuses the live profile (see restart()); only
        # the FIRST launch of a run starts from a clean directory.
        if self.restarts == 0:
            discard_profile(self.profile_dir)
        os.makedirs(self.profile_dir, exist_ok=True)
        os.makedirs(self.download_dir, exist_ok=True)
        args = list(C.CHROMIUM_LEAN_ARGS)
        if C.LOW_MEMORY_MODE:
            args += list(C.CHROMIUM_LOW_MEMORY_ARGS)
        if with_extension:
            self.validate_extension()
            args += ["--disable-extensions-except=%s" % self.extension_dir,
                     "--load-extension=%s" % self.extension_dir]
        self.playwright = await async_playwright().start()
        self.context = await self.playwright.chromium.launch_persistent_context(
            self.profile_dir,
            channel="chromium",
            headless=self.headless,
            accept_downloads=True,
            downloads_path=self.download_dir,
            viewport={"width": 1440, "height": 960},
            args=args,
        )
        self.context.set_default_navigation_timeout(C.NAVIGATION_TIMEOUT)
        self._context_closed = False
        self._page = None
        # Chromium dying is not an exception anywhere -- it is a context
        # close event.  Without listening for it the next call simply
        # raises "Target page, context or browser has been closed".
        self.context.on("close", self._on_context_close)
        # The GitHub session, as naukri loads naukri_auth.json.  Done here
        # rather than through storage_state= because launch_persistent_context
        # has no such argument.  Harmless when the profile is already logged
        # in: the same cookies are simply set again.
        loaded = await load_github_auth(self.context)
        if loaded:
            log.info("loaded %d cookie(s) from %s", loaded,
                     os.path.basename(C.AUTH_FILE))
        if with_extension:
            self.extension_id = await self._wait_for_extension()
        return self.context

    def _on_context_close(self, _context=None):
        self._context_closed = True
        self._page = None
        log.warning("browser context closed")

    # -- health ------------------------------------------------------------
    async def is_alive(self):
        """True only if the browser can still execute something.

        Checking attributes is not enough: a dead Chromium leaves the
        Python objects in place and only fails when they are used.
        """
        if self.context is None or self._context_closed:
            return False
        try:
            pages = [page for page in self.context.pages
                     if not page.is_closed()]
            if not pages:
                probe = await self.context.new_page()
                await probe.close()
                return True
            await pages[0].evaluate("1")
            return True
        except Exception as exc:
            log.warning("health check failed: %s", exc)
            return False

    def page_count(self):
        try:
            return len(self.context.pages) if self.context else 0
        except Exception:
            return -1

    async def restart(self, with_extension=True):
        """Relaunch Chromium on the SAME persistent profile.

        The profile carries the GitHub login and the extension state, so a
        restart costs a page load, not a re-login and not another crawl.
        """
        self.restarts += 1
        self._healthy_since_restart = 0
        log.warning("restarting Chromium (restart #%d)", self.restarts)
        try:
            await self.close()
        except Exception as exc:
            log.warning("could not close the dead browser cleanly: %s", exc)
        self.playwright = None
        self.context = None
        self._context_closed = False
        self._page = None
        self._clear_stale_singleton()
        await asyncio.sleep(2)
        await self.start(with_extension=with_extension)
        return await self.page()

    def note_healthy(self):
        """Credit one student that finished without the browser dying.

        Hands the restart budget back once C.BROWSER_RESTART_FORGIVE_AFTER
        students have been collected since the last relaunch, so the budget
        counts a STREAK of failures rather than the whole life of the
        process.  Cheap to call and safe to call often: it does nothing at
        all while the budget is already full.
        """
        if self.restarts == 0:
            self._healthy_since_restart = 0
            return
        self._healthy_since_restart += 1
        if self._healthy_since_restart < C.BROWSER_RESTART_FORGIVE_AFTER:
            return
        log.info("browser healthy for %d students; restart budget reset "
                 "(was %d of %d used)", self._healthy_since_restart,
                 self.restarts, C.BROWSER_MAX_RESTARTS)
        self.restarts = 0
        self._healthy_since_restart = 0

    def _clear_stale_singleton(self):
        """Remove a SingletonLock left by a Chromium that is gone.

        Chromium refuses to launch while the lock exists.  It is only
        removed when the process it names is definitely not running, so a
        live browser is never disturbed.
        """
        lock = os.path.join(self.profile_dir, "SingletonLock")
        if not os.path.islink(lock) and not os.path.exists(lock):
            return
        try:
            target = os.readlink(lock)
        except OSError:
            return
        pid = target.rsplit("-", 1)[-1]
        if pid.isdigit() and os.path.exists("/proc/%s" % pid):
            log.warning("SingletonLock belongs to live pid %s; leaving it",
                        pid)
            return
        for name in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            path = os.path.join(self.profile_dir, name)
            try:
                if os.path.islink(path) or os.path.exists(path):
                    os.unlink(path)
                    log.info("removed stale %s", name)
            except OSError as exc:
                log.warning("could not remove %s: %s", name, exc)

    async def _wait_for_extension(self):
        deadline = C.EXTENSION_TIMEOUT / 1000.0
        waited = 0.0
        while waited < deadline:
            worker = self.service_worker
            if worker is not None:
                return worker.url.split("/")[2]
            await asyncio.sleep(0.5)
            waited += 0.5
        raise SystemExit(
            "Instant Data Scraper did not start.\n"
            "Extension directory: %s\n"
            "Run: python3 student_scraper.py --check-extension"
            % self.extension_dir)

    @property
    def service_worker(self):
        for worker in self.context.service_workers:
            if worker.url.startswith("chrome-extension://"):
                return worker
        return None

    async def page(self):
        """Reuse the profile's first tab so only one window is ever open."""
        if self.context.pages:
            return self.context.pages[0]
        return await self.context.new_page()

    async def get_live_page(self):
        """The authoritative working page, or None if the browser is gone.

        A page object captured before a restart points at a target that no
        longer exists, and using it raises TargetClosedError.  Callers must
        ask for the page again at every phase rather than holding one.
        """
        if self.context is None or self._context_closed:
            return None
        try:
            if self._page is not None and not self._page.is_closed():
                return self._page
            for candidate in self.context.pages:
                if candidate.is_closed():
                    continue
                # The extension's own options page is not a working tab.
                if candidate.url.startswith("chrome-extension://"):
                    continue
                self._page = candidate
                return candidate
            self._page = await self.context.new_page()
            return self._page
        except Exception as exc:
            log.warning("could not obtain a live page: %s", exc)
            return None

    async def close_extra_pages(self, keep=()):
        """Close every page that is not actively in use.

        Instant Data Scraper opens its own options/onboarding page
        (chrome-extension://<id>/src/options/options.html) and leaves it
        open for the whole session, which costs a renderer process for
        nothing.  Portfolio and resume pages close themselves, so anything
        else still open here is surplus.

        `keep` must list the pages currently being driven -- the working
        tab and, during collection, the extension popup.
        """
        kept = {id(page) for page in keep if page is not None}
        closed = 0
        for page in list(self.context.pages):
            if id(page) in kept or page.is_closed():
                continue
            try:
                await page.close()
                closed += 1
                log.debug("closed surplus page %s", page.url[:80])
            except Exception as exc:
                log.debug("could not close %s: %s", page.url[:60], exc)
        return closed

    async def release_page_memory(self, page):
        """Drop the current document so its renderer memory can be reused.

        Navigating to the next profile replaces the DOM anyway; doing it
        explicitly between students stops one heavy page (a profile with a
        long README, or a portfolio) from staying resident while the next
        network round-trip happens.
        """
        try:
            await page.goto("about:blank", wait_until="domcontentloaded",
                            timeout=10000)
            return True
        except Exception as exc:
            log.debug("could not release page memory: %s", exc)
            return False

    async def chrome_tab_id(self, page):
        """The real chrome tabs API id for a Playwright page.

        Needed because background.js addresses the popup with ?tabid=<id>.
        Playwright never exposes it, so the extension's own service worker is
        asked through chrome.tabs.query.
        """
        worker = self.service_worker
        if worker is None:
            raise RuntimeError("extension service worker is not available")
        target = page.url.split("#")[0]
        tabs = await worker.evaluate(
            "(async () => (await chrome.tabs.query({}))"
            ".map(t => ({id: t.id, url: t.url, active: t.active})))()")
        for tab in tabs:
            if (tab.get("url") or "").split("#")[0] == target:
                return tab["id"]
        for tab in tabs:
            if "github.com" in (tab.get("url") or ""):
                return tab["id"]
        raise RuntimeError("no chrome tab matched %s" % target)

    # -- github -------------------------------------------------------------
    async def github_login_name(self, page):
        await page.goto("https://github.com", wait_until="domcontentloaded")
        try:
            return await page.get_attribute('meta[name="user-login"]',
                                            "content", timeout=5000)
        except Exception:
            return None

    async def close(self):
        # THE SESSION LEAVES WITH THE PROFILE, so it is written out first.
        # GitHub rotates user_session during a long run; that refreshed cookie
        # used to survive only inside .pw_profile/Default/Cookies, which meant
        # github_auth.json went stale while the profile quietly held the real
        # login.  Now the portable file is the one that carries it.
        if self.context is not None and not self._context_closed:
            try:
                count = await save_github_auth(self.context)
                log.info("saved %d github cookie(s) to %s", count,
                         os.path.basename(C.AUTH_FILE))
            except Exception as exc:
                log.warning("could not save the session (%s); %s is unchanged",
                            exc, os.path.basename(C.AUTH_FILE))
        for closer in (getattr(self.context, "close", None),
                       getattr(self.playwright, "stop", None)):
            if closer is None:
                continue
            try:
                await closer()
            except Exception:
                pass


# ==========================================================================
# BUNDLE
# ==========================================================================
# Portable device result bundles: rank stores + manifest + checksums.
#
# A bundle is what one of the 8 devices hands to the master laptop for the final
# merge.  It contains ONLY what the merge needs -- the rank SQLite stores, the
# campaign manifest, and a bundle manifest with a checksum per store -- and
# deliberately NOT the Playwright profiles, browser caches, temporary downloads,
# or logs (those are large, machine-specific and irrelevant to the data).
#
#     device_1_results/
#         json                 device range, campaign, checksums
#         campaign.json               the shared campaign manifest
#         stores/
#             rank_001.students.db
#             rank_002.students.db
#             ...
#
# The bundle manifest lets the master merge REFUSE to mix incompatible data: a
# corrupted store (checksum mismatch), a store from a different campaign, a
# different plan fingerprint, a different total_ranks, a missing rank, or the
# same rank supplied twice.
# --------------------------------------------------------------------------
BUNDLE_VERSION = 1


class BundleError(RuntimeError):
    """A bundle cannot be created, read or trusted."""


def file_checksum(path, chunk=1 << 20):
    """sha256 of a file, streamed so a large DB is not read into memory."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(chunk), b""):
            digest.update(block)
    return digest.hexdigest()


def create_bundle(destination, ranks, store_paths, manifest,
                  total_ranks, include_logs=False, log_paths=None,
                  stamp=None):
    """Assemble a bundle at `destination`.

    ranks         the rank numbers this device owns and is exporting
    store_paths   {rank: path to that rank's students.db}
    manifest      the campaign manifest (dict) every device shares
    Returns the bundle manifest dict that was written.
    """
    if os.path.exists(destination):
        raise BundleError("%s already exists; choose a new bundle path or "
                          "remove the old one" % destination)
    stores_dir = os.path.join(destination, "stores")
    os.makedirs(stores_dir)

    entries = {}
    for rank in ranks:
        source = store_paths.get(rank)
        if not source or not os.path.exists(source):
            # A rank with no store yet (never ran) is recorded as absent
            # rather than fabricated -- the master merge reports it.
            entries[str(rank)] = {"present": False}
            continue
        name = "rank_%03d.students.db" % rank
        shutil.copyfile(source, os.path.join(stores_dir, name))
        entries[str(rank)] = {
            "present": True,
            "file": "stores/%s" % name,
            "sha256": file_checksum(os.path.join(stores_dir, name)),
            "bytes": os.path.getsize(os.path.join(stores_dir, name)),
        }

    # The campaign manifest travels with the data so the master can check it.
    campaign_copy = os.path.join(destination, "campaign.json")
    with open(campaign_copy, "w", encoding="utf-8") as handle:
        json.dump(manifest, handle, indent=2, ensure_ascii=False)

    if include_logs and log_paths:
        logs_dir = os.path.join(destination, "logs")
        os.makedirs(logs_dir, exist_ok=True)
        for rank, log_path in log_paths.items():
            if log_path and os.path.exists(log_path):
                shutil.copyfile(log_path,
                                os.path.join(logs_dir, "rank_%03d.log" % rank))

    bundle_manifest = {
        "bundle_version": BUNDLE_VERSION,
        "created_at": stamp or datetime.datetime.now().replace(
            microsecond=0).isoformat(),
        "campaign_id": manifest.get("campaign_id", ""),
        "plan_fingerprint": manifest.get("plan_fingerprint", ""),
        "keywords_fingerprint": manifest.get("keywords_fingerprint", ""),
        "locations_fingerprint": manifest.get("locations_fingerprint", ""),
        "total_ranks": int(total_ranks),
        "ranks": sorted(int(rank) for rank in ranks),
        "stores": entries,
        "includes_logs": bool(include_logs and log_paths),
    }
    tmp = os.path.join(destination, ".bundle.json.tmp")
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(bundle_manifest, handle, indent=2, ensure_ascii=False)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(tmp, os.path.join(destination, "json"))
    return bundle_manifest


def read_bundle(path):
    """Load and structurally validate a bundle's manifest.  Returns the dict.

    Does NOT verify checksums (that is verify_bundle, which touches the DB
    files); this only proves the manifest itself is well formed.
    """
    manifest_path = os.path.join(path, "json")
    if not os.path.isdir(path) or not os.path.exists(manifest_path):
        raise BundleError("%s is not a device bundle (no json)" % path)
    try:
        with open(manifest_path, encoding="utf-8") as handle:
            manifest = json.load(handle)
    except Exception as exc:
        raise BundleError("%s/json is unreadable (%s)" % (path, exc))
    for key in ("campaign_id", "plan_fingerprint", "total_ranks", "ranks",
                "stores"):
        if key not in manifest:
            raise BundleError("%s/json is missing '%s'" % (path, key))
    return manifest


def verify_bundle(path):
    """Check every store's checksum against the manifest.  Returns the manifest.

    Raises BundleError on a corrupted or missing store, so a damaged transfer
    can never be merged silently.
    """
    manifest = read_bundle(path)
    for rank, entry in manifest["stores"].items():
        if not entry.get("present"):
            continue
        store_path = os.path.join(path, entry["file"])
        if not os.path.exists(store_path):
            raise BundleError(
                "%s: rank %s claims a store (%s) that is missing"
                % (path, rank, entry["file"]))
        actual = file_checksum(store_path)
        if actual != entry.get("sha256"):
            raise BundleError(
                "%s: rank %s store is corrupted (checksum mismatch)"
                % (path, rank))
    return manifest


def bundle_store_paths(path):
    """{rank: absolute store path} for the PRESENT stores in a bundle."""
    manifest = read_bundle(path)
    stores = {}
    for rank, entry in manifest["stores"].items():
        if entry.get("present"):
            stores[int(rank)] = os.path.join(path, entry["file"])
    return stores


def combine_bundles(paths):
    """Validate a set of bundles together for the master merge.

    Returns (store_paths_by_rank, report).  Raises BundleError on:
      * a corrupted or missing store
      * two bundles supplying the same rank
      * bundles from different campaigns / plans / total_ranks
    """
    if not paths:
        raise BundleError("no bundles were supplied")
    campaign_id = plan_fp = None
    total_ranks = None
    store_paths = {}
    owner = {}
    covered = set()
    per_bundle = []

    for path in paths:
        manifest = verify_bundle(path)
        if campaign_id is None:
            campaign_id = manifest["campaign_id"]
            plan_fp = manifest["plan_fingerprint"]
            total_ranks = manifest["total_ranks"]
        else:
            if manifest["campaign_id"] != campaign_id:
                raise BundleError(
                    "campaign mismatch: %s is campaign %s, expected %s"
                    % (path, manifest["campaign_id"], campaign_id))
            if manifest["plan_fingerprint"] != plan_fp:
                raise BundleError(
                    "plan fingerprint mismatch in %s; these bundles are not "
                    "from the same search plan" % path)
            if manifest["total_ranks"] != total_ranks:
                raise BundleError(
                    "total-ranks mismatch: %s has %d, expected %d"
                    % (path, manifest["total_ranks"], total_ranks))
        for rank, store in bundle_store_paths(path).items():
            if rank in owner:
                raise BundleError(
                    "rank %d is supplied by two bundles (%s and %s)"
                    % (rank, owner[rank], path))
            owner[rank] = path
            store_paths[rank] = store
        covered.update(manifest["ranks"])
        per_bundle.append((path, manifest))

    missing = [rank for rank in range(1, (total_ranks or 0) + 1)
               if rank not in covered]
    report = {
        "campaign_id": campaign_id,
        "plan_fingerprint": plan_fp,
        "total_ranks": total_ranks,
        "bundles": len(paths),
        "ranks_covered": sorted(covered),
        "ranks_missing": missing,
        "stores": store_paths,
        "per_bundle": per_bundle,
    }
    return store_paths, report


def _compact_ranks(numbers):
    """Readable rank spans for a report: [1,2,3,7] -> '1-3, 7'."""
    if not numbers:
        return ""
    numbers = sorted(numbers)
    spans, start, prev = [], numbers[0], numbers[0]
    for value in numbers[1:]:
        if value == prev + 1:
            prev = value
            continue
        spans.append((start, prev))
        start = prev = value
    spans.append((start, prev))
    return ", ".join("%d" % a if a == b else "%d-%d" % (a, b)
                     for a, b in spans)


# ==========================================================================
# CYCLE WRITER
# ==========================================================================
# current_data.xlsx and its sealed cycle files.
#
#     append_detail_to_excel(record)
#         current_data.xlsx            <- appended to, row by row
#             reaches CYCLE_SIZE rows?
#                 copy  -> current_data_cycle1.xlsx    (sealed, never reopened)
#                 empty -> current_data.xlsx           (keeps going)
#
#         current_data_cycle1.xlsx, current_data_cycle2.xlsx, ...
#
# This is the same output contract the job fetching scripts already use, so
# whatever collects from them can collect from this folder without being told
# anything new.
#
# WHY A SEALED FILE MATTERS
# -------------------------
# A collector copying current_data.xlsx while a row is being appended would get
# a half-written workbook.  A *cycle* file is finished the moment it is created
# and is never opened for writing again, so it can be copied away at any time.
# The collector therefore only ever looks for names ending `cycleN.xlsx` and
# leaves current_data.xlsx alone.  That is the whole reason for the rollover.
#
# WHY THE CYCLE NUMBER IS IN A FILE
# ---------------------------------
# cycle_state.json remembers the last number used.  Deriving it by listing the
# directory instead would restart at 1 as soon as a collector removed the files
# it had collected, and current_data_cycle1.xlsx would then be written twice --
# the second one silently replacing a batch of 100 people that had already been
# handed over.
#
# EVERY WRITE IS ATOMIC
# ---------------------
# Workbooks are built in full, saved to a temporary file in the same directory
# and then os.replace()d into position.  A crash during a save therefore leaves
# either the previous workbook or the new one, never a truncated file -- which
# matters here because the previous workbook is the only copy of up to
# CYCLE_SIZE people.
#
# openpyxl, not pandas: the rest of this project already reads and writes xlsx
# with openpyxl, and the behaviour is identical.
# --------------------------------------------------------------------------
SHEET_TITLE = "Students"


# ------------------------------------------------------------- reading ----
def _move_aside(path):
    """Preserve an unreadable workbook instead of deleting it.

    A corrupt current_data.xlsx still holds up to CYCLE_SIZE real people.
    Deleting it to keep going -- which is the obvious thing to do -- throws
    them away silently, so it is renamed instead and the loss is visible.
    """
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    spoiled = "%s.corrupt_%s" % (path, stamp)
    try:
        os.replace(path, spoiled)
        log.warning("unreadable workbook moved aside: %s", spoiled)
    except OSError as exc:
        log.warning("could not move aside %s (%s); removing it", path, exc)
        try:
            os.remove(path)
        except OSError:
            pass


def read_rows(path=None, columns=None):
    """Every data row of `path` as a list of dicts.  Missing file -> [].

    A file that cannot be read is moved aside and treated as empty, so one
    bad workbook can never stop collection.
    """
    path = path or C.DETAIL_OUTPUT_FILE
    columns = columns or C.FINAL_COLUMNS
    if not os.path.exists(path):
        return []
    try:
        if os.path.getsize(path) == 0:
            raise ValueError("workbook is empty")
        book = load_workbook(path, read_only=True, data_only=True)
    except (zipfile.BadZipFile, InvalidFileException, ValueError, OSError,
            EOFError, KeyError) as exc:
        # An .xlsx IS a zip archive, so a truncated or non-workbook file
        # surfaces as BadZipFile rather than anything openpyxl defines.
        log.warning("%s is unreadable (%s)", path, exc)
        _move_aside(path)
        return []

    try:
        sheet = book.active
        rows = sheet.iter_rows(values_only=True)
        try:
            header = list(next(rows))
        except StopIteration:
            return []
        header = [str(name) if name is not None else "" for name in header]
        records = []
        for row in rows:
            if row is None or all(cell is None for cell in row):
                continue
            record = {}
            for index, name in enumerate(header):
                if not name:
                    continue
                value = row[index] if index < len(row) else None
                record[name] = "" if value is None else value
            records.append(record)
        return records
    finally:
        book.close()


def row_count(path=None):
    """How many data rows current_data.xlsx holds right now."""
    return len(read_rows(path))


# ------------------------------------------------------------- writing ----
def _save_atomically(records, path, columns):
    book = Workbook()
    sheet = book.active
    sheet.title = SHEET_TITLE
    sheet.append(list(columns))
    for record in records:
        sheet.append([_cell(record.get(name, "")) for name in columns])

    directory = os.path.dirname(path) or "."
    os.makedirs(directory, exist_ok=True)
    tmp = os.path.join(directory, ".%s.tmp" % os.path.basename(path))
    try:
        book.save(tmp)
        os.replace(tmp, path)
    finally:
        book.close()
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def _cell(value):
    """openpyxl refuses unknown types; everything odd becomes text."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool,
                          datetime.datetime, datetime.date)):
        return value
    return str(value)


# -------------------------------------------------------- cycle numbers ----
def load_cycle_num():
    """The last cycle number written, or 0 when none has been."""
    if not os.path.exists(C.CYCLE_STATE_FILE):
        return 0
    try:
        with open(C.CYCLE_STATE_FILE, encoding="utf-8") as handle:
            return int((json.load(handle) or {}).get("cycle_num", 0))
    except Exception:
        return 0


def save_cycle_num(cycle_num):
    """Record the cycle number just written."""
    tmp = C.CYCLE_STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump({"cycle_num": int(cycle_num)}, handle, indent=4)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(tmp, C.CYCLE_STATE_FILE)


def get_next_cycle_file(base_file=None):
    """(path, number) of the next cycle file.

        cycle_num = 0 -> current_data_cycle1.xlsx
        cycle_num = 1 -> current_data_cycle2.xlsx
        cycle_num = 2 -> current_data_cycle3.xlsx

    The name is never reused: if a file with the computed name somehow still
    exists -- a collector that has not taken it yet -- the number is advanced
    until a free one is found rather than overwriting a sealed batch.
    """
    base_file = base_file or C.DETAIL_OUTPUT_FILE
    stem = os.path.splitext(base_file)[0]
    number = load_cycle_num() + 1
    while os.path.exists("%s_cycle%d.xlsx" % (stem, number)):
        log.warning("cycle file %d already exists; skipping to %d",
                    number, number + 1)
        number += 1
    return "%s_cycle%d.xlsx" % (stem, number), number


# -------------------------------------------------------------- append ----
def append_detail_to_excel(record, output_file=None, cycle_size=None,
                           columns=None, say=None):
    """Append one person to current_data.xlsx, rolling over at cycle_size.

    Returns the sealed cycle file's path when this row completed a cycle, and
    "" otherwise, so a caller can report it.

    The order below is the part that matters.  The cycle file is written
    FIRST and current_data.xlsx is emptied SECOND, so a crash between the two
    leaves the same 100 rows in both places -- visible and recoverable --
    rather than in neither.
    """
    output_file = output_file or C.DETAIL_OUTPUT_FILE
    cycle_size = int(cycle_size or C.CYCLE_SIZE)
    columns = columns or C.FINAL_COLUMNS
    talk = say or (lambda message: log.info("%s", message))

    records = read_rows(output_file, columns)
    records.append(dict(record))

    if len(records) < cycle_size:
        _save_atomically(records, output_file, columns)
        talk("[CYCLE] %s | rows: %d/%d"
             % (os.path.basename(output_file), len(records), cycle_size))
        return ""

    cycle_file, cycle_num = get_next_cycle_file(output_file)
    _save_atomically(records, cycle_file, columns)
    save_cycle_num(cycle_num)
    _save_atomically([], output_file, columns)
    talk("[CYCLE] completed -> %s | rows: %d"
         % (os.path.basename(cycle_file), len(records)))
    talk("[CYCLE] emptied   -> %s" % os.path.basename(output_file))
    return cycle_file


def cycle_files(directory=None, base_file=None):
    """Every sealed cycle file present, ordered by cycle number."""
    base_file = base_file or C.DETAIL_OUTPUT_FILE
    directory = directory or os.path.dirname(base_file) or "."
    stem = os.path.basename(os.path.splitext(base_file)[0]) + "_cycle"
    found = []
    try:
        names = os.listdir(directory)
    except OSError:
        return []
    for name in names:
        if not name.startswith(stem) or not name.endswith(".xlsx"):
            continue
        digits = name[len(stem):-len(".xlsx")]
        if digits.isdigit():
            found.append((int(digits), os.path.join(directory, name)))
    return [path for _number, path in sorted(found)]




# ==========================================================================
# DEFERRED PROFILES
# ==========================================================================
# Profiles set aside after repeated transient failures -- the deferred queue.
#
# WHY THIS EXISTS
# ---------------
# A GitHub rate-limit / challenge on ONE profile used to stop the whole rank:
# the search never completed, so no later candidate or search was processed.
# After MAX_DEFERRED_FAILURES separate transient failures a profile is moved
# here instead, so the rank keeps going.  It is:
#
#   * NOT a blacklist -- it is retried at end-of-rank and by --retry-deferred,
#     and removed the moment it finally succeeds;
#   * NOT counted as a collected student;
#   * NOT deleted -- its identity and failure history are preserved.
#
# This is a sibling of py: transient holds profiles still
# under their retry budget; deferred holds profiles that exhausted it.  Both are
# rank-local (runtime/rank_NNN/), atomic and crash-safe.
# --------------------------------------------------------------------------
class DeferredQueue:
    """Atomic, crash-safe JSON of deferred profiles, keyed by lower username."""

    def __init__(self, path=None):
        # One folder is one rank, so this folder's json is
        # already rank-local; another copy cannot reach it.
        self.path = path or C.DEFERRED_FILE
        self.data = self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return {"version": 1, "deferred": {}}
        try:
            with open(self.path, encoding="utf-8") as handle:
                loaded = json.load(handle)
        except Exception as exc:
            log.warning("deferred queue unreadable (%s); starting fresh", exc)
            return {"version": 1, "deferred": {}}
        loaded.setdefault("version", 1)
        loaded.setdefault("deferred", {})
        return loaded

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, self.path)

    # ------------------------------------------------------------------
    def has(self, username):
        return (username or "").lower() in self.data["deferred"]

    def count(self):
        return len(self.data["deferred"])

    def add(self, candidate, reason, failures):
        """Park one candidate.  Returns its entry.  Idempotent per username:
        a second defer updates the count, reason and last time."""
        key = (candidate.username or "").lower()
        entry = self.data["deferred"].get(key, {})
        entry.update({
            "username": candidate.username,
            "github_url": candidate.profile_url,
            "keyword": candidate.keyword,
            "source_xlsx": candidate.source_file,
            "source_row": candidate.source_row_number,
            "failures": int(failures),
            "last_reason": str(reason)[:300],
            "first_failure": entry.get("first_failure") or _now(),
            "last_failure": _now(),
        })
        self.data["deferred"][key] = entry
        self.save()
        log.warning("deferred %s after %d failure(s): %s",
                    candidate.username, entry["failures"], reason)
        return entry

    def touch(self, username, reason, failures=None):
        """Record another failed retry of an already-deferred profile."""
        key = (username or "").lower()
        entry = self.data["deferred"].get(key)
        if not entry:
            return None
        entry["last_failure"] = _now()
        entry["last_reason"] = str(reason)[:300]
        entry["failures"] = int(entry.get("failures", 0) + 1
                                if failures is None else failures)
        self.save()
        return entry

    def remove(self, username):
        """Drop a profile once it finally succeeds."""
        key = (username or "").lower()
        if key in self.data["deferred"]:
            del self.data["deferred"][key]
            self.save()

    def all_entries(self):
        return sorted(self.data["deferred"].values(),
                      key=lambda entry: entry.get("first_failure", ""))


def candidate_from(entry):
    """Rebuild a minimal Candidate from a deferred entry, for a retry.

    Only identity fields are needed: the profile phase re-fetches everything
    from GitHub.  The xlsx_* enrichment fields are intentionally left blank.
    """
    return Candidate(
        source_row_number=int(entry.get("source_row", 0) or 0),
        source_file=entry.get("source_xlsx", ""),
        keyword=entry.get("keyword", ""),
        username=entry.get("username", ""),
        profile_url=entry.get("github_url", ""))


# ==========================================================================
# DEVICE HEALTH
# ==========================================================================
# Device health gates: battery and network reachability.
#
# Two conditions must hold for a long unattended run to be SAFE to continue:
#
#     battery    a laptop that dies mid-profile loses nothing already saved,
#                but it loses the browser session and stops the device without
#                warning.  Below BATTERY_PAUSE_PERCENT on battery power the run
#                PAUSES and waits for a charger; below BATTERY_CRITICAL_PERCENT
#                it STOPS cleanly while there is still time to write state.
#
#     network    no connectivity means every search would fail and be recorded
#                as a transient failure, which could push good profiles into
#                the deferred queue for no reason.  The run pauses while the
#                network is down and stops cleanly if it stays down.
#
# Both are read WITHOUT new dependencies:
#
#     battery    /sys/class/power_supply/  (Linux sysfs).  new_mylib.py's
#                check_battery() uses psutil, which is NOT installed on this
#                machine, so sysfs is used instead and psutil is only tried as
#                a fallback.
#     network    a TCP connect to github.com:443.  new_mylib.py's isinternet()
#                tests connectivity by calling gTTS text-to-speech and PLAYING
#                AUDIO, which is unusable for an unattended 80-rank run.
#
# Every reader is injectable so the tests never touch real hardware, the real
# clock or the real network.
# --------------------------------------------------------------------------
# ------------------------------------------------------------------ battery --
def _read_first(pattern, name):
    """Contents of <dir>/<name> for the first matching power-supply dir."""
    for directory in sorted(glob.glob(pattern)):
        path = os.path.join(directory, name)
        try:
            with open(path, encoding="utf-8") as handle:
                return handle.read().strip()
        except OSError:
            continue
    return ""


def battery_state(root="/sys/class/power_supply"):
    """(percent, plugged) for this device.

    percent is None when the device has no battery at all (a desktop or a
    server), which is NOT a reason to pause -- such a machine simply has no
    battery risk.  plugged is True when mains power is present.
    """
    percent_text = _read_first(os.path.join(root, "BAT*"), "capacity")
    if not percent_text:
        return _battery_state_psutil()
    try:
        percent = int(percent_text)
    except ValueError:
        return None, None
    online = _read_first(os.path.join(root, "A[CD]*"), "online")
    status = _read_first(os.path.join(root, "BAT*"), "status").lower()
    plugged = online == "1" or status in ("charging", "full")
    return percent, plugged


def _battery_state_psutil():
    """Fallback for a machine whose battery is not exposed through sysfs."""
    try:
        import psutil
    except Exception:
        return None, None
    try:
        battery = psutil.sensors_battery()
    except Exception:
        return None, None
    if battery is None:
        return None, None
    return int(battery.percent), bool(battery.power_plugged)


def battery_verdict(percent, plugged):
    """"ok" | "pause" | "stop" for a battery reading.

    Mains power is always ok: a plugged-in laptop at 5% is charging, not
    dying.  Hysteresis is deliberate -- once paused, the battery must reach
    BATTERY_RESUME_PERCENT rather than merely creep back to the pause
    threshold, so the run cannot flap on and off around one number.
    """
    if percent is None:
        return "ok"                      # no battery: nothing to protect
    if plugged:
        return "ok"
    if percent <= C.BATTERY_CRITICAL_PERCENT:
        return "stop"
    if percent <= C.BATTERY_PAUSE_PERCENT:
        return "pause"
    return "ok"




def battery_resumed(percent, plugged):
    """True when a PAUSED run may continue (charger back, or recharged)."""
    if percent is None or plugged:
        return True
    return percent >= C.BATTERY_RESUME_PERCENT


# ------------------------------------------------------------------ network --
def network_ok(host=None, port=None, timeout=None, connect=None):
    """True when GitHub is reachable over TCP.

    A plain TCP connect, not an HTTP request: it answers "is the network up"
    without sending anything to GitHub that resembles a scrape, and it costs
    no rate-limit budget.
    """
    host = host or C.HEALTH_HOST
    port = port or C.HEALTH_PORT
    timeout = timeout or C.HEALTH_TIMEOUT_SECONDS
    opener = connect or socket.create_connection
    try:
        connection = opener((host, port), timeout)
    except Exception:
        return False
    try:
        connection.close()
    except Exception:
        pass
    return True


# -------------------------------------------------------------------- report --
def describe_health(percent, plugged, online):
    """One health block.  online=None means the network was not checked --
    the mid-search gate skips it, and reporting "up" on a reading nobody
    took would be a lie in the run log."""
    lines = ["[HEALTH]"]
    if percent is None:
        lines.append("Battery    : none (desktop/server)")
    else:
        lines.append("Battery    : %d%% (%s)"
                     % (percent, "charging" if plugged else "on battery"))
    if online is None:
        lines.append("Network    : not checked")
    else:
        lines.append("Network    : %s" % ("up" if online else "DOWN"))
    return "\n".join(lines)


# ==========================================================================
# NORMALIZER
# ==========================================================================
# Canonical forms: degrees, specialisations, locations, skills, roles.
#
# Alias tables do the work that regex cannot (spec 42/43/46/53).  Everything
# here is pure and side-effect free so it can be unit tested without a browser.
# --------------------------------------------------------------------------
# ------------------------------------------------------------- cleaning ----
_PUNCT_MAP = {
    "‘": "'", "’": "'", "‚": "'", "‛": "'",
    "“": '"', "”": '"', "„": '"',
    "–": "-", "—": "-", "―": "-", "−": "-",
    " ": " ", " ": " ", " ": " ", "​": "",
    "﻿": "", "•": "*", "·": "*",
}

_NAV_NOISE = re.compile(
    r"^(skip to content|sign in|sign up|toggle navigation|menu|home|"
    r"cookie|accept all|manage cookies|loading|footer|search or jump to|"
    r"we use cookies|privacy policy|terms of service|all rights reserved)\b",
    re.I)


def clean_text(text):
    """Normalise Unicode punctuation and whitespace; keep line structure."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    for source, target in _PUNCT_MAP.items():
        text = text.replace(source, target)
    lines = []
    for line in text.splitlines():
        line = re.sub(r"[ \t]+", " ", line).strip()
        if not line or _NAV_NOISE.match(line):
            continue
        lines.append(line)
    return "\n".join(lines)


def flatten(text):
    return re.sub(r"\s+", " ", clean_text(text)).strip()


# ---------------------------------------------------------- delegation ----
# Degrees, specialisations, locations, skills, roles and industries now live
# in dedicated extractors.  These names stay available here so existing
# callers keep working, but they forward rather than holding a second copy
# of the tables -- two copies drift, and a fix applied to one of them is a
# fix that silently did not happen.
#
# The imports are lazy (PEP 562 module __getattr__) because those modules
# import THIS one for clean_text/flatten.















def find_skills(text, vocabulary):
    """Ordered vocabulary hits, with the single-letter guard applied."""
    if not text:
        return []
    found = []
    for skill in vocabulary:
        pattern = _PATTERNS.get(skill)
        if pattern is None:
            pattern = _pattern(skill)
        if skill in SINGLE_LETTER:
            if not _single_letter_allowed(skill, text,
                                                           False):
                continue
        if not pattern.search(text):
            continue
        if _is_negated(pattern, text):
            continue
        found.append(skill)
    return found


def join_list(values, limit=None):
    seen, output = set(), []
    for value in values:
        key = str(value).strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        output.append(str(value).strip())
        if limit and len(output) >= limit:
            break
    return ", ".join(output)


# ==========================================================================
# SECTIONS
# ==========================================================================
# Split free text into labelled sections.
#
# A resume, a profile README and a portfolio page all say "here comes my
# education" in their own way, but they agree on one thing: a short line that
# names a topic, followed by the content for that topic.  Recognising those
# headings is what lets an extractor say "this degree came from the Education
# section" instead of "this degree appeared somewhere on a 4000-character
# page", which is the difference between a fact and a coincidence.
#
# Markdown is not required: a heading may be "## Education", "EDUCATION",
# "Education:" or "Technical Skills".
# --------------------------------------------------------------------------
# Canonical section -> the words that introduce it.
SECTION_ALIASES = {
    "education": [
        "education", "educational background", "academics", "academic",
        "academic background", "qualification", "qualifications",
        "educational qualification", "schooling", "studies",
    ],
    "experience": [
        "experience", "work experience", "professional experience",
        "employment", "employment history", "work history", "career",
        "professional background", "internship", "internships",
        "work", "positions",
    ],
    "skills": [
        "skills", "technical skills", "core skills", "key skills",
        "tech stack", "technology stack", "technologies", "tools",
        "tools and technologies", "expertise", "competencies",
        "technical expertise", "programming languages", "languages and tools",
        "technical arsenal", "what i work with", "tech i use",
        "languages & libraries", "languages and libraries",
    ],
    "projects": [
        "projects", "personal projects", "selected projects", "portfolio",
        "notable projects", "featured projects", "work samples",
    ],
    "contact": [
        "contact", "contact me", "contact info", "contact information",
        "get in touch", "reach me", "reach out", "connect", "let's connect",
        "find me", "socials", "social links",
    ],
    "about": [
        "about", "about me", "summary", "profile", "professional summary",
        "objective", "career objective", "introduction", "intro", "bio",
        "who am i", "overview", "my mission", "my philosophy",
    ],
    "certifications": [
        "certifications", "certificates", "certification", "licenses",
        "courses", "credentials",
    ],
    "achievements": [
        "achievements", "awards", "honors", "honours", "accomplishments",
        "recognition",
    ],
}

_ALIAS_TO_SECTION = {}
for _section, _aliases in SECTION_ALIASES.items():
    for _alias in _aliases:
        _ALIAS_TO_SECTION[_alias] = _section

# A heading is short, mostly letters, and optionally decorated.
_HEADING_CLEAN = re.compile(r"^[\s#*_=~>\-•●▪️◦·|]+|[\s#*_=~:\-•]+$")
# Markdown heading syntax, whatever the heading happens to say.
_MARKDOWN_HEADING = re.compile(r"^\s{0,3}#{1,6}\s")
# A short line that ends in a colon and carries no content after it is a
# heading, whatever it says: "GitHub Stats:", "Top Contributed Repo:".
_BARE_HEADING = re.compile(r"^[A-Za-z][A-Za-z0-9 &/'()-]{1,48}:\s*$")
_EMOJI = re.compile(
    "[\U0001F300-\U0001FAFF\U00002600-\U000027BF\U0001F1E6-\U0001F1FF️]")


def _heading_of(line):
    """Canonical section name if `line` is a heading, else ''."""
    text = _EMOJI.sub(" ", line)
    text = _HEADING_CLEAN.sub("", text).strip()
    if not text or len(text) > 60:
        return ""
    # A heading rarely ends in a sentence.
    if text.endswith((".", "!", "?")):
        return ""
    key = re.sub(r"[^a-z& ]+", " ", text.lower())
    key = re.sub(r"\s+", " ", key).strip()
    if not key:
        return ""
    if key in _ALIAS_TO_SECTION:
        return _ALIAS_TO_SECTION[key]
    # "Technical Skills & Tools" -> skills
    for alias, section in _ALIAS_TO_SECTION.items():
        if len(alias) < 5:
            continue
        if key.startswith(alias + " ") or key == alias:
            return section
    return ""


def parse_sections(text):
    """{section: text} for every section found, plus '_head' for the preamble.

    Unrecognised headings end the previous section rather than being folded
    into it, so a skills list never leaks into the education block.
    """
    sections = {}
    current = "_head"
    buffer = []

    def flush():
        body = "\n".join(buffer).strip()
        if body:
            sections[current] = (sections.get(current, "") + "\n" + body
                                 ).strip()

    for raw in clean_text(text).splitlines():
        # "Skills: Python, SQL" is checked FIRST.  Treating it as a plain
        # heading would match the "skills ..." prefix rule and throw the
        # content away.
        inline = re.match(r"^\s*([A-Za-z][A-Za-z &/]{2,40})\s*:\s*(.+)$", raw)
        if inline:
            named = _heading_of(inline.group(1))
            if named:
                flush()
                buffer = [inline.group(2).strip()]
                current = named
                continue
        heading = _heading_of(raw)
        if heading:
            flush()
            buffer = []
            current = heading
            continue
        # An UNRECOGNISED heading still ends the section it follows.
        # Without this, "GitHub Stats:" stays inside the Tech Stack block
        # and its text is read as part of the skills list -- which is how
        # "GitHub" became a claimed tool.  Both forms matter: the Markdown
        # source writes "# GitHub Stats:", the rendered page writes the
        # same line with no "#" at all.
        if _MARKDOWN_HEADING.match(raw) or _BARE_HEADING.match(
                _EMOJI.sub(" ", raw).strip()):
            flush()
            buffer = []
            current = "_other"
            continue
        buffer.append(raw)
    flush()
    return sections


def section_text(sections, *names):
    """Concatenated text of the named sections, in the order given."""
    return "\n".join(sections.get(name, "") for name in names
                     if sections.get(name)).strip()


def has_section(sections, name):
    return bool(sections.get(name, "").strip())


# ==========================================================================
# EDUCATION EXTRACTOR
# ==========================================================================
# Education fields, decided by context rather than by one big pattern.
#
# The failure this module exists to prevent, taken from a real profile whose
# bio reads "an ordinary data science college student at a university.":
#
#     College Name = "an ordinary data science college student at a university"
#     University   = "stand their strengths and limitations. Pragmatic ..."
#
# Both are regex hits on words that merely *sound* academic.  An institution
# is only accepted when a real NAME is present -- a proper noun or a known
# acronym next to an institution word.  "a university" is not a name.
# --------------------------------------------------------------------------
# --------------------------------------------------------------- degrees --
DEGREE_ALIASES = [
    ("B.Tech", ["bachelor of technology", "b tech", "b.tech", "btech",
                "b-tech", "bachelors of technology", "b.tech."]),
    ("M.Tech", ["master of technology", "m tech", "m.tech", "mtech",
                "m-tech", "masters of technology", "m.tech."]),
    ("B.E.", ["bachelor of engineering", "b.e.", "b e", "b.e"]),
    ("M.E.", ["master of engineering", "m.e.", "m e", "m.e"]),
    ("BCA", ["bachelor of computer applications",
             "bachelor of computer application", "bca"]),
    ("MCA", ["master of computer applications",
             "master of computer application", "mca"]),
    ("B.Sc", ["bachelor of science", "b.sc", "bsc", "b sc", "b.s."]),
    ("M.Sc", ["master of science", "m.sc", "msc", "m sc", "m.s."]),
    ("MBA", ["master of business administration", "mba"]),
    ("BBA", ["bachelor of business administration", "bba"]),
    ("B.Com", ["bachelor of commerce", "b.com", "bcom"]),
    ("M.Com", ["master of commerce", "m.com", "mcom"]),
    ("PhD", ["doctor of philosophy", "ph.d", "phd", "doctorate"]),
    ("Diploma", ["polytechnic diploma", "polytechnic", "diploma"]),
]

UG_DEGREES = {"B.Tech", "B.E.", "BCA", "B.Sc", "BBA", "B.Com"}
PG_DEGREES = {"M.Tech", "M.E.", "MCA", "M.Sc", "MBA", "M.Com"}

EDUCATION_LEVEL = {
    "B.Tech": "Undergraduate", "B.E.": "Undergraduate",
    "BCA": "Undergraduate", "B.Sc": "Undergraduate",
    "BBA": "Undergraduate", "B.Com": "Undergraduate",
    "Diploma": "Diploma",
    "M.Tech": "Postgraduate", "M.E.": "Postgraduate",
    "MCA": "Postgraduate", "M.Sc": "Postgraduate",
    "MBA": "Postgraduate", "M.Com": "Postgraduate",
    "PhD": "Doctorate",
}

# Bare two-letter degrees are ordinary English words in lower case.
_BARE_BE = re.compile(r"(?<![A-Za-z.])BE(?![A-Za-z])")
_BARE_ME = re.compile(r"(?<![A-Za-z.])ME(?![A-Za-z])")


def normalize_degree(value):
    text = re.sub(r"\s+", " ", value or "")
    low = " %s " % text.lower()
    for canonical, aliases in DEGREE_ALIASES:
        for alias in aliases:
            if re.search(r"(?<![a-z])%s(?![a-z])" % re.escape(alias), low):
                return canonical
    if _BARE_BE.search(text):
        return "B.E."
    if _BARE_ME.search(text):
        return "M.E."
    return ""


# -------------------------------------------------------- specialisations --
SPECIALIZATION_ALIASES = [
    ("Artificial Intelligence & Machine Learning",
     ["artificial intelligence and machine learning",
      "artificial intelligence & machine learning", "ai and ml", "ai & ml",
      "ai/ml", "aiml", "ai-ml", "ai ml"]),
    ("Computer Science Engineering",
     ["computer science and engineering", "computer science & engineering",
      "computer science engineering", "computer science", "cse", "comp sci"]),
    ("Information Technology", ["information technology", "info tech", "it"]),
    ("Artificial Intelligence", ["artificial intelligence", "ai"]),
    ("Data Science", ["data science", "data analytics"]),
    ("Software Engineering", ["software engineering"]),
    ("Cyber Security", ["cyber security", "cybersecurity",
                        "information security", "infosec"]),
    ("Electronics and Communication Engineering",
     ["electronics and communication engineering",
      "electronics & communication", "electronics and communication", "ece"]),
    ("Electrical and Electronics Engineering",
     ["electrical and electronics engineering", "eee"]),
    ("Electrical Engineering", ["electrical engineering"]),
    ("Mechanical Engineering", ["mechanical engineering",
                                "mech engineering"]),
    ("Civil Engineering", ["civil engineering"]),
    ("Chemical Engineering", ["chemical engineering"]),
    ("Biotechnology", ["biotechnology", "bio technology"]),
    ("Cloud Computing", ["cloud computing"]),
    ("Internet of Things", ["internet of things", "iot"]),
]

# Short forms that are ordinary English words or unrelated jargon, so they
# need a branch cue nearby.  ECE/EEE/IoT are unambiguous engineering
# abbreviations and do not.
AMBIGUOUS = {"it", "ai", "cs"}

# Only these establish that a branch is being NAMED, not merely mentioned.
_BRANCH_CUE = re.compile(
    r"\b(b\.?tech|b\.?e\.?|bca|mca|m\.?tech|b\.?sc|m\.?sc|mba|bba|diploma|"
    r"branch|stream|department|dept|major|majoring|honou?rs|"
    r"specializ\w+|specialis\w+|pursuing|studying|degree in|"
    r"bachelor|master|undergrad\w*|postgrad\w*)\b", re.I)


def normalize_specialization(value, require_context=True):
    """Canonical branch, or "".

    "AI enthusiast" must not become a degree specialisation, so a short
    ambiguous alias needs a degree/branch word within 40 characters.
    """
    text = re.sub(r"\s+", " ", (value or "").lower())
    low = " %s " % text
    for canonical, aliases in SPECIALIZATION_ALIASES:
        for alias in aliases:
            pattern = re.compile(r"(?<![a-z-])%s(?![a-z-])"
                                 % re.escape(alias))
            for match in pattern.finditer(low):
                if alias in AMBIGUOUS and require_context:
                    window = low[max(0, match.start() - 40):match.end() + 40]
                    if not _BRANCH_CUE.search(window):
                        continue
                return canonical
    return ""


# ------------------------------------------------------------ institutions --
# A word that makes a phrase an institution...
INSTITUTION_WORDS = [
    "college", "university", "universiti", "universitas", "institute",
    "institut", "institue", "polytechnic", "academy", "vidyapeeth",
    "vishwavidyalaya", "school of engineering", "school of technology",
    "school of business", "technical university", "institute of technology",
]
# ...and acronyms that ARE institutions on their own.
INSTITUTION_ACRONYMS = [
    "IIT", "NIT", "IIIT", "IIM", "BITS", "VIT", "SRM", "LNCT", "LNCTS",
    "MIT", "LPU", "KIIT", "MANIT", "DTU", "NSUT", "IGDTUW", "JNTU", "AKTU",
    "RGPV", "VTU", "SPPU", "MAKAUT", "COEP", "PICT", "PES", "RVCE", "BMS",
    "MSRIT", "SSN", "VJTI", "NITK", "NITT", "NITW", "IISC", "IISER", "NSIT",
]
_ACRONYM_RE = re.compile(
    r"\b(%s)\b" % "|".join(sorted(INSTITUTION_ACRONYMS, key=len,
                                  reverse=True)))


# Online course providers use "Academy"/"School"/"Institute" in their
# brand names.  A certification badge from one is not where somebody
# studies, so they never become a College Name.
LEARNING_PLATFORMS = [
    "dicoding", "coursera", "udemy", "udacity", "edx", "bisa ai",
    "great learning", "simplilearn", "upgrad", "scaler", "newton school",
    "masai", "almabetter", "codecademy", "datacamp", "pluralsight",
    "linkedin learning", "freecodecamp", "w3schools", "geeksforgeeks",
    "hackerrank", "leetcode", "kaggle", "google developer", "gdsc",
    "microsoft learn", "aws academy", "cisco networking academy", "nptel",
    "swayam", "internshala", "board infinity", "skill india",
]
_PLATFORM_RE = re.compile(
    r"\b(%s)\b" % "|".join(re.escape(p) for p in LEARNING_PLATFORMS), re.I)

# A line only names where somebody STUDIES when it says so.
_ENROLLMENT_CUE = re.compile(
    r"\b(student\s+at|student\s+of|studying\s+at|studying\s+in|"
    r"study\s+at|pursuing|enrolled\s+(at|in)|graduated\s+from|"
    r"alumni\s+of|alumnus\s+of|undergraduate\s+at|graduate\s+at|"
    r"education|degree\s+(at|from)|attending|attended|"
    r"b\.?tech\s+at|be\s+at|mca\s+at|bca\s+at|m\.?tech\s+at)\b", re.I)

# Phrases where an institution word is NOT an institution name.
_INSTITUTION_NOISE = re.compile(
    r"\b(hackathon|contest|competition|fest|festival|event|meetup|"
    r"workshop|conference|summit|club|chapter|community|society|"
    r"ambassador|program|programme|bootcamp|course|certification|"
    r"ranking|rankings|admission|admissions|life|days|project|projects)\b",
    re.I)

# An article before the institution word means it is generic:
# "a university", "the college", "any institute".
_GENERIC_INSTITUTION = re.compile(
    r"^(a|an|the|my|our|some|any|this|that|another)\s+"
    r"(top\s+|local\s+|private\s+|public\s+|good\s+|reputed\s+)?"
    r"(college|university|institute|school|academy|polytechnic)\b",
    re.I)

_STOPWORD = {
    "a", "an", "the", "of", "at", "in", "and", "for", "to", "from", "my",
    "our", "is", "am", "was", "with", "on", "by", "as", "ordinary",
}


def _looks_like_name(phrase, allow_lowercase=False):
    """A real institution name carries proper nouns or a known acronym.

    `allow_lowercase` is set when an explicit enrolment cue already
    established that this IS where somebody studies.  Plenty of profiles
    are typed entirely in lower case -- "Student at malla reddy
    engineering college" names a real institution, and rejecting it for
    its capitalisation loses a fact the sentence states outright.
    """
    phrase = phrase.strip()
    if not phrase or len(phrase) > 90:
        return False
    if _GENERIC_INSTITUTION.match(phrase):
        return False
    if _ACRONYM_RE.search(phrase):
        return True
    words = [word for word in re.split(r"\s+", phrase) if word]
    content = [word for word in words
               if word.lower() not in _STOPWORD
               and word.lower() not in INSTITUTION_WORDS]
    if allow_lowercase:
        # Still needs a name of its own, not just the institution word:
        # "a university" and "the college" have no content words.
        return len(content) >= 1
    proper = [word for word in content if word[:1].isupper()]
    # "Lakshmi Narain College of Technology" -> 4 proper nouns.
    # "an ordinary data science college student" -> none.
    return len(proper) >= 1


def _institution_from_line(line, require_enrollment=True):
    """The institution name inside one line, or ""."""
    text = flatten(line).strip(" -*|•")
    if not text or len(text) > 160:
        return ""
    low = text.lower()
    if not any(word in low for word in INSTITUTION_WORDS) and \
            not _ACRONYM_RE.search(text):
        return ""
    if _INSTITUTION_NOISE.search(text):
        return ""
    if _PLATFORM_RE.search(text):
        return ""
    # Outside an Education section, a line must actually say that this is
    # where the person studies.  Without this, any "Academy" in a
    # certification badge becomes their college.
    if require_enrollment and not _ENROLLMENT_CUE.search(text):
        return ""

    # An explicit cue ("student at X") means the words after it name the
    # institution, whatever their capitalisation.
    cue = _ENROLLMENT_CUE.search(text)

    # Prefer the phrase that actually contains the institution word.
    for word in sorted(INSTITUTION_WORDS, key=len, reverse=True):
        position = low.find(word)
        if position == -1:
            continue
        # "college student" is a description of a person, not the name of
        # an institution.
        after = low[position + len(word):].lstrip()
        if re.match(r"student|students|life|days|years?\b", after):
            continue
        # Lower-case names are only trusted when the cue comes FIRST:
        # "student at malla reddy engineering college" names one,
        # "data science college student at a university" does not.
        lowercase_ok = bool(cue and cue.end() <= position)
        # Take the name words immediately before the institution word...
        before = text[:position].rstrip()
        if cue and cue.end() <= position:
            # Everything between "student at" and "college" is the name.
            before = text[cue.end():position].rstrip()
        tokens = re.split(r"[,;:|]|\s+-\s+", before)
        head = tokens[-1].strip() if tokens else ""
        head_words = [w for w in head.split() if w]
        kept = []
        for token in reversed(head_words):
            if token.lower() in _STOPWORD and kept:
                break
            if token[:1].isupper() or lowercase_ok \
                    or token.lower() in ("of", "and"):
                kept.insert(0, token)
            else:
                break
        # ...and whatever qualifies it afterwards ("of Technology").
        after = text[position + len(word):]
        tail = re.match(r"^(\s+of\s+[A-Z][\w.&-]*(\s+[A-Z][\w.&-]*){0,3})",
                        after)
        name = " ".join(kept + [text[position:position + len(word)]])
        if tail:
            name += tail.group(1)
        name = flatten(name).strip(" .,-")
        if _looks_like_name(name, allow_lowercase=lowercase_ok) \
                and len(name) >= 4:
            return _titlecase_institution(name)

    acronym = _ACRONYM_RE.search(text)
    if acronym:
        window = text[acronym.start():acronym.start() + 70]
        window = re.split(r"[,;|]|\s+-\s+", window)[0].strip(" .,")
        if 2 <= len(window) <= 80:
            return window
    return ""


def _titlecase_institution(name):
    out = []
    for word in name.split():
        if word.isupper() and len(word) <= 6:
            out.append(word)
        elif word.lower() in ("of", "and", "the"):
            out.append(word.lower())
        else:
            out.append(word[:1].upper() + word[1:])
    return " ".join(out)


def extract_institution(text, want_university=False,
                        require_enrollment=True):
    """College / university name, or "".

    `want_university` restricts the answer to lines that really name a
    university or an affiliation.
    """
    body = clean_text(text)
    if not body:
        return ""
    lines = body.splitlines()

    # An explicit label beats everything.
    for line in lines:
        label = re.match(
            r"^\s*(college|university|institute|school|affiliated\s+to|"
            r"affiliation)\s*[:\-]\s*(.+)$", line, re.I)
        if label:
            value = flatten(label.group(2)).strip(" .,-")
            if _looks_like_name(value):
                return _titlecase_institution(value)

    for line in lines:
        low = line.lower()
        if want_university and not re.search(
                r"\b(university|universitas|universiti|affiliated|"
                r"affiliation|vishwavidyalaya)\b", low):
            continue
        if want_university and re.search(r"\bcollege\b", low) and \
                not re.search(r"\buniversity\b", low):
            continue
        name = _institution_from_line(
            line, require_enrollment=require_enrollment)
        if name:
            if want_university and "university" not in name.lower() and \
                    not _ACRONYM_RE.search(name):
                continue
            return name
    return ""


# ------------------------------------------------------- year of study ----
_ORDINAL = {
    "1st": "1st", "first": "1st", "2nd": "2nd", "second": "2nd",
    "3rd": "3rd", "third": "3rd", "4th": "4th", "fourth": "4th",
    "5th": "5th", "fifth": "5th", "6th": "6th", "sixth": "6th",
}
_YEAR_WORD = r"(1st|2nd|3rd|4th|5th|6th|first|second|third|fourth|fifth|sixth)"

# "3rd year student" and also "3rd Data Science honours student", where the
# noun "year" is implied.
YEAR_OF_STUDY_RE = re.compile(
    r"\b%s\s+year\b" % _YEAR_WORD, re.I)
# "3rd Data Science honours student" -- the noun "year" is implied.  The
# negative lookahead keeps "5th semester" out: that is a semester, not a
# year of study.
YEAR_IMPLIED_RE = re.compile(
    r"\b%s\s+(?!semester\b)(?:[A-Za-z&/ ]{0,40}?)\b"
    r"(student|undergrad\w*)\b" % _YEAR_WORD, re.I)
FINAL_YEAR_RE = re.compile(r"\b(pre[-\s]?final|final)\s+year\b", re.I)
SEMESTER_RE = re.compile(
    r"\b(?:semester|sem)\s*[-:]?\s*([1-9]|1[0-2])\b"
    r"|\b([1-9]|1[0-2])(?:st|nd|rd|th)?\s+semester\b", re.I)


def extract_year_of_study(text):
    """'3rd Year', 'Final Year', 'Semester 5' or ''."""
    body = flatten(text)
    parts = []
    final = FINAL_YEAR_RE.search(body)
    if final:
        word = final.group(1).lower().replace(" ", "-")
        parts.append("Pre-Final Year" if word.startswith("pre")
                     else "Final Year")
    else:
        match = YEAR_OF_STUDY_RE.search(body) or YEAR_IMPLIED_RE.search(body)
        if match:
            key = match.group(1).lower()
            parts.append("%s Year" % _ORDINAL.get(key, key))
    sem = SEMESTER_RE.search(body)
    if sem:
        parts.append("Semester %s" % (sem.group(1) or sem.group(2)))
    return ", ".join(parts)


# ------------------------------------------------------ graduation year ----
YEAR_RE = re.compile(r"\b(20(?:1[8-9]|2[0-9]|3[0-5]))\b")
SHORT_YEAR_RE = re.compile(r"(?<!\d)'(2[0-9])(?!\d)")
YEAR_RANGE_RE = re.compile(r"\b(20\d{2})\s*(?:-|--|to)\s*(20\d{2})\b", re.I)

_STRONG_GRAD = re.compile(
    r"\b(class\s+of|batch|graduat\w*|passing\s+out|pass\s*out|expected)\b",
    re.I)
_EDU_NEAR = re.compile(
    r"\b(b\.?tech|b\.?e\.?|bca|mca|m\.?tech|b\.?sc|m\.?sc|mba|bba|diploma|"
    r"degree|college|university|institute|student|semester|undergrad\w*|"
    r"postgrad\w*|honou?rs)\b", re.I)
_COPYRIGHT = re.compile(r"(copyright|©|\(c\)|all rights reserved|"
                        r"version|release|build)", re.I)


def extract_graduation_year(text):
    body = flatten(text)
    if not body:
        return ""

    for match in YEAR_RANGE_RE.finditer(body):
        window = body[max(0, match.start() - 70):match.end() + 70]
        if _COPYRIGHT.search(window):
            continue
        if _STRONG_GRAD.search(window) or _EDU_NEAR.search(window):
            end = int(match.group(2))
            if int(match.group(1)) < end <= 2035:
                return str(end)

    best = ""
    for match in YEAR_RE.finditer(body):
        window = body[max(0, match.start() - 45):match.end() + 45]
        if _COPYRIGHT.search(window):
            continue
        if _STRONG_GRAD.search(window):
            return match.group(1)
        if _EDU_NEAR.search(window) and not best:
            best = match.group(1)

    for match in SHORT_YEAR_RE.finditer(body):
        window = body[max(0, match.start() - 60):match.end() + 60]
        before = body[max(0, match.start() - 12):match.start()]
        if _STRONG_GRAD.search(window) or _EDU_NEAR.search(window) or \
                re.search(r"\b[A-Z]{2,8}\s?$", before):
            return "20%s" % match.group(1)
    return best


# ------------------------------------------------------------ the bundle ---
COLLEGE_STUDENT_RE = re.compile(
    r"\b(college|university)\s+student\b|\b(undergrad\w*|bachelor\w*)\b",
    re.I)
SCHOOL_RE = re.compile(r"\b(high\s+school|secondary\s+school|"
                       r"school\s+student|class\s+(9|10|11|12)\b)", re.I)


def education_level(degree, text):
    """Normalised level; never over-inferred."""
    if degree and degree in EDUCATION_LEVEL:
        return EDUCATION_LEVEL[degree]
    body = flatten(text)
    if SCHOOL_RE.search(body):
        return "School Student"
    if re.search(r"\b(phd|doctoral|doctorate)\b", body, re.I):
        return "Doctorate"
    if re.search(r"\b(postgrad\w*|master'?s?\s+(degree|student))\b", body,
                 re.I):
        return "Postgraduate"
    if COLLEGE_STUDENT_RE.search(body):
        return "Undergraduate"
    # Enrolled somewhere, but no degree stated: "Student at malla reddy
    # engineering college".  Saying "College Student" is honest; calling
    # it Undergraduate would assert a level the profile never gives.
    if re.search(r"\b(student|studying|pursuing|enrolled)\b", body, re.I) \
            and re.search(r"\b(college|university|institute|"
                          r"polytechnic)\b", body, re.I):
        return "College Student"
    return ""


def extract_education(sources):
    """Run every education field over sources in priority order.

    `sources` is an ordered list of (text, confidence); the first source
    that yields a field wins, which is what makes a resume outrank a
    README and a README outrank a search snippet.
    """
    result = {
        "degree": "", "specialization": "", "level": "", "year": "",
        "graduation": "", "college": "", "university": "",
        "graduation_sources": [],
    }
    for text, _score in sources:
        if not text:
            continue
        parsed = parse_sections(text)
        edu = section_text(parsed, "education") or text

        result["degree"] = result["degree"] or normalize_degree(edu)
        result["specialization"] = (result["specialization"]
                                    or normalize_specialization(edu))
        result["year"] = result["year"] or extract_year_of_study(edu)
        # Inside a genuine Education section the heading IS the cue, so a
        # bare "Lakshmi Narain College of Technology" line counts there.
        in_section = bool(has_section(parsed, "education"))
        result["college"] = result["college"] or extract_institution(
            edu, require_enrollment=not in_section)
        result["university"] = result["university"] or extract_institution(
            edu, want_university=True, require_enrollment=not in_section)
        year = extract_graduation_year(edu)
        if year:
            result["graduation"] = result["graduation"] or year
            result["graduation_sources"].append(year)

    all_text = "\n".join(text for text, _score in sources if text)
    result["level"] = education_level(result["degree"], all_text)
    # A course with no branch: "Data Science honours student".
    if not result["specialization"]:
        result["specialization"] = normalize_specialization(
            all_text, require_context=True)
    return result


# ==========================================================================
# EMPLOYMENT EXTRACTOR
# ==========================================================================
# Company, role, employment status, experience, internships, aspirations.
#
# Two mistakes this module is built to avoid:
#
#   * "Student at LNCT" becoming Current Job Role = Student,
#     Current Company = LNCT.  Studying somewhere is not employment.
#   * an internship duration becoming Total Experience.
#
# Both are context problems, not pattern problems.
# --------------------------------------------------------------------------
# ------------------------------------------------------------ vocabulary --
ROLE_ALIASES = [
    ("Full Stack Developer", ["full stack developer", "fullstack developer",
                              "full-stack developer", "full stack engineer",
                              "mern stack developer",
                              "mean stack developer"]),
    ("Frontend Developer", ["frontend developer", "front end developer",
                            "front-end developer", "frontend engineer",
                            "ui developer"]),
    ("Backend Developer", ["backend developer", "back end developer",
                           "back-end developer", "backend engineer"]),
    ("DevOps Engineer", ["devops engineer", "dev ops engineer", "devops"]),
    ("Site Reliability Engineer", ["site reliability engineer", "sre"]),
    ("Cloud Engineer", ["cloud engineer", "cloud architect",
                        "cloud developer"]),
    ("Data Scientist", ["data scientist"]),
    ("Data Analyst", ["data analyst", "business analyst"]),
    ("Data Engineer", ["data engineer"]),
    ("ML Engineer", ["machine learning engineer", "ml engineer",
                     "ai engineer", "mlops engineer"]),
    ("Android Developer", ["android developer"]),
    ("iOS Developer", ["ios developer"]),
    ("Mobile App Developer", ["mobile app developer", "flutter developer",
                              "react native developer"]),
    ("Cybersecurity Engineer", ["cyber security engineer",
                                "cybersecurity engineer",
                                "security engineer", "penetration tester",
                                "ethical hacker", "soc analyst"]),
    ("Network Engineer", ["network engineer", "network administrator",
                          "noc engineer"]),
    ("System Administrator", ["system administrator",
                              "systems administrator", "sysadmin"]),
    ("IT Support Engineer", ["it support engineer", "it support specialist",
                             "desktop support engineer", "help desk",
                             "service desk", "technical support engineer",
                             "it support"]),
    ("QA Engineer", ["qa engineer", "test engineer", "sdet",
                     "quality assurance engineer", "automation tester"]),
    ("Web Developer", ["web developer", "website developer"]),
    ("Software Engineer", ["software engineer", "software developer",
                           "software development engineer", "sde",
                           "application developer", "programmer"]),
]

INDUSTRY_ALIASES = [
    ("Cybersecurity", ["cyber security", "cybersecurity", "infosec"]),
    ("Cloud Computing", ["cloud computing", "cloud infrastructure"]),
    ("AI/ML", ["artificial intelligence", "machine learning",
               "deep learning", "ai/ml", "generative ai"]),
    ("Data Science", ["data science", "data analytics"]),
    ("FinTech", ["fintech", "financial technology"]),
    ("EdTech", ["edtech", "education technology"]),
    ("HealthTech", ["healthtech", "health technology"]),
    ("Networking", ["networking", "network infrastructure"]),
    ("IT Services", ["it services", "information technology services"]),
    ("Software Development", ["software development",
                              "product engineering"]),
]

ROLE_WORD = re.compile(
    r"\b(engineer|developer|analyst|administrator|admin|consultant|manager|"
    r"architect|scientist|designer|specialist|technician|lead|intern|"
    r"trainee|associate|executive|officer|programmer|tester|sre|devops|"
    r"sde|sdet|qa|support|researcher)\b", re.I)

# Studying is not employment: these block company extraction on a line.
EDUCATION_BLOCK = re.compile(
    r"\b(student|studying|pursuing|undergrad\w*|postgrad\w*|freshman|"
    r"sophomore|junior\s+year|senior\s+year|alumni|alumnus|graduated|"
    r"b\.?tech|b\.?e\.?|bca|mca|m\.?tech|b\.?sc|m\.?sc|college|university|"
    r"school|semester|batch|cgpa|gpa)\b", re.I)

PAST_ROLE = re.compile(r"\b(ex[-\s]|former(ly)?|previously|past)\b", re.I)



def normalize_role(value):
    low = " %s " % re.sub(r"\s+", " ", (value or "").lower())
    for canonical, aliases in ROLE_ALIASES:
        for alias in aliases:
            if re.search(r"(?<![a-z])%s(?![a-z])" % re.escape(alias), low):
                return canonical
    return ""


def normalize_industry(value):
    low = " %s " % re.sub(r"\s+", " ", (value or "").lower())
    for canonical, aliases in INDUSTRY_ALIASES:
        for alias in aliases:
            if re.search(r"(?<![a-z])%s(?![a-z])" % re.escape(alias), low):
                return canonical
    return ""


# --------------------------------------------------------- company + role --
# "Lead Engineer @natwestgroup" -- GitHub lower-cases org mentions, so the
# company after "@" may start with any letter.
_AT_MENTION = re.compile(
    r"([A-Z][A-Za-z0-9 .&/+-]{2,40}?)\s*@\s*"
    r"([A-Za-z][A-Za-z0-9.&_-]{1,40}(?:\s[A-Z][A-Za-z0-9.&_-]{1,25}){0,3})")
# "Software Engineer at Infosys" -- after the word "at" a company is a
# proper noun, otherwise "engineer at the office" would match.
_AT_WORD = re.compile(
    r"([A-Z][A-Za-z0-9 .&/+-]{2,40}?)\s+\bat\b\s*"
    r"@?([A-Z][A-Za-z0-9.&_-]{1,40}(?:\s[A-Z][A-Za-z0-9.&_-]{1,25}){0,3})")
_DASH = re.compile(
    r"([A-Z][A-Za-z0-9 .&/+-]{2,40}?)\s+-\s+([A-Z][A-Za-z0-9.&_-]{2,40})")
_WORKING_AS = re.compile(
    r"\b(?:working|work)\s+as\s+(?:an?\s+)?([A-Za-z][A-Za-z ./+-]{2,40}?)"
    r"\s+(?:at|@|for|with)\s+@?([A-Za-z][A-Za-z0-9.&_ -]{1,40})", re.I)
_CURRENTLY_AT = re.compile(
    r"\bcurrently\s+(?:working\s+)?(?:at|@|with)\s+@?"
    r"([A-Za-z][A-Za-z0-9.&_-]{1,40}(?:\s[A-Z][A-Za-z0-9.&_-]{1,25}){0,3})",
    re.I)

_COMPANY_NOISE = {"the", "and", "a", "an", "my", "our", "this", "home",
                  "work", "school", "college", "university", "github",
                  "linkedin", "twitter", "gmail", "email"}


def _clean_company(name):
    name = flatten(name).strip(" .,;:-|")
    if not name or name.lower() in _COMPANY_NOISE or len(name) < 2:
        return ""
    return name


def extract_company_role(text):
    """(role, company) stated as CURRENT employment, else ("", "")."""
    body = clean_text(text)
    for line in body.splitlines():
        for segment in re.split(r"\s*[|•]\s*", line):
            stripped = segment.strip()
            if not stripped or len(stripped) > 160:
                continue
            if PAST_ROLE.search(stripped):
                continue
            # Studying somewhere is not working somewhere.
            if EDUCATION_BLOCK.search(stripped):
                continue

            match = _WORKING_AS.search(stripped)
            if match:
                role = flatten(match.group(1)).strip()
                company = _clean_company(match.group(2))
                if company and ROLE_WORD.search(role):
                    return role, company

            for pattern in (_AT_MENTION, _AT_WORD, _DASH):
                match = pattern.search(stripped)
                if not match:
                    continue
                role = flatten(match.group(1)).strip()
                company = _clean_company(match.group(2))
                if not company or not ROLE_WORD.search(role):
                    continue
                return role, company

            match = _CURRENTLY_AT.search(stripped)
            if match:
                company = _clean_company(match.group(1))
                if company:
                    return "", company
    return "", ""


# ------------------------------------------------------ employment status --
_STATUS_RULES = [
    ("Intern", re.compile(
        r"\b(software|data|research|summer|winter|web|ml|ai)?\s*intern\b"
        r"|\binternship\s+at\b|\bcurrently\s+interning\b", re.I)),
    ("Freelancer", re.compile(r"\bfreelanc\w+\b", re.I)),
    ("Self-Employed", re.compile(
        r"\b(self[- ]employed|founder|co[- ]?founder|entrepreneur)\b", re.I)),
    ("Student", re.compile(
        r"\b(student|studying|pursuing|undergrad\w*|postgrad\w*|"
        r"final\s+year|b\.?tech|bca|mca|college\s+student)\b", re.I)),
    ("Fresher", re.compile(r"\bfresher\b", re.I)),
    ("Job Seeker", re.compile(
        r"\b(seeking|looking\s+for)\s+(a\s+)?(job|role|opportunit\w+|"
        r"position)\b|\bopen\s+to\s+(work|opportunit\w+)\b", re.I)),
]


def extract_employment_status(text, has_company=False):
    """Normalised status.  'Open to Work' alone never means unemployed."""
    body = flatten(text)
    if has_company:
        # Employed people may still describe themselves as open to work.
        if _STATUS_RULES[0][1].search(body):
            return "Intern"
        return "Employed"
    for status, pattern in _STATUS_RULES:
        if pattern.search(body):
            return status
    return ""


# ------------------------------------------------------------ experience --
EXPERIENCE_RE = re.compile(
    r"\b("
    r"\d+(?:\.\d+)?\s*(?:-|--|to)\s*\d+(?:\.\d+)?\s*(?:years?|yrs?|yr)"
    r"|\d+(?:\.\d+)?\s*(?:years?|yrs?|yr)"
    r"(?:\s*(?:and\s*)?\d+\s*(?:months?|mos?|mo))?"
    r"|\d+\s*(?:months?|mos?|mo)"
    r"|fresher|no\s+experience"
    r")\b", re.I)

WORK_CONTEXT = re.compile(
    r"\b(experience|exp|worked|working|work|employed|employment|intern\w*|"
    r"professional|industry|career|job|role|company|freelanc\w*|"
    r"developer|engineer|analyst|consultant)\b", re.I)
EDU_DURATION = re.compile(
    r"\b(b\.?tech|b\.?e\.?|bca|mca|m\.?tech|b\.?sc|m\.?sc|mba|bba|diploma|"
    r"degree|course|programme|program|curriculum)\b", re.I)
_INTERN_CONTEXT = re.compile(r"\b(intern\w*|trainee|apprentice)\b", re.I)
_EXPERIENCE_WORD = re.compile(r"\b(experience|exp\.?)\b", re.I)


def parse_experience_to_months(value):
    low = (value or "").lower().strip()
    if re.search(r"\bfresher\b|\bno\s+experience\b", low):
        return 0
    months = 0
    years = re.search(r"(\d+(?:\.\d+)?)\s*(?:years?|yrs?|yr)", low)
    if years:
        months += int(round(float(years.group(1)) * 12))
    month = re.search(r"(\d+)\s*(?:months?|mos?|mo)\b", low)
    if month:
        months += int(month.group(1))
    if not years and not month:
        return -1
    return months


def format_experience(months):
    if months < 0:
        return ""
    if months == 0:
        return "Fresher"
    years, rest = divmod(months, 12)
    if years and rest:
        return "%d year%s %d month%s" % (years, "s" if years > 1 else "",
                                         rest, "s" if rest > 1 else "")
    if years:
        return "%d year%s" % (years, "s" if years > 1 else "")
    return "%d month%s" % (rest, "s" if rest > 1 else "")


def _line_of(text, position):
    start = text.rfind("\n", 0, position) + 1
    end = text.find("\n", position)
    return text[start:end if end != -1 else len(text)]


def extract_experience(text):
    """Total WORK experience in months; -1 unknown, 0 fresher.

    An internship duration never wins over a stated total, and an education
    duration is never work experience at all.
    """
    body = clean_text(text)
    fresher = False
    ranked = []
    for match in EXPERIENCE_RE.finditer(body):
        value = match.group(1)
        window = body[max(0, match.start() - 70):match.end() + 70]
        if re.fullmatch(r"fresher|no\s+experience", value, re.I):
            fresher = True
            continue
        if EDU_DURATION.search(window) and not re.search(
                r"\b(experience|intern\w*|worked|working|employ\w*)\b",
                window, re.I):
            continue
        if not WORK_CONTEXT.search(window):
            continue
        range_match = re.match(
            r"(\d+(?:\.\d+)?)\s*(?:-|--|to)\s*(\d+(?:\.\d+)?)\s*"
            r"(?:years?|yrs?|yr)", value, re.I)
        months = (int(round(float(range_match.group(1)) * 12))
                  if range_match else parse_experience_to_months(value))
        if months < 0:
            continue
        line = _line_of(body, match.start())
        ranked.append((( bool(_INTERN_CONTEXT.search(line)),
                         not bool(_EXPERIENCE_WORD.search(line)),
                         len(ranked)), months))
    if ranked:
        return min(ranked)[1]
    return 0 if fresher else -1


_ASPIRATION = re.compile(
    r"\b(seeking|looking\s+for|open\s+to|interested\s+in|want|wish|"
    r"hoping|aspiring|apply|applying|available\s+for)\b", re.I)

INTERNSHIP_RE = re.compile(
    r"\b(software\s+intern|developer\s+intern|summer\s+intern|"
    r"research\s+intern|winter\s+intern|web\s+intern|data\s+intern|"
    r"ml\s+intern|ai\s+intern|intern(?:ship)?|trainee|apprentice)\b", re.I)


def extract_internships(text, limit=3):
    """Concise internship lines: role, company and duration where present."""
    body = clean_text(text)
    parsed = parse_sections(body)
    scope = section_text(parsed, "experience") or body
    found = []
    for line in scope.splitlines():
        stripped = line.strip(" -*•|")
        if not stripped or len(stripped) > 200:
            continue
        if not INTERNSHIP_RE.search(stripped):
            continue
        if stripped.lower() in {"intern", "internship", "internships",
                                "experience"}:
            continue
        # "seeking an internship" is a wish, not experience held.
        if _ASPIRATION.search(stripped):
            continue
        if stripped not in found:
            found.append(stripped)
        if len(found) >= limit:
            break
    return found


# ----------------------------------------------------------- aspirations --
INTEREST_RE = re.compile(
    r"\b(?:interested\s+in|passionate\s+about|aspiring|looking\s+for|"
    r"seeking|open\s+to|career\s+interest|focused\s+on|exploring|"
    r"want\s+to\s+become|aiming\s+to\s+be)\b[\s:]*([^.\n;|]{3,90})", re.I)


def extract_career_interest(text):
    for match in INTEREST_RE.finditer(clean_text(text)):
        phrase = flatten(match.group(1))
        phrase = re.sub(r"\b(opportunities|opportunity|roles?|positions?|"
                        r"jobs?)\b\s*$", "", phrase, flags=re.I)
        phrase = phrase.strip(" ,-")
        if len(phrase) >= 3:
            return phrase[:90]
    return ""


def extract_preferred_role(text):
    """Only a STATED preference, never the job somebody already has."""
    for match in INTEREST_RE.finditer(clean_text(text)):
        role = normalize_role(match.group(1))
        if role:
            return role
    return ""


def extract_preferred_industry(career_interest):
    """Explicit evidence only -- never inferred from a language or tool."""
    return normalize_industry(career_interest)


# ==========================================================================
# EXCEL WRITER
# ==========================================================================
# Final workbook: ONE student_data.xlsx, written by the export only.
#
# WHAT THIS FILE NO LONGER DOES
# -----------------------------
# It used to rotate output across student_data_1.xlsx, student_data_2.xlsx, ...
# at 100 data rows each, and every scraped student appended to it directly.
# Both of those are gone:
#
#     no rotation      there is one workbook.  ROWS_PER_FILE, _target_file(),
#                      the numbering-gap logic and RotationError are removed
#                      rather than left switched off, so nothing half-dead
#                      remains to be re-enabled by accident.
#
#     no per-student   scraper ranks write to their own SQLite store.  With 80
#     appends          ranks across 8 laptops, appending to one shared workbook
#                      would corrupt it -- openpyxl rewrites the whole file on
#                      every save, so two writers means one of them loses.
#
#     no duplicate     the duplicate index used to be three in-memory sets
#     index            rebuilt by re-reading every workbook at startup.  That
#                      does not scale past a few thousand rows and cannot be
#                      shared.  py + py own it now, and
#                      the database enforces it.
#
# WHAT IT STILL DOES, UNCHANGED
# -----------------------------
#     LibreOffice locks    find_lock_files()/describe_locks() are untouched.  A
#                          lock is NEVER removed automatically -- not even when
#                          it looks stale -- because the cost of being wrong is
#                          somebody's spreadsheet.
#
#     atomic saves         written to a sibling temp file, fsynced, then
#                          os.replace()d over the real one.  An interrupted
#                          export cannot corrupt the existing workbook.
#
#     the 65 columns       exact order, alias mirroring, same meanings.  The
#                          schema did not change and campaign metadata is NOT
#                          written into it: that lives in the store.
#
#     legacy migration     the pre-redesign 16-column reader is unchanged, and
#                          the numbered student_data_N.xlsx files are still
#                          READ -- for --migrate-output -- but never written.
#
# THE XLSX ROW LIMIT IS REAL
# --------------------------
# A worksheet holds 1,048,576 rows and one is the header.  The store is the
# dataset and can hold far more, so an export that would exceed the limit stops
# and says so.  Silently truncating would be the worst possible behaviour:
# the file would look complete.
# --------------------------------------------------------------------------
# Pre-redesign rotated workbooks.  Matched so they can be READ for migration
# and reported by --validate-output; never written.
_NAME_RE = re.compile(r"^student_data_(\d+)\.xlsx$", re.I)
_LOCK_RE = re.compile(r"^\.~lock\.(.+)#$")


class WorkbookLocked(RuntimeError):
    """A final workbook is open in LibreOffice; writing must not start."""


def alias_fill(record):
    """Make every alias column mirror its canonical column.

    The single place this is done, so a migrated row, an exported row and a
    freshly scraped row cannot disagree about it.
    """
    for alias, canonical in C.ALIAS_COLUMNS.items():
        record[alias] = record.get(canonical, "")
    return record

def find_lock_files(directory):
    """LibreOffice lock files in `directory`, classified.

    A lock is written by LibreOffice as `.~lock.<name>#` holding
    `,user,host,DD.MM.YYYY HH:MM,file:///...;`.  It is NEVER removed here:
    deleting an active lock would let two writers corrupt the workbook, and
    even a lock that looks stale may belong to a session we cannot see.
    """
    locks = []
    if not os.path.isdir(directory):
        return locks
    for name in sorted(os.listdir(directory)):
        match = _LOCK_RE.match(name)
        if not match:
            continue
        target = match.group(1)
        path = os.path.join(directory, name)
        info = {"lock_file": path, "workbook": target,
                "workbook_path": os.path.join(directory, target),
                "user": "", "host": "", "when": "",
                "orphan": not os.path.exists(os.path.join(directory, target))}
        try:
            with open(path, encoding="utf-8", errors="replace") as handle:
                parts = handle.read().strip().strip(";").split(",")
            if len(parts) >= 4:
                info["user"], info["host"], info["when"] = (
                    parts[1], parts[2], parts[3])
        except OSError:
            pass
        locks.append(info)
    return locks


def describe_locks(locks):
    lines = []
    for lock in locks:
        who = " ".join(part for part in (lock["user"], lock["host"]) if part)
        when = (" since %s" % lock["when"]) if lock["when"] else ""
        if lock["orphan"]:
            lines.append(
                "  %s - lock present but the workbook does not exist "
                "(stale)%s" % (lock["workbook"], when))
        else:
            lines.append("  %s appears to be open in LibreOffice%s%s"
                         % (lock["workbook"],
                            (" by %s" % who) if who else "", when))
    return lines


LEGACY_COLUMNS = ["ID", "Keyword", "Username", "Name", "Email", "Phone",
                  "Location", "Bio", "Links", "Profile URL"]


_EMPTY_VALUES = {"", "none", "not found", "n/a", "na", "null", "-"}


def _legacy_cell(header, index, row, name):
    position = index.get(name)
    if position is None or position >= len(row):
        return ""
    value = row[position]
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in _EMPTY_VALUES else text


def legacy_row_to_record(header, index, row, source_file):
    """One pre-redesign row -> one 65-column record.

    The row is fed through the SAME StudentMerger the live pipeline uses,
    so a migrated row and a freshly scraped row are built by identical
    rules -- including alias mirroring and the validation statuses.  Only
    data present in the old row is used; missing fields stay blank.
    """

    def cell(name):
        return _legacy_cell(header, index, row, name)

    profile = cell("Profile URL")
    username = cell("Username")
    if not profile and username:
        profile = "https://github.com/%s" % username
    if not profile:
        return None
    if not username:
        username = profile.rstrip("/").rsplit("/", 1)[-1]

    links = [part.strip() for part in re.split(r"[,\s]+", cell("Links"))
             if part.strip().lower().startswith("http")]
    raw_cells = [value for value in
                 (cell("Source Row"), cell("Repo URL"), cell("Repo Name"),
                  cell("Stars")) if value]

    candidate = Candidate(
        source_row_number=0,
        source_file=cell("Source File") or source_file,
        keyword=cell("Keyword"),
        username=username,
        profile_url=profile,
        xlsx_name=cell("Name"),
        xlsx_bio=cell("Bio"),
        xlsx_location=cell("Location"),
        raw_cells=raw_cells,
    )
    bundle = SourceBundle(candidate=candidate)
    # The profile really was visited when these rows were written.
    bundle.github_ok = True
    bundle.github_fields = {
        "fullname": cell("Name"),
        "nickname": username,
        "bio": cell("Bio"),
        "location": cell("Location"),
        "email": cell("Email"),
        "website": "",
        "links": links,
        "readme": "",
    }
    bundle.github_text = "\n".join(
        value for value in (cell("Name"), cell("Bio"), cell("Location"),
                            cell("Phone"), cell("Source Row")) if value)
    bundle.github_links = links

    record = StudentMerger().build(bundle)

    # Facts the old row states directly always win over anything derived.
    phone = cell("Phone")
    if phone:
        R = sys.modules[__name__]
        normalised = R.normalize_phone(phone)
        record["Phone Number"] = normalised or phone
        record["Phone Validation Status"] = ("Valid Format" if normalised
                                             else "Invalid Format")
    record["GitHub Profile"] = profile
    record["Source URL"] = profile
    record["source_file"] = cell("Source File") or source_file
    scraped = cell("Scraped At")
    if scraped:
        record["Data Collection Date"] = scraped[:10]

    note = "Migrated from the pre-redesign 16-column schema"
    if cell("Keyword"):
        note += "; search keyword: %s" % cell("Keyword")
    existing = record.get("Remarks / Notes", "")
    record["Remarks / Notes"] = ("%s; %s" % (existing, note)
                                 if existing else note)

    # Alias columns must mirror after these overrides.
    for alias, canonical in C.ALIAS_COLUMNS.items():
        record[alias] = record.get(canonical, "")
    return record


# ============================================================== validation ==
def validate_output(path=None, directory=None):
    """Check the final workbook against the 65-column contract.

    Returns (ok, problems, summary).  Read-only: production data is never
    altered to make a check pass.  Legacy numbered workbooks in the same
    directory are reported as needing migration, not treated as output.
    """
    path = path or C.FINAL_WORKBOOK_PATH
    directory = directory or os.path.dirname(path)
    problems, summary = [], []

    legacy = legacy_workbooks(directory)
    for legacy_path in legacy:
        problems.append(
            "%s is a pre-redesign rotated workbook. It is preserved for "
            "migration and is not written to; import it with "
            "--migrate-output." % os.path.basename(legacy_path))

    if not os.path.exists(path):
        return (not problems), problems, summary

    name = os.path.basename(path)
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        problems.append("%s: cannot be opened (%s)" % (name, exc))
        return False, problems, summary

    sheet = book.active
    rows = sheet.iter_rows(values_only=True)
    header = [str(value or "").strip() for value in (next(rows, None) or ())]

    if header[:len(LEGACY_COLUMNS)] == LEGACY_COLUMNS:
        problems.append("%s: still uses the pre-redesign 16-column schema"
                        % name)
        book.close()
        return False, problems, summary
    if len(header) != len(C.FINAL_COLUMNS):
        problems.append("%s: has %d columns, expected %d"
                        % (name, len(header), len(C.FINAL_COLUMNS)))
    if header != C.FINAL_COLUMNS:
        for position, (got, want) in enumerate(zip(header, C.FINAL_COLUMNS)):
            if got != want:
                problems.append("%s: column %d is %r, expected %r"
                                % (name, position + 1, got, want))
                break

    index = {value: position for position, value in enumerate(header)}
    seen_usernames = {}
    count = 0
    for position, row in enumerate(rows, start=2):
        if not any(value not in (None, "") for value in row):
            continue
        count += 1

        def cell(column):
            spot = index.get(column)
            if spot is None or spot >= len(row):
                return ""
            value = row[spot]
            return "" if value is None else str(value).strip()

        for alias, canonical in C.ALIAS_COLUMNS.items():
            if cell(alias) != cell(canonical):
                problems.append(
                    "%s row %d: alias %r=%r does not mirror %r=%r"
                    % (name, position, alias, cell(alias), canonical,
                       cell(canonical)))

        github = cell("GitHub Profile")
        if github:
            username = username_from_url(github) or github.lower()
            if username in seen_usernames:
                problems.append(
                    "%s row %d: duplicate GitHub username %r (also row %d)"
                    % (name, position, username, seen_usernames[username]))
            else:
                seen_usernames[username] = position

    if count > C.XLSX_MAX_DATA_ROWS:
        problems.append("%s: %d data rows exceeds the xlsx maximum of %d"
                        % (name, count, C.XLSX_MAX_DATA_ROWS))

    file_problems = [p for p in problems if p.startswith(name)]
    summary.append((name, len(header), count,
                    "FAIL" if file_problems else "PASS"))
    book.close()
    return (not problems), problems, summary


# ============================================================ legacy input ==
def legacy_workbooks(directory=None):
    """student_data_N.xlsx files, sorted.  Read for migration, never written.

    These are PRESERVED on purpose: they hold real collected students and are
    the backup of the pre-redesign output.  Nothing in this project deletes
    them.
    """
    directory = directory or C.FINAL_OUTPUT_DIR
    if not os.path.isdir(directory):
        return []
    found = []
    for name in os.listdir(directory):
        match = _NAME_RE.match(name)
        if match:
            found.append((int(match.group(1)),
                          os.path.join(directory, name)))
    return [path for _number, path in sorted(found)]


def read_workbook_records(path):
    """Every data row of one workbook as a 65-column record.

    Handles both schemas: a current-schema file is read straight, and a
    pre-redesign 16-column file is fed through legacy_row_to_record() so it
    is built by exactly the same rules a live run uses.
    """
    name = os.path.basename(path)
    records, problems = [], []
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        return [], ["%s: cannot be opened (%s)" % (name, exc)]
    sheet = book.active
    rows = sheet.iter_rows(values_only=True)
    header = [str(value or "").strip() for value in (next(rows, None) or ())]
    index = {value: position for position, value in enumerate(header)}
    current_schema = header[:len(C.FINAL_COLUMNS)] == C.FINAL_COLUMNS

    for number, row in enumerate(rows, start=2):
        if not any(value not in (None, "") for value in row):
            continue
        if current_schema:
            record = {}
            for column in C.FINAL_COLUMNS:
                spot = index.get(column)
                value = (row[spot] if spot is not None and spot < len(row)
                         else "")
                record[column] = "" if value is None else str(value).strip()
            records.append(alias_fill(record))
            continue
        record = legacy_row_to_record(header, index, row, name)
        if record is None:
            problems.append("%s row %d: no GitHub profile URL" % (name, number))
            continue
        records.append(alias_fill(record))
    book.close()
    return records, problems


# ============================================================ final output ==
def _atomic_save(book, path):
    """Write to a sibling temp file, flush it, then replace the real one.

    A crash therefore leaves either the previous workbook or the new one --
    never a half-written file.  fsync before replace, because os.replace is
    atomic with respect to the directory entry but not to the data.
    """
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    tmp = os.path.join(directory, ".%s.tmp" % os.path.basename(path))
    book.save(tmp)
    with open(tmp, "rb") as handle:
        os.fsync(handle.fileno())
    os.replace(tmp, path)
    return path


class RowLimitExceeded(RuntimeError):
    """The export would need more rows than a worksheet can hold."""


def guard_locks(directory=None, say=None):
    """Refuse to write while a workbook is open elsewhere.

    An active LibreOffice lock means somebody has the workbook open; saving
    over it would lose their edits.  The lock is never removed automatically
    -- not even when it looks stale -- because the cost of being wrong is
    somebody's data.  Returns the stale locks it reported.
    """
    directory = directory or C.FINAL_OUTPUT_DIR
    locks = find_lock_files(directory)
    active = [lock for lock in locks if not lock["orphan"]]
    stale = [lock for lock in locks if lock["orphan"]]
    if say is not None:
        for line in describe_locks(stale):
            say("WARNING:%s" % line)
        if stale:
            say("         Nothing was deleted. Remove the stale lock file "
                "yourself if you are sure.")
    if active:
        raise WorkbookLocked("\n".join(describe_locks(active)))
    return stale


def write_final(records, path=None):
    """Write ONE workbook holding every record, atomically.

    Returns (path, rows_written).  Raises RowLimitExceeded rather than
    truncating: a workbook that quietly dropped the tail of the dataset would
    look complete and be wrong.
    """
    path = path or C.FINAL_WORKBOOK_PATH
    records = list(records)
    if len(records) > C.XLSX_MAX_DATA_ROWS:
        raise RowLimitExceeded(
            "%d records will not fit in one xlsx worksheet (maximum %d data "
            "rows).\n"
            "Nothing was written -- the existing workbook is untouched.\n"
            "The store remains the complete dataset.  Export a subset (for "
            "example by state) or to CSV instead."
            % (len(records), C.XLSX_MAX_DATA_ROWS))

    book = Workbook(write_only=False)
    sheet = book.active
    sheet.title = "Students"
    sheet.append(C.FINAL_COLUMNS)
    for record in records:
        alias_fill(record)
        sheet.append([record.get(column, "") for column in C.FINAL_COLUMNS])
    _atomic_save(book, path)
    book.close()
    log.info("wrote %d record(s) to %s", len(records), path)
    return path, len(records)


# ==========================================================================
# GITHUB PROFILE
# ==========================================================================
# Public GitHub profile extraction.
#
# Structured DOM fields are read first, then the visible text as a fallback
# (spec 23).  The Playwright equivalent of Ctrl+A/Ctrl+C is
# `page.locator("body").inner_text()`, which never touches the clipboard.
# --------------------------------------------------------------------------
# Several selector generations are tried because GitHub's markup changes.
PROFILE_FIELDS_JS = r"""
() => {
  const clean = v => (v || '').replace(/\s+/g, ' ').trim();
  const pick = (...sels) => {
    for (const sel of sels) {
      const el = document.querySelector(sel);
      if (el) { const v = clean(el.textContent); if (v) return v; }
    }
    return '';
  };

  const pronouns = pick('[itemprop="pronouns"]');
  const strip = v => {
    let out = v || '';
    if (pronouns) out = out.split(pronouns).join(' ');
    return clean(out).replace(/^[\s·•|,\/]+|[\s·•|,\/]+$/g, '');
  };

  const isUser = !!document.querySelector(
      '.vcard-username, [itemprop="additionalName"], .h-card');
  const isOrg = !!document.querySelector(
      '.orghead, [data-hovercard-type="organization"] .org-name') ||
      /\/orgs\//.test(location.pathname);

  const mail = document.querySelector('li[itemprop="email"] a[href^="mailto:"]')
            || document.querySelector('.vcard-details a[href^="mailto:"]')
            || document.querySelector('a[href^="mailto:"]');

  const links = Array.from(document.querySelectorAll(
      '.vcard-details a[href], [data-test-selector="profile-website-url"] a'))
      .map(a => a.getAttribute('href') || '')
      .filter(h => h && !h.startsWith('mailto:'));

  // The website row GitHub renders for the profile's own "website" field.
  let website = '';
  const site = document.querySelector(
      '[data-test-selector="profile-website-url"] a, li[itemprop="url"] a');
  if (site) website = site.getAttribute('href') || clean(site.textContent);

  // The profile README must keep its line structure: textContent runs every
  // heading and list item together, which makes section detection
  // impossible.  innerText preserves the rendered line breaks.
  let readmeText = '';
  const readmeEl = document.querySelector(
      '.Box-body .markdown-body, article.markdown-body, .markdown-body');
  if (readmeEl) readmeText = (readmeEl.innerText || '').trim();

  const social = Array.from(document.querySelectorAll(
      '.vcard-details a[href^="http"], .js-profile-editable-area a[href^="http"]'))
      .map(a => a.href);

  return {
    isUser: isUser,
    isOrg: isOrg && !isUser,
    notFound: /Not Found|404/.test(document.title || ''),
    nickname: clean((location.pathname.split('/')[1] || '')),
    fullname: strip(pick('.vcard-fullname', '[itemprop="name"]',
                         'h1.vcard-names span.p-name')),
    bio: pick('.user-profile-bio', '[data-bio-text]', '.p-note'),
    location: pick('[itemprop="homeLocation"]', '.p-label'),
    email: mail ? clean(mail.getAttribute('href').replace('mailto:', '')) : '',
    website: website,
    links: Array.from(new Set(links.concat(social))),
    readme: readmeText
  };
}
"""


class GitHubProfileError(RuntimeError):
    def __init__(self, message, permanent=False):
        super().__init__(message)
        self.permanent = permanent


def blocked_marker(text):
    low = (text or "").lower()
    for marker in C.BLOCK_MARKERS:
        if marker in low:
            return marker
    return None


# A profile README lives in the repository <username>/<username>.  The
# default branch is not always "main", so the raw URL uses the HEAD ref,
# which GitHub resolves to whatever the default branch actually is.
RAW_README_URL = "https://raw.githubusercontent.com/%s/%s/HEAD/%s"
README_NAMES = ("README.md", "readme.md", "README.MD", "Readme.md",
                "README", "README.markdown")


class GitHubProfile:
    def __init__(self, browser, say=print):
        self.browser = browser
        self.say = say

    async def fetch_readme_raw(self, username):
        """The profile README as Markdown, or "".

        The rendered page is not enough: a tech stack written as badges is
        a row of <img> elements, and images contribute nothing to
        innerText.  The names only exist in the Markdown source.
        """
        context = getattr(self.browser, "context", None)
        if context is None:
            return ""
        for name in README_NAMES:
            url = RAW_README_URL % (username, username, name)
            try:
                response = await context.request.get(url, timeout=20000)
            except Exception as exc:
                log.debug("raw README request failed (%s): %s", name, exc)
                continue
            if response.status == 200:
                try:
                    body = await response.text()
                except Exception:
                    continue
                if body.strip():
                    log.info("raw README fetched for %s (%s, %d chars)",
                             username, name, len(body))
                    return body
            elif response.status not in (404,):
                log.debug("raw README %s -> HTTP %d", name, response.status)
        return ""

    async def fetch(self, page, username):
        """Return (fields, visible_text).  Raises GitHubProfileError."""
        url = "https://github.com/%s" % username
        last_error = None
        for attempt in range(1, C.PROFILE_RETRIES + 1):
            try:
                response = await page.goto(url, wait_until="domcontentloaded",
                                           timeout=C.PROFILE_TIMEOUT)
                status = response.status if response else 0
                if status == 404:
                    raise GitHubProfileError("404 not found", permanent=True)
                if status in C.BLOCK_STATUS:
                    raise GitHubProfileError(
                        "GitHub returned HTTP %d" % status)

                body = await page.locator("body").inner_text(
                    timeout=C.PROFILE_TIMEOUT)
                fields = await page.evaluate(PROFILE_FIELDS_JS)
                if fields.get("notFound"):
                    raise GitHubProfileError("404 not found", permanent=True)
                if fields.get("isOrg") and not fields.get("isUser"):
                    raise GitHubProfileError("organisation, not a user",
                                             permanent=True)

                # BLOCK DETECTION IS ASKED LAST, AND ONLY OF A PAGE THAT IS
                # NOT A PROFILE.
                #
                # The body text of a profile page includes the person's own
                # profile README, so scanning it for words like "rate limit"
                # reads THEIR content as if it were GitHub's.  Two live
                # examples, both of which stopped a rank:
                #     "JWT auth, bcrypt hashing, rate limiting, env ..."
                #     "and rate limiting (300 GET / 60 write req/min)"
                # Neither was a rate limit; both were skills lists, hit at
                # ~4 requests/minute -- hundreds of times under any GitHub
                # limit.
                #
                # A genuine block or challenge page does NOT render a user
                # profile, so a page that parsed as one cannot be a block
                # however its README is worded.  HTTP 429/403 above stays
                # authoritative and is checked first.
                if not fields.get("isUser"):
                    marker = blocked_marker(body)
                    if marker:
                        raise GitHubProfileError(
                            "GitHub rate limit / challenge: %s" % marker)
                    # Layout change or a partial render -- retry, never a
                    # permanent skip.
                    raise GitHubProfileError("profile markup not recognised")
                # Separate source: GITHUB_README_RAW.
                fields["readme_raw"] = await self.fetch_readme_raw(username)
                return fields, body
            except GitHubProfileError as exc:
                if exc.permanent:
                    raise
                last_error = exc
            except Exception as exc:
                if browser_died(exc):
                    raise BrowserLost(str(exc))
                last_error = exc
            if attempt < C.PROFILE_RETRIES:
                # Jittered, so two copies retrying the same profile minute do
                # not retry together (naukri jitters every backoff).
                await sleep_jitter(2 * attempt)
        raise GitHubProfileError(str(last_error) or "profile unavailable")


async def pace(span):
    await asyncio.sleep(random.uniform(*span))


# ==========================================================================
# IDENTITY
# ==========================================================================
# One definition of "the same person", shared by every component.
#
# The duplicate rules live here and nowhere else.  The store's UNIQUE indexes,
# the scraper's pre-extraction skip and the final export all call these
# functions, so the three can never drift apart and disagree about whether two
# rows are the same human being.
#
# IDENTITY RULES (unchanged in meaning from the pre-redesign behaviour)
#
#     primary    normalised GitHub username, case-insensitively
#                normalised GitHub profile URL
#
#     secondary  a reliable normalised email address
#
#     never      name, city, college, skills, or any combination of them.
#                Two students called Rahul Sharma studying CSE in Pune are two
#                students.
#
# WHAT NORMALISATION FIXES
#
#     usernames   GitHub usernames are case-insensitive, so
#                     AGAUTHAMRAJU  agouthamraju  AgouthamRaju
#                 are one account.
#
#     URLs        scheme, host case, a "www." prefix, a trailing slash and a
#                 trailing ".git" are all noise:
#                     https://github.com/User
#                     https://github.com/user/
#                     http://github.com/user
#                     https://www.github.com/user
#                 are one account.  The pre-redesign index lower-cased and
#                 stripped the trailing slash but kept the scheme, so
#                 http:// and https:// forms of one profile were two keys; it
#                 only worked because the username tail was indexed too.
#
#     emails      lower-cased, trimmed.  Placeholders are rejected outright:
#                 thousands of students copy a portfolio template and never
#                 replace "your_email@example.com", and treating that as an
#                 identity would collapse all of them into one record.
# --------------------------------------------------------------------------
# Single-segment github.com paths that are never a user profile.  config
# carries the full list for the scraper; this is the subset that can appear
# where a profile URL is expected.
_NOT_A_USER = {
    "about", "apps", "blog", "collections", "contact", "dashboard",
    "enterprise", "events", "explore", "features", "issues", "join",
    "login", "logout", "marketplace", "new", "notifications", "orgs",
    "organizations", "pricing", "pulls", "readme", "search", "security",
    "sessions", "settings", "signup", "site", "sponsors", "stars", "team",
    "topics", "trending", "users", "watching",
}

# A GitHub username: alphanumeric and single hyphens, 1-39 characters, no
# leading or trailing hyphen.  Anything else is not a username, so it is
# never used as an identity.
_USERNAME_RE = re.compile(r"^[A-Za-z0-9](?:[A-Za-z0-9]|-(?=[A-Za-z0-9])){0,38}$")

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")

# Local parts and domains that mean "the template was never filled in".
_PLACEHOLDER_LOCALS = {
    "email", "youremail", "your_email", "your-email", "myemail", "my_email",
    "example", "test", "testing", "sample", "demo", "dummy", "user",
    "username", "name", "yourname", "your_name", "firstname", "lastname",
    "someone", "somebody", "anybody", "abc", "xyz", "asdf", "qwerty",
    "noreply", "no-reply", "donotreply", "do-not-reply", "admin", "mail",
    "info", "contact", "hello", "support", "changeme", "placeholder",
}
_PLACEHOLDER_DOMAINS = {
    "example.com", "example.org", "example.net", "example.co",
    "test.com", "test.org", "test.net", "testing.com",
    "domain.com", "yourdomain.com", "mydomain.com", "somedomain.com",
    "email.com", "youremail.com", "mail.com", "sample.com", "demo.com",
    "dummy.com", "abc.com", "xyz.com", "company.com", "yourcompany.com",
    "website.com", "yourwebsite.com", "site.com", "localhost",
    "gmail.co", "gmial.com", "gmai.com",
}



def normalize_username(value):
    """'AgouthamRaju' -> 'agouthamraju'.  '' when it is not a username."""
    text = (value or "").strip().lstrip("@")
    if not text:
        return ""
    # Tolerate a full URL being passed in by mistake rather than silently
    # producing a nonsense key.
    if "/" in text or "github.com" in text.lower():
        return normalize_profile_url(text).rsplit("/", 1)[-1]
    if not _USERNAME_RE.match(text):
        return ""
    if text.lower() in _NOT_A_USER:
        return ""
    return text.lower()


def normalize_profile_url(value):
    """Any spelling of a GitHub profile -> 'https://github.com/<lowercase>'.

    Returns "" for anything that is not a single-segment github.com user
    path, so a repository URL or an organisation page never becomes a
    person's identity.
    """
    text = (value or "").strip()
    if not text:
        return ""
    text = re.sub(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", "", text)
    text = text.split("?", 1)[0].split("#", 1)[0]
    text = text.rstrip("/")
    if not text:
        return ""
    host, _, path = text.partition("/")
    host = host.lower()
    if host.startswith("www."):
        host = host[4:]
    if host not in ("github.com", "gist.github.com"):
        return ""
    segments = [segment for segment in path.split("/") if segment]
    if len(segments) != 1:
        return ""
    username = segments[0]
    if username.lower().endswith(".git"):
        username = username[:-4]
    username = normalize_username(username)
    if not username:
        return ""
    return "https://github.com/%s" % username


def username_from_url(value):
    """The username inside a profile URL, normalised.  '' when there is none."""
    url = normalize_profile_url(value)
    return url.rsplit("/", 1)[-1] if url else ""


def identity_is_placeholder_email(value):
    """True for template addresses that must never establish identity."""
    text = (value or "").strip().lower()
    if not text or not _EMAIL_RE.match(text):
        return True
    local, _, domain = text.partition("@")
    if domain in _PLACEHOLDER_DOMAINS:
        return True
    if local in _PLACEHOLDER_LOCALS:
        return True
    # "your.email", "your-email-here", "email_here"
    bare = re.sub(r"[^a-z]", "", local)
    if bare in _PLACEHOLDER_LOCALS:
        return True
    if domain.split(".")[0] in ("example", "test", "domain", "yourdomain"):
        return True
    return False


def normalize_email(value):
    """A trustworthy lower-cased address, or '' for absent/placeholder."""
    text = (value or "").strip().lower()
    if identity_is_placeholder_email(text):
        return ""
    return text


def identity_keys(username="", profile_url="", email=""):
    """The three keys used by the store and the export, all normalised.

    Order matters: a profile URL is the stronger signal, so it fills in a
    missing or malformed username rather than the other way round.
    """
    url_key = normalize_profile_url(profile_url)
    name_key = normalize_username(username) or (
        url_key.rsplit("/", 1)[-1] if url_key else "")
    if name_key and not url_key:
        url_key = "https://github.com/%s" % name_key
    return {
        "github_username_key": name_key,
        "profile_url_key": url_key,
        "email_key": normalize_email(email),
    }


# ==========================================================================
# INSTANT DATA SCRAPER
# ==========================================================================
# Drives the Instant Data Scraper extension for ONE keyword.
#
# One keyword == one crawl == one XLSX (spec 9/16).  The extension does its
# own pagination; this module never clicks GitHub's Next while a crawl is
# running.
#
# The UI is the extension's own popup page.  background.js builds it as
#
#     windows.create({url: runtime.getURL(
#         "src/popup.html?" + URLSearchParams({tabid, url})), type: "popup"})
#
# on action.onClicked.  Ctrl+G is tried first so the configured shortcut is
# used when the browser delivers it; because the manifest has no "commands"
# key the shortcut can only be a browser-level _execute_action binding, which
# page.keyboard.press() cannot reach, so the same URL is opened directly as
# the fallback.  Both paths produce the identical page.
#
# Control ids come from src/popup.js.  A control is rendered only while it is
# usable (`e.startScraping && jsx(button id:"startScraping")`), so presence in
# the DOM *is* the availability signal.
# --------------------------------------------------------------------------
_STAT_RE = re.compile(
    r"Pages scraped:\s*(\d+).*?Rows collected:\s*(\d+)", re.S)

# GitHub's pagination control, most specific first.  Used both to click
# Next during the one-off learn flow and, via config.STABLE_NEXT_SELECTORS,
# to seed the extension with a selector that does not rot.
NEXT_LINK_SELECTORS = list(C.STABLE_NEXT_SELECTORS) + [
    'nav[aria-label="Pagination"] a:has-text("Next")',
    'a:has-text("Next")',
]


# Everything needed to tell apart the four reasons the extension can say
# "Table not found".  Read-only: it never clicks, navigates or changes state.
PAGE_STATE_JS = r"""
() => {
  const text = (document.body ? document.body.innerText : '') || '';
  const norm  = text.replace(/\s+/g, ' ').trim();

  const nextEl = document.querySelector('a[rel="next"]');
  const nextDisabled = nextEl
      ? (nextEl.getAttribute('aria-disabled') === 'true'
         || nextEl.hasAttribute('disabled')
         || nextEl.getAttribute('aria-hidden') === 'true')
      : null;

  // User cards, counted by shape rather than by GitHub's hashed classes:
  // a results area anchor pointing at a single-segment profile path.
  const reserved = new Set(['search','login','settings','topics','orgs',
      'marketplace','features','pricing','about','security','enterprise',
      'notifications','issues','pulls','codespaces','sponsors','explore',
      'trending','collections','events','new','organizations','users']);
  const profileLinks = new Set();
  document.querySelectorAll('a[href]').forEach(a => {
    const href = a.getAttribute('href') || '';
    const m = href.match(/^\/([A-Za-z0-9][A-Za-z0-9-]{0,38})$/);
    if (m && !reserved.has(m[1].toLowerCase())) profileLinks.add(m[1]);
  });

  // The same links, but ONLY those outside the chrome.  A signed-in GitHub
  // page carries the viewer's own avatar, "Your profile", sponsor links and
  // footer links, every one of which is a single-segment profile path.
  // Counting those as search results is how an empty results page could be
  // mistaken for a page full of users, so the chrome is excluded explicitly.
  const CHROME = 'header, nav, footer, [role="banner"], [role="navigation"],'
               + ' [role="contentinfo"], .Header, .header, .footer,'
               + ' .js-header-wrapper, [data-testid="site-header"],'
               + ' [aria-label="Global"], [aria-label="User navigation"]';
  const resultProfileLinks = new Set();
  document.querySelectorAll('a[href]').forEach(a => {
    const href = a.getAttribute('href') || '';
    const m = href.match(/^\/([A-Za-z0-9][A-Za-z0-9-]{0,38})$/);
    if (!m || reserved.has(m[1].toLowerCase())) return;
    if (a.closest(CHROME)) return;
    resultProfileLinks.add(m[1]);
  });

  // Is this actually a Users search, and did a results region render?
  const params = new URLSearchParams(location.search || '');
  const isUsersSearch = /\/search$/.test(location.pathname)
                        && params.get('type') === 'users';
  // Declared here, BEFORE anything that reads it: `region` below and the
  // return object both reference resultsList, and a `const` used before its
  // declaration is a temporal-dead-zone ReferenceError at RUNTIME (it parses
  // fine, which is why a parse-only check missed it).
  const resultsList = document.querySelector('[data-testid="results-list"]');
  const mainRegion = document.querySelector(
      '[data-testid="results-list"], main, [role="main"]');

  // A cheap content fingerprint, so "has this stopped changing?" can be
  // answered without diffing the whole DOM.
  const region = resultsList || mainRegion;

  // GitHub's own empty-state copy.  The apostrophe is typographic in
  // GitHub's UI, so it is matched as a character class rather than assumed
  // to be ASCII -- an exact-substring check with a straight quote silently
  // fails against "couldn\u2019t".
  const APOS = '[\u2018\u2019\u0027\u00b4`]?';
  const zeroUsersCopy =
      new RegExp('we\\s+couldn' + APOS + 't\\s+find\\s+any\\s+users', 'i')
          .test(norm)
      || /no users (matched|found)/i.test(norm)
      || /your search did not match any users/i.test(norm);
  // A standalone zero count, never a digit-prefixed one: "0 users" is empty,
  // "10 users" is not.  The old /0 results/ alternative matched BOTH.
  const zeroUsersCount = /(^|[^\d.,])0\s+users\b/i.test(norm);

  // Login / authentication wall.  A login page has no results either, and
  // must never be read as "this search has no users".
  const loginForm = !!document.querySelector(
      'form[action="/session"], input[name="login"], #login_field,'
      + ' input[name="password"], #password');
  const loginCopy = /sign in to github|you must be signed in|please sign in/i
      .test(norm);

  // GitHub states the size of the result set: "Users 19 (19) results".
  // Used to verify a finished multi-page crawl collected everything.
  const totalMatch = norm.match(/Users\s+([\d.,]+k?)\s*\(/i);

  // Anything that says the page is still rendering.  A crawl must not open
  // the extension while the results are mid-swap.
  let skeletonCount = 0;
  ['[class*="keleton"]', '[aria-busy="true"]', '[role="progressbar"]',
   'progress', '[data-testid*="loading"]'].forEach(s => {
    try { skeletonCount += document.querySelectorAll(s).length; }
    catch (e) { /* an unsupported selector is not evidence */ }
  });



  // Does GitHub offer ANY next-page control?  hasNext above looks only for
  // a[rel="next"], which is the shape we prefer to teach the extension.
  // Deciding "this is a single page" needs the broader question answered,
  // because a differently shaped Next must not be mistaken for none at all.
  const nextShapes = ['a[rel="next"]', 'a[aria-label="Next Page"]',
                      'a[aria-label="Next page"]', 'a.next_page'];
  let nextByShape = 0;
  nextShapes.forEach(s => { nextByShape += document.querySelectorAll(s).length; });

  const pagination = document.querySelector(
      'nav[aria-label="Pagination"], nav[aria-label="pagination"], '
      + '.pagination, [class*="Pagination"]');
  let nextByText = 0;
  const scope = pagination || document;
  scope.querySelectorAll('a').forEach(a => {
    const t = (a.textContent || '').trim().toLowerCase();
    if (t === 'next' || t === 'next page' || t === 'next >' || t === '>') {
      nextByText += 1;
    }
  });

  return {
    nextByShape: nextByShape,
    nextByText: nextByText,
    paginationPresent: !!pagination,
    url: location.href,
    title: document.title || '',
    bodyChars: text.length,
    bodyHead: norm.slice(0, 600),
    bodyTail: norm.slice(-400),
    hasNext: !!nextEl,
    nextDisabled: nextDisabled,
    nextHref: nextEl ? (nextEl.getAttribute('href') || '') : '',
    profileLinkCount: profileLinks.size,
    resultsListPresent: !!resultsList,
    resultsListChildren: resultsList ? resultsList.children.length : 0,
    // Deliberately NOT the old
    //   /We couldn.t find any users|no users matched|0 results/i
    // whose "0 results" alternative matched "10 results", "100 results" and
    // "230 results" as substrings, so a page FULL of users reported itself
    // empty.
    reportedTotalUsers: totalMatch ? totalMatch[1] : null,
    skeletonCount: skeletonCount,
    domSignature: region ? region.innerHTML.length : 0,
    noResults: zeroUsersCopy || zeroUsersCount,
    zeroUsersCopy: zeroUsersCopy,
    zeroUsersCount: zeroUsersCount,
    resultProfileLinkCount: resultProfileLinks.size,
    resultProfileUsers: Array.from(resultProfileLinks).slice(0, 300),
    isUsersSearch: isUsersSearch,
    mainRegionPresent: !!mainRegion,
    loginForm: loginForm,
    loginCopy: loginCopy,
    saysError: /this page is taking too long|unicorn|something went wrong|server error|bad gateway|service unavailable|try again/i.test(norm),
  };
}
"""


# The extension's own preview grid (handsontable, id "hot").  Reading the
# cells it currently DISPLAYS is how the selected table is judged BEFORE any
# export -- in this environment export can kill Chromium, so cycling tables by
# exporting each candidate is not an option.
#
# It now returns the grid STRUCTURE as well -- column headers and one array of
# cell texts per data row -- because the verified preview, not the extension's
# #xlsx button, is the export source for GitHub Users pages: the button
# sometimes writes a table that diverged from the validated preview (and can
# kill Chromium on click), so Python writes the page XLSX straight from these
# rows.  `cells` (the flat list) is kept unchanged so identity reading and the
# change-signature are exactly as before.
IDS_PREVIEW_JS = r"""
() => {
  const host = document.querySelector('#hot')
            || document.querySelector('.handsontable');
  if (!host) return {hasTable: false, grids: [], cells: []};

  // `cells` is the LEGACY flat list, kept only for the parser tests; the
  // production identity path uses `grids` and never unions across them.
  const cells = [];
  host.querySelectorAll('td, th').forEach(node => {
    const text = (node.textContent || '').trim();
    if (text) cells.push(text);
  });

  const grids = [];

  function gridFromCore(core) {
    let headers = [];
    core.querySelectorAll('thead tr th').forEach(th => {
      headers.push((th.textContent || '').trim());
    });
    const rows = [];
    core.querySelectorAll('tbody tr').forEach(tr => {
      const row = [];
      // Data cells only: the leading <th> is Handsontable's row-number header.
      tr.querySelectorAll('td').forEach(td => {
        row.push((td.textContent || '').trim());
      });
      if (row.length) rows.push(row);
    });
    if (!rows.length) return null;
    // Align headers to data columns: with row headers on, thead has a leading
    // corner <th>, so keep the last N headers (N = data columns).
    let numCols = 0;
    rows.forEach(r => { if (r.length > numCols) numCols = r.length; });
    if (numCols && headers.length > numCols) {
      headers = headers.slice(headers.length - numCols);
    }
    return {headers: headers, rows: rows};
  }

  // 1. Authoritative path: the Handsontable INSTANCE model, if reachable.
  //    One instance == one dataset, so there is no clone ambiguity.
  try {
    const candidates = [host];
    host.querySelectorAll('.handsontable').forEach(el => candidates.push(el));
    for (const el of candidates) {
      const inst = el.hotInstance;
      if (inst && typeof inst.getData === 'function') {
        const data = inst.getData() || [];
        let headers = [];
        if (typeof inst.getColHeader === 'function') {
          const h = inst.getColHeader();
          if (Array.isArray(h)) headers = h.map(x => (x == null ? '' : String(x)));
        }
        const rows = data.map(r => (r || []).map(
            x => (x == null ? '' : String(x).trim())));
        if (rows.length) { grids.push({headers: headers, rows: rows}); break; }
      }
    }
  } catch (e) { /* instance not reachable; fall through to DOM grids */ }

  // 2. DOM path: EACH concrete .htCore grid on its own -- overlay clones are
  //    separate entries, never merged.  (A clone with no <tbody> data rows is
  //    skipped by gridFromCore.)
  let cores = host.querySelectorAll('.htCore');
  if (!cores || !cores.length) cores = host.querySelectorAll('table');
  cores.forEach(core => {
    const g = gridFromCore(core);
    if (g) grids.push(g);
  });

  return {hasTable: true, grids: grids, cells: cells};
}
"""


def _short(selector, limit=70):
    return selector if len(selector) <= limit else selector[:limit] + "..."


def page_number(url):
    """GitHub's own page number from a search URL.  1 when there is no ?p=.

    This is the ground truth for "which page is the browser on", and it is
    what the extension's "Pages scraped" counter is checked against: if the
    browser reached p=2 and the extension only ever scraped 1 page, a page of
    results was navigated to and never collected.
    """
    try:
        query = urllib.parse.parse_qs(urllib.parse.urlparse(url or "").query)
    except Exception:
        return 1
    for key in ("p", "page"):
        values = query.get(key)
        if values and str(values[0]).isdigit():
            return max(1, int(values[0]))
    return 1


#: GitHub stops returning user-search results past this many.  A total above
#  it can never be collected in full, so numeric equality must not be asked
#  for.  (GitHub documents 1,000 for search results.)
GITHUB_SEARCH_RESULT_CAP = 1000

#: Users per results page.
GITHUB_USERS_PER_PAGE = 10


class ValidationMode:
    """How a finished multi-page crawl is checked for completeness.

        EXACT_TOTAL        GitHub stated a small, exact figure ("19 users"),
                           so the unique rows collected must equal it.
        PAGE_COMPLETENESS  the figure was abbreviated ("9.1k"), absent, or
                           above the result cap, so equality is impossible;
                           instead every reachable page through the last one
                           must have been exported.
    """

    EXACT_TOTAL = "exact-total"
    PAGE_COMPLETENESS = "page-completeness"


def expected_total_users(state):
    """(total, mode) from GitHub's own stated result count.

    total is None when GitHub did not state a usable figure.  "9.1k" is
    recognised as abbreviated and deliberately NOT turned into 9100: the real
    number is unknown, and inventing one would make a completeness check fail
    for ever.
    """
    raw = state.get("reportedTotalUsers")
    if raw in (None, ""):
        return (None, ValidationMode.PAGE_COMPLETENESS)
    text = str(raw).strip().lower().replace(",", "")
    if text.endswith("k") or text.endswith("m"):
        return (None, ValidationMode.PAGE_COMPLETENESS)
    if "." in text:
        # "9.1" only ever appears as part of an abbreviated figure.
        return (None, ValidationMode.PAGE_COMPLETENESS)
    if not text.isdigit():
        return (None, ValidationMode.PAGE_COMPLETENESS)
    total = int(text)
    if total > GITHUB_SEARCH_RESULT_CAP:
        return (total, ValidationMode.PAGE_COMPLETENESS)
    return (total, ValidationMode.EXACT_TOTAL)


def expected_page_count(total):
    """How many pages a stated total implies, or None."""
    if not total:
        return None
    return -(-int(total) // GITHUB_USERS_PER_PAGE)


class CrawlProgress:
    """Positive evidence that pagination actually happened.

    WHY THIS EXISTS -- the live failure it is built from
    ---------------------------------------------------
    "University Student Visakhapatnam" has 19 users over 2 pages.  The crawl
    was declared complete after 10 rows and ONE page, and the search unit was
    accepted as successfully collected.  Nine people were silently lost.

    The captured state at the decision point:

        url                  ...&type=users&p=2   <- the browser DID navigate
        nextDisabled         True                 <- p2 really is the last page
        resultsListChildren  9                    <- 9 users ON SCREEN
        ext_pages            1                    <- but only 1 page scraped
        ext_error            Table not found      <- the extension lost p2
        ext_instructions     "Crawling stopped. Please download data or
                              continue crawling."

    classify_page_state() saw "last page, results present" and called it
    END_OF_RESULTS -- which is indistinguishable, FROM THE PAGE ALONE, from
    "the extension failed to scrape the last page".  The page cannot answer
    it; only the crawl's own history can.

    Note that a bare "did any page transition happen?" check would have
    PASSED this case, because a transition did happen.  Two independent
    guarantees are needed:

        transitions >= 1            the crawler actually moved (catches a
                                    crawler that never navigated at all)
        ext_pages >= browser page   every page the browser reached was also
                                    SCRAPED (catches this failure)
    """

    def __init__(self, had_next=False, start_url=""):
        #: did GitHub offer a live next page BEFORE Start Crawling?
        self.had_next = bool(had_next)
        self.start_url = start_url or ""
        self.start_page = page_number(start_url)
        #: highest GitHub page number the browser was ever seen on
        self.max_browser_page = self.start_page
        #: how many times the browser's page number went up
        self.transitions = 0
        self.ext_pages = 0
        self.ext_rows = 0
        self.last_browser_page = self.start_page

    def note_url(self, url):
        """Record the browser's page number.  Returns the new page, or 0."""
        number = page_number(url)
        moved = 0
        if number > self.last_browser_page:
            self.transitions += 1
            moved = number
        if number > self.max_browser_page:
            self.max_browser_page = number
        self.last_browser_page = number
        return moved

    def note_stats(self, pages, rows):
        if pages and pages > 0:
            self.ext_pages = max(self.ext_pages, pages)
        if rows and rows > 0:
            self.ext_rows = max(self.ext_rows, rows)

    def verify_complete(self):
        """(ok, why) -- may a crawl be declared COMPLETE right now?

        Refuses in exactly two situations, both of which mean rows are
        missing:

          1  a live Next existed before the crawl and the browser never moved
          2  the browser reached a page the extension never scraped
        """
        if self.had_next and self.transitions == 0:
            return (False,
                    "GitHub offered a next page but the crawler never left "
                    "page %d" % self.start_page)
        if self.ext_pages and self.max_browser_page > self.ext_pages:
            return (False,
                    "the browser reached page %d but the extension scraped "
                    "only %d page(s); page %d was navigated to and never "
                    "collected"
                    % (self.max_browser_page, self.ext_pages,
                       self.ext_pages + 1))
        return (True, "")

    def describe(self):
        return ("pages visited %d, transitions %d, extension pages %d, "
                "rows %d" % (self.max_browser_page, self.transitions,
                             self.ext_pages, self.ext_rows))


class CrawlOutcome:
    """Why wait_for_crawl_complete() returned."""

    FINISHED = "finished"           # extension reported the crawl stopped
    END_OF_RESULTS = "end-of-results"   # no more pages; data is complete
    SINGLE_PAGE = "single-page"     # GitHub served ONE page; nothing to crawl
    RATE_LIMITED = "rate-limited"   # GitHub throttled or challenged us
    TABLE_LOST = "table-lost"       # valid page, extension lost detection
    INCOMPLETE = "incomplete"       # pagination did NOT finish; rows missing
    ERROR = "error"                 # extension's own internal error
    TIMEOUT = "timeout"             # exceeded CRAWL_MAX_SECONDS
    NEVER_STARTED = "never-started"

    #: outcomes whose collected rows are worth exporting AS COMPLETE.
    #  INCOMPLETE is deliberately absent: its rows are real and are saved as a
    #  .partial.xlsx, but the search must never be marked finished.
    EXPORTABLE = {FINISHED, END_OF_RESULTS, SINGLE_PAGE}


class NextControl:
    """What ensure_next_selector() found.  THREE answers, not a bool.

    A single bool conflated three genuinely different situations, and the
    caller could only report the worst of them:

        PAGINATED    GitHub offers a next page and the extension can reach
                     it.  Crawl normally.
        SINGLE_PAGE  GitHub offers NO next control and the page really does
                     hold results.  There is nothing to paginate, so this is
                     a complete result set -- success, not failure.
        FAILED       either GitHub offers a next control the extension cannot
                     learn or use, or the page is not a usable results page
                     at all.  Preserve everything and stop.

    Telling SINGLE_PAGE from FAILED is the whole point: both leave the
    extension without a #startScraping button, but one is a finished job and
    the other is a broken one.
    """

    PAGINATED = "paginated"
    SINGLE_PAGE = "single-page"
    FAILED = "failed"


class PageState:
    """What the GitHub Users search page in front of us actually IS.

    Zero user results and a broken page both render zero result elements, and
    conflating them cost a whole rank: a legitimate zero-result query was
    reported as a collection failure, preserved as unresolved, and retried
    forever on every run.

        RESULTS_MULTI_PAGE   users present, a next page exists
        RESULTS_SINGLE_PAGE  users present, no next page
        ZERO_RESULTS         valid Users search, GitHub says no users
        LOGIN_REQUIRED       an authentication wall
        HTTP_ERROR           429/403/5xx rendered as a page
        BLOCKED              throttle / abuse / challenge wording
        GITHUB_ERROR         GitHub's own error page ("unicorn", 500 ...)
        UNKNOWN_INVALID      not recognisable as a Users search at all

    Only the two RESULTS_* states collect.  ZERO_RESULTS COMPLETES the search
    unit with no candidates.  Everything else fails and preserves.
    """

    RESULTS_MULTI_PAGE = "results-multi-page"
    RESULTS_SINGLE_PAGE = "results-single-page"
    ZERO_RESULTS = "zero-results"
    LOGIN_REQUIRED = "login-required"
    HTTP_ERROR = "http-error"
    BLOCKED = "blocked"
    GITHUB_ERROR = "github-error"
    UNKNOWN_INVALID = "unknown-invalid"

    #: states that mean "go and collect"
    COLLECTABLE = {RESULTS_MULTI_PAGE, RESULTS_SINGLE_PAGE}
    #: the one state that completes a search unit without collecting
    COMPLETE_EMPTY = {ZERO_RESULTS}
    #: everything else preserves and stops
    FAILURES = {LOGIN_REQUIRED, HTTP_ERROR, BLOCKED, GITHUB_ERROR,
                UNKNOWN_INVALID}


# Minimum non-chrome profile links before a page counts as holding results.
# Three, because GitHub's own chrome is already excluded and a real results
# page shows ten per page; one stray link must never look like a result set.
MIN_RESULT_LINKS = 3


def _http_error_code(state):
    """An HTTP status this page is really reporting, or None.

    The SPA renders "Error 429" as a title and "Status: 429" in the body
    instead of failing the navigation, so the code has to be read off the
    rendered page.
    """
    title = str(state.get("title") or "")
    body = "%s %s" % (state.get("bodyHead") or "", state.get("bodyTail") or "")
    match = re.match(C.BLOCK_TITLE_RE, title, re.I)
    if match:
        return int(match.group(1))
    match = re.search(C.BLOCK_STATUS_RE, body, re.I)
    if match:
        return int(match.group(1))
    return None


def classify_search_page(state):
    """THE page classifier.  Returns (PageState, human explanation).

    ORDER MATTERS, and this is the order:

        1  login / authentication wall
        2  HTTP error rendered as a page
        3  blocked / throttled / challenged
        4  GitHub's own error page
        5  is this even a Users search page?
        6  are there user results?          -> RESULTS_*
        7  does GitHub explicitly say none? -> ZERO_RESULTS
        8  otherwise                        -> UNKNOWN_INVALID

    Steps 1-4 come first because every one of them also renders zero user
    results.  Asking "how many users?" before asking "is this page real?" is
    exactly how a rate-limit page became "this search has no users".

    Step 6 before step 7 matters too: if users are visibly present, no amount
    of empty-state text may override them.
    """
    if state.get("page_error"):
        return (PageState.UNKNOWN_INVALID,
                "the page could not be inspected: %s" % state["page_error"])

    # -- 1. login ----------------------------------------------------------
    if state.get("loginForm") or state.get("loginCopy"):
        return (PageState.LOGIN_REQUIRED,
                "GitHub is asking for a sign-in")

    # -- 2. HTTP error -----------------------------------------------------
    code = _http_error_code(state)
    if code is not None and code in C.BLOCK_STATUS:
        return (PageState.HTTP_ERROR, "GitHub returned HTTP %d" % code)
    if code is not None and code >= 400:
        return (PageState.HTTP_ERROR, "GitHub returned HTTP %d" % code)

    # -- 3. blocked / throttled / challenged -------------------------------
    marker = state.get("blocked_marker")
    if marker:
        return (PageState.BLOCKED, "GitHub says: %s" % marker)

    # -- 4. GitHub's own error page ----------------------------------------
    if state.get("saysError"):
        return (PageState.GITHUB_ERROR, "GitHub rendered an error page")

    # -- 5. is this a Users search at all? ---------------------------------
    # isUsersSearch is absent from older captures, so its absence is not
    # treated as a negative -- only an explicit False is.
    if state.get("isUsersSearch") is False:
        return (PageState.UNKNOWN_INVALID,
                "not a GitHub Users search page (%s)"
                % (state.get("url") or "unknown url"))

    # -- 6. are there results? --------------------------------------------
    # Non-chrome profile links when the capture has them, falling back to the
    # older field so a state captured by an earlier build still classifies.
    result_links = state.get("resultProfileLinkCount")
    if result_links is None:
        result_links = state.get("profileLinkCount")
    result_links = int(result_links or 0)
    listed = int(state.get("resultsListChildren") or 0)
    has_results = listed > 0 or result_links >= MIN_RESULT_LINKS

    if has_results:
        has_next = (bool(state.get("hasNext"))
                    or int(state.get("nextByShape") or 0) > 0
                    or int(state.get("nextByText") or 0) > 0)
        if has_next and state.get("nextDisabled") is not True:
            return (PageState.RESULTS_MULTI_PAGE,
                    "%d result(s) and a next page" % max(listed, result_links))
        return (PageState.RESULTS_SINGLE_PAGE,
                "%d result(s), no next page" % max(listed, result_links))

    # -- 7. does GitHub explicitly say there are none? ---------------------
    if state.get("noResults"):
        return (PageState.ZERO_RESULTS,
                "GitHub reports no users for this search")

    # -- 8. no results, and no statement that there are none ---------------
    return (PageState.UNKNOWN_INVALID,
            "no user results and no empty-state message; the page did not "
            "render as expected")


def classify_pagination(state):
    """Does this page have a next page, no next page, or is it not usable?

    Answered from GITHUB'S OWN DOM, never from the extension failing to do
    something.  That distinction is the fix: the extension has no
    #startScraping button both when there is nothing to paginate and when it
    could not learn a Next control that exists, and those must not be the
    same answer.

    Returns (verdict, why) where verdict is a NextControl value.  PAGINATED
    here means only "a next control exists", not "the extension can use it".
    """
    # ONE authority on what this page is.  classify_pagination() used to
    # repeat the blocked/error/empty checks itself, which is how it came to
    # disagree with collect() about the same page.
    page, why = classify_search_page(state)
    if page == PageState.RESULTS_MULTI_PAGE:
        return (NextControl.PAGINATED, why)
    if page == PageState.RESULTS_SINGLE_PAGE:
        return (NextControl.SINGLE_PAGE, why)
    # ZERO_RESULTS never reaches here: collect() classifies the page before
    # the extension is opened, so a zero-result search is completed without
    # any pagination question being asked at all.
    return (NextControl.FAILED, why)


def classify_page_state(state):
    """Why did the extension stop?  Four distinct answers.

    "Table not found" is a symptom, not a diagnosis: the same message
    appears when GitHub returns an error page, when the results genuinely
    ran out, and when a re-render momentarily emptied the list.  Treating
    them alike either throws away a complete crawl or hammers a throttled
    endpoint.

    Returns (outcome, human explanation).
    """
    # An uninspectable page tells us NOTHING, and must never be read as an
    # answer.  Without this, a dead renderer gave hasNext=None and
    # resultsListChildren=None, which fell through to "last page reached" --
    # a second, independent way to silently truncate a crawl.
    if state.get("page_error"):
        return (CrawlOutcome.TABLE_LOST,
                "the page could not be inspected (%s)" % state["page_error"])

    title = str(state.get("title") or "")
    head = str(state.get("bodyHead") or "")
    tail = str(state.get("bodyTail") or "")
    body = ("%s %s" % (head, tail)).lower()

    # 2. GitHub error / rate limit / challenge -- checked FIRST, because an
    #    error page also has no results and no Next link, and would
    #    otherwise look exactly like the end of the results.
    title_match = re.match(C.BLOCK_TITLE_RE, title, re.I)
    status_match = re.search(C.BLOCK_STATUS_RE, body, re.I)
    code = None
    if title_match:
        code = int(title_match.group(1))
    elif status_match:
        code = int(status_match.group(1))
    if code in C.BLOCK_STATUS:
        return (CrawlOutcome.RATE_LIMITED,
                "GitHub returned HTTP %d" % code)
    marker = state.get("blocked_marker")
    if marker:
        return (CrawlOutcome.RATE_LIMITED,
                "GitHub says: %s" % marker)
    if code and code >= 400:
        return (CrawlOutcome.RATE_LIMITED,
                "GitHub returned HTTP %d" % code)

    # 1. genuine end of results
    if state.get("noResults"):
        return (CrawlOutcome.END_OF_RESULTS, "GitHub reported no more users")
    has_next = state.get("hasNext")
    next_dead = state.get("nextDisabled") is True
    no_results_shown = (not state.get("resultsListChildren")
                        and int(state.get("profileLinkCount") or 0) < 3)
    if (not has_next or next_dead) and no_results_shown:
        return (CrawlOutcome.END_OF_RESULTS, "last page reached")
    if (not has_next or next_dead) and state.get("resultsListChildren"):
        return (CrawlOutcome.END_OF_RESULTS,
                "last page reached (results present, no Next)")

    # 3. a perfectly good Users page the extension stopped seeing
    if state.get("resultsListPresent") or \
            int(state.get("profileLinkCount") or 0) >= 5:
        return (CrawlOutcome.TABLE_LOST,
                "valid results page; the extension lost table detection")

    # 4. anything else is the extension's own problem
    return (CrawlOutcome.ERROR,
            "extension error: %s" % (state.get("ext_error") or "unknown"))


class InstantDataScraper:
    def __init__(self, browser, state, say=print):
        #: the most recent CrawlProgress, for troubleshooting and for tests
        self.progress = None
        #: set by export_xlsx when the last export was rescued from disk
        self.last_export_recovered = False
        self.browser = browser
        self.state = state
        self.say = say
        self.last_state = None

    # ------------------------------------------------------------------ UI
    async def open_instant_data_scraper(self, page):
        """Return the extension's popup Page, opened for `page`'s tab."""
        await page.bring_to_front()

        # 1. the configured shortcut, exactly as done by hand.
        before = set(self.browser.context.pages)
        try:
            await page.keyboard.press("Control+G")
        except Exception as exc:
            log.debug("Ctrl+G could not be delivered: %s", exc)
        popup = await self._await_popup(before, seconds=3)
        if popup is not None:
            self.say("[SCRAPER] Instant Data Scraper opened (Ctrl+G)")
            await self._prepare_popup(popup)
            return popup

        # 2. the same window background.js would have created.
        tab_id = await self.browser.chrome_tab_id(page)
        url = "chrome-extension://%s/%s?%s" % (
            self.browser.extension_id, C.EXTENSION_POPUP_PATH,
            urllib.parse.urlencode({"tabid": str(tab_id), "url": page.url}))
        popup = await self.browser.context.new_page()
        # The extension closes its own window as soon as a download starts.
        # If that happens before save_as() runs Playwright loses the page
        # that owns the download, so window.close is neutralised and the page
        # is closed explicitly instead.
        await popup.add_init_script("window.close = function(){};")
        await popup.goto(url, wait_until="domcontentloaded",
                         timeout=C.EXTENSION_TIMEOUT)
        self.say("[SCRAPER] Instant Data Scraper opened (extension page)")
        await self._prepare_popup(popup)
        return popup

    async def _await_popup(self, before, seconds):
        marker = "chrome-extension://%s/%s" % (self.browser.extension_id,
                                               C.EXTENSION_POPUP_PATH)
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            for candidate in self.browser.context.pages:
                if candidate in before:
                    continue
                if candidate.url.startswith(marker):
                    await candidate.add_init_script(
                        "window.close = function(){};")
                    return candidate
            await asyncio.sleep(0.25)
        return None

    async def _prepare_popup(self, popup, need_table=True):
        """Wait for the React UI, dismiss the opt-in banner, set delays.

        Waiting for the table matters after a reload: the popup rebuilds its
        state from scratch and only re-detects rows once the content script
        answers.  Pressing Start before that yields "Table not found".
        """
        await popup.wait_for_selector("#root *", timeout=C.EXTENSION_TIMEOUT)
        # The privacy opt-in banner is optional and must not be accepted for
        # the user; it is only dismissed if a close control is present.
        for selector in ('#optInModalTitle ~ * button[aria-label="Close"]',
                         '.modal button.close'):
            try:
                node = popup.locator(selector)
                if await node.count() and await node.first.is_visible():
                    await node.first.click(timeout=2000)
            except Exception:
                pass
        await self._set_delays(popup)
        if need_table:
            await self.wait_for_table(popup)

    async def _set_delays(self, popup):
        for selector, value in ((C.SEL_CRAWL_DELAY, C.CRAWL_MIN_DELAY_SEC),
                                (C.SEL_MAX_WAIT, C.CRAWL_MAX_DELAY_SEC)):
            try:
                node = popup.locator(selector)
                if await node.count():
                    await node.first.fill(str(value), timeout=3000)
            except Exception as exc:
                log.debug("could not set %s: %s", selector, exc)

    # -------------------------------------------------------- pagination --
    async def ensure_next_selector(self, popup, page):
        """Decide how this result set can be collected.  Returns NextControl.

        THREE outcomes, because the caller has three different jobs to do:

            PAGINATED    the extension can reach page 2 -- crawl
            SINGLE_PAGE  GitHub served one page -- export it, do not crawl
            FAILED       preserve everything and stop

        This used to return a bool, and "no Next control" returned False, so
        a perfectly good one-page result set was reported as "the extension
        has no way to reach page 2" and the whole search failed.  The
        diagnosis was right; the verdict was wrong.

        The extension's own mechanism is preserved exactly: popup.js reads
        localStorage["nextSelector:" + hostname] during init and asks the
        content script to mark that element, enabling #startScraping.  Only
        the VALUE is improved.

        Priority (spec item 3):
          1. a[rel="next"]                      -- semantic, one match
          2. a stable accessible selector       -- aria-label / .next_page
          3. the selector remembered in state   -- may be a learned CSS path
          4. re-learn through the extension's own "Locate Next" flow

        A candidate is only used once it has been confirmed to match exactly
        one element on the live page AND the extension has accepted it.
        """
        host = urllib.parse.urlparse(page.url).hostname or "github.com"
        key = C.NEXT_SELECTOR_STORAGE_KEY % host

        candidates = []
        for selector in C.STABLE_NEXT_SELECTORS:
            if await self._matches_one(page, selector):
                candidates.append((selector, "stable"))
        remembered = self.state.next_selector_for(host)
        if remembered and remembered not in [c for c, _ in candidates]:
            candidates.append((remembered, "remembered"))

        for selector, kind in candidates:
            if await self._apply_next_selector(popup, key, selector):
                self.state.remember_next_selector(host, selector)
                self.say("[SCRAPER] pagination selector (%s): %s"
                         % (kind, _short(selector)))
                return NextControl.PAGINATED

        # Already usable without any of them (the extension kept its own).
        if await self._has(popup, C.SEL_START_CRAWL):
            return NextControl.PAGINATED

        # No selector we know of works.  Before blaming the extension, ask
        # GITHUB whether there is anything to paginate at all.
        state = await self.capture_state(page, popup, reason="pagination check")
        self.last_state = state
        verdict, why = classify_pagination(state)

        if verdict == NextControl.SINGLE_PAGE:
            self.say('[SCRAPER] GitHub showed no "Next" control '
                     '(single page of results)')
            self.say("[SCRAPER] single-page mode: %s" % why)
            return NextControl.SINGLE_PAGE

        if verdict == NextControl.FAILED:
            self.say("[SCRAPER] this is not a usable results page: %s" % why)
            self.log_state(state)
            return NextControl.FAILED

        # GitHub DOES offer a next page, so the extension must be taught it.
        # Run the extension's own Locate flow once: clicking #nextButton puts
        # the content script into pick-an-element mode, and clicking GitHub's
        # Next control then teaches it.
        if not await self._has(popup, C.SEL_NEXT_BUTTON):
            self.say("[SCRAPER] GitHub offers a next page but the extension "
                     "has no Locate control to learn it with")
            return NextControl.FAILED

        self.say('[SCRAPER] learning the "Next" control through the '
                 'extension')
        await popup.click(C.SEL_NEXT_BUTTON, timeout=C.EXTENSION_TIMEOUT)
        await page.bring_to_front()
        clicked = False
        for selector in NEXT_LINK_SELECTORS:
            node = page.locator(selector)
            try:
                if await node.count() and await node.first.is_visible():
                    await node.first.click(timeout=5000)
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            # classify_pagination() saw a next control a moment ago and now
            # it cannot be clicked.  That is the extension or the page
            # changing under us -- NOT a single page.  Guessing "single page"
            # here would export one page of a many-page result set and then
            # delete the search as finished.
            self.say("[SCRAPER] GitHub offers a next page but it could not "
                     "be clicked to teach the extension")
            return NextControl.FAILED

        await popup.bring_to_front()
        try:
            await popup.wait_for_selector(C.SEL_START_CRAWL,
                                          timeout=C.EXTENSION_TIMEOUT)
        except Exception:
            self.say("[SCRAPER] the extension did not accept a Next control")
            return NextControl.FAILED
        learned = await popup.evaluate("k => localStorage.getItem(k) || ''",
                                       key)
        if learned:
            self.state.remember_next_selector(host, learned)
            self.say("[SCRAPER] pagination selector learned (extension path)")
        return NextControl.PAGINATED

    async def wait_for_stable_results(self, page, timeout=None,
                                      stable_polls=3, poll=0.5):
        """Block until GitHub's results have stopped changing.

        The diagnostic proved GitHub REPLACES the results DOM when it
        paginates, so the extension must not be opened while that is in
        flight -- it would bind to nodes that are about to be discarded.

        Stable means, for `stable_polls` consecutive polls: readyState
        complete, at least one result on the page, no loading/skeleton node,
        and an unchanged (url, result count, content length).

        Always reads the LIVE page handle.  Returns (ok, state).
        """
        limit = (timeout or C.NAVIGATION_TIMEOUT) / 1000.0
        waited = 0.0
        signature = None
        steady = 0
        state = {}
        while waited < limit:
            state = await self.capture_page_state(page, reason="stability")
            ready = state.get("readyState") in (None, "complete",
                                               "interactive")
            results = (int(state.get("resultsListChildren") or 0) > 0
                       or int(state.get("resultProfileLinkCount") or 0) > 0)
            quiet = int(state.get("skeletonCount") or 0) == 0
            current = (state.get("url"),
                       int(state.get("resultsListChildren") or 0),
                       int(state.get("resultProfileLinkCount") or 0),
                       int(state.get("domSignature") or 0))
            if ready and results and quiet and current == signature:
                steady += 1
                if steady >= stable_polls:
                    return True, state
            else:
                steady = 0
                signature = current
            await asyncio.sleep(poll)
            waited += poll
        log.warning("results did not stabilise within %.0fs", limit)
        return False, state

    async def wait_for_table(self, popup, timeout=None):
        """Block until the extension reports at least one collected row."""
        limit = (timeout or C.EXTENSION_TIMEOUT) / 1000.0
        waited = 0.0
        while waited < limit:
            _pages, rows = await self.read_stats(popup)
            if rows and rows > 0:
                return True
            error = await self._text(popup, C.SEL_ERROR)
            if "table not found" in error.lower():
                # Ask the extension to re-scan rather than giving up.
                if await self._has(popup, C.SEL_WRONG_TABLE):
                    try:
                        await popup.click(C.SEL_WRONG_TABLE, timeout=3000)
                    except Exception:
                        pass
            await asyncio.sleep(0.5)
            waited += 0.5
        log.warning("extension reported no table within %.0fs", limit)
        return False

    @staticmethod
    async def _matches_one(page, selector):
        """True when `selector` identifies exactly one visible element.

        Ambiguity matters: the class the extension learns for GitHub
        matches the Previous link too.
        """
        try:
            node = page.locator(selector)
            if await node.count() != 1:
                return False
            return await node.first.is_visible()
        except Exception:
            return False

    async def _apply_next_selector(self, popup, key, selector):
        """Seed the extension's own store, reload, and see if it took.

        popup.js reads the selector only during init, so the popup has to be
        reloaded for a new value to be picked up.
        """
        try:
            current = await popup.evaluate(
                "k => localStorage.getItem(k) || ''", key)
            already_ready = await self._has(popup, C.SEL_START_CRAWL)
            if current == selector and already_ready:
                # Nothing to change: avoid a reload, which would make the
                # extension rebuild its table state for no reason.
                return True
            await popup.evaluate("([k, v]) => localStorage.setItem(k, v)",
                                 [key, selector])
            await popup.reload(wait_until="domcontentloaded",
                               timeout=C.EXTENSION_TIMEOUT)
            await self._prepare_popup(popup)
        except Exception as exc:
            log.warning("could not seed selector %s: %s", selector, exc)
            return False
        return await self._has(popup, C.SEL_START_CRAWL)

    # ------------------------------------------------------------ crawling
    async def start_crawling(self, popup):
        if not await self._has(popup, C.SEL_START_CRAWL):
            return False
        if not await self.wait_for_table(popup):
            self.say("[SCRAPER] the extension found no table to collect")
            return False
        await popup.click(C.SEL_START_CRAWL, timeout=C.EXTENSION_TIMEOUT)
        self.say("[SCRAPER] Start Crawling clicked")
        return True

    # --------------------------------------------- correct-table check ----
    MAX_TABLE_CANDIDATES = 8

    @staticmethod
    def dom_result_identities(state):
        """The normalised GitHub identities GitHub's own results DOM shows.

        The ground truth a page's IDS table must reproduce.  Built from
        resultProfileUsers, which PAGE_STATE_JS collects from the results
        region with the page chrome excluded.
        """
        _identity = sys.modules[__name__]
        users = state.get("resultProfileUsers") or []
        ids = set()
        for name in users:
            key = _identity.normalize_username(name)
            if key:
                ids.add(key)
        return ids

    @staticmethod
    def dom_usernames(state):
        """The valid, unique raw github usernames the results DOM shows.

        Deduped by normalised identity but returned in raw/original case so the
        profile phase navigates the real handles; identity comparison still
        normalises.  The order follows resultProfileUsers.  Used by the exact
        single-page DOM fallback, where IDS has too few rows to infer a table.
        """
        _identity = sys.modules[__name__]
        _xr = sys.modules[__name__]
        users = state.get("resultProfileUsers") or []
        seen = set()
        valid = []
        for name in users:
            if not name or not _xr.valid_username(name):
                continue
            key = _identity.normalize_username(name)
            if key and key not in seen:
                seen.add(key)
                valid.append(name)
        return valid

    @staticmethod
    def write_dom_candidates_xlsx(destination, usernames):
        """Write a keyword XLSX with ONE verified DOM identity per row.

        Used only for a small exact single-page result set that IDS cannot
        infer a repeating table for.  Each row is just a user's profile URL, in
        a column XlsxReader already parses to the identity -- it is NOT scraped
        profile data (the profile phase collects every detail normally).
        Reuses write_preview_xlsx so the serialization is identical to the
        verified path.  Returns the number of rows written.
        """
        rows = [["https://github.com/%s" % name] for name in usernames]
        return InstantDataScraper.write_preview_xlsx(
            destination, ["prc-Link-Link href"], rows)

    @staticmethod
    def _identities_from_cells(cells):
        """Normalised GitHub identities found in preview cell texts.

        Each cell is tested on its own with the SAME parser the export uses
        (avatar URL, profile URL, repo owner), so a profile-link cell yields
        its user and a page-chrome cell ("Skip to content") yields nothing --
        which is exactly how a wrong table is spotted.
        """
        _identity = sys.modules[__name__]
        _xr = sys.modules[__name__]
        ids = set()
        for cell in cells:
            user = _xr.XlsxReader._username_from([cell])
            if user and _xr.valid_username(user):
                key = _identity.normalize_username(user)
                if key:
                    ids.add(key)
        return ids

    async def _read_preview(self, popup):
        """(identities, signature) from the flat cell list -- LEGACY helper.

        Kept only for the preview-parser unit tests.  The establishment,
        stability, verification and WRITE paths do NOT use this: a flat union
        of every td/th across Handsontable's overlay clones can contain the
        expected identities even when no single concrete grid does -- the exact
        bug that let a 4/4 "match" write a 1-row file.  Those paths use
        read_preview_grids(), which never unions cells across grids.
        """
        try:
            data = await popup.evaluate(IDS_PREVIEW_JS)
        except Exception as exc:
            log.warning("could not read the IDS preview: %s", exc)
            return set(), "unreadable"
        cells = data.get("cells", []) if isinstance(data, dict) else []
        ids = self._identities_from_cells(cells)
        signature = hashlib.sha1(
            ("\u0001".join(cells)).encode("utf-8", "replace")).hexdigest()
        return ids, signature

    # -------------------------------------- ONE SOURCE OF TRUTH: grids ----
    # Every identity decision -- establishment, stability, final pre-write
    # verification and the WRITE itself -- comes from the SAME single concrete
    # grid's rows.  read_preview_grids() enumerates each concrete Handsontable
    # grid separately and NEVER unions cells across overlay clones or stale
    # grids, so a table is "correct" only when ONE grid's own rows reproduce
    # the exact expected identity set.
    async def read_preview_grids(self, popup):
        """Every concrete preview grid: [{headers, rows, identities, signature}].

        identities are parsed ONLY from that one grid's rows; signature is a
        content hash of its rows (so a row-level shift is detectable even when
        the identity set is unchanged).  Grids are never merged.
        """
        try:
            data = await popup.evaluate(IDS_PREVIEW_JS)
        except Exception as exc:
            log.warning("could not read the IDS preview grids: %s", exc)
            return []
        if not isinstance(data, dict):
            return []
        grids = []
        raw = data.get("grids")
        if raw:
            for entry in raw:
                if not isinstance(entry, dict):
                    continue
                headers = [("" if h is None else str(h))
                           for h in (entry.get("headers") or [])]
                rows = [[("" if c is None else str(c)) for c in (row or [])]
                        for row in (entry.get("rows") or [])]
                grids.append(self._grid_record(headers, rows))
        else:
            # Legacy popups (older fakes) return only a flat cell list; treat
            # it as ONE grid so parser-level tests still work.  Production
            # always returns `grids`.
            cells = [("" if c is None else str(c))
                     for c in (data.get("cells") or [])]
            if cells:
                grids.append(self._grid_record([], [list(cells)]))
        return grids

    @staticmethod
    def _grid_record(headers, rows):
        """One grid: {headers, rows, identities (from THESE rows), signature}."""
        flat = [cell for row in rows for cell in row]
        blob = "\n".join("\t".join(row) for row in rows)
        signature = hashlib.sha1(blob.encode("utf-8", "replace")).hexdigest()
        return {"headers": headers, "rows": rows,
                "identities": InstantDataScraper._identities_from_cells(flat),
                "signature": signature}

    @staticmethod
    def _select_grid(grids, expected_ids):
        """The single grid whose OWN rows' identities == expected, else None.

        Identical identities across two different grids never combine: a match
        is one grid alone reproducing the exact expected set.
        """
        for grid in grids:
            if grid["identities"] == expected_ids:
                return grid
        return None

    @staticmethod
    def _grids_signature(grids):
        """A change-detector over ALL grids (only to tell whether #wrongTable
        changed anything -- never an identity source)."""
        return hashlib.sha1(
            ("|".join(g["signature"] for g in grids))
            .encode("utf-8", "replace")).hexdigest()

    async def _selected_grid(self, popup, expected_ids):
        """(matching_grid_or_None, all_grids_signature, grids)."""
        grids = await self.read_preview_grids(popup)
        return (self._select_grid(grids, expected_ids),
                self._grids_signature(grids), grids)

    async def _await_grid_change(self, popup, previous_signature,
                                 timeout=15.0, poll=0.5):
        waited = 0.0
        while waited < timeout:
            grids = await self.read_preview_grids(popup)
            if self._grids_signature(grids) != previous_signature:
                return True
            await asyncio.sleep(poll)
            waited += poll
        return False

    async def preview_identities(self, popup, expected_ids=None):
        """Row-derived identities of a SINGLE concrete grid -- never a union.

        With expected_ids: the identities of the grid whose OWN rows match them
        exactly, or the empty set when no single grid matches.  Without: the
        identities of the largest single grid (best effort, still no union).
        """
        grids = await self.read_preview_grids(popup)
        if expected_ids is not None:
            grid = self._select_grid(grids, expected_ids)
            return grid["identities"] if grid is not None else set()
        if not grids:
            return set()
        return max(grids, key=lambda g: len(g["identities"]))["identities"]

    @staticmethod
    def export_identities(path):
        """Normalised GitHub identities in an exported XLSX.

        Uses the same parser downstream uses (XlsxReader), so what this counts
        is exactly what will become candidates.
        """
        _identity = sys.modules[__name__]
        _xr = sys.modules[__name__]
        ids = set()
        try:
            for candidate in _xr.XlsxReader().read(path, ""):
                key = _identity.normalize_username(candidate.username)
                if key:
                    ids.add(key)
        except Exception as exc:
            log.warning("could not read export identities from %s: %s",
                        path, exc)
        return ids

    # ----------------------------------- verified-preview page export -----
    # The extension's #xlsx button is NOT used to collect a GitHub Users
    # results page: it can export a table that diverged from the validated
    # preview (the Engineering p2 case -- preview 4/4, XLSX 1 user + chrome),
    # and clicking it can kill Chromium in this environment.  Instead the
    # VERIFIED preview grid (already stabilised and matched to the DOM) is the
    # export source, and Python writes the page XLSX directly.

    @staticmethod
    def _rows_one_identity_each(rows):
        """Reshape grid rows so EACH output row carries exactly one github
        identity -- the serialization bridge that makes the round-trip
        invariant.

        The grid's identities are found PER CELL (a cell that parses to a
        github user), but XlsxReader parses ONE identity per ROW and dedupes by
        username.  A grid row can pack several users' links (Handsontable can
        render a wide/merged cell) or none, so writing rows verbatim can
        collapse four users to one on read -- the live p2 bug.  Here every
        distinct identity present in the grid becomes its OWN row, carrying its
        own url cell(s) with OTHER users' url cells blanked (so XlsxReader
        reads THIS identity) and all non-url text kept for enrichment.  A row
        with no identity (page chrome) is dropped, exactly as XlsxReader would
        ignore it.  XlsxReader's own identity logic is untouched.
        """
        _identity = sys.modules[__name__]
        _xr = sys.modules[__name__]
        out = []
        seen = set()
        for row in rows:
            present = []                       # (key, raw_username) first-seen
            for cell in row:
                user = _xr.XlsxReader._username_from([cell])
                if user and _xr.valid_username(user):
                    key = _identity.normalize_username(user)
                    if key and key not in {k for k, _u in present}:
                        present.append((key, user))
            for key, _user in present:
                if key in seen:
                    continue                   # same identity already written
                seen.add(key)
                isolated = []
                for cell in row:
                    other = _xr.XlsxReader._username_from([cell])
                    if (other and _xr.valid_username(other)
                            and _identity.normalize_username(other) != key):
                        isolated.append("")    # blank the OTHER users' urls
                    else:
                        isolated.append(cell)
                out.append(isolated)
        return out

    @staticmethod
    def write_preview_xlsx(destination, headers, rows):
        """Write the verified grid to `destination` so XlsxReader recovers
        EXACTLY its identities.

        The header row is the extension's column names (or synthesised names
        when none came back, so XlsxReader never mistakes the first data row
        for the header).  The data rows are reshaped to one github identity per
        row (see _rows_one_identity_each), which keeps the clean one-user-per-
        row case byte-identical while making a packed row round-trip instead of
        collapsing.  Returns the number of DATA rows written.
        """
        from openpyxl import Workbook
        data_rows = InstantDataScraper._rows_one_identity_each(rows)
        width = max([len(headers)] + [len(r) for r in data_rows], default=0)
        if width == 0:
            width = 1
        if headers:
            header_row = list(headers) + [
                "col_%d" % i for i in range(len(headers), width)]
        else:
            # No column names came back (a bare grid); synthesise a header so
            # XlsxReader does not mistake the first data row for the header.
            header_row = ["col_%d" % i for i in range(width)]
        book = Workbook()
        sheet = book.active
        sheet.append(header_row)
        for row in data_rows:
            sheet.append(list(row) + [""] * (width - len(row)))
        folder = os.path.dirname(os.path.abspath(destination))
        if folder:
            os.makedirs(folder, exist_ok=True)
        book.save(destination)
        book.close()
        return len(data_rows)

    async def export_page_from_preview(self, popup, destination, expected_ids):
        """Export ONE GitHub Users page by writing ONE verified grid's rows.

        Never clicks #xlsx.  It selects the single concrete grid whose OWN rows
        reproduce the DOM identities (`expected_ids`) exactly and writes THAT
        grid's headers+rows -- the same rows whose identities were verified.
        If no single grid matches it writes nothing and returns -1, and the
        caller keeps everything and does not advance.  The caller still
        verifies the written file's identities after (PROTECTION 2).  Returns
        the number of rows written, or -1.
        """
        if not expected_ids:
            self.say("[EXPORT] no DOM identities to match; refusing to write")
            return -1
        grid, _all_sig, grids = await self._selected_grid(popup, expected_ids)
        if grid is None:
            best = max((len(g["identities"]) for g in grids), default=0)
            self.say("[EXPORT] no single structured grid reproduces the DOM "
                     "(best grid has %d of %d expected across %d grid(s)); "
                     "refusing to write"
                     % (best, len(expected_ids), len(grids)))
            return -1
        written = self.write_preview_xlsx(destination, grid["headers"],
                                          grid["rows"])
        self.say("[EXPORT] wrote %d row(s) from the verified structured grid "
                 "(no #xlsx click) -> %s"
                 % (written, os.path.basename(destination)))
        return written

    async def establish_results_table(self, popup, page, expected_ids):
        """Ensure ONE concrete grid's rows match the DOM before any export.

        The extension sometimes auto-selects the wrong table on a small
        results page -- yielding one real user plus page-chrome rows.  A table
        is correct ONLY when a single concrete grid's OWN rows reproduce the
        exact expected identity set: not when a flat union of cells across
        Handsontable's overlay clones happens to contain them.  'Try another
        table' (#wrongTable) is cycled until such a grid appears.

        Bounded: stops when the grids stop changing, when #wrongTable is gone,
        or after MAX_TABLE_CANDIDATES.  Returns True only on an exact
        single-grid match; on failure the caller preserves everything.
        """
        if not expected_ids:
            # Nothing to match against (a zero-result page never reaches here);
            # refuse rather than accept an unvalidated table.
            self.say("[TABLE] no DOM identities to validate against")
            return False
        seen = set()
        for attempt in range(1, self.MAX_TABLE_CANDIDATES + 1):
            grid, all_sig, grids = await self._selected_grid(popup, expected_ids)
            best = max((len(g["identities"]) for g in grids), default=0)
            self.say("[TABLE] candidate %d: best structured grid has %d usable "
                     "identity(ies) of %d expected (%d grid(s))"
                     % (attempt, best, len(expected_ids), len(grids)))
            if grid is not None:
                self.say("[TABLE] correct: one structured grid's rows exactly "
                         "match the %d DOM user(s)" % len(expected_ids))
                return True
            if all_sig in seen:
                self.say("[TABLE] tables cycled without an exact structured "
                         "match")
                return False
            seen.add(all_sig)
            if not await self._has(popup, C.SEL_WRONG_TABLE):
                self.say("[TABLE] 'Try another table' is unavailable; cannot "
                         "correct the selection")
                return False
            try:
                await popup.click(C.SEL_WRONG_TABLE, timeout=5000)
            except Exception as exc:
                log.warning("could not click 'Try another table': %s", exc)
                return False
            if not await self._await_grid_change(popup, all_sig):
                self.say("[TABLE] the preview did not change after 'Try "
                         "another table'")
                return False
        self.say("[TABLE] exhausted %d candidate table(s) without an exact "
                 "structured match" % self.MAX_TABLE_CANDIDATES)
        return False

    # ------------------------------------------- table STABILITY wait -----
    # After 'Try another table' the extension keeps mutating its model
    # (mn.data) asynchronously -- an async getTableData round-trip plus a
    # possible server-processing step.  A table that reads correctly at one
    # instant can therefore be replaced a fraction of a second later, and the
    # XLSX button exports whatever mn.data holds at click time.  So a correct
    # read is not enough: the preview must HOLD the exact expected identities
    # for several consecutive reads before we trust it.
    STABILITY_READS = 3          # consecutive matching reads required
    STABILITY_POLL = 0.4         # seconds between reads (~300-500 ms)
    STABILITY_TIMEOUT = 5.0      # give up on stability after this long
    MAX_ESTABLISH_ATTEMPTS = 2   # establish+stabilise cycles per page

    @staticmethod
    def _identity_projection(grid):
        """The stable, IDENTITY-bearing view of a grid: its normalised github
        identity set as a sorted tuple.

        Non-identity cell text -- result timing like "(29 ms)", counters,
        button/state text -- is deliberately excluded, so it can change from
        read to read without being mistaken for the table shifting.  A removed
        or replaced identity DOES change this projection.
        """
        return tuple(sorted(grid["identities"])) if grid is not None else None

    async def _await_stable_preview(self, popup, expected_ids):
        """True once, for STABILITY_READS reads in a row, ONE structured grid's
        identity set EXACTLY equals expected AND its identity projection is
        unchanged.

        Stability is judged on IDENTITY only, not the full row text: GitHub/IDS
        rows carry dynamic non-identity cells (timing "(29 ms)", counters,
        button text) that change while the users stay the same, and those must
        NOT read as instability.  A removed or replaced identity resets the
        streak.  Bounded by a fixed read budget so it can't hang.
        """
        budget = int(self.STABILITY_TIMEOUT / self.STABILITY_POLL) + 1
        streak = 0
        last_projection = None
        for _ in range(max(budget, self.STABILITY_READS)):
            grid, _all_sig, _grids = await self._selected_grid(popup,
                                                               expected_ids)
            projection = self._identity_projection(grid)
            if grid is None:
                if streak:
                    self.say("[TABLE] no matching structured grid mid-"
                             "stability-check; resetting")
                streak = 0
                last_projection = None
            elif projection == last_projection:
                streak += 1
            else:
                # first matching read (or the identity set changed): (re)start
                streak = 1
                last_projection = projection
            if streak >= self.STABILITY_READS:
                self.say("[TABLE] identity projection stable: %d consecutive "
                         "reads of the %d expected user(s) (non-identity text "
                         "ignored)" % (streak, len(expected_ids)))
                return True
            await asyncio.sleep(self.STABILITY_POLL)
        self.say("[TABLE] no structured grid held the expected identities for "
                 "%d reads in a row" % self.STABILITY_READS)
        return False

    async def stabilize_results_table(self, popup, page, expected_ids):
        """Establish the correct table AND require it to stay put.

        The required sequence per attempt is: establish the correct table
        (identity-set match) -> wait until the preview holds those exact
        identities for STABILITY_READS consecutive reads -> one final
        pre-click re-read.  If any step fails the whole cycle repeats, up to
        MAX_ESTABLISH_ATTEMPTS.  Returns True only when a table was correct
        AND stable AND still correct at the final check; on failure the caller
        preserves everything and does not export.
        """
        if not expected_ids:
            self.say("[TABLE] no DOM identities to validate against")
            return False
        for attempt in range(1, self.MAX_ESTABLISH_ATTEMPTS + 1):
            if not await self.establish_results_table(popup, page,
                                                      expected_ids):
                self.say("[TABLE] could not establish the correct table "
                         "(attempt %d of %d)"
                         % (attempt, self.MAX_ESTABLISH_ATTEMPTS))
                continue
            if not await self._await_stable_preview(popup, expected_ids):
                self.say("[TABLE] table did not remain stable; re-establishing "
                         "(attempt %d of %d)"
                         % (attempt, self.MAX_ESTABLISH_ATTEMPTS))
                continue
            # Final pre-write confirmation: one more read the instant before
            # the caller exports.  If no single grid matches here, re-establish
            # within the bound rather than exporting a table that just changed.
            final_grid, _sig, _grids = await self._selected_grid(popup,
                                                                 expected_ids)
            if final_grid is not None:
                return True
            self.say("[TABLE] no matching structured grid at the final check; "
                     "re-establishing (attempt %d of %d)"
                     % (attempt, self.MAX_ESTABLISH_ATTEMPTS))
        self.say("[TABLE] IDS result table did not remain stable after %d "
                 "attempt(s)" % self.MAX_ESTABLISH_ATTEMPTS)
        return False

    async def finish_single_page(self, popup, page):
        """Complete a result set that has exactly one page.

        No crawl runs, because there is nothing to crawl: the extension has
        already detected the table on the page in front of it, and its XLSX
        button exports what it holds.  Start Crawling is not merely skipped
        but UNAVAILABLE -- the extension renders it only once a Next selector
        is marked -- which is why this path exists at all.

        Returns (outcome, pages, rows) so the caller treats it exactly like a
        finished crawl.
        """
        if not await self.wait_for_table(popup):
            self.say("[SCRAPER] the extension found no table on the single "
                     "page")
            state = await self.capture_state(page, popup,
                                             reason="single page, no table")
            self.log_state(state)
            self.last_state = state
            return CrawlOutcome.NEVER_STARTED, 0, 0

        _pages, rows = await self.read_stats(popup)
        if rows is None or rows <= 0:
            self.say("[SCRAPER] the single page reported no rows")
            return CrawlOutcome.NEVER_STARTED, 0, 0

        # One page, by definition: the extension never navigated anywhere.
        self.say("[SCRAPER] single page complete: %d row(s), no pagination "
                 "needed" % rows)
        return CrawlOutcome.SINGLE_PAGE, 1, rows

    async def read_stats(self, popup):
        """(pages, rows) from the extension's own #stats block."""
        try:
            text = await popup.inner_text(C.SEL_STATS, timeout=3000)
        except Exception:
            return (-1, -1)
        match = _STAT_RE.search(re.sub(r"\s+", " ", text))
        if not match:
            return (-1, -1)
        return int(match.group(1)), int(match.group(2))

    async def wait_for_crawl_complete(self, popup, page, max_pages=None,
                                      had_next=False):
        """Block until the extension genuinely stopped crawling.

        Never decided by a single signal (spec 14).  The crawl is finished
        only when ALL of these hold together:

          * #stopScraping is gone   -- popup.js renders it only while
                                       mn.scraping is true
          * #startScraping is back  -- Mr() re-enables it when a crawl ends
          * #instructions says "Crawling stopped"
          * pages AND rows unchanged for CRAWL_STABLE_SECONDS

        A frozen network looks like "no change" but never clears the
        crawling flag or writes that message, so a stall can never be
        mistaken for success.

        AND -- since the "University Student Visakhapatnam" loss -- a crawl
        that started with a live Next must also PROVE it paginated before any
        completion is accepted.  had_next carries that pre-crawl fact in;
        CrawlProgress checks it against what actually happened.  Without it,
        "the extension stopped and the page has no further Next" was accepted
        as success even when the extension had failed to scrape the page it
        was standing on.
        """
        started = time.monotonic()
        last_pages, last_rows = await self.read_stats(popup)
        last_change = time.monotonic()
        last_url = page.url
        ever_crawled = False
        recoveries = 0
        shown = set()

        progress = CrawlProgress(had_next=had_next, start_url=last_url)
        progress.note_stats(last_pages, last_rows)
        self.progress = progress
        self.say("[CRAWL] before start: page %d, rows %d, Next present: %s"
                 % (progress.start_page, max(last_rows, 0),
                    "yes" if had_next else "no"))

        while True:
            elapsed = time.monotonic() - started
            if elapsed > C.CRAWL_MAX_SECONDS:
                return CrawlOutcome.TIMEOUT, last_pages, last_rows

            error = await self._text(popup, C.SEL_ERROR)
            if error.strip():
                state = await self.capture_state(
                    page, popup, reason=error.strip()[:60])
                self.log_state(state)
                self.last_state = state
                outcome, why = classify_page_state(state)
                pages = state.get("ext_pages", last_pages)
                rows = state.get("ext_rows", last_rows)
                if pages and pages > 0:
                    last_pages, last_rows = pages, rows

                progress.note_url(state.get("url") or last_url)
                progress.note_stats(last_pages, last_rows)

                if outcome == CrawlOutcome.END_OF_RESULTS:
                    ok, complaint = progress.verify_complete()
                    if not ok:
                        # The page says "last page" and the extension has
                        # stopped -- but rows are provably missing.  Try the
                        # extension's own recovery before giving up, because
                        # it reports "continue crawling" in exactly this
                        # state and re-establishing the table can finish the
                        # job.
                        if recoveries < C.CRAWL_TABLE_RECOVERIES:
                            recoveries += 1
                            self.say("[CRAWL] INCOMPLETE: %s" % complaint)
                            self.say("[CRAWL] recovery %d/%d: re-establishing "
                                     "the table and continuing"
                                     % (recoveries,
                                        C.CRAWL_TABLE_RECOVERIES))
                            if await self._recover_table(popup, page):
                                last_change = time.monotonic()
                                continue
                        self.say("[CRAWL] INCOMPLETE: %s" % complaint)
                        self.say("[CRAWL] %s" % progress.describe())
                        self.say("[CRAWL] refusing to call this search "
                                 "complete; nothing will be deleted")
                        return (CrawlOutcome.INCOMPLETE, last_pages,
                                last_rows)
                    self.say("[SCRAPER] %s - %d row(s) collected"
                             % (why, last_rows))
                    self.say("[CRAWL] final: %s" % progress.describe())
                    return outcome, last_pages, last_rows

                if outcome == CrawlOutcome.RATE_LIMITED:
                    self.say("[BLOCKED] %s" % why)
                    self.say("          %d row(s) were already collected "
                             "over %d page(s)." % (last_rows, last_pages))
                    return outcome, last_pages, last_rows

                if outcome == CrawlOutcome.TABLE_LOST:
                    if recoveries < C.CRAWL_TABLE_RECOVERIES:
                        recoveries += 1
                        self.say("[SCRAPER] %s; recovery %d/%d"
                                 % (why, recoveries,
                                    C.CRAWL_TABLE_RECOVERIES))
                        if await self._recover_table(popup, page):
                            last_change = time.monotonic()
                            continue
                    self.say("[SCRAPER] could not restore table detection "
                             "after %d attempt(s)" % recoveries)
                    return outcome, last_pages, last_rows

                self.say("[SCRAPER] %s" % why)
                return CrawlOutcome.ERROR, last_pages, last_rows

            crawling = await self._has(popup, C.SEL_STOP_CRAWL)
            ever_crawled = ever_crawled or crawling

            pages, rows = await self.read_stats(popup)
            # The live page handle, re-read every poll: a handle captured
            # before a restart is dead and its .url is frozen at whatever it
            # last was, which would make a stalled crawl look stable.
            try:
                current_url = page.url
            except Exception as exc:
                log.warning("could not read the page url: %s", exc)
                current_url = last_url
            if (pages, rows) != (last_pages, last_rows) \
                    or current_url != last_url:
                moved_to = progress.note_url(current_url)
                progress.note_stats(pages, rows)
                if moved_to:
                    self.say("[CRAWL] transition %d: page %d -> %d, "
                             "rows %d -> %d"
                             % (progress.transitions,
                                moved_to - 1, moved_to,
                                max(last_rows, 0), max(rows, 0)))
                elif pages > last_pages and pages not in shown:
                    shown.add(pages)
                    self.say("[SCRAPER] pages: %d  rows: %d" % (pages, rows))
                last_pages, last_rows = pages, rows
                last_url = current_url
                last_change = time.monotonic()

            if not ever_crawled and elapsed > C.CRAWL_NO_START_SECONDS:
                return CrawlOutcome.NEVER_STARTED, last_pages, last_rows

            if max_pages and crawling and last_pages >= max_pages:
                # Testing cap: press the extension's own Stop so the rows
                # already collected stay in its table and remain exportable.
                self.say("[SCRAPER] page cap %d reached; stopping crawl"
                         % max_pages)
                try:
                    await popup.click(C.SEL_STOP_CRAWL, timeout=5000)
                except Exception as exc:
                    log.warning("could not press Stop: %s", exc)
                last_change = time.monotonic()
                max_pages = None

            if ever_crawled and not crawling:
                instructions = await self._text(popup, C.SEL_INSTRUCTIONS)
                said_stopped = C.CRAWL_STOPPED_TEXT.lower() \
                    in instructions.lower()
                can_restart = await self._has(popup, C.SEL_START_CRAWL)
                stable_for = time.monotonic() - last_change
                if said_stopped and can_restart \
                        and stable_for >= C.CRAWL_STABLE_SECONDS:
                    progress.note_stats(last_pages, last_rows)
                    ok, complaint = progress.verify_complete()
                    if not ok:
                        self.say("[SCRAPER] crawling stopped")
                        self.say("[CRAWL] INCOMPLETE: %s" % complaint)
                        self.say("[CRAWL] %s" % progress.describe())
                        self.say("[CRAWL] refusing to call this search "
                                 "complete; nothing will be deleted")
                        return (CrawlOutcome.INCOMPLETE, last_pages,
                                last_rows)
                    self.say("[SCRAPER] crawling stopped")
                    self.say("[SCRAPER] stable for %d seconds"
                             % C.CRAWL_STABLE_SECONDS)
                    self.say("[CRAWL] final: Next present: no, %s"
                             % progress.describe())
                    return CrawlOutcome.FINISHED, last_pages, last_rows

            await asyncio.sleep(C.CRAWL_POLL_SECONDS)

    # ------------------------------------------------------------- export
    async def export_xlsx(self, popup, destination, download_dir=None):
        """Click XLSX once and land a validated export at `destination`.

        Chromium writes downloads to its OWN configured directory, which is
        NOT necessarily where the caller wants the file: a page-by-page export
        is destined for crawl_staging/<slug>/page_NNN.xlsx, but Chromium drops
        it in keyword_downloads/.  So recovery scans the real download
        directory -- the browser's, when there is one, else the destination's
        own -- never blindly `dirname(destination)`.

        The extension holds its rows only while the popup lives, so a lost
        popup means a lost crawl; a completed download, though, survives the
        browser dying mid-save (Chromium writes the bytes, then renames).
        That file is recovered from disk and validated rather than discarded
        -- the proven single-page behaviour, now correct for staged exports.

        Recovery is snapshotted and stability-checked so it can never adopt a
        previous page's file, another rank's file, or a half-written one:
          * the download directory is snapshotted BEFORE the click, so only a
            genuinely NEW file is ever considered;
          * more than one new export refuses (ambiguous), never guesses;
          * a .crdownload is only recovered if it OPENS as a valid workbook
            (in this environment a finished download stays unrenamed);
          * the result must pass validate_export() before it is accepted.
        """
        directory = download_dir or self._download_dir_for(destination)
        self.last_export_recovered = False
        for attempt in range(1, C.EXPORT_RETRIES + 1):
            before = self._dir_snapshot(directory)
            started = time.time()
            try:
                await popup.bring_to_front()
                await popup.wait_for_selector(C.SEL_XLSX,
                                              timeout=C.EXTENSION_TIMEOUT)
                async with popup.expect_download(
                        timeout=C.DOWNLOAD_TIMEOUT) as info:
                    await popup.click(C.SEL_XLSX)
                download = await info.value
                try:
                    await download.save_as(destination)
                except Exception as exc:
                    # save_as failed -- the completed bytes are usually still
                    # on disk (this environment leaves them unrenamed).  Try
                    # the quick path() copy; otherwise fall through to the one
                    # disk recovery below.
                    log.warning("save_as failed: %s", exc)
                    try:
                        source = await download.path()
                    except Exception as path_exc:
                        log.warning("download.path failed: %s", path_exc)
                        source = None
                    if not (source and os.path.exists(source)):
                        raise
                    shutil.copyfile(source, destination)
                rows = validate_export(destination)
                if rows < 0:
                    raise RuntimeError("downloaded file did not validate")
                return rows
            except Exception as exc:
                # ONE disk recovery per attempt: the export may be complete on
                # disk even though the page that owned it is gone.
                recovered = await self._recover_export(
                    destination, directory, started, before)
                if recovered >= 0:
                    self.last_export_recovered = True
                    self.say("[DOWNLOAD] recovered the export from disk "
                             "after a browser hiccup")
                    return recovered
                if browser_died(exc):
                    raise BrowserLost(str(exc))
                log.warning("XLSX export attempt %d/%d failed: %s",
                            attempt, C.EXPORT_RETRIES, exc)
                if attempt == C.EXPORT_RETRIES:
                    self.say("[FAIL] XLSX export failed: %s" % exc)
                    return -1
                await sleep_jitter(3 * attempt)
        return -1

    def _download_dir_for(self, destination):
        """Where Chromium actually writes: the browser's dir, else the dest's.

        The browser's configured download directory is the truth; falling
        back to the destination's own directory keeps single-page callers --
        whose destination already lives in that directory -- working.
        """
        browser_dir = getattr(self.browser, "download_dir", None)
        if browser_dir:
            return browser_dir
        return os.path.dirname(os.path.abspath(destination))

    @staticmethod
    def _dir_snapshot(directory):
        """The download-like filenames already present.  Empty when absent."""
        if not os.path.isdir(directory):
            return set()
        return {name for name in os.listdir(directory)
                if name.lower().endswith((".xlsx", ".crdownload"))}

    @staticmethod
    def _new_exports(directory, since, before):
        """(completed_xlsx_paths, pending_crdownload) -- kept for the shim.

        A conservative view used only by the sync _recover_download: a file
        counts as completed only when it is a new, non-empty .xlsx.  The
        content-aware recovery below is what production uses.
        """
        if not os.path.isdir(directory):
            return [], False
        completed, pending = [], False
        for name in os.listdir(directory):
            lower = name.lower()
            if not lower.endswith((".xlsx", ".crdownload")):
                continue
            if name in before:
                continue
            path = os.path.join(directory, name)
            try:
                stamp = os.path.getmtime(path)
                size = os.path.getsize(path)
            except OSError:
                continue
            if stamp < since - 5:
                continue
            if lower.endswith(".crdownload"):
                pending = True
            elif size > 0:
                completed.append(path)
        return completed, pending

    @staticmethod
    def _candidate_downloads(directory, since, before):
        """Every NEW download artifact since the click -- any extension.

        Both .xlsx and .crdownload are returned, because in this environment
        Playwright leaves a COMPLETED download as an unrenamed <guid>.crdownload
        in its downloads directory (verified: those files open as valid
        workbooks).  Completeness is therefore decided by OPENING the file,
        not by its extension -- an extension check is exactly what discarded a
        finished export and turned a survivable browser hiccup into a dead
        end.  A file present before the click, or older than the click, is not
        this export and is never returned.
        """
        if not os.path.isdir(directory):
            return []
        found = []
        for name in os.listdir(directory):
            lower = name.lower()
            if not (lower.endswith(".xlsx") or lower.endswith(".crdownload")):
                continue
            if name in before:
                continue
            path = os.path.join(directory, name)
            try:
                stamp = os.path.getmtime(path)
                size = os.path.getsize(path)
            except OSError:
                continue
            if stamp < since - 5 or size == 0:
                continue
            found.append(path)
        return found

    async def _recover_export(self, destination, directory, since, before,
                              settle_timeout=None):
        """Rescue THIS export from disk.  Rows on success, -1 otherwise.

        Completeness is proven by opening the file, so a finished download
        still named <guid>.crdownload is recovered while a genuinely partial
        one is refused.  Never guesses:

          * wait until a candidate's size is stable (it has stopped growing);
          * a stable candidate that opens as a valid workbook is recovered;
          * a stable candidate that will NOT open is a real partial -> refuse;
          * more than one recoverable candidate -> refuse (ambiguous);
          * nothing at all within the window -> -1.

        Capped well below DOWNLOAD_TIMEOUT: the bytes are already on disk, so
        this only waits out a size-stability check, never a fresh download.
        """
        limit = min((settle_timeout or C.DOWNLOAD_TIMEOUT) / 1000.0, 20.0)
        waited, poll = 0.0, 0.5
        sizes = {}
        while True:
            candidates = self._candidate_downloads(directory, since, before)
            stable, growing = [], False
            for path in candidates:
                try:
                    size = os.path.getsize(path)
                except OSError:
                    continue
                if sizes.get(path) == size and size > 0:
                    stable.append(path)
                else:
                    growing = True
                sizes[path] = size

            # Among stable candidates, which actually open as workbooks?
            recoverable = [path for path in stable
                           if self._opens_as_workbook(path)]
            if len(recoverable) > 1:
                self.say("[DOWNLOAD] refusing to recover: %d complete export "
                         "files appeared, cannot tell which is this page's"
                         % len(recoverable))
                log.warning("ambiguous recovery: %s",
                            [os.path.basename(p) for p in recoverable])
                return -1
            if len(recoverable) == 1:
                return self._claim_export(recoverable[0], destination)

            # Stable but none open -> the download really is partial.
            stable_unopenable = [path for path in stable
                                 if path not in recoverable]
            if stable_unopenable and not growing:
                self.say("[DOWNLOAD] the download is present but incomplete; "
                         "treating the page as incomplete")
                return -1

            if waited >= limit:
                if candidates:
                    self.say("[DOWNLOAD] the export did not finish in time; "
                             "treating the page as incomplete")
                return -1
            # A brief grace when nothing has appeared yet.
            if not candidates and waited >= 5.0:
                return -1
            await asyncio.sleep(poll)
            waited += poll

    @staticmethod
    def _opens_as_workbook(path):
        """True when `path` opens as a valid export, whatever its extension.

        validate_export() insists on an .xlsx name, so a .crdownload is copied
        to a temp .xlsx first -- this is only a completeness probe, it moves
        nothing into place.
        """
        if path.lower().endswith(".xlsx"):
            return validate_export(path) >= 0
        tmp = path + ".probe.xlsx"
        try:
            shutil.copyfile(path, tmp)
            return validate_export(tmp) >= 0
        except OSError:
            return False
        finally:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass

    def _claim_export(self, source, destination):
        """Validate `source`'s contents, then place it atomically at dest.

        Validation happens on a temp copy BEFORE anything lands at
        `destination`, so a file that fails to open can never be left sitting
        where a caller would treat it as a finished export.
        """
        target_dir = os.path.dirname(os.path.abspath(destination))
        os.makedirs(target_dir, exist_ok=True)
        tmp = os.path.join(target_dir,
                           ".%s.recover.xlsx" % os.path.basename(destination))
        try:
            shutil.copyfile(source, tmp)
            with open(tmp, "rb") as handle:
                os.fsync(handle.fileno())
        except OSError as exc:
            log.warning("could not stage recovered export: %s", exc)
            self._quiet_remove(tmp)
            return -1
        rows = validate_export(tmp)
        if rows < 0:
            self.say("[DOWNLOAD] the recovered file did not validate; "
                     "treating the page as incomplete")
            self._quiet_remove(tmp)
            return -1
        try:
            os.replace(tmp, destination)
        except OSError as exc:
            log.warning("could not place recovered export: %s", exc)
            self._quiet_remove(tmp)
            return -1
        self._quiet_remove(source)
        log.info("recovered export from %s", os.path.basename(source))
        return rows

    @staticmethod
    def _quiet_remove(path):
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            pass

    @staticmethod
    def _recover_download(destination, since, download_dir=None, before=None):
        """Sync one-shot recovery, kept for callers that do not await.

        A thin, non-waiting form of _recover_export: scans the download
        directory once, refuses ambiguity, and copies a single new completed
        export.  download_dir defaults to the destination's own directory so
        existing single-page callers are unchanged.
        """
        directory = download_dir or os.path.dirname(
            os.path.abspath(destination))
        completed, _pending = InstantDataScraper._new_exports(
            directory, since, before or set())
        if len(completed) != 1:
            return False
        source = completed[0]
        try:
            shutil.copyfile(source, destination)
        except OSError as exc:
            log.warning("could not recover %s: %s", source, exc)
            return False
        if validate_export(destination) < 0:
            log.warning("recovered file %s did not validate", source)
            return False
        try:
            os.remove(source)
        except OSError:
            pass
        log.info("recovered export from %s", os.path.basename(source))
        return True

    async def _recover_table(self, popup, page):
        """Re-establish table detection WITHOUT losing collected rows.

        The popup keeps its collected rows in memory for as long as the
        page lives, so recovery must never reload or reopen it.  Only the
        GitHub tab is refreshed, then the crawl is resumed with the same
        Start control, which appends to the existing data.
        """
        try:
            await page.reload(wait_until="domcontentloaded",
                              timeout=C.NAVIGATION_TIMEOUT)
            try:
                await page.wait_for_load_state("networkidle", timeout=8000)
            except Exception:
                pass
            # Let the extension re-scan the refreshed DOM.
            if await self._has(popup, C.SEL_WRONG_TABLE):
                try:
                    await popup.click(C.SEL_WRONG_TABLE, timeout=3000)
                except Exception:
                    pass
            if not await self.wait_for_table(popup, timeout=15000):
                return False
            if await self._has(popup, C.SEL_START_CRAWL):
                await popup.click(C.SEL_START_CRAWL,
                                  timeout=C.EXTENSION_TIMEOUT)
                self.say("[SCRAPER] crawling resumed")
                return True
        except Exception as exc:
            log.warning("table recovery failed: %s", exc)
        return False

    # --------------------------------------------------------- diagnostics
    async def capture_page_state(self, page, reason=""):
        """Snapshot GITHUB only -- no popup, no extension.

        This is what lets collect() decide what the page is BEFORE opening
        Instant Data Scraper: a zero-result search should not pay for the
        extension, and a blocked page should not be poked at further.
        """
        state = {"reason": reason}
        try:
            state.update(await page.evaluate(PAGE_STATE_JS))
        except Exception as exc:
            state["page_error"] = str(exc)
            try:
                state["url"] = page.url
            except Exception:
                pass
        try:
            state["blocked_marker"] = blocked_marker(
                "%s %s" % (state.get("bodyHead", ""), state.get("bodyTail", "")))
        except Exception:
            state["blocked_marker"] = None
        state["browserPage"] = page_number(state.get("url") or "")
        return state

    async def capture_state(self, page, popup, reason=""):
        """Snapshot of GitHub AND the extension. Never changes anything."""
        state = {"reason": reason}
        try:
            state.update(await page.evaluate(PAGE_STATE_JS))
        except Exception as exc:
            state["page_error"] = str(exc)
            try:
                state["url"] = page.url
            except Exception:
                pass
        try:
            response_marker = blocked_marker(state.get("bodyHead", "")
                                             + state.get("bodyTail", ""))
        except Exception:
            response_marker = None
        state["blocked_marker"] = response_marker

        pages, rows = await self.read_stats(popup)
        state["ext_pages"] = pages
        state["ext_rows"] = rows
        state["ext_error"] = (await self._text(popup, C.SEL_ERROR)).strip()
        state["ext_instructions"] = (
            await self._text(popup, C.SEL_INSTRUCTIONS)).strip()
        state["ext_has_start"] = await self._has(popup, C.SEL_START_CRAWL)
        state["ext_has_stop"] = await self._has(popup, C.SEL_STOP_CRAWL)
        state["ext_has_wrong_table"] = await self._has(popup,
                                                       C.SEL_WRONG_TABLE)
        state["browserPage"] = page_number(state.get("url") or "")
        return state

    @staticmethod
    def log_state(state):
        log.error("---- page state at '%s' ----", state.get("reason", ""))
        for key in ("url", "title", "bodyChars", "hasNext", "nextDisabled",
                    "nextHref", "profileLinkCount", "resultsListPresent",
                    "resultsListChildren", "noResults", "saysError",
                    "blocked_marker", "ext_pages", "ext_rows", "ext_error",
                    "ext_instructions", "ext_has_start", "ext_has_stop",
                    "ext_has_wrong_table", "page_error"):
            if key in state:
                log.error("   %-20s %s", key, str(state[key])[:300])
        log.error("   body head: %s", str(state.get("bodyHead", ""))[:300])
        log.error("   body tail: %s", str(state.get("bodyTail", ""))[:300])

    # ------------------------------------------------------------- helpers
    @staticmethod
    async def _has(page, selector):
        try:
            return await page.locator(selector).count() > 0
        except Exception:
            return False

    @staticmethod
    async def _text(page, selector):
        try:
            node = page.locator(selector)
            if await node.count() == 0:
                return ""
            return await node.first.inner_text(timeout=2000)
        except Exception:
            return ""


def write_dom_fallback(state, destination, expected_ids, page_count):
    """Last-resort serializer: write ONE page's verified DOM identities.

    Shared by the whole-single-page fallback and the per-page multi-page-crawl
    fallback, so both use the SAME authoritative-DOM validation and the SAME
    one-identity-per-row serialization.  DOM identities are the authoritative
    set; this is used ONLY when IDS cannot establish the exact table.

    Returns (ok, count, reason).  Writes and accepts a file ONLY when every
    check holds:
      * 1 <= page_count <= GITHUB_USERS_PER_PAGE (a small page);
      * the DOM's unique identity set equals exactly page_count entries
        (expected_ids is that set) -- an ambiguous/duplicate/short DOM refuses;
      * exactly page_count valid unique usernames are extractable;
      * the written file reads back through XlsxReader as EXACTLY expected_ids.
    On any failure ok is False and the caller must NOT accept the page.
    """
    if not 1 <= page_count <= GITHUB_USERS_PER_PAGE:
        return (False, 0, "page count %d is outside 1..%d"
                % (page_count, GITHUB_USERS_PER_PAGE))
    if len(expected_ids) != page_count:
        return (False, 0,
                "DOM shows %d unique identity(ies) but the page count is %d"
                % (len(expected_ids), page_count))
    usernames = InstantDataScraper.dom_usernames(state)
    if len(usernames) != page_count:
        return (False, 0,
                "only %d valid unique DOM username(s) for a %d-user page"
                % (len(usernames), page_count))
    InstantDataScraper.write_dom_candidates_xlsx(destination, usernames)
    got = InstantDataScraper.export_identities(destination)
    if got != expected_ids:
        return (False, 0, "written file read back as %s, not %s"
                % (sorted(got), sorted(expected_ids)))
    return (True, len(usernames), "")


def validate_export(path):
    """Row count of a downloaded export, or -1 when it is not usable.

    Checks existence, non-zero size, the .xlsx extension, that openpyxl can
    open it, that a worksheet and a header exist.  Zero data rows is a valid
    answer (GitHub really can return no users), so 0 is returned, not -1.
    """
    if not path.lower().endswith(".xlsx"):
        return -1
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return -1
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        log.warning("export will not open: %s", exc)
        return -1
    try:
        if not book.worksheets:
            return -1
        sheet = book.active
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, None)
        if not header or not any(value not in (None, "") for value in header):
            return -1
        return sum(1 for row in rows
                   if any(value not in (None, "") for value in row))
    finally:
        book.close()


# ==========================================================================
# LOCATION EXTRACTOR
# ==========================================================================
# Location parsing driven by gazetteers, not by regex.
#
# A location string is a short list of place names.  Deciding what they mean
# is a lookup problem: "Special Region of Yogyakarta" is a province, and
# provinces belong to exactly one country.  No pattern over characters can
# know that, so the knowledge lives in tables here.
#
# Two rules matter more than coverage:
#
#   * country is never guessed.  It is set only when a country name, a known
#     first-level subdivision or a known city establishes it.  A person's
#     name never contributes.
#   * the raw string is preserved, so a normalisation that turns out to be
#     wrong can always be audited against what the profile actually said.
# --------------------------------------------------------------------------
# ------------------------------------------------------------- countries --
COUNTRY_ALIASES = {
    "india": "India", "bharat": "India", "ind": "India", "in": "India",
    "indonesia": "Indonesia", "id": "Indonesia",
    "united states": "United States", "united states of america":
        "United States", "usa": "United States", "u.s.a": "United States",
    "u.s.": "United States", "us": "United States", "america":
        "United States",
    "united kingdom": "United Kingdom", "uk": "United Kingdom",
    "great britain": "United Kingdom", "england": "United Kingdom",
    "scotland": "United Kingdom", "wales": "United Kingdom",
    "northern ireland": "United Kingdom",
    "canada": "Canada", "australia": "Australia", "new zealand":
        "New Zealand",
    "germany": "Germany", "deutschland": "Germany", "france": "France",
    "spain": "Spain", "espana": "Spain", "italy": "Italy",
    "netherlands": "Netherlands", "holland": "Netherlands",
    "belgium": "Belgium", "switzerland": "Switzerland", "austria": "Austria",
    "poland": "Poland", "portugal": "Portugal", "ireland": "Ireland",
    "sweden": "Sweden", "norway": "Norway", "denmark": "Denmark",
    "finland": "Finland", "greece": "Greece", "czechia": "Czechia",
    "czech republic": "Czechia", "romania": "Romania", "hungary": "Hungary",
    "ukraine": "Ukraine", "russia": "Russia", "turkey": "Turkey",
    "turkiye": "Turkey",
    "china": "China", "japan": "Japan", "south korea": "South Korea",
    "korea": "South Korea", "singapore": "Singapore", "malaysia": "Malaysia",
    "thailand": "Thailand", "vietnam": "Vietnam", "viet nam": "Vietnam",
    "philippines": "Philippines", "taiwan": "Taiwan",
    "hong kong": "Hong Kong",
    "pakistan": "Pakistan", "bangladesh": "Bangladesh", "nepal": "Nepal",
    "sri lanka": "Sri Lanka", "afghanistan": "Afghanistan",
    "uae": "United Arab Emirates", "united arab emirates":
        "United Arab Emirates", "saudi arabia": "Saudi Arabia",
    "qatar": "Qatar", "kuwait": "Kuwait", "bahrain": "Bahrain",
    "oman": "Oman", "israel": "Israel", "jordan": "Jordan",
    "egypt": "Egypt", "nigeria": "Nigeria", "kenya": "Kenya",
    "ghana": "Ghana", "south africa": "South Africa", "ethiopia": "Ethiopia",
    "morocco": "Morocco", "tunisia": "Tunisia", "algeria": "Algeria",
    "brazil": "Brazil", "brasil": "Brazil", "argentina": "Argentina",
    "chile": "Chile", "colombia": "Colombia", "peru": "Peru",
    "mexico": "Mexico", "venezuela": "Venezuela", "ecuador": "Ecuador",
    "uruguay": "Uruguay", "bolivia": "Bolivia",
    "guatemala": "Guatemala", "costa rica": "Costa Rica",
    "panama": "Panama", "cuba": "Cuba", "honduras": "Honduras",
    "el salvador": "El Salvador", "nicaragua": "Nicaragua",
    "dominican republic": "Dominican Republic", "jamaica": "Jamaica",
    "paraguay": "Paraguay", "tanzania": "Tanzania", "uganda": "Uganda",
    "rwanda": "Rwanda", "zambia": "Zambia", "zimbabwe": "Zimbabwe",
    "botswana": "Botswana", "namibia": "Namibia", "senegal": "Senegal",
    "cameroon": "Cameroon", "ivory coast": "Ivory Coast",
    "sudan": "Sudan", "libya": "Libya", "lebanon": "Lebanon",
    "iraq": "Iraq", "iran": "Iran", "syria": "Syria", "yemen": "Yemen",
    "kazakhstan": "Kazakhstan", "uzbekistan": "Uzbekistan",
    "azerbaijan": "Azerbaijan", "armenia": "Armenia", "belarus": "Belarus",
    "lithuania": "Lithuania", "latvia": "Latvia", "estonia": "Estonia",
    "slovakia": "Slovakia", "slovenia": "Slovenia", "croatia": "Croatia",
    "serbia": "Serbia", "bulgaria": "Bulgaria", "albania": "Albania",
    "bosnia and herzegovina": "Bosnia and Herzegovina",
    "north macedonia": "North Macedonia", "moldova": "Moldova",
    "iceland": "Iceland", "luxembourg": "Luxembourg", "malta": "Malta",
    "cyprus": "Cyprus", "myanmar": "Myanmar", "cambodia": "Cambodia",
    "laos": "Laos", "mongolia": "Mongolia", "bhutan": "Bhutan",
    "maldives": "Maldives", "brunei": "Brunei", "fiji": "Fiji",
    "papua new guinea": "Papua New Guinea", "mauritius": "Mauritius",
    "madagascar": "Madagascar", "mozambique": "Mozambique",
    "angola": "Angola", "ghana": "Ghana", "mali": "Mali",
    "burkina faso": "Burkina Faso", "niger": "Niger", "chad": "Chad",
    "somalia": "Somalia", "malawi": "Malawi",
}

# ------------------------------------------- first-level subdivisions -----
# region alias -> (canonical region, country)
_SUBDIVISIONS_RAW = {
    "India": {
        "Andhra Pradesh": ["andhra pradesh", "ap"],
        "Arunachal Pradesh": ["arunachal pradesh"],
        "Assam": ["assam"], "Bihar": ["bihar"],
        "Chhattisgarh": ["chhattisgarh", "chattisgarh"],
        "Delhi": ["delhi", "nct of delhi", "national capital territory"],
        "Goa": ["goa"], "Gujarat": ["gujarat"],
        "Haryana": ["haryana"], "Himachal Pradesh": ["himachal pradesh", "hp"],
        "Jharkhand": ["jharkhand"], "Karnataka": ["karnataka", "ka"],
        "Kerala": ["kerala"],
        "Madhya Pradesh": ["madhya pradesh", "mp"],
        "Maharashtra": ["maharashtra", "mh"], "Manipur": ["manipur"],
        "Meghalaya": ["meghalaya"], "Mizoram": ["mizoram"],
        "Nagaland": ["nagaland"], "Odisha": ["odisha", "orissa"],
        "Punjab": ["punjab"], "Rajasthan": ["rajasthan", "rj"],
        "Sikkim": ["sikkim"], "Tamil Nadu": ["tamil nadu", "tn"],
        "Telangana": ["telangana", "ts"], "Tripura": ["tripura"],
        "Uttar Pradesh": ["uttar pradesh", "up"],
        "Uttarakhand": ["uttarakhand", "uttaranchal"],
        "West Bengal": ["west bengal", "wb"],
        "Chandigarh": ["chandigarh"],
        "Jammu and Kashmir": ["jammu and kashmir", "jammu & kashmir", "j&k"],
        "Puducherry": ["puducherry", "pondicherry"],
    },
    "Indonesia": {
        "Special Region of Yogyakarta": [
            "special region of yogyakarta", "daerah istimewa yogyakarta",
            "di yogyakarta", "diy yogyakarta", "yogyakarta special region"],
        "Jakarta": ["jakarta", "dki jakarta", "special capital region of "
                    "jakarta"],
        "West Java": ["west java", "jawa barat"],
        "Central Java": ["central java", "jawa tengah"],
        "East Java": ["east java", "jawa timur"],
        "Banten": ["banten"], "Bali": ["bali"],
        "North Sumatra": ["north sumatra", "sumatera utara"],
        "West Sumatra": ["west sumatra", "sumatera barat"],
        "South Sumatra": ["south sumatra", "sumatera selatan"],
        "Riau": ["riau"], "Lampung": ["lampung"],
        "South Sulawesi": ["south sulawesi", "sulawesi selatan"],
        "North Sulawesi": ["north sulawesi", "sulawesi utara"],
        "East Kalimantan": ["east kalimantan", "kalimantan timur"],
        "West Kalimantan": ["west kalimantan", "kalimantan barat"],
        "Aceh": ["aceh"], "Papua": ["papua"],
    },
    "United States": {
        "California": ["california", "ca"], "New York": ["new york", "ny"],
        "Texas": ["texas", "tx"], "Florida": ["florida", "fl"],
        "Washington": ["washington", "wa"], "Massachusetts":
            ["massachusetts", "ma"], "Illinois": ["illinois", "il"],
        "Pennsylvania": ["pennsylvania", "pa"], "Georgia": ["georgia", "ga"],
        "New Jersey": ["new jersey", "nj"], "Virginia": ["virginia", "va"],
        "North Carolina": ["north carolina", "nc"], "Ohio": ["ohio", "oh"],
        "Michigan": ["michigan", "mi"], "Arizona": ["arizona", "az"],
        "Colorado": ["colorado", "co"], "Oregon": ["oregon", "or"],
        "Maryland": ["maryland", "md"], "Minnesota": ["minnesota", "mn"],
        "Utah": ["utah", "ut"], "Tennessee": ["tennessee", "tn"],
        "Indiana": ["indiana"], "Wisconsin": ["wisconsin", "wi"],
        "Missouri": ["missouri", "mo"], "Nevada": ["nevada", "nv"],
    },
    "Canada": {
        "Ontario": ["ontario", "on"], "Quebec": ["quebec", "qc"],
        "British Columbia": ["british columbia", "bc"],
        "Alberta": ["alberta", "ab"], "Manitoba": ["manitoba", "mb"],
        "Nova Scotia": ["nova scotia", "ns"],
        "Saskatchewan": ["saskatchewan", "sk"],
    },
    "Australia": {
        "New South Wales": ["new south wales", "nsw"],
        "Victoria": ["victoria", "vic"], "Queensland": ["queensland", "qld"],
        "Western Australia": ["western australia"],
        "South Australia": ["south australia"],
        "Australian Capital Territory": ["australian capital territory",
                                         "act"],
        "Tasmania": ["tasmania"],
    },
    "Germany": {
        "Bavaria": ["bavaria", "bayern"], "Berlin": ["berlin"],
        "Hesse": ["hesse", "hessen"],
        "North Rhine-Westphalia": ["north rhine-westphalia", "nrw"],
        "Baden-Wurttemberg": ["baden-wurttemberg", "baden wurttemberg"],
        "Hamburg": ["hamburg"], "Saxony": ["saxony", "sachsen"],
    },
    "Pakistan": {
        "Punjab (PK)": ["punjab pakistan"], "Sindh": ["sindh"],
        "Khyber Pakhtunkhwa": ["khyber pakhtunkhwa", "kpk"],
        "Balochistan": ["balochistan"],
    },
    "Nigeria": {"Lagos State": ["lagos state"], "Abuja": ["abuja", "fct"]},
    "Brazil": {"Sao Paulo": ["sao paulo", "sp"],
               "Rio de Janeiro": ["rio de janeiro", "rj"],
               "Minas Gerais": ["minas gerais"]},
}

SUBDIVISIONS = {}
for _country, _regions in _SUBDIVISIONS_RAW.items():
    for _canonical, _aliases in _regions.items():
        for _alias in _aliases:
            # Two-letter codes collide across countries (TN, PA, CA...);
            # the first table wins and ambiguous codes stay low-confidence.
            SUBDIVISIONS.setdefault(_alias, (_canonical, _country))

# --------------------------------------------------------------- cities --
# city alias -> (canonical city, region or "", country)
_CITIES_RAW = [
    # India (canonical renames matter for deduplication)
    ("bangalore", "Bengaluru", "Karnataka", "India"),
    ("banglore", "Bengaluru", "Karnataka", "India"),
    ("bengaluru", "Bengaluru", "Karnataka", "India"),
    ("bombay", "Mumbai", "Maharashtra", "India"),
    ("mumbai", "Mumbai", "Maharashtra", "India"),
    ("calcutta", "Kolkata", "West Bengal", "India"),
    ("kolkata", "Kolkata", "West Bengal", "India"),
    ("madras", "Chennai", "Tamil Nadu", "India"),
    ("chennai", "Chennai", "Tamil Nadu", "India"),
    ("gurgaon", "Gurugram", "Haryana", "India"),
    ("gurugram", "Gurugram", "Haryana", "India"),
    ("allahabad", "Prayagraj", "Uttar Pradesh", "India"),
    ("prayagraj", "Prayagraj", "Uttar Pradesh", "India"),
    ("mysore", "Mysuru", "Karnataka", "India"),
    ("mangalore", "Mangaluru", "Karnataka", "India"),
    ("trivandrum", "Thiruvananthapuram", "Kerala", "India"),
    ("cochin", "Kochi", "Kerala", "India"),
    ("calicut", "Kozhikode", "Kerala", "India"),
    ("baroda", "Vadodara", "Gujarat", "India"),
    ("vizag", "Visakhapatnam", "Andhra Pradesh", "India"),
    ("poona", "Pune", "Maharashtra", "India"),
    ("pune", "Pune", "Maharashtra", "India"),
    ("hyderabad", "Hyderabad", "Telangana", "India"),
    ("bhopal", "Bhopal", "Madhya Pradesh", "India"),
    ("indore", "Indore", "Madhya Pradesh", "India"),
    ("gwalior", "Gwalior", "Madhya Pradesh", "India"),
    ("jabalpur", "Jabalpur", "Madhya Pradesh", "India"),
    ("noida", "Noida", "Uttar Pradesh", "India"),
    ("ghaziabad", "Ghaziabad", "Uttar Pradesh", "India"),
    ("lucknow", "Lucknow", "Uttar Pradesh", "India"),
    ("kanpur", "Kanpur", "Uttar Pradesh", "India"),
    ("varanasi", "Varanasi", "Uttar Pradesh", "India"),
    ("agra", "Agra", "Uttar Pradesh", "India"),
    ("mathura", "Mathura", "Uttar Pradesh", "India"),
    ("jaipur", "Jaipur", "Rajasthan", "India"),
    ("jodhpur", "Jodhpur", "Rajasthan", "India"),
    ("udaipur", "Udaipur", "Rajasthan", "India"),
    ("kota", "Kota", "Rajasthan", "India"),
    ("ahmedabad", "Ahmedabad", "Gujarat", "India"),
    ("surat", "Surat", "Gujarat", "India"),
    ("rajkot", "Rajkot", "Gujarat", "India"),
    ("gandhinagar", "Gandhinagar", "Gujarat", "India"),
    ("nagpur", "Nagpur", "Maharashtra", "India"),
    ("nashik", "Nashik", "Maharashtra", "India"),
    ("thane", "Thane", "Maharashtra", "India"),
    ("aurangabad", "Aurangabad", "Maharashtra", "India"),
    ("patna", "Patna", "Bihar", "India"),
    ("ranchi", "Ranchi", "Jharkhand", "India"),
    ("jamshedpur", "Jamshedpur", "Jharkhand", "India"),
    ("raipur", "Raipur", "Chhattisgarh", "India"),
    ("bhubaneswar", "Bhubaneswar", "Odisha", "India"),
    ("cuttack", "Cuttack", "Odisha", "India"),
    ("guwahati", "Guwahati", "Assam", "India"),
    ("dehradun", "Dehradun", "Uttarakhand", "India"),
    ("coimbatore", "Coimbatore", "Tamil Nadu", "India"),
    ("madurai", "Madurai", "Tamil Nadu", "India"),
    ("salem", "Salem", "Tamil Nadu", "India"),
    ("vellore", "Vellore", "Tamil Nadu", "India"),
    ("warangal", "Warangal", "Telangana", "India"),
    ("vijayawada", "Vijayawada", "Andhra Pradesh", "India"),
    ("tirupati", "Tirupati", "Andhra Pradesh", "India"),
    ("amritsar", "Amritsar", "Punjab", "India"),
    ("ludhiana", "Ludhiana", "Punjab", "India"),
    ("jalandhar", "Jalandhar", "Punjab", "India"),
    ("mohali", "Mohali", "Punjab", "India"),
    ("faridabad", "Faridabad", "Haryana", "India"),
    ("new delhi", "New Delhi", "Delhi", "India"),
    # City-states: naming one names both the city and its division.
    ("delhi", "Delhi", "Delhi", "India"),
    ("jakarta", "Jakarta", "Jakarta", "Indonesia"),
    ("chandigarh", "Chandigarh", "Chandigarh", "India"),
    ("puducherry", "Puducherry", "Puducherry", "India"),
    ("hong kong", "Hong Kong", "", "Hong Kong"),
    ("shimla", "Shimla", "Himachal Pradesh", "India"),
    ("srinagar", "Srinagar", "Jammu and Kashmir", "India"),
    ("siliguri", "Siliguri", "West Bengal", "India"),
    ("howrah", "Howrah", "West Bengal", "India"),
    ("durgapur", "Durgapur", "West Bengal", "India"),
    # Indonesia
    ("bantul", "Bantul", "Special Region of Yogyakarta", "Indonesia"),
    ("sleman", "Sleman", "Special Region of Yogyakarta", "Indonesia"),
    ("yogyakarta", "Yogyakarta", "Special Region of Yogyakarta", "Indonesia"),
    ("jogja", "Yogyakarta", "Special Region of Yogyakarta", "Indonesia"),
    ("bandung", "Bandung", "West Java", "Indonesia"),
    ("surabaya", "Surabaya", "East Java", "Indonesia"),
    ("semarang", "Semarang", "Central Java", "Indonesia"),
    ("medan", "Medan", "North Sumatra", "Indonesia"),
    ("denpasar", "Denpasar", "Bali", "Indonesia"),
    ("makassar", "Makassar", "South Sulawesi", "Indonesia"),
    ("malang", "Malang", "East Java", "Indonesia"),
    ("depok", "Depok", "West Java", "Indonesia"),
    ("bekasi", "Bekasi", "West Java", "Indonesia"),
    ("tangerang", "Tangerang", "Banten", "Indonesia"),
    # elsewhere
    ("london", "London", "", "United Kingdom"),
    ("manchester", "Manchester", "", "United Kingdom"),
    ("birmingham", "Birmingham", "", "United Kingdom"),
    ("edinburgh", "Edinburgh", "", "United Kingdom"),
    ("san francisco", "San Francisco", "California", "United States"),
    ("los angeles", "Los Angeles", "California", "United States"),
    ("san jose", "San Jose", "California", "United States"),
    ("seattle", "Seattle", "Washington", "United States"),
    ("boston", "Boston", "Massachusetts", "United States"),
    ("chicago", "Chicago", "Illinois", "United States"),
    ("austin", "Austin", "Texas", "United States"),
    ("toronto", "Toronto", "Ontario", "Canada"),
    ("vancouver", "Vancouver", "British Columbia", "Canada"),
    ("montreal", "Montreal", "Quebec", "Canada"),
    ("sydney", "Sydney", "New South Wales", "Australia"),
    ("melbourne", "Melbourne", "Victoria", "Australia"),
    ("berlin", "Berlin", "Berlin", "Germany"),
    ("munich", "Munich", "Bavaria", "Germany"),
    ("paris", "Paris", "", "France"),
    ("amsterdam", "Amsterdam", "", "Netherlands"),
    ("madrid", "Madrid", "", "Spain"),
    ("barcelona", "Barcelona", "", "Spain"),
    ("lisbon", "Lisbon", "", "Portugal"),
    ("dublin", "Dublin", "", "Ireland"),
    ("warsaw", "Warsaw", "", "Poland"),
    ("athens", "Athens", "", "Greece"),
    ("istanbul", "Istanbul", "", "Turkey"),
    ("dubai", "Dubai", "", "United Arab Emirates"),
    ("abu dhabi", "Abu Dhabi", "", "United Arab Emirates"),
    ("singapore", "Singapore", "", "Singapore"),
    ("kuala lumpur", "Kuala Lumpur", "", "Malaysia"),
    ("bangkok", "Bangkok", "", "Thailand"),
    ("manila", "Manila", "", "Philippines"),
    ("hanoi", "Hanoi", "", "Vietnam"),
    ("ho chi minh city", "Ho Chi Minh City", "", "Vietnam"),
    ("tokyo", "Tokyo", "", "Japan"),
    ("seoul", "Seoul", "", "South Korea"),
    ("beijing", "Beijing", "", "China"),
    ("shanghai", "Shanghai", "", "China"),
    ("lahore", "Lahore", "Punjab (PK)", "Pakistan"),
    ("karachi", "Karachi", "Sindh", "Pakistan"),
    ("islamabad", "Islamabad", "", "Pakistan"),
    ("dhaka", "Dhaka", "", "Bangladesh"),
    ("kathmandu", "Kathmandu", "", "Nepal"),
    ("colombo", "Colombo", "", "Sri Lanka"),
    ("lagos", "Lagos", "Lagos State", "Nigeria"),
    ("nairobi", "Nairobi", "", "Kenya"),
    ("cairo", "Cairo", "", "Egypt"),
    ("accra", "Accra", "", "Ghana"),
    ("sao paulo", "Sao Paulo", "Sao Paulo", "Brazil"),
    ("rio de janeiro", "Rio de Janeiro", "Rio de Janeiro", "Brazil"),
    ("buenos aires", "Buenos Aires", "", "Argentina"),
    ("mexico city", "Mexico City", "", "Mexico"),
    ("bogota", "Bogota", "", "Colombia"),
    ("lima", "Lima", "", "Peru"),
    ("santiago", "Santiago", "", "Chile"),
    # Seen in real collected rows that produced no country before.
    ("dar es salaam", "Dar es Salaam", "", "Tanzania"),
    ("houston", "Houston", "Texas", "United States"),
    ("philadelphia", "Philadelphia", "Pennsylvania", "United States"),
    ("new york city", "New York City", "New York", "United States"),
    ("washington dc", "Washington", "Washington", "United States"),
    ("atlanta", "Atlanta", "Georgia", "United States"),
    ("dallas", "Dallas", "Texas", "United States"),
    ("miami", "Miami", "Florida", "United States"),
    ("denver", "Denver", "Colorado", "United States"),
    ("phoenix", "Phoenix", "Arizona", "United States"),
    ("detroit", "Detroit", "Michigan", "United States"),
    ("nelspruit", "Nelspruit", "", "South Africa"),
    ("durban", "Durban", "", "South Africa"),
    ("johannesburg", "Johannesburg", "", "South Africa"),
    ("cape town", "Cape Town", "", "South Africa"),
    ("pretoria", "Pretoria", "", "South Africa"),
    ("cartagena", "Cartagena", "", "Colombia"),
    ("medellin", "Medellin", "", "Colombia"),
    ("bahawalnagar", "Bahawalnagar", "Punjab (PK)", "Pakistan"),
    ("rawalpindi", "Rawalpindi", "Punjab (PK)", "Pakistan"),
    ("faisalabad", "Faisalabad", "Punjab (PK)", "Pakistan"),
    ("peshawar", "Peshawar", "Khyber Pakhtunkhwa", "Pakistan"),
    ("chittagong", "Chittagong", "", "Bangladesh"),
    ("sylhet", "Sylhet", "", "Bangladesh"),
    ("kampala", "Kampala", "", "Uganda"),
    ("addis ababa", "Addis Ababa", "", "Ethiopia"),
    ("casablanca", "Casablanca", "", "Morocco"),
    ("tunis", "Tunis", "", "Tunisia"),
    ("algiers", "Algiers", "", "Algeria"),
    ("kigali", "Kigali", "", "Rwanda"),
    ("abuja", "Abuja", "Abuja", "Nigeria"),
    ("ibadan", "Ibadan", "", "Nigeria"),
    ("port harcourt", "Port Harcourt", "", "Nigeria"),
    ("guadalajara", "Guadalajara", "", "Mexico"),
    ("monterrey", "Monterrey", "", "Mexico"),
    ("curitiba", "Curitiba", "", "Brazil"),
    ("belo horizonte", "Belo Horizonte", "Minas Gerais", "Brazil"),
    ("brasilia", "Brasilia", "", "Brazil"),
    ("porto alegre", "Porto Alegre", "", "Brazil"),
    ("cordoba", "Cordoba", "", "Argentina"),
    ("quito", "Quito", "", "Ecuador"),
    ("guayaquil", "Guayaquil", "", "Ecuador"),
    ("la paz", "La Paz", "", "Bolivia"),
    ("montevideo", "Montevideo", "", "Uruguay"),
    ("caracas", "Caracas", "", "Venezuela"),
    ("kyiv", "Kyiv", "", "Ukraine"),
    ("kiev", "Kyiv", "", "Ukraine"),
    ("moscow", "Moscow", "", "Russia"),
    ("saint petersburg", "Saint Petersburg", "", "Russia"),
    ("bucharest", "Bucharest", "", "Romania"),
    ("budapest", "Budapest", "", "Hungary"),
    ("prague", "Prague", "", "Czechia"),
    ("krakow", "Krakow", "", "Poland"),
    ("milan", "Milan", "", "Italy"),
    ("rome", "Rome", "", "Italy"),
    ("zurich", "Zurich", "", "Switzerland"),
    ("vienna", "Vienna", "", "Austria"),
    ("brussels", "Brussels", "", "Belgium"),
    ("stockholm", "Stockholm", "", "Sweden"),
    ("oslo", "Oslo", "", "Norway"),
    ("copenhagen", "Copenhagen", "", "Denmark"),
    ("helsinki", "Helsinki", "", "Finland"),
    ("rotterdam", "Rotterdam", "", "Netherlands"),
    ("frankfurt", "Frankfurt", "Hesse", "Germany"),
    ("stuttgart", "Stuttgart", "Baden-Wurttemberg", "Germany"),
    ("cologne", "Cologne", "North Rhine-Westphalia", "Germany"),
    ("lyon", "Lyon", "", "France"),
    ("marseille", "Marseille", "", "France"),
    ("porto", "Porto", "", "Portugal"),
    ("valencia", "Valencia", "", "Spain"),
    ("leeds", "Leeds", "", "United Kingdom"),
    ("glasgow", "Glasgow", "", "United Kingdom"),
    ("bristol", "Bristol", "", "United Kingdom"),
    ("nottingham", "Nottingham", "", "United Kingdom"),
    ("sheffield", "Sheffield", "", "United Kingdom"),
    ("liverpool", "Liverpool", "", "United Kingdom"),
    ("brisbane", "Brisbane", "Queensland", "Australia"),
    ("adelaide", "Adelaide", "South Australia", "Australia"),
    ("canberra", "Canberra", "Australian Capital Territory", "Australia"),
    ("auckland", "Auckland", "", "New Zealand"),
    ("wellington", "Wellington", "", "New Zealand"),
    ("calgary", "Calgary", "Alberta", "Canada"),
    ("ottawa", "Ottawa", "Ontario", "Canada"),
    ("edmonton", "Edmonton", "Alberta", "Canada"),
    ("osaka", "Osaka", "", "Japan"),
    ("kyoto", "Kyoto", "", "Japan"),
    ("busan", "Busan", "", "South Korea"),
    ("shenzhen", "Shenzhen", "", "China"),
    ("guangzhou", "Guangzhou", "", "China"),
    ("hangzhou", "Hangzhou", "", "China"),
    ("taipei", "Taipei", "", "Taiwan"),
    ("jeddah", "Jeddah", "", "Saudi Arabia"),
    ("riyadh", "Riyadh", "", "Saudi Arabia"),
    ("doha", "Doha", "", "Qatar"),
    ("sharjah", "Sharjah", "", "United Arab Emirates"),
    ("amman", "Amman", "", "Jordan"),
    ("tel aviv", "Tel Aviv", "", "Israel"),
    ("ankara", "Ankara", "", "Turkey"),
    ("izmir", "Izmir", "", "Turkey"),
    ("hanoi", "Hanoi", "", "Vietnam"),
    ("da nang", "Da Nang", "", "Vietnam"),
    ("cebu", "Cebu", "", "Philippines"),
    ("quezon city", "Quezon City", "", "Philippines"),
    ("penang", "Penang", "", "Malaysia"),
    ("johor bahru", "Johor Bahru", "", "Malaysia"),
    ("chiang mai", "Chiang Mai", "", "Thailand"),
    ("pokhara", "Pokhara", "", "Nepal"),
    ("kandy", "Kandy", "", "Sri Lanka"),
    # Exposed as unresolved by the real collected rows.
    ("trikala", "Trikala", "", "Greece"),
    ("thessaloniki", "Thessaloniki", "", "Greece"),
    ("khulna", "Khulna", "", "Bangladesh"),
    ("rajshahi", "Rajshahi", "", "Bangladesh"),
    ("recife", "Recife", "", "Brazil"),
    ("maceio", "Maceio", "", "Brazil"),
    ("fortaleza", "Fortaleza", "", "Brazil"),
    ("salvador", "Salvador", "", "Brazil"),
    ("balneario camboriu", "Balneario Camboriu", "", "Brazil"),
    ("florianopolis", "Florianopolis", "", "Brazil"),
    ("salamanca", "Salamanca", "", "Spain"),
    ("seville", "Seville", "", "Spain"),
    ("zaragoza", "Zaragoza", "", "Spain"),
    ("cebu city", "Cebu City", "", "Philippines"),
    ("davao", "Davao", "", "Philippines"),
    # Indian cities the gazetteer was missing.
    ("thrissur", "Thrissur", "Kerala", "India"),
    ("kollam", "Kollam", "Kerala", "India"),
    ("kannur", "Kannur", "Kerala", "India"),
    ("alappuzha", "Alappuzha", "Kerala", "India"),
    ("palakkad", "Palakkad", "Kerala", "India"),
    ("kottayam", "Kottayam", "Kerala", "India"),
    ("tirupur", "Tiruppur", "Tamil Nadu", "India"),
    ("erode", "Erode", "Tamil Nadu", "India"),
    ("thanjavur", "Thanjavur", "Tamil Nadu", "India"),
    ("tirunelveli", "Tirunelveli", "Tamil Nadu", "India"),
    ("tiruchirappalli", "Tiruchirappalli", "Tamil Nadu", "India"),
    ("trichy", "Tiruchirappalli", "Tamil Nadu", "India"),
    ("guntur", "Guntur", "Andhra Pradesh", "India"),
    ("nellore", "Nellore", "Andhra Pradesh", "India"),
    ("kakinada", "Kakinada", "Andhra Pradesh", "India"),
    ("rajahmundry", "Rajahmundry", "Andhra Pradesh", "India"),
    ("karimnagar", "Karimnagar", "Telangana", "India"),
    ("nizamabad", "Nizamabad", "Telangana", "India"),
    ("hubli", "Hubballi", "Karnataka", "India"),
    ("belgaum", "Belagavi", "Karnataka", "India"),
    ("davangere", "Davangere", "Karnataka", "India"),
    ("shimoga", "Shivamogga", "Karnataka", "India"),
    ("udupi", "Udupi", "Karnataka", "India"),
    ("solapur", "Solapur", "Maharashtra", "India"),
    ("kolhapur", "Kolhapur", "Maharashtra", "India"),
    ("navi mumbai", "Navi Mumbai", "Maharashtra", "India"),
    ("amravati", "Amravati", "Maharashtra", "India"),
    ("bhavnagar", "Bhavnagar", "Gujarat", "India"),
    ("jamnagar", "Jamnagar", "Gujarat", "India"),
    ("anand", "Anand", "Gujarat", "India"),
    ("ujjain", "Ujjain", "Madhya Pradesh", "India"),
    ("sagar", "Sagar", "Madhya Pradesh", "India"),
    ("satna", "Satna", "Madhya Pradesh", "India"),
    ("rewa", "Rewa", "Madhya Pradesh", "India"),
    ("bilaspur", "Bilaspur", "Chhattisgarh", "India"),
    ("bhilai", "Bhilai", "Chhattisgarh", "India"),
    ("durg", "Durg", "Chhattisgarh", "India"),
    ("gaya", "Gaya", "Bihar", "India"),
    ("bhagalpur", "Bhagalpur", "Bihar", "India"),
    ("muzaffarpur", "Muzaffarpur", "Bihar", "India"),
    ("darbhanga", "Darbhanga", "Bihar", "India"),
    ("dhanbad", "Dhanbad", "Jharkhand", "India"),
    ("bokaro", "Bokaro", "Jharkhand", "India"),
    ("rourkela", "Rourkela", "Odisha", "India"),
    ("sambalpur", "Sambalpur", "Odisha", "India"),
    ("puri", "Puri", "Odisha", "India"),
    ("meerut", "Meerut", "Uttar Pradesh", "India"),
    ("bareilly", "Bareilly", "Uttar Pradesh", "India"),
    ("aligarh", "Aligarh", "Uttar Pradesh", "India"),
    ("gorakhpur", "Gorakhpur", "Uttar Pradesh", "India"),
    ("moradabad", "Moradabad", "Uttar Pradesh", "India"),
    ("saharanpur", "Saharanpur", "Uttar Pradesh", "India"),
    ("greater noida", "Greater Noida", "Uttar Pradesh", "India"),
    ("ajmer", "Ajmer", "Rajasthan", "India"),
    ("bikaner", "Bikaner", "Rajasthan", "India"),
    ("alwar", "Alwar", "Rajasthan", "India"),
    ("panipat", "Panipat", "Haryana", "India"),
    ("hisar", "Hisar", "Haryana", "India"),
    ("karnal", "Karnal", "Haryana", "India"),
    ("sonipat", "Sonipat", "Haryana", "India"),
    ("ambala", "Ambala", "Haryana", "India"),
    ("patiala", "Patiala", "Punjab", "India"),
    ("bathinda", "Bathinda", "Punjab", "India"),
    ("roorkee", "Roorkee", "Uttarakhand", "India"),
    ("haridwar", "Haridwar", "Uttarakhand", "India"),
    ("asansol", "Asansol", "West Bengal", "India"),
    ("jorhat", "Jorhat", "Assam", "India"),
    ("silchar", "Silchar", "Assam", "India"),
    ("dibrugarh", "Dibrugarh", "Assam", "India"),
    ("imphal", "Imphal", "Manipur", "India"),
    ("shillong", "Shillong", "Meghalaya", "India"),
    ("aizawl", "Aizawl", "Mizoram", "India"),
    ("kohima", "Kohima", "Nagaland", "India"),
    ("itanagar", "Itanagar", "Arunachal Pradesh", "India"),
    ("agartala", "Agartala", "Tripura", "India"),
    ("gangtok", "Gangtok", "Sikkim", "India"),
    ("panaji", "Panaji", "Goa", "India"),
    ("jammu", "Jammu", "Jammu and Kashmir", "India"),
]
KNOWN_CITIES = {alias: (city, region, country)
          for alias, city, region, country in _CITIES_RAW}

# Names that exist in more than one country.  A bare one of these is not
# evidence of a country: "Victoria" is an Australian state AND a Canadian
# city, "Salem" is in Tamil Nadu AND in Oregon.  Each entry lists the
# plausible readings; when they disagree on the country and nothing else in
# the string settles it, the country is left blank.
AMBIGUOUS_PLACES = {
    "victoria": [("Victoria", "Victoria", "Australia"),
                 ("Victoria", "British Columbia", "Canada")],
    "cambridge": [("Cambridge", "", "United Kingdom"),
                  ("Cambridge", "Massachusetts", "United States")],
    "springfield": [("Springfield", "Illinois", "United States"),
                    ("Springfield", "Missouri", "United States")],
    "salem": [("Salem", "Tamil Nadu", "India"),
              ("Salem", "Oregon", "United States")],
    "hyderabad": [("Hyderabad", "Telangana", "India"),
                  ("Hyderabad", "Sindh", "Pakistan")],
    "birmingham": [("Birmingham", "", "United Kingdom"),
                   ("Birmingham", "Alabama", "United States")],
    "london": [("London", "", "United Kingdom"),
               ("London", "Ontario", "Canada")],
    "punjab": [("", "Punjab", "India"), ("", "Punjab (PK)", "Pakistan")],
    "georgia": [("", "Georgia", "United States"),
                ("", "", "Georgia")],
    "san jose": [("San Jose", "California", "United States"),
                 ("San Jose", "", "Costa Rica")],
    "santiago": [("Santiago", "", "Chile"),
                 ("Santiago", "", "Dominican Republic")],
    "vancouver": [("Vancouver", "British Columbia", "Canada"),
                  ("Vancouver", "Washington", "United States")],
    "newcastle": [("Newcastle", "", "United Kingdom"),
                  ("Newcastle", "New South Wales", "Australia")],
    "perth": [("Perth", "Western Australia", "Australia"),
              ("Perth", "", "United Kingdom")],
    "hamilton": [("Hamilton", "Ontario", "Canada"),
                 ("Hamilton", "", "New Zealand")],
    "richmond": [("Richmond", "Virginia", "United States"),
                 ("Richmond", "British Columbia", "Canada")],
    "aurangabad": [("Aurangabad", "Maharashtra", "India"),
                   ("Aurangabad", "Bihar", "India")],
}

# Words that are never a place.
_NOT_A_PLACE = {
    "remote", "worldwide", "earth", "internet", "everywhere", "anywhere",
    "the world", "planet earth", "localhost", "127.0.0.1", "void", "space",
    "somewhere", "n/a", "na", "unknown", "home", "here", "your city",
}


@dataclass
class Location:
    city: str = ""
    state: str = ""
    country: str = ""
    raw: str = ""
    confidence: float = 0.0
    notes: list = field(default_factory=list)

    def as_tuple(self):
        return self.city, self.state, self.country


def _key(value):
    """Lower-case, accent-folded lookup key.

    Without folding, "Maceió" and "Balneário Camboriú" lose their accented
    letters to the character filter and can never match a table entry.
    """
    value = flatten(value).lower().strip(" .,")
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^a-z0-9&' -]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _build_index():
    """Every known alias, longest first, as one scanning table.

    Real GitHub locations are not reliably comma-separated: the collected
    data contains "Nairobi Kenya", "Aurangabad Bihar" and
    "Bahawalnagar Punjab Pakistan".  Splitting on commas alone finds
    nothing in those, so aliases are matched as PHRASES anywhere in the
    string, longest first so "west java" wins over "java".
    """
    entries = []
    for alias, country in COUNTRY_ALIASES.items():
        entries.append((alias, "country", (None, None, country)))
    for alias, (region, country) in SUBDIVISIONS.items():
        # A city-state ("Delhi", "Jakarta", "Berlin") is indexed in both
        # tables.  Carry the city with it, so matching the name fills the
        # city as well as the region instead of only one of them.
        city = KNOWN_CITIES.get(alias, ("", "", ""))[0]
        entries.append((alias, "region", (city or None, region, country)))
    for alias, (city, region, country) in KNOWN_CITIES.items():
        entries.append((alias, "city", (city, region, country)))
    # An ambiguous name may appear in no other table ("Cambridge",
    # "Springfield"); without this it would never be scanned at all and the
    # ambiguity logic would never run.
    for alias, readings in AMBIGUOUS_PLACES.items():
        if alias in KNOWN_CITIES or alias in SUBDIVISIONS:
            continue
        first = readings[0]
        kind = "city" if first[0] else "region"
        entries.append((alias, kind, first))
    # Longest alias first; ties resolved alphabetically for determinism.
    entries.sort(key=lambda item: (-len(item[0]), item[0]))
    return entries


_INDEX = _build_index()
# Two-letter aliases only count as a standalone token next to something
# else we recognised -- "in", "us", "ca", "or" and "it" are ordinary words.
_SHORT_ALIAS = {alias for alias, _kind, _payload in _INDEX if len(alias) <= 2}


def _scan(text):
    """Non-overlapping alias matches in `text`, longest first."""
    found = []
    taken = []

    def overlaps(start, end):
        return any(not (end <= s or start >= e) for s, e in taken)

    for alias, kind, payload in _INDEX:
        pattern = re.compile(r"(?<![a-z0-9])%s(?![a-z0-9])"
                             % re.escape(alias))
        for match in pattern.finditer(text):
            if overlaps(match.start(), match.end()):
                continue
            taken.append((match.start(), match.end()))
            found.append({"alias": alias, "kind": kind, "payload": payload,
                          "start": match.start(), "end": match.end()})
    found.sort(key=lambda item: item["start"])
    return found


def parse_location_tuple(value):
    """(city, state, country) -- the tuple shape older callers expect.

    parse_location() returns a Location object.  This was normalizer's
    one-line adapter and it is NOT redundant: callers that unpack three
    values need this, and a Location is not iterable.
    """
    return parse_location(value).as_tuple()


def parse_location(raw, confidence=0.6):
    """Parse a location string into city / state / country.

    Enrichment runs in both directions: a known city supplies its region
    and country, and a known region supplies its country.  Country is only
    ever filled from geographic evidence found in the text itself.
    """
    result = Location(raw=flatten(raw))
    if not result.raw:
        return result
    key = _key(result.raw)
    if not key or key in _NOT_A_PLACE:
        result.notes.append("not a real place")
        return result
    parts = [part.strip() for part in re.split(r"[,/|;]|\s+-\s+",
                                               result.raw) if part.strip()]
    if parts and all(_key(part) in _NOT_A_PLACE for part in parts):
        result.notes.append("not a real place")
        return result

    matches = _scan(key)
    # Drop bare two-letter hits unless something else was recognised too.
    if len(matches) == 1 and matches[0]["alias"] in _SHORT_ALIAS:
        matches = []

    countries = {m["payload"][2] for m in matches if m["payload"][2]}
    explicit_country = next(
        (m["payload"][2] for m in matches if m["kind"] == "country"), "")
    explicit_region = next(
        (m["payload"][1] for m in matches
         if m["kind"] == "region" and m["alias"] not in AMBIGUOUS_PLACES), "")
    if explicit_region and not explicit_country:
        known = SUBDIVISIONS.get(_key(explicit_region))
        if known:
            explicit_country = known[1]

    ambiguous_hit = None
    for match in matches:
        readings = AMBIGUOUS_PLACES.get(match["alias"])
        if not readings:
            continue
        options = {country for _city, _region, country in readings}
        # Context elsewhere in the string may settle it.
        # A region named in the same string settles it: "Aurangabad Bihar"
        # is the Bihar one, not the Maharashtra one.
        settled = [r for r in readings
                   if explicit_region and r[1] == explicit_region]
        if not settled:
            settled = [r for r in readings
                       if explicit_country and r[2] == explicit_country]
        if settled:
            match["payload"] = settled[0]
            continue
        if len(options) == 1:
            match["payload"] = readings[0]
            continue
        ambiguous_hit = match
        match["payload"] = (readings[0][0] or match["alias"].title(), "", "")

    for match in matches:
        city, region, country = match["payload"]
        if match["kind"] == "country":
            result.country = result.country or country
            continue
        # A resolved ambiguous reading can carry a city even when the alias
        # was indexed as a region ("Victoria, Canada" is the BC city).
        result.city = result.city or (city or "")
        result.state = result.state or (region or "")
        result.country = result.country or (country or "")

    if ambiguous_hit and not explicit_country and not explicit_region:
        options = {country
                   for _city, _region, country
                   in AMBIGUOUS_PLACES[ambiguous_hit["alias"]]}
        if len(options) > 1:
            # More than one country is plausible and nothing in the string
            # settles it, so no country is better than a coin flip.
            result.country = ""
            result.state = ""
            result.notes.append(
                "ambiguous place %r: could be %s"
                % (ambiguous_hit["alias"], ", ".join(sorted(options))))

    # Whatever no table recognised, taken from the first comma-part, is the
    # most likely city name.
    if not result.city:
        for part in parts:
            leftover = _key(part)
            if not leftover or leftover in _NOT_A_PLACE:
                continue
            for match in matches:
                leftover = leftover.replace(match["alias"], " ")
            leftover = re.sub(r"\s+", " ", leftover).strip()
            if (2 <= len(leftover) <= 40 and not leftover.isdigit()
                    and len(leftover.split()) <= 3):
                result.city = part.strip(" .").title() \
                    if part.strip(" .").islower() else part.strip(" .")
                result.notes.append("city not in gazetteer")
                break

    # Final enrichment pass: city -> region -> country.
    if result.city and (not result.state or not result.country):
        known = KNOWN_CITIES.get(_key(result.city))
        if known and _key(result.city) not in AMBIGUOUS_PLACES:
            result.state = result.state or known[1]
            result.country = result.country or known[2]
    if result.state and not result.country:
        known = SUBDIVISIONS.get(_key(result.state))
        if known:
            result.country = known[1]

    if result.city or result.state or result.country:
        result.confidence = confidence
    return result


def best_location(candidates):
    """Pick the most trustworthy location from (value, confidence) pairs.

    Candidates arrive in source-priority order; a later one is only used to
    fill in a field the earlier one left blank.
    """
    best = Location()
    for raw, score in candidates:
        if not raw:
            continue
        parsed = parse_location(raw, confidence=score)
        if not (parsed.city or parsed.state or parsed.country):
            continue
        if not best.raw:
            best = parsed
            continue
        best.city = best.city or parsed.city
        best.state = best.state or parsed.state
        best.country = best.country or parsed.country
        for note in parsed.notes:
            if note not in best.notes:
                best.notes.append(note)

    # Final enrichment across the merged result: a city contributed by one
    # source can still supply the region and country another source lacked.
    if best.city and (not best.state or not best.country):
        known = KNOWN_CITIES.get(_key(best.city))
        if known and _key(best.city) not in AMBIGUOUS_PLACES:
            best.state = best.state or known[1]
            best.country = best.country or known[2]
    if best.state and not best.country:
        known = SUBDIVISIONS.get(_key(best.state))
        if known:
            best.country = known[1]
    return best


# ==========================================================================
# MODELS
# ==========================================================================
# Data carriers shared by the whole pipeline.
# --------------------------------------------------------------------------
@dataclass
class Candidate:
    """One row of an Instant Data Scraper XLSX, before any web work."""

    source_row_number: int = 0
    source_file: str = ""
    keyword: str = ""

    username: str = ""
    profile_url: str = ""

    xlsx_name: str = ""
    xlsx_bio: str = ""
    xlsx_location: str = ""

    raw_cells: list = field(default_factory=list)

    def merge(self, other):
        """Fold a duplicate row for the same username into this one."""
        self.xlsx_name = self.xlsx_name or other.xlsx_name
        self.xlsx_location = self.xlsx_location or other.xlsx_location
        if other.xlsx_bio and other.xlsx_bio not in self.xlsx_bio:
            self.xlsx_bio = (self.xlsx_bio + " " + other.xlsx_bio).strip()
        for cell in other.raw_cells:
            if cell not in self.raw_cells:
                self.raw_cells.append(cell)


@dataclass
class SourceBundle:
    """Everything publicly gathered for one person, before extraction."""

    candidate: Candidate = None

    github_ok: bool = False
    github_permanent_fail: bool = False
    github_fields: dict = field(default_factory=dict)
    github_text: str = ""
    github_links: list = field(default_factory=list)

    portfolio_url: str = ""
    portfolio_text: str = ""
    portfolio_links: list = field(default_factory=list)
    portfolio_pages: list = field(default_factory=list)

    resume_url: str = ""
    resume_file: str = ""
    resume_text: str = ""

    remarks: list = field(default_factory=list)

    def note(self, message):
        if message and message not in self.remarks:
            self.remarks.append(message)

    def sources_used(self):
        used = []
        if self.github_ok:
            used.append("GitHub")
        if self.portfolio_text.strip():
            used.append("Portfolio")
        if self.resume_text.strip():
            used.append("Resume")
        return used or ["Instant Data Scraper"]


# ==========================================================================
# PROGRESS
# ==========================================================================
# Terminal progress rendering: live countdowns, cooldown display, phase cards.
#
# Presentation ONLY.  This module never sleeps, never decides how long to wait,
# and never touches scraping / pacing / cooldown policy -- callers pass in the
# seconds and the facts, and this decides only how to DRAW them.  It is safe to
# change freely: nothing here alters behaviour.
#
# Two output modes, chosen from the stream itself:
#
#   * a real terminal (isatty)     -> one line updated in place with '\\r'
#   * anything else (a log file,   -> a simple start/finish message, no animation
#     redirected output, CI, tests)   so runner.log never fills with tick spam.
#
# A countdown's TICKS are terminal-only and are NEVER written to a log; only the
# start and finish of a wait are logged.  That keeps log files small while the
# terminal still shows a live count.  The underlying interruptible sleep keeps
# stepping once a second regardless, so Ctrl+C stays responsive even when a long
# cooldown only redraws every 30 seconds.
# --------------------------------------------------------------------------


def is_tty(stream):
    """True only for a real interactive terminal."""
    try:
        return bool(stream.isatty())
    except Exception:
        return False


def fmt_clock(seconds):
    """Whole seconds as MM:SS (or H:MM:SS once past an hour)."""
    seconds = int(max(0, round(seconds)))
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    if hours:
        return "%d:%02d:%02d" % (hours, minutes, secs)
    return "%02d:%02d" % (minutes, secs)


def fmt_hms(seconds):
    """Whole seconds as HH:MM:SS (used for the long-cooldown 'Remaining')."""
    seconds = int(max(0, round(seconds)))
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    return "%02d:%02d:%02d" % (hours, minutes, secs)


def refresh_for(total_seconds):
    """How often (seconds) to REDRAW a countdown of this length.

    Short pacing waits tick every second; a cooldown up to ten minutes redraws
    every ten seconds; a longer one every thirty, so a two-hour wait does not
    scroll the screen.  This governs only the redraw cadence -- the sleep still
    steps every second underneath, so Ctrl+C is never delayed.
    """
    if total_seconds <= 60:
        return 1
    if total_seconds <= 600:
        return 10
    return 30


class Countdown:
    """A single-line live countdown for a normal pacing wait.

        cd = Countdown(seconds, "SEARCH")
        cd.start()
        await sleep_interruptible(seconds, check_stop, on_tick=cd.tick)
        cd.finish()

    `tick` is exactly the on_tick the interruptible sleep already calls once per
    step; it represents the EXISTING delay and adds no sleeping of its own.  It
    redraws nothing unless at least `refresh` seconds have passed since the last
    draw.
    """

    def __init__(self, total, tag, prefix="[PACING]", stream=None,
                 logger=None, refresh=None):
        self.total = float(total)
        self.tag = tag
        self.stream = stream if stream is not None else sys.stdout
        self.tty = is_tty(self.stream)
        self.log = logger or log
        self.refresh = refresh or refresh_for(self.total)
        self._label = "%s[%s]" % (prefix, tag)
        self._last_sec = None
        self._width = 0

    def start(self):
        self.log.info("%s wait started: %.1fs", self._label, self.total)
        if not self.tty:
            print("%s waiting %s..." % (self._label, fmt_clock(self.total)),
                  file=self.stream, flush=True)

    def tick(self, remaining):
        if not self.tty:
            return
        sec = int(remaining)
        if sec <= 0:
            return
        if (self._last_sec is not None
                and self._last_sec - sec < self.refresh):
            return
        self._last_sec = sec
        self._draw("%s %s remaining" % (self._label, fmt_clock(sec)))

    def finish(self):
        self.log.info("%s wait finished", self._label)
        if self.tty:
            self._draw("%s Done. Continuing..." % self._label)
            self.stream.write("\n")
            self.stream.flush()
        else:
            print("%s done." % self._label, file=self.stream, flush=True)

    def _draw(self, text):
        pad = max(0, self._width - len(text))
        self.stream.write("\r" + text + (" " * pad))
        self.stream.flush()
        self._width = len(text)


class CooldownDisplay:
    """The multi-line rate-limit cooldown block with an updating 'Remaining'.

    On a terminal the header is printed once and only the last line (Remaining)
    is rewritten in place; elsewhere the block is printed once and left alone.
    Redraw cadence follows refresh_for(total) -- 10s up to ten minutes, 30s
    beyond -- so a long cooldown never scrolls.
    """

    def __init__(self, level, reason, retry_at, total, stream=None,
                 logger=None, refresh=None):
        self.level = level
        self.reason = reason or "GitHub rate limit"
        self.retry_at = retry_at or "-"
        self.total = float(total)
        self.stream = stream if stream is not None else sys.stdout
        self.tty = is_tty(self.stream)
        self.log = logger or log
        self.refresh = refresh or refresh_for(self.total)
        self._last_sec = None
        self._width = 0

    def start(self):
        minutes = int(round(self.total / 60.0))
        self.log.info("[RATE LIMIT] cooldown started: level %d, %d minute(s)",
                      self.level, minutes)
        s = self.stream
        print("", file=s)
        print("[RATE LIMIT COOLDOWN]", file=s)
        print("Level      : %d" % self.level, file=s)
        print("Reason     : %s" % self.reason, file=s)
        print("Retry at   : %s" % self.retry_at, file=s)
        line = "Remaining  : %s" % fmt_hms(self.total)
        if self.tty:
            self._draw(line)
        else:
            print(line, file=s, flush=True)

    def tick(self, remaining):
        if not self.tty:
            return
        sec = int(remaining)
        if sec < 0:
            sec = 0
        if (self._last_sec is not None
                and self._last_sec - sec < self.refresh and sec > 0):
            return
        self._last_sec = sec
        self._draw("Remaining  : %s" % fmt_hms(sec))

    def finish(self):
        self.log.info("[RATE LIMIT] cooldown finished")
        if self.tty:
            self._draw("Remaining  : %s" % fmt_hms(0))
            self.stream.write("\n")
            self.stream.flush()
        print("[RATE LIMIT] Cooldown finished.", file=self.stream, flush=True)

    def _draw(self, text):
        pad = max(0, self._width - len(text))
        self.stream.write("\r" + text + (" " * pad))
        self.stream.flush()
        self._width = len(text)


# ------------------------------------------------------ retry / phase cards ---
def retry_banner(source, query="", page=None, username=""):
    """What is about to be retried once a cooldown finishes (req 6).

    A plain multi-line string; the caller prints it through its own say()/log
    so it appears once, not per tick.
    """
    lines = ["[RATE LIMIT] Retrying same operation..."]
    if source == "pagination":
        lines.append("[RATE LIMIT] Retry pagination:")
        lines.append("Search: %s" % query)
        if page is not None:
            lines.append("Page: %s" % page)
    elif source == "profile":
        # A profile rate limit retries the whole SEARCH from its saved resume
        # point, so say that rather than implying only one profile is redone.
        lines.append("[RATE LIMIT] Retry search (a profile was "
                     "rate-limited):")
        if query:
            lines.append(query)
        if username:
            lines.append("Profile: %s" % username)
    else:
        lines.append("[RATE LIMIT] Retry search:")
        lines.append(query)
    return "\n".join(lines)


def rank_card(rank, total_ranks, search_no, search_total, query,
              students, deferred):
    """Compact per-rank context, printed at a search transition (req 4)."""
    return "\n".join([
        "[RANK %02d/%02d]" % (rank, total_ranks),
        "Search      : %d/%d" % (search_no, search_total),
        "Query       : %s" % query,
        "Students    : %d" % students,
        "Deferred    : %d" % deferred,
    ])


def profile_card(candidate_no, candidate_total, username, search_no,
                 search_total):
    """Compact profile-phase context (req 4)."""
    return "\n".join([
        "[PROFILE]",
        "Candidate   : %d/%d" % (candidate_no, candidate_total),
        "Username    : %s" % username,
        "Search      : %d/%d" % (search_no, search_total),
    ])


def crawl_card(search_no, search_total, page, pages, collected, expected):
    """Compact pagination context (req 4)."""
    page_text = "%d/%d" % (page, pages) if pages else str(page)
    return "\n".join([
        "[CRAWL]",
        "Search      : %d/%d" % (search_no, search_total),
        "Page        : %s" % page_text,
        "Collected   : %d" % collected,
        "Expected    : %s" % (expected if expected else "?"),
    ])


def _ranges(numbers):
    """Collapse a sorted rank list into compact ranges: [1,2,3,5] -> '1-3, 5'."""
    numbers = sorted(set(numbers))
    if not numbers:
        return "-"
    spans = []
    start = prev = numbers[0]
    for value in numbers[1:]:
        if value == prev + 1:
            prev = value
            continue
        spans.append((start, prev))
        start = prev = value
    spans.append((start, prev))
    return ", ".join("%d" % a if a == b else "%d-%d" % (a, b)
                     for a, b in spans)


def device_card(owned, completed, current_rank, total_ranks,
                rank_progress="", students=0):
    """The device-wide progress block for the distributed runner (req 5).

    `owned`/`completed` are rank lists; `current_rank` is the rank now running
    (or None); `rank_progress` is that rank's own "done/total" search count if
    known.  Returns a multi-line string; the runner prints it at each rank
    transition, not continuously.
    """
    owned = sorted(owned)
    completed = sorted(set(completed))
    pending = [r for r in owned if r not in set(completed)]
    lines = [
        "=" * 60,
        "DEVICE PROGRESS",
        "=" * 60,
        "Owned ranks       : %s" % _ranges(owned),
        "Finished ranks    : %d/%d" % (len(completed), len(owned)),
    ]
    if current_rank:
        lines.append("Current rank      : %d/%d" % (current_rank, total_ranks))
        lines.append("Rank progress     : %s" % (rank_progress or "-"))
    lines.append("Students collected: %d" % students)
    lines.append("")
    lines.append("Completed:")
    lines.append(_ranges(completed) if completed else "-")
    lines.append("")
    lines.append("Pending:")
    lines.append(_ranges(pending) if pending else "-")
    lines.append("=" * 60)
    return "\n".join(lines)


# ==========================================================================
# RATELIMIT
# ==========================================================================
# Conservative rate-limit cooldown -- respect GitHub, never bypass it.
#
# Normal pacing (a few tens of seconds between searches, a few seconds between
# pages) is not enough once GitHub actually returns a rate-limit/challenge.  Then
# this escalating, PERSISTED cooldown takes over:
#
# every hit waits a FRESH RANDOM draw of RATE_LIMIT_COOLDOWN_MIN..MAX minutes
# (top of this file), then retries the SAME search from its saved resume
# point.  Hit 1 and hit 20 are treated alike -- there is no ladder that runs
# out and no level that stops the device, so an unattended multi-day run
# survives any number of rate limits:
#
#     rate limit -> wait MIN..MAX min -> retry -> rate limit -> wait again...
#
# The escalation LEVEL is still counted and shown, because a level climbing
# into the dozens is how the log tells you a limit is not clearing; it just
# no longer selects the wait or ends the run.
#
# The cooldown is rank-local (runtime/rank_NNN/rate_limit_state.json) and
# persisted, so restarting the command cannot bypass it.  A server-provided
# Retry-After is taken AS GIVEN (config.TRUST_RETRY_AFTER) -- GitHub stating
# its own reset time is better information than our own draw, in either
# direction.  A successful retry of the blocked operation resets the
# escalation to 0.
#
# This module NEVER circumvents a restriction: it only decides to wait or stop.
# --------------------------------------------------------------------------
def cooldown_seconds(level):
    """A fresh random cooldown in SECONDS, drawn MIN..MAX minutes.

    Read RATE_LIMIT_COOLDOWN_MIN / RATE_LIMIT_COOLDOWN_MAX from the top of
    this file, exactly like the pacing delays and LAP_COOLDOWN_MINUTES, and
    draw a new wait on every call -- so two rate limits in a row do not wait
    the same number of minutes, and the 20th waits no longer than the 1st.

    `level` is still accepted and still decides NOTHING but "is there a
    cooldown at all": level < 1 means none and returns 0.  The level is kept
    for the log and the on-screen display.
    """
    if level < 1:
        return 0
    low = int(RATE_LIMIT_COOLDOWN_MIN)
    high = int(RATE_LIMIT_COOLDOWN_MAX)
    if high < low:                     # a swapped pair is an operator typo,
        low, high = high, low          # not a reason to crash a 3-day run
    return random.randrange(low, high + 1) * 60


def _parse(text):
    try:
        return datetime.datetime.fromisoformat(text)
    except Exception:
        return None


def parse_retry_after(value, now=None):
    """Seconds to wait from a GitHub Retry-After header value, or None.

    Accepts either a delta-seconds integer or an HTTP date; a date in the past
    yields 0.  Never negative.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.isdigit():
        return max(0, int(text))
    try:
        from email.utils import parsedate_to_datetime
        when = parsedate_to_datetime(text)
    except Exception:
        return None
    if when is None:
        return None
    reference = now or datetime.datetime.now(when.tzinfo)
    return max(0, int((when - reference).total_seconds()))


class RateLimitCooldown:
    """Rank-local, persisted rate-limit escalation + cooldown window."""

    def __init__(self, path=None):
        self.path = path or C.RATE_LIMIT_FILE
        self.data = self._load()

    def _load(self):
        blank = {"level": 0, "reason": "", "source": "",
                 "cooldown_until": "", "last_hit_at": "", "retry_after": None}
        if not os.path.exists(self.path):
            return blank
        try:
            with open(self.path, encoding="utf-8") as handle:
                loaded = json.load(handle)
        except Exception as exc:
            log.warning("rate-limit state unreadable (%s); starting fresh", exc)
            return blank
        for key, value in blank.items():
            loaded.setdefault(key, value)
        return loaded

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, self.path)

    # ------------------------------------------------------------------
    @property
    def level(self):
        return int(self.data.get("level", 0) or 0)

    @property
    def source(self):
        return self.data.get("source", "") or ""

    @property
    def reason(self):
        return self.data.get("reason", "") or ""

    def cooldown_until(self):
        return _parse(self.data.get("cooldown_until", "") or "")

    def _now(self, now=None):
        return now or datetime.datetime.now()

    def active(self, now=None):
        until = self.cooldown_until()
        return until is not None and until > self._now(now)

    def remaining(self, now=None):
        until = self.cooldown_until()
        if until is None:
            return 0
        return max(0, int((until - self._now(now)).total_seconds()))

    @property
    def kind(self):
        """"rate-limit" (the default) or "outage" (a GitHub 5xx)."""
        return self.data.get("kind", "") or "rate-limit"

    def exhausted(self):
        """True once escalation has passed the last defined level.

        NEVER true for an outage.  A 5xx is GitHub being broken, not GitHub
        restricting us, so the run waits at the top of the ladder and keeps
        retrying instead of stopping the device.  The kind is persisted, so a
        restart in the middle of an outage does not become a stop either."""
        if self.kind == "outage":
            return False
        if not C.STOP_WHEN_COOLDOWN_EXHAUSTED:
            # Same treatment as an outage: wait at the top of the ladder and
            # keep retrying rather than ending an unattended run.  See
            # C.STOP_WHEN_COOLDOWN_EXHAUSTED for the trade-off.
            return False
        return self.level > C.MAX_COOLDOWN_LEVEL

    def escalate(self, source, reason, retry_after=None, now=None,
                 kind="rate-limit"):
        """Record a fresh rate-limit hit: bump the level and set the cooldown
        window.  Returns the new level.  A server Retry-After (seconds) wins
        outright when config.TRUST_RETRY_AFTER is set.

        kind="outage" marks a GitHub 5xx, which never exhausts: the wait
        stays at the top of the ladder (cooldown_seconds clamps there) and
        the unit is retried for as long as the outage lasts."""
        now = self._now(now)
        new_level = self.level + 1
        seconds = cooldown_seconds(new_level)
        if retry_after is not None:
            # GitHub stating its own reset time is better information than our
            # ladder, so it is used as given when TRUST_RETRY_AFTER is set --
            # including when it is SHORTER.  max() overrode the server in the
            # wrong direction: a "retry after 60 seconds" waited 15 minutes.
            # Direct attribute access on purpose: a getattr default would
            # turn a deleted constant into a silently disabled feature.
            if C.TRUST_RETRY_AFTER:
                seconds = max(1, int(retry_after))
            else:
                seconds = max(seconds, int(retry_after))
        until = now + datetime.timedelta(seconds=seconds)
        self.data.update({
            "level": new_level, "source": source, "reason": str(reason)[:300],
            "last_hit_at": now.isoformat(timespec="seconds"),
            "cooldown_until": until.isoformat(timespec="seconds"),
            "retry_after": int(retry_after) if retry_after is not None else None,
            "kind": kind,
        })
        self.save()
        log.warning("%s cooldown level %d (%s): %s; until %s",
                    kind, new_level, source, reason,
                    self.data["cooldown_until"])
        return new_level

    def clear(self):
        """Reset the escalation after the blocked operation finally works."""
        if self.level or self.data.get("cooldown_until"):
            self.data.update({"level": 0, "reason": "", "source": "",
                              "cooldown_until": "", "last_hit_at": "",
                              "retry_after": None, "kind": ""})
            self.save()
            return True
        return False


async def sleep_interruptible(seconds, check_stop=None, step=1.0, sleep=None,
                              on_tick=None):
    """asyncio.sleep(`seconds`) that notices a Ctrl+C between short steps.

    `check_stop` (a no-arg callable) is invoked before each step and may raise
    to abort the wait cleanly.  `on_tick`, when given, is called with the
    seconds still remaining before each step -- a display hook only, it adds no
    sleeping and cannot change the total wait.  `sleep` overrides the sleep
    primitive for tests.
    """
    remaining = float(seconds)
    naptime = sleep or asyncio.sleep
    while remaining > 0:
        if check_stop is not None:
            check_stop()
        if on_tick is not None:
            on_tick(remaining)
        await naptime(min(step, remaining))
        remaining -= step


async def sleep_jitter(base_seconds, check_stop=None, label="", say=None):
    """A computed wait (a retry backoff) spread up by 0..JITTER_FRACTION, so
    copies retrying the same minute do not retry in lockstep."""
    base = float(base_seconds or 0)
    if base <= 0:
        return 0.0
    delay = base * (1.0 + random.uniform(0.0, float(JITTER_FRACTION)))
    if label and say is not None:
        say("%s %.1fs" % (label, delay))
    await sleep_interruptible(delay, check_stop)
    return delay


# ==========================================================================
# THROTTLE CIRCUIT
# ==========================================================================
# Global-throttle circuit breaker: tell an account-wide block from a bad page.
#
# ONE unavailable profile is a profile problem -- it gets deferred and the rank
# keeps going.  But when several DIFFERENT profiles all hit a GitHub
# rate-limit/challenge in a short window, that is an account/session/IP throttle,
# and continuing to probe GitHub only deepens it.  This records those events
# (rank-local, persisted, time-windowed) and reports when the breaker has
# tripped, so the rank -- and the whole device -- can stop and respect the block.
#
# It NEVER bypasses GitHub's restrictions: it only decides to STOP sooner.  A
# successful profile load clears the streak; the file persists across restarts so
# rerunning every few minutes cannot slip past the breaker.
# --------------------------------------------------------------------------
# Reasons that count as a GitHub rate-limit / challenge (as opposed to an
# ordinary timeout or a one-off 5xx).  Matched case-insensitively against the
# failure reason string.
_THROTTLE_MARKERS = (
    "rate limit", "secondary rate limit", "abuse detection",
    "temporarily blocked", "unusual traffic", "verify your identity",
    "security check", "too many requests", "challenge",
    "unable to load", "please reload page", "http 429", "http 403",
    "returned http 429", "returned http 403",
)


def is_throttle_reason(reason):
    """True when a failure reason indicates a GitHub rate-limit/challenge."""
    text = (reason or "").lower()
    return any(marker in text for marker in _THROTTLE_MARKERS)


# Reasons that mean GitHub ITSELF is broken (5xx), as opposed to GitHub
# throttling us.  Deliberately NOT added to _THROTTLE_MARKERS: a one-off 5xx
# must not count toward the global throttle circuit breaker.  Checked BEFORE
# the throttle markers, because GitHub's SPA wraps a 503 in the same generic
# "Unable to load page" screen it uses for a 429 -- the explicit 5xx is the
# stronger signal.  A wrapped 429 carries no 5xx marker, so it stays a rate
# limit.
_OUTAGE_MARKERS = tuple(
    ["http %d" % code for code in sorted(C.OUTAGE_STATUS)]
    + ["status: %d" % code for code in sorted(C.OUTAGE_STATUS)]
    + ["bad gateway", "service unavailable", "internal server error",
       "server error", "unicorn"])


def is_outage_reason(reason):
    """True when a failure reason indicates a GitHub-side outage (5xx)."""
    text = (reason or "").lower()
    return any(marker in text for marker in _OUTAGE_MARKERS)




class ThrottleCircuit:
    """Rank-local, persisted, windowed record of rate-limit/challenge events."""

    def __init__(self, path=None, distinct=None, window_minutes=None):
        self.path = path or C.THROTTLE_FILE
        self.distinct_limit = (C.GLOBAL_THROTTLE_DISTINCT_PROFILES
                               if distinct is None else distinct)
        self.window = datetime.timedelta(
            minutes=(C.GLOBAL_THROTTLE_WINDOW_MINUTES
                     if window_minutes is None else window_minutes))
        self.data = self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return {"version": 1, "events": []}
        try:
            with open(self.path, encoding="utf-8") as handle:
                loaded = json.load(handle)
        except Exception as exc:
            log.warning("throttle state unreadable (%s); starting fresh", exc)
            return {"version": 1, "events": []}
        loaded.setdefault("version", 1)
        loaded.setdefault("events", [])
        return loaded

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, self.path)

    # ------------------------------------------------------------------
    def _now(self, now=None):
        return now or datetime.datetime.now()

    def _recent(self, now=None):
        cutoff = self._now(now) - self.window
        fresh = []
        for event in self.data.get("events", []):
            moment = _parse(event.get("at", ""))
            if moment is not None and moment >= cutoff:
                fresh.append(event)
        return fresh

    def record_throttle(self, username, now=None):
        """Record ONE rate-limit/challenge event.  Prunes stale events and
        persists.  Returns the number of DISTINCT profiles in the window."""
        events = self._recent(now)
        events.append({"username": username or "",
                       "at": self._now(now).isoformat(timespec="seconds")})
        self.data["events"] = events
        self.save()
        distinct = self.distinct_recent(now)
        log.warning("throttle event %s (%d distinct profile(s) in window)",
                    username, distinct)
        return distinct

    def record_success(self, now=None):
        """A profile loaded normally -> the account is not blocked right now, so
        the streak is cleared."""
        if self.data.get("events"):
            self.data["events"] = []
            self.save()

    def streak_usernames(self, now=None):
        seen = []
        for event in self._recent(now):
            user = (event.get("username") or "").lower()
            if user and user not in seen:
                seen.append(user)
        return seen

    def distinct_recent(self, now=None):
        return len(self.streak_usernames(now))

    def tripped(self, now=None):
        return self.distinct_recent(now) >= self.distinct_limit


# ==========================================================================
# XLSX READER
# ==========================================================================
# Reads one Instant Data Scraper XLSX into Candidate rows.
#
# Nothing is read by fixed position (spec 19).  Instant Data Scraper names its
# columns after the CSS classes it found, e.g.
#
#     prc-Avatar-Avatar-0xaUi src | prc-Link-Link-9ZwDx href | search-match | ...
#
# Those hashes change whenever GitHub ships CSS, and the layout differs between
# a user search and a repository search, so every cell is classified by its own
# content instead.
#
# GitHub also splits a highlighted search term into its own column, so the bio
# arrives in pieces:
#
#     "3rd year student at"  |  "LNCT"  |  "BHOPAL."
#
# Every meaningful text cell is therefore preserved in raw_cells and the bio is
# rebuilt by joining the fragments in sheet order.
# --------------------------------------------------------------------------
AVATAR_RE = re.compile(
    r"github(?:usercontent)?\.com/(?:u/)?([A-Za-z0-9][A-Za-z0-9-]{0,38})\.png",
    re.I)
PROFILE_RE = re.compile(
    r"^https?://(?:www\.)?github\.com/([A-Za-z0-9][A-Za-z0-9-]{0,38})/?$",
    re.I)
REPO_RE = re.compile(
    r"^https?://(?:www\.)?github\.com/([A-Za-z0-9][A-Za-z0-9-]{0,38})"
    r"/([^/\s?#]+)", re.I)
ANY_GH_RE = re.compile(
    r"https?://(?:www\.)?github\.com/([A-Za-z0-9][A-Za-z0-9-]{0,38})"
    r"(?:[/?#]|$)", re.I)

# Cells that are page furniture rather than information about the person.
NOISE_CELLS = {
    "·", "-", "updated", "sponsor", "star", "follow", "public", "private",
    "archived", "fork", "forked", "loading",
}
NOISE_PREFIX_RE = re.compile(r"^(updated\s+on\b|on\s+\w{3}\s+\d)", re.I)

# Repository exports carry a language label where a user export carries a
# location; these must never be mistaken for a city.
LANGUAGE_LABELS = {
    "python", "java", "javascript", "typescript", "c", "c++", "c#", "go",
    "rust", "php", "ruby", "kotlin", "swift", "dart", "html", "css", "scss",
    "shell", "bash", "powershell", "jupyter notebook", "r", "matlab",
    "objective-c", "scala", "perl", "lua", "haskell", "elixir", "clojure",
    "vue", "svelte", "dockerfile", "makefile", "cmake", "assembly", "vim script",
    "batchfile", "tex", "sql", "plpgsql", "hack", "groovy", "julia", "solidity",
}

LOCATION_HEADER_HINTS = ("location", "homelocation")
NAME_HEADER_HINTS = ("titlename", "fullname", "vcard-fullname")
LOGIN_HEADER_HINTS = ("titlelogin", "username", "vcard-username")


def valid_username(name):
    if not name:
        return False
    if name.lower() in C.RESERVED_PATHS:
        return False
    if name.startswith("-") or name.endswith("-"):
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9-]{0,38}", name))


def _looks_like_location(text):
    """City-ish: short, comma-separated words, no sentence punctuation."""
    if not text or len(text) > 60:
        return False
    if any(ch in text for ch in "@/|") or text.count(".") > 1:
        return False
    words = text.replace(",", " ").split()
    if not 1 <= len(words) <= 5:
        return False
    return all(word[0].isupper() or word.isupper()
               for word in words if word[:1].isalpha())


def _is_noise(text):
    low = text.strip().lower()
    if not low or low in NOISE_CELLS:
        return True
    if NOISE_PREFIX_RE.match(low):
        return True
    return low.isdigit() and len(low) <= 6


class XlsxReader:
    """Turns an export file into ordered, de-duplicated Candidate rows."""

    def read(self, path, keyword):
        source_file = os.path.basename(path)
        try:
            book = load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            log.error("cannot open export %s: %s", source_file, exc)
            return []

        candidates = []
        by_username = {}
        row_number = 0
        try:
            for sheet in book.worksheets:
                headers = []
                for raw in sheet.iter_rows(values_only=True):
                    cells = ["" if value is None else str(value).strip()
                             for value in raw]
                    if not headers:
                        headers = [cell.lower() for cell in cells]
                        continue
                    if not any(cells):
                        continue
                    row_number += 1
                    candidate = self._to_candidate(headers, cells, keyword,
                                                   source_file, row_number)
                    if candidate is None:
                        continue
                    key = candidate.username.lower()
                    if key in by_username:
                        by_username[key].merge(candidate)
                        continue
                    by_username[key] = candidate
                    candidates.append(candidate)
        finally:
            book.close()
        return candidates

    # ------------------------------------------------------------------
    def _to_candidate(self, headers, cells, keyword, source_file, row_number):
        username = self._username_from(cells)
        if not username:
            return None

        name, location = "", ""
        parts = []                      # (header, cell) in sheet order
        for position, cell in enumerate(cells):
            if not cell or cell.lower().startswith("http"):
                continue
            header = headers[position] if position < len(headers) else ""
            if any(hint in header for hint in LOCATION_HEADER_HINTS):
                location = location or cell
                continue
            if any(hint in header for hint in NAME_HEADER_HINTS):
                name = name or cell
                continue
            if any(hint in header for hint in LOGIN_HEADER_HINTS):
                continue
            # The login is already known; repeating it is not biography.
            if cell.lower() == username.lower():
                continue
            if _is_noise(cell):
                continue
            parts.append((header, cell))
        bio_parts = [cell for _header, cell in parts]

        # Shape-based fallbacks are only safe on a USER search.  A
        # repository export has no person name and no location -- its
        # equivalent columns hold a repo description and a language label --
        # so guessing there would invent data.  A row counts as a user row
        # only when some cell is a bare https://github.com/<user> URL.
        if self._is_user_row(cells):
            if not name:
                for cell in bio_parts:
                    if "/" in cell or cell.lower() == username.lower():
                        continue
                    name = cell
                    break
            if not location:
                for header, cell in reversed(parts):
                    # "search-match" columns are highlighted *bio* fragments
                    # by construction, so a city can never be read out of
                    # one; doing so turns "...at LNCT BHOPAL." into a
                    # location.
                    if header.startswith("search-match"):
                        continue
                    if cell == name or cell.lower() in LANGUAGE_LABELS:
                        continue
                    if _looks_like_location(cell):
                        location = cell
                        break
        if location.lower() in LANGUAGE_LABELS:
            location = ""

        # Drop the owner/repo slug: it is a repository id, not biography.
        slug = re.compile(r"^%s/" % re.escape(username), re.I)
        bio = " ".join(part for part in bio_parts
                       if part not in (name, location)
                       and not slug.match(part)).strip()
        bio = re.sub(r"\s+", " ", bio)

        raw_cells = [cell for cell in cells if cell]
        return Candidate(
            source_row_number=row_number,
            source_file=source_file,
            keyword=keyword,
            username=username,
            profile_url="https://github.com/%s" % username,
            xlsx_name=name,
            xlsx_bio=bio,
            xlsx_location=location,
            raw_cells=raw_cells,
        )

    @staticmethod
    def _is_user_row(cells):
        """True when the row links to a profile, not to a repository."""
        return any(PROFILE_RE.match(cell) for cell in cells)

    @staticmethod
    def _username_from(cells):
        """Canonical account, tried from the most reliable evidence down."""
        # 1. the avatar URL is always the account itself
        for cell in cells:
            match = AVATAR_RE.search(cell)
            if match and valid_username(match.group(1)):
                return match.group(1)
        # 2. a bare profile URL (user search)
        for cell in cells:
            match = PROFILE_RE.match(cell)
            if match and valid_username(match.group(1)):
                return match.group(1)
        # 3. the owner of a repository URL (repository search)
        for cell in cells:
            match = REPO_RE.match(cell)
            if match and valid_username(match.group(1)):
                return match.group(1)
        # 4. any github.com link at all
        for cell in cells:
            match = ANY_GH_RE.search(cell)
            if match and valid_username(match.group(1)):
                return match.group(1)
        return ""


# ==========================================================================
# PAGE CRAWLER
# ==========================================================================
# Page-by-page collection: one fresh Instant Data Scraper per results page.
#
# WHY THIS REPLACES "START CRAWLING" FOR MULTI-PAGE SEARCHES
# ---------------------------------------------------------
# The diagnostic settled it: GitHub REPLACES the results DOM when it paginates,
# and the popup stays bound to the page it was opened on.  So IDS's own
# cross-page crawl loses its table the moment it turns a page -- while a FRESH
# popup opened on page 2 detects all 9 users there without difficulty.
#
#     page 1   fresh IDS -> 10 rows  -> export
#     Next     Playwright clicks, waits for the DOM to settle
#     page 2   fresh IDS ->  9 rows  -> export
#     merge    19 unique rows -> ONE keyword XLSX
#
# Instant Data Scraper is still the extractor.  Playwright only does what it is
# better at: deterministic navigation and synchronisation.
#
# WHY PAGES MUST BE EXPORTED, NOT ACCUMULATED
# -------------------------------------------
# A popup holds its collected rows in memory only for as long as it lives, so
# reinitialising IDS for page 2 necessarily discards page 1.  Page 1 therefore
# has to be written out BEFORE the popup is replaced.  That is not a design
# preference; it is forced by how the extension works.
#
# ATOMICITY
# ---------
# Page exports go to a rank-local staging directory:
#
#     runtime/rank_NNN/crawl_staging/<search-slug>/page_001.xlsx
#                                                  page_002.xlsx
#
# The final keyword XLSX is written ONCE, after every page has succeeded.  If
# any page fails, staging is preserved, no keyword XLSX is created, and the
# search unit stays incomplete.  SUCCESS -> cleanup, FAILURE -> preserve, the
# same rule as everywhere else.
#
# RETRY
# -----
# A retry restarts from page 1 (option B).  Staging is archived under an attempt
# id first, so a fresh crawl can never merge yesterday's page_002 with today's
# page_001.
# --------------------------------------------------------------------------
#: A crawl that wants more pages than this has gone wrong.  GitHub caps user
#  search at 1,000 results, which is 100 pages.
MAX_PAGES = 120


class PageCrawlError(RuntimeError):
    """The crawl could not be completed.  Staging is preserved."""


class PaginationRateLimit(PageCrawlError):
    """A GitHub rate-limit/challenge hit a results page mid-crawl.  Staging is
    preserved; the caller cools down and retries the search later."""

    def __init__(self, message, source="pagination", retry_after=None):
        super().__init__(message)
        self.source = source
        self.retry_after = retry_after


def with_page_number(url, number):
    """`url` rewritten to request results page `number` (drops ?p= for 1)."""
    parts = urllib.parse.urlsplit(url or "")
    query = urllib.parse.parse_qs(parts.query, keep_blank_values=True)
    if number <= 1:
        query.pop("p", None)
    else:
        query["p"] = [str(number)]
    new_query = urllib.parse.urlencode(query, doseq=True)
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, parts.path,
                                    new_query, parts.fragment))


class PageResult:
    """One page's outcome."""

    def __init__(self, number, url, rows, path):
        self.number = number
        self.url = url
        self.rows = rows
        self.path = path

    def __repr__(self):
        return "PageResult(%d, rows=%d)" % (self.number, self.rows)


# ============================================================== staging =====
def staging_dir(slug):
    """crawl_staging/<search-slug>/ for one search, in this folder."""
    return C.staging_for(slug)


def page_path(directory, number):
    return os.path.join(directory, "page_%03d.xlsx" % number)


def staged_pages(directory):
    """Staged page files in page order."""
    if not os.path.isdir(directory):
        return []
    found = []
    for name in os.listdir(directory):
        if not name.startswith("page_") or not name.endswith(".xlsx"):
            continue
        digits = name[len("page_"):-len(".xlsx")]
        if digits.isdigit():
            found.append((int(digits), os.path.join(directory, name)))
    return [path for _number, path in sorted(found)]


def archive_staging(directory, say=log.info):
    """Move a previous attempt's staging aside, so it cannot be merged in.

    Archived rather than deleted: if a retry was a mistake, the rows are still
    there.  Returns the archive path, or "" when there was nothing to move.
    """
    if not os.path.isdir(directory) or not os.listdir(directory):
        return ""
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    target = "%s.attempt_%s" % (directory.rstrip(os.sep), stamp)
    serial = 1
    while os.path.exists(target):
        target = "%s.attempt_%s-%d" % (directory.rstrip(os.sep), stamp, serial)
        serial += 1
    shutil.move(directory, target)
    say("[STAGING] previous attempt archived to %s"
        % os.path.basename(target))
    return target


# ================================================================ merging ===
def _row_identity(cells):
    """The GitHub account a raw export row belongs to, normalised.

    Reuses XlsxReader's own detection -- avatar URL, then bare profile URL,
    then repository owner -- so a merged row and a parsed candidate can never
    disagree about who a row is.
    """
    username = XlsxReader._username_from(cells)
    return normalize_username(username) if username else ""


def merge_staged_pages(paths, destination, say=log.info):
    """Combine staged page exports into ONE export, de-duplicated.

    Rows keep the Instant Data Scraper shape, because XlsxReader parses the
    result: merging is concatenation plus dedupe, never reinterpretation.
    Columns are aligned BY NAME, since a later page can in principle expose a
    column the first did not.

    Returns (unique_rows, duplicates_dropped, unusable_rows).
    """
    header = []
    seen_header = {}
    merged = []
    seen_identity = set()
    duplicates = 0
    unusable = 0

    for path in paths:
        try:
            book = load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            raise PageCrawlError("cannot read staged page %s: %s"
                                 % (os.path.basename(path), exc))
        try:
            for sheet in book.worksheets:
                local = []
                for raw in sheet.iter_rows(values_only=True):
                    cells = ["" if value is None else str(value).strip()
                             for value in raw]
                    if not local:
                        local = cells
                        for name in local:
                            if name not in seen_header:
                                seen_header[name] = len(header)
                                header.append(name)
                        continue
                    if not any(cells):
                        continue
                    key = _row_identity(cells)
                    if not key:
                        unusable += 1
                        continue
                    if key in seen_identity:
                        duplicates += 1
                        continue
                    seen_identity.add(key)
                    aligned = [""] * len(header)
                    for position, name in enumerate(local):
                        if position < len(cells):
                            aligned[seen_header[name]] = cells[position]
                    merged.append(aligned)
        finally:
            book.close()

    # Re-pad in case a later page widened the header.
    width = len(header)
    for row in merged:
        if len(row) < width:
            row.extend([""] * (width - len(row)))

    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    sheet.append(header)
    for row in merged:
        sheet.append(row)
    directory = os.path.dirname(os.path.abspath(destination))
    os.makedirs(directory, exist_ok=True)
    tmp = os.path.join(directory, ".%s.tmp" % os.path.basename(destination))
    book.save(tmp)
    with open(tmp, "rb") as handle:
        os.fsync(handle.fileno())
    os.replace(tmp, destination)
    book.close()
    say("[MERGE] %d unique row(s) from %d page file(s); %d duplicate(s), "
        "%d unusable" % (len(merged), len(paths), duplicates, unusable))
    return len(merged), duplicates, unusable


# =========================================================== verification ===
def verify_collection(unique_rows, expected_total, mode, pages_done,
                      expected_pages, last_page_had_next):
    """(ok, why) -- may this multi-page collection be accepted as complete?

    EXACT_TOTAL       the unique rows must equal GitHub's stated figure.  Both
                      a shortfall and an excess refuse: a shortfall means rows
                      were lost, and an excess means the result set moved
                      under us, so neither is silently accepted.
    PAGE_COMPLETENESS equality is impossible, so completeness is judged by
                      pagination instead: the crawl must have run out of Next,
                      not out of patience.
    """
    if mode == ValidationMode.EXACT_TOTAL and expected_total:
        if unique_rows == expected_total:
            return (True, "collected %d of %d stated user(s)"
                    % (unique_rows, expected_total))
        if unique_rows < expected_total:
            return (False,
                    "GitHub states %d user(s) but only %d unique row(s) were "
                    "collected over %d page(s); %d are missing"
                    % (expected_total, unique_rows, pages_done,
                       expected_total - unique_rows))
        return (False,
                "GitHub states %d user(s) but %d unique row(s) were collected "
                "over %d page(s); the result set changed while paginating"
                % (expected_total, unique_rows, pages_done))

    if last_page_had_next:
        return (False,
                "stopped after %d page(s) while GitHub still offered a next "
                "page" % pages_done)
    if expected_pages and pages_done < expected_pages:
        return (False,
                "only %d of about %d page(s) were exported"
                % (pages_done, expected_pages))
    return (True, "every reachable page exported (%d page(s), %d unique row(s))"
            % (pages_done, unique_rows))


# ============================================================== the crawl ===
class PageByPageCollector:
    """Walks a multi-page search, one fresh IDS per page."""

    def __init__(self, browser, scraper, say=print, check_stop=None,
                 health_gate=None):
        self.browser = browser
        self.scraper = scraper
        self.say = say
        # A no-arg callable that raises to abort a wait cleanly on Ctrl+C.
        self._check_stop = check_stop or (lambda: None)
        # Battery/health gate, awaited BETWEEN result pages.  A crawl can hold
        # the run for many minutes, so without this a laptop that reached the
        # pause threshold mid-crawl would keep paging GitHub until the whole
        # search finished -- the same bug the profile loop had.  Default None:
        # callers that do not pass one are completely unaffected.
        self._health_gate = health_gate
        self.pages = []
        self.expected_total = None
        self.validation_mode = ValidationMode.PAGE_COMPLETENESS
        self.expected_pages = None
        self.last_page_had_next = False
        # True when --max-pages cut the crawl short.  The result set
        # is then incomplete by construction; completeness is still
        # judged by the usual rule (GitHub was still offering a next
        # page), so a capped crawl fails validation on its own.
        self.truncated = False
        self.staging = ""
        self.slug = ""
        self.attempt_archive = ""

    # -------------------------------------------------------------- pages --
    def _dom_fallback_page(self, number, page, destination, expected_ids,
                           state, page_count):
        """Write ONE page from its exact DOM identities when IDS can't infer
        the table.  Returns a PageResult on success, or None when the fallback
        does not apply (the caller then stops safely, preserving staging).

        Uses the shared authoritative-DOM validator/serializer -- identical to
        the whole-single-page fallback.  DOM identities stay authoritative; a
        wrong/ambiguous DOM (count mismatch, duplicate, short) does NOT apply.
        """
        if state is None or not expected_ids:
            return None
        count = page_count or (
            int(state.get("resultsListChildren") or 0)
            or int(state.get("resultProfileLinkCount") or 0))
        ok, written, reason = write_dom_fallback(
            state, destination, expected_ids, count)
        if not ok:
            self.say("[PAGE %d] IDS could not establish the table; exact-DOM "
                     "fallback did NOT apply: %s" % (number, reason))
            return None
        if validate_export(destination) < 0:
            self.say("[PAGE %d] exact-DOM fallback file did not validate"
                     % number)
            return None
        self.say("[PAGE %d] IDS could not establish the table; exact-DOM "
                 "fallback wrote %d verified user(s) -> %s"
                 % (number, written, os.path.basename(destination)))
        return PageResult(number, page.url, written, destination)

    async def _export_current_page(self, page, number, expected_ids=None,
                                   state=None, page_count=0):
        """Fresh popup bound to the CURRENT url, establish the table, export.

        A new popup every page is the whole point: the old one is bound to the
        page it opened on and cannot see a replaced DOM.  Before exporting,
        the SELECTED table is validated against this page's DOM identities --
        IDS can auto-pick a wrong (page-chrome) table on a small page.

        IDS is always preferred.  When IDS cannot establish the exact table for
        a SMALL page (e.g. page 2 of a 12-user search with 2 users), the exact
        DOM identity set for THAT page is written directly -- the same
        last-resort serializer used for a whole single-page search -- rather
        than discarding the pages already staged.
        """
        popup = None
        try:
            popup = await self.scraper.open_instant_data_scraper(page)
            bound = getattr(popup, "url", "") or ""
            if not await self.scraper.wait_for_table(popup):
                raise PageCrawlError(
                    "the extension found no table on page %d (%s)"
                    % (number, page.url))
            _pages, rows = await self.scraper.read_stats(popup)
            if not rows or rows <= 0:
                raise PageCrawlError("page %d reported no rows" % number)
            destination = page_path(self.staging, number)
            if expected_ids:
                # Establish the correct table AND require it to stay stable:
                # after "Try another table" the extension keeps mutating its
                # model asynchronously, so a table correct at one instant can
                # shift before the XLSX click.  stabilize_results_table
                # requires several consecutive matching reads and a final
                # pre-click re-read, bounded to a couple of attempts.
                if not await self.scraper.stabilize_results_table(
                        popup, page, expected_ids):
                    # IDS could not establish the exact table for this page.
                    # Try the exact-DOM fallback for THIS page before giving up
                    # (which would discard the pages already staged).
                    result = self._dom_fallback_page(
                        number, page, destination, expected_ids, state,
                        page_count)
                    if result is not None:
                        return result
                    raise PageCrawlError(
                        "IDS result table did not remain stable on page %d "
                        "(could not hold the %d DOM user(s)) and the exact-DOM "
                        "fallback did not apply" % (number, len(expected_ids)))
                # PROTECTION 1 (pre-export): one last read the instant before
                # the XLSX click; require it still matches the DOM.
                now = await self.scraper.preview_identities(popup, expected_ids)
                if now != expected_ids:
                    raise PageCrawlError(
                        "the IDS table on page %d changed after validation; "
                        "refusing to export a shifted table" % number)
            if expected_ids:
                # Export from the VERIFIED PREVIEW, not the #xlsx button: the
                # button can write a table that diverged from the validated
                # preview and can kill Chromium on click.  Python writes the
                # page XLSX straight from the stabilised grid into staging.
                exported = await self.scraper.export_page_from_preview(
                    popup, destination, expected_ids)
            else:
                # No DOM identities to verify against; fall back to the
                # extension export.  Chromium writes to keyword_downloads/, not
                # into staging, so the download directory is passed explicitly.
                exported = await self.scraper.export_xlsx(
                    popup, destination,
                    download_dir=C.KEYWORD_DOWNLOAD_DIR)
            if exported < 0:
                raise PageCrawlError("page %d could not be exported" % number)
            # PROTECTION 2 (post-export): the written file's identities must
            # EXACTLY equal this page's DOM identities -- a set check, not a
            # count.  Catches an export whose model diverged from the
            # validated preview.
            if expected_ids:
                got = self.scraper.export_identities(destination)
                if got != expected_ids:
                    missing = sorted(expected_ids - got)
                    extra = sorted(got - expected_ids)
                    raise PageCrawlError(
                        "page %d export identities do not match the %d DOM "
                        "user(s): %d exported, missing %s, unexpected %s"
                        % (number, len(expected_ids), len(got),
                           missing[:5] or "none", extra[:5] or "none"))
            if getattr(self.scraper, "last_export_recovered", False):
                self.say("[PAGE %d] recovered export from disk after browser "
                         "hiccup" % number)
            if exported != rows:
                self.say("[PAGE %d] extension showed %d row(s), export holds "
                         "%d" % (number, rows, exported))
            self.say("[PAGE %d] %d row(s) -> %s"
                     % (number, exported, os.path.basename(destination)))
            log.info("page %d popup bound to %s", number, bound)
            # A page counts as complete ONLY once its staged file validates.
            if validate_export(destination) < 0:
                raise PageCrawlError(
                    "page %d staged file did not validate" % number)
            return PageResult(number, page.url, exported, destination)
        finally:
            if popup is not None and not popup.is_closed():
                try:
                    await popup.close()
                except Exception as exc:
                    log.debug("could not close the popup: %s", exc)

    async def _click_next(self, page, from_number):
        """Turn the page with Playwright.  (moved, selector_used)."""
        for selector in C.STABLE_NEXT_SELECTORS:
            node = page.locator(selector)
            try:
                if await node.count() != 1 or not await node.first.is_visible():
                    continue
            except Exception as exc:
                # A dead browser must surface as BrowserLost so the caller can
                # restart -- not be mistaken for "this selector is absent".
                if browser_died(exc):
                    raise BrowserLost(str(exc))
                continue
            # A failure to READ the attribute is not evidence that the control
            # is disabled, and must not discard an otherwise usable Next.
            try:
                disabled = await node.first.get_attribute("aria-disabled")
            except Exception as exc:
                log.debug("could not read aria-disabled on %s: %s",
                          selector, exc)
                disabled = None
            if disabled == "true":
                continue
            try:
                await node.first.click(timeout=15000)
            except Exception as exc:
                log.warning("clicking %s failed: %s", selector, exc)
                continue
            return True, selector
        return False, ""

    async def _has_live_next(self, state):
        has_any = (bool(state.get("hasNext"))
                   or int(state.get("nextByShape") or 0) > 0
                   or int(state.get("nextByText") or 0) > 0)
        return has_any and state.get("nextDisabled") is not True

    # ------------------------------------------------------- restart -----
    async def _browser_dead(self):
        """True only when there is a browser manager AND it is dead.

        A collector built without a manager (the unit tests drive navigation
        through a fake page) is never "dead": there is nothing to restart, so
        the ordinary click path is used and a genuine stall still raises.
        """
        if self.browser is None:
            return False
        try:
            return not await self.browser.is_alive()
        except Exception:
            return True

    async def _restart_to(self, page, number):
        """Relaunch Chromium on the SAME profile and land on results page N.

        Used when the browser dies between pages.  The crawl does not require
        one browser process to survive the whole search: a page already staged
        and validated stays valid across a restart, so collection resumes at
        the page that was interrupted.  Returns (page, state).
        """
        if self.browser is None:
            raise PageCrawlError("the browser died and cannot be restarted")
        if self.browser.restarts >= C.BROWSER_MAX_RESTARTS:
            raise PageCrawlError(
                "the browser died and the restart budget (%d) is exhausted"
                % C.BROWSER_MAX_RESTARTS)
        if not self.search_url:
            raise PageCrawlError(
                "cannot restart: the search URL was never recorded")
        self.say("[CRAWL] the browser died; restarting on the same profile")
        page = await self.browser.restart(with_extension=True)
        try:
            login = await self.browser.github_login_name(page)
        except Exception:
            login = None
        self.say("[CRAWL] browser restarted%s"
                 % (" (logged in as %s)" % login if login else ""))
        url = with_page_number(self.search_url, number)
        await page.goto(url, wait_until="domcontentloaded",
                        timeout=C.NAVIGATION_TIMEOUT)
        ok, state = await self.scraper.wait_for_stable_results(page)
        if not ok:
            raise PageCrawlError(
                "results page %d never stabilised after a restart" % number)
        landed = page_number(state.get("url") or page.url)
        if landed != number:
            raise PageCrawlError(
                "after a restart the browser landed on page %d, not %d"
                % (landed, number))
        return page, state

    async def _stable_page(self, page, number):
        """A stable results page `number`, reloading it before giving up.

        A dead browser is restarted, as before.  A LIVE browser that simply
        served a slow page used to raise immediately and lose the whole
        search -- see C.PAGE_STABLE_ATTEMPTS for the run that prompted this.
        The page is now reloaded and re-waited instead.

        Returns (page, state).  Raises PageCrawlError once the attempts are
        spent, so a genuinely unreachable page still stops the search.
        """
        for attempt in range(1, C.PAGE_STABLE_ATTEMPTS + 1):
            try:
                ok, state = await self.scraper.wait_for_stable_results(page)
            except BrowserLost:
                return await self._restart_to(page, number)
            if ok:
                if attempt > 1:
                    self.say("[CRAWL] page %d settled on attempt %d"
                             % (number, attempt))
                return page, state
            if await self._browser_dead():
                return await self._restart_to(page, number)
            if attempt >= C.PAGE_STABLE_ATTEMPTS:
                break
            self.say("[CRAWL] page %d did not settle (attempt %d of %d); "
                     "reloading" % (number, attempt, C.PAGE_STABLE_ATTEMPTS))
            try:
                await page.reload(wait_until="domcontentloaded",
                                  timeout=C.NAVIGATION_TIMEOUT)
            except Exception as exc:
                # A reload that fails because the browser died is a restart;
                # any other reload failure just costs this attempt.
                if browser_died(exc):
                    return await self._restart_to(page, number)
                log.warning("could not reload page %d: %s", number, exc)
        raise PageCrawlError("page %d results never stabilised" % number)

    async def _export_page(self, page, number, expected_ids=None,
                           state=None, page_count=0):
        """Export page `number`, restarting and retrying once if the browser
        dies before the page is proven staged.

        A page is complete only when its staged file validates; a browser
        death before that means the page is retried (option B semantics, but
        scoped to the one interrupted page).  A death that export recovery
        already survived does not reach here -- export_xlsx returns the rows.
        Returns (page, PageResult).
        """
        attempts = 0
        while True:
            attempts += 1
            try:
                result = await self._export_current_page(
                    page, number, expected_ids, state, page_count)
                return page, result
            except BrowserLost:
                # The page was NOT staged (recovery inside export_xlsx would
                # have returned rows instead of raising).  Restart and retry
                # this same page.
                if attempts >= 2:
                    raise PageCrawlError(
                        "page %d could not be exported after a restart"
                        % number)
                page, state = await self._restart_to(page, number)
                page_count = (int(state.get("resultsListChildren") or 0)
                              or int(state.get("resultProfileLinkCount") or 0))

    # --------------------------------------------------------------- run ---
    async def collect(self, page, slug, first_state=None, max_pages=None):
        """Export every page into staging.  Raises PageCrawlError on failure.

        Survives the browser dying between pages: a validated staged page is
        kept, the browser is restarted on the same profile, and collection
        continues from the interrupted page.  Returns the PageResult list.

        max_pages is the --max-pages TESTING cap.  It stops the crawl early,
        so the result is deliberately INCOMPLETE and says so -- a capped crawl
        must never be reported as "every reachable page exported", because a
        later run would then trust a partial result set.
        """
        self.slug = slug
        self.staging = staging_dir(slug)
        # A retry must never merge a previous attempt's pages with this one.
        self.attempt_archive = archive_staging(self.staging, self.say)
        os.makedirs(self.staging, exist_ok=True)

        state = first_state
        if state is None:
            page, state = await self._stable_page(page, page_number(page.url))
        # The base search URL, so a restart can navigate straight to any page.
        self.search_url = with_page_number(state.get("url") or page.url, 1)
        self.expected_total, self.validation_mode = expected_total_users(state)
        self.expected_pages = expected_page_count(self.expected_total)
        self.say("[CRAWL] page-by-page mode")
        self.say("[CRAWL] GitHub states %s user(s); validation: %s%s"
                 % (self.expected_total if self.expected_total else "no exact "
                    "total", self.validation_mode,
                    "" if not self.expected_pages
                    else " (about %d page(s))" % self.expected_pages))

        number = page_number(state.get("url") or page.url)
        have_state = True
        while True:
            # Safe point: the previous page is fully staged and the next one
            # has not been fetched.  Pausing here loses nothing.
            if self._health_gate is not None:
                await self._health_gate()
            if max_pages and len(self.pages) >= max_pages:
                self.truncated = True
                self.say("[CRAWL] --max-pages %d reached; stopping early. "
                         "This result set is INCOMPLETE." % max_pages)
                break
            if len(self.pages) >= MAX_PAGES:
                raise PageCrawlError(
                    "stopped after %d pages; a user search cannot be this long"
                    % MAX_PAGES)
            if not have_state:
                page, state = await self._stable_page(page, number)
            have_state = False
            # A rate-limit/challenge on THIS results page: keep the staged pages,
            # do not build the final XLSX, cool down and retry the search later.
            verdict, why = classify_search_page(state)
            if (verdict in (PageState.BLOCKED, PageState.HTTP_ERROR)
                    and is_throttle_reason(why)):
                raise PaginationRateLimit(
                    "page %d rate-limited: %s" % (number, why))
            results = (int(state.get("resultsListChildren") or 0)
                       or int(state.get("resultProfileLinkCount") or 0))
            if not results:
                raise PageCrawlError(
                    "page %d shows no user results" % number)
            self.say("[PAGE %d] stable: %d result(s) on %s"
                     % (number, results, state.get("url")))

            # `had_next` is read from THIS page's stable state, before export;
            # export may kill the browser, after which the page cannot be
            # re-read, so the decision is captured now.
            had_next = await self._has_live_next(state)
            expected_ids = self.scraper.dom_result_identities(state)
            page, result = await self._export_page(page, number, expected_ids,
                                                   state, results)
            self.pages.append(result)
            self.last_page_had_next = had_next
            if not had_next:
                self.say("[CRAWL] no further Next; %d page(s) exported"
                         % len(self.pages))
                break

            await self._page_delay()          # conservative delay before Next
            page, number = await self._advance(page, number)
        return list(self.pages)

    async def _page_delay(self):
        """A random, interruptible delay before turning to the next results
        page.  Python controls pagination; IDS only extracts the current page."""
        seconds = random.uniform(PAGE_DELAY_MIN, PAGE_DELAY_MAX)
        cd = Countdown(seconds, "PAGE")
        cd.start()
        await sleep_interruptible(seconds, self._check_stop,
                                            on_tick=cd.tick)
        cd.finish()

    async def _advance(self, page, number):
        """Move to the next results page, restarting the browser if it died.

        The current page is already staged and validated, so a browser death
        here is recoverable: restart and navigate straight to page N+1.
        Returns (page, landed_number).
        """
        if await self._browser_dead():
            page, _state = await self._restart_to(page, number + 1)
            self.say("[CRAWL] transition: page %d -> %d (after restart)"
                     % (number, number + 1))
            return page, number + 1
        try:
            moved, selector = await self._click_next(page, number)
            if not moved:
                raise PageCrawlError(
                    "GitHub offered a next page after page %d but it could "
                    "not be clicked" % number)
            try:
                await page.wait_for_function(
                    "expected => new URLSearchParams(location.search)"
                    ".get('p') !== expected",
                    arg=str(number) if number > 1 else None,
                    timeout=C.NAVIGATION_TIMEOUT)
            except BrowserLost:
                raise
            except Exception:
                log.debug("URL change was not observed directly")
            page, state = await self._stable_page(page, number + 1)
        except BrowserLost:
            page, _state = await self._restart_to(page, number + 1)
            self.say("[CRAWL] transition: page %d -> %d (after restart)"
                     % (number, number + 1))
            return page, number + 1
        landed = page_number(state.get("url") or page.url)
        if landed <= number:
            raise PageCrawlError(
                "clicking Next did not advance past page %d (still on "
                "page %d)" % (number, landed))
        self.say("[CRAWL] transition: page %d -> %d" % (number, landed))
        return page, landed

    # ------------------------------------------------------------- finish --
    def merge_result(self, destination):
        """Merge staging + verify, WITHOUT raising.

        Returns (unique, duplicates, unusable, ok, why).  The merged file is
        written to `destination`; the caller decides whether to keep it (an
        accepted collection) or remove it (a rejected one).  This is the
        non-raising core shared by merge_into and the dynamic-result
        reconciliation.
        """
        paths = staged_pages(self.staging)
        if not paths:
            return 0, 0, 0, False, "no staged page files to merge"
        if len(paths) != len(self.pages):
            return (0, 0, 0, False,
                    "%d page(s) were exported but %d staged file(s) are present"
                    % (len(self.pages), len(paths)))
        unique, duplicates, unusable = merge_staged_pages(
            paths, destination, self.say)
        ok, why = verify_collection(
            unique, self.expected_total, self.validation_mode,
            len(self.pages), self.expected_pages, self.last_page_had_next)
        return unique, duplicates, unusable, ok, why

    def merge_into(self, destination):
        """Merge staging into ONE export and verify it.  (rows, why).

        Raises PageCrawlError on any failure, removing the merged file so it is
        never mistaken for a finished collection.  Unchanged contract.
        """
        unique, _dupes, _unusable, ok, why = self.merge_result(destination)
        if not ok:
            try:
                os.remove(destination)
            except OSError:
                pass
            raise PageCrawlError(why)
        self.say("[CRAWL] verified (%s): %s" % (self.validation_mode, why))
        return unique, why

    def is_exact_total_mismatch(self, unique, duplicates, unusable):
        """True when a rejected merge is the RECONCILABLE case: GitHub stated an
        exact total, every page was DOM-validated, the rows are unique with no
        unusable rows, pagination ended naturally, and only the count differs.
        A shortfall/excess caused by dupes, unusable rows or an unfinished crawl
        is NOT reconcilable."""
        return (self.validation_mode == ValidationMode.EXACT_TOTAL
                and bool(self.expected_total)
                and unique != self.expected_total
                and duplicates == 0 and unusable == 0
                and not self.last_page_had_next)

    async def read_current_total(self, page):
        """Reopen the FIRST search page and re-read GitHub's stated total.
        Returns (total, mode) -- total is None when GitHub gives no exact
        figure."""
        url = with_page_number(self.search_url, 1)
        await page.goto(url, wait_until="domcontentloaded",
                        timeout=C.NAVIGATION_TIMEOUT)
        _ok, state = await self.scraper.wait_for_stable_results(page)
        return expected_total_users(state)

    async def reconcile(self, page, destination, unique, first_ids,
                        accept_stable):
        """One controlled dynamic-result reconciliation after a reconcilable
        EXACT_TOTAL mismatch.  `destination` holds the first crawl's merged set.

        Returns (accepted, info).  info["kind"] is:
          reconciled     GitHub's re-read total now equals the collected unique
          stable         a second fully-validated crawl produced the SAME set
                         (accepted only when `accept_stable`)
          changed        the second crawl produced a different set  -> reject
          verify_failed  the second crawl could not complete        -> reject
        """
        info = {"initial": self.expected_total, "unique": unique,
                "final": None, "second": None, "kind": ""}

        # Step 1: re-read GitHub's current stated total.
        final_total, mode = await self.read_current_total(page)
        info["final"] = final_total
        if mode == ValidationMode.EXACT_TOTAL and final_total == unique:
            info["kind"] = "reconciled"
            return True, info

        # Step 2: exactly ONE clean verification re-crawl into fresh staging.
        verifier = PageByPageCollector(self.browser, self.scraper,
                                       self.say)
        verify_dest = destination + ".verify.xlsx"
        try:
            await verifier.collect(page, self.slug)
        except (PageCrawlError, BrowserLost):
            info["kind"] = "verify_failed"
            return False, info
        v_unique, v_dupes, v_unusable, _ok, _why = verifier.merge_result(
            verify_dest)
        v_ids = (InstantDataScraper.export_identities(verify_dest)
                 if os.path.exists(verify_dest) else set())
        try:
            os.remove(verify_dest)
        except OSError:
            pass
        info["second"] = v_unique
        stable = (v_ids == first_ids and v_dupes == 0 and v_unusable == 0
                  and not verifier.last_page_had_next
                  and verifier.is_exact_total_mismatch(v_unique, v_dupes,
                                                       v_unusable))
        if not stable:
            info["kind"] = "changed"
            return False, info
        info["kind"] = "stable"
        return (bool(accept_stable), info)

    def cleanup(self):
        """Remove staging.  Only ever called after a verified success."""
        if self.staging and os.path.isdir(self.staging):
            try:
                shutil.rmtree(self.staging)
                self.say("[CLEANUP] crawl staging removed")
            except OSError as exc:
                self.say("[WARN] could not remove staging: %s" % exc)

    def state_block(self, attempt_id=""):
        """The `crawl` block recorded in scraper_state.json."""
        return {
            "attempt_id": attempt_id,
            "mode": "page-by-page",
            "validation": self.validation_mode,
            "expected_total": self.expected_total,
            "expected_pages": self.expected_pages,
            "current_page": self.pages[-1].number if self.pages else 0,
            "completed_pages": len(self.pages),
            "staged_page_files": [os.path.basename(result.path)
                                  for result in self.pages],
            "last_page_complete": bool(self.pages) and
            not self.last_page_had_next,
            "staging_dir": self.staging,
        }


# ==========================================================================
# REGEX EXTRACTOR
# ==========================================================================
# Pattern candidates plus the context rules that validate them.
#
# Regex proposes, context disposes (spec 31).  Every extractor here returns
# values that survived a validation step, not raw regex hits: a ten-digit
# number is only a phone when something nearby says so, six digits are only a
# PIN near an address word, and a year is only a graduation year near
# education wording.
# --------------------------------------------------------------------------
# ----------------------------------------------------------------- email ---
# The trailing guard is (?![\w-]) rather than (?![\w.-]): a dot may legally
# follow an address that ends a sentence ("write to me@site.com."), and with
# the dot in the guard that whole address matches nothing at all -- the TLD
# cannot be shortened either, because any shorter ending is followed by a
# word character.  A dot can never be part of the match anyway, since the
# pattern must end in letters, and extract_emails() strips a trailing dot.
EMAIL_RE = re.compile(
    r'(?<![\w.+-])'
    r'[A-Za-z0-9._%+-]+'
    r'@'
    r'[A-Za-z0-9.-]+'
    r'\.[A-Za-z]{2,}'
    r'(?![\w-])',
    re.I,
)

PLACEHOLDER_EMAILS = {
    "example@example.com", "test@test.com", "name@email.com",
    "email@example.com", "you@example.com", "user@example.com",
    "your@email.com", "hello@example.com", "info@example.com",
    "someone@example.com", "john@doe.com", "johndoe@example.com",
    "name@domain.com", "your.email@example.com", "admin@example.com",
    "contact@example.com", "mail@example.com", "sample@email.com",
}
PLACEHOLDER_DOMAINS = {"example.com", "example.org", "example.net",
                       "domain.com", "email.com", "yourdomain.com",
                       "test.com", "sentry.io", "wixpress.com"}
# Addresses GitHub itself generates or that belong to infrastructure.
NON_PERSONAL_EMAIL_RE = re.compile(
    r"(@users\.noreply\.github\.com$|^noreply@|^no-reply@|^donotreply@"
    r"|@sentry\.|@wordpress\.|@godaddy\.)", re.I)


def is_placeholder_email(value):
    low = (value or "").strip().lower()
    if low in PLACEHOLDER_EMAILS:
        return True
    domain = low.rsplit("@", 1)[-1]
    if domain in PLACEHOLDER_DOMAINS:
        return True
    return bool(NON_PERSONAL_EMAIL_RE.search(low))


# Real pages glue contact details together.  A portfolio that renders an icon
# and its text as adjacent inline elements gives inner_text() one run with no
# separator, and a live example looked like this:
#
#     82830-89459eshanvohra@gmail.comFirozpur, Punjab, India
#
# EMAIL_RE matches that whole run: "82830-89459eshanvohra" is a legal local
# part, and `[A-Za-z]{2,}` happily eats "comFirozpur" as a TLD.  The address
# is then wrong in a way nobody notices until they try to use it.
#
# The two halves are trimmed back separately, each by a rule that is safe:
# a TLD is cut at the longest spelling that is actually a TLD, and a local
# part that STARTS with a phone number has the phone removed.
KNOWN_TLDS = (
    "com", "org", "net", "edu", "gov", "mil", "int", "info", "biz", "name",
    "io", "ai", "co", "dev", "app", "me", "xyz", "tech", "online", "site",
    "in", "uk", "us", "ca", "au", "de", "fr", "nl", "it", "es", "se", "no",
    "fi", "dk", "pl", "ru", "br", "jp", "cn", "kr", "sg", "ae", "za", "nz",
    "ch", "at", "be", "cz", "gr", "ie", "pt", "tr", "mx", "id", "my", "ph",
    "edu.in", "ac.in", "co.in", "org.in", "net.in", "gov.in", "co.uk",
    "ac.uk", "org.uk", "com.au", "edu.au", "co.nz", "com.br", "co.za",
)
# Longest first, so "edu.in" wins over "in" and "com" over "co".
_TLD_BY_LENGTH = tuple(sorted(KNOWN_TLDS, key=len, reverse=True))

# A phone glued to the front: 6+ digits, optionally broken by - . ( ) or
# space, immediately followed by the letters that start the real local part.
_GLUED_PHONE_PREFIX = re.compile(
    r"^(?:\+?\d[\d\-.()\s]{4,}\d)(?=[A-Za-z])")


def trim_glued_email(value):
    """Cut a glued run back to the address inside it.

    Returns "" when what is left is not usable.  Conservative on purpose: if
    the tail after the TLD is not plain letters, or the leading digits are not
    phone-shaped, the value is left exactly as it was rather than guessed at.
    """
    local, at, domain = value.partition("@")
    if not at or not domain:
        return value

    # -- the domain: cut the TLD at a spelling that really is a TLD ---------
    labels = domain.split(".")
    if len(labels) >= 2:
        tail = labels[-1]
        # A tail that IS a TLD is already correct.  Checking this first
        # matters: "com" starts with "co", and trimming it to "co" would
        # corrupt every ordinary address.
        if tail not in KNOWN_TLDS:
            for tld in _TLD_BY_LENGTH:
                if "." in tld or not tail.startswith(tld):
                    continue
                rest = tail[len(tld):]
                # Only when the surplus is plain letters -- that is glued
                # prose.  Anything else might be a TLD this list lacks.
                if rest.isalpha():
                    labels[-1] = tld
                    domain = ".".join(labels)
                break

    # -- the local part: drop a phone number stuck on the front ------------
    trimmed_local = _GLUED_PHONE_PREFIX.sub("", local)
    if trimmed_local and trimmed_local != local:
        local = trimmed_local

    return "%s@%s" % (local, domain)


# ==========================================================================
# THE MASTER INDEX -- the same shape find_naukri_job.py uses
# ==========================================================================
# naukri keeps no database.  It loads ONE master spreadsheet into a dict at
# startup, asks that dict whether a row is new, and writes the master back as
# it goes:
#
#     master_df     = load_existing_master()
#     existing_data = {make_unique_key(...): row}
#     ...
#     if key in existing_data:  seen_count += 1
#     else:                     existing_data[key] = {...}
#     save_master(updated_df)
#
# This is that, for people instead of jobs.  The master is
# "Student Data original/student_data.xlsx" -- the same 65 columns the output
# uses, plus naukri's three bookkeeping columns.
#
# WHAT THIS COSTS, SO IT IS NOT A SURPRISE
# ----------------------------------------
# The whole master is held in memory and rewritten in full each time it is
# saved.  That is exactly what naukri does, and it is fine at naukri's scale.
# Here the master grows with every person ever collected, and an .xlsx sheet
# stops at 1,048,576 rows.  Writing gets noticeably slow long before that --
# expect tens of seconds a save once the master passes a few hundred thousand
# rows.  save_master() is therefore called once per SEARCH, not once per
# person, and MASTER_SAVE_EVERY exists to loosen that further if needed.
# Its own file, NOT student_data.xlsx.  naukri keeps its master
# (naukri_it_jobs_by_city.xlsx) separate from current_data.xlsx for the same
# reason: student_data.xlsx is what --export-final writes, and an export
# rebuilding that file would silently destroy the record of who has already
# been collected.
MASTER_NAME = "master_index.xlsx"


def master_path():
    """Where the duplicate index lives, resolved on every call.

    NOT a module constant: config's FINAL_OUTPUT_DIR is redirected by
    tests/sandbox.py AFTER this module is imported, and a constant baked at
    import time would keep pointing at the real "Student Data original", so a
    test calling load_existing_master() or save_master() with no explicit path
    would read and rewrite the developer's own collected people.
    """
    return os.path.join(C.FINAL_OUTPUT_DIR, MASTER_NAME)
MASTER_EXTRA_COLUMNS = ["first_seen_at", "last_seen_at", "seen_count"]
MASTER_SAVE_EVERY = 1          # searches between master writes


def make_unique_key(record):
    """The identity of one person, as naukri's make_unique_key does for jobs.

    The GitHub username is the identity when there is one: it is unique on
    GitHub and cannot drift.  A record without one falls back to the profile
    URL, then to a real email address.  Returns "" when there is nothing to
    identify the person by, and such a record is never treated as a duplicate
    of anything.
    """
    username = normalize_username(str(record.get("GitHub Profile", "") or "")
                              .rstrip("/").rsplit("/", 1)[-1])
    if username:
        return "github:%s" % username
    url = normalize_username(str(record.get("Source URL", "") or ""))
    if url:
        return "url:%s" % url
    email = str(record.get("Email Address", "") or "").strip().lower()
    if email and not is_placeholder_email(email):
        return "email:%s" % email
    return ""


def load_existing_master(path=None):
    """{unique key: row} from the master spreadsheet.  Missing file -> {}."""
    path = path or master_path()
    existing = {}
    for row in read_rows(path, C.FINAL_COLUMNS + MASTER_EXTRA_COLUMNS):
        key = make_unique_key(row)
        if key:
            existing[key] = row
    return existing


def save_master(existing, path=None):
    """Write the whole master back, atomically.

    Atomic because this is the only record of who has been collected: a kill
    during the write must leave the previous master, never a half-written one.
    """
    path = path or master_path()
    rows = list(existing.values())
    _save_atomically(rows, path, C.FINAL_COLUMNS + MASTER_EXTRA_COLUMNS)
    return len(rows)


def extract_emails(text):
    """Valid, de-duplicated, lower-cased personal addresses in order."""
    found, seen = [], set()
    for match in EMAIL_RE.finditer(text or ""):
        value = trim_glued_email(match.group(0).strip().lower()).rstrip(".")
        if value in seen or is_placeholder_email(value):
            continue
        local, _, domain = value.partition("@")
        if not local or "." not in domain or domain.startswith("."):
            continue
        if len(value) > 254 or ".." in value:
            continue
        # A trailing image/file extension means this came out of a filename.
        if re.search(r"\.(png|jpg|jpeg|gif|svg|webp|css|js)$", domain):
            continue
        seen.add(value)
        found.append(value)
    return found


# ----------------------------------------------------------------- phone ---
PHONE_RE = re.compile(
    r'(?<!\d)'
    r'(?:\+?91[\s.-]?)?'
    r'([6-9]\d{4}[\s.-]?\d{5})'
    r'(?!\d)'
)

PHONE_CONTEXT_RE = re.compile(
    r'\b(phone|mobile|mob|contact|call|tel|telephone|cell|whats\s?app|'
    r'whatsapp|wa|reach\s+me|ph\.?|no\.?)\b', re.I)
WHATSAPP_CONTEXT_RE = re.compile(r'\b(whats\s?app|whatsapp|wa\s*(no|number))\b',
                                 re.I)
# Contexts that produce ten-digit numbers which are NOT phone numbers.
PHONE_ANTI_CONTEXT_RE = re.compile(
    r'\b(id|uid|order|invoice|account|acc|pan|aadha?ar|gst|ifsc|pin\s?code|'
    r'zip|issue|pr|commit|sha|version|build|port|repo|stars?|followers?|'
    r'views?|downloads?)\b', re.I)


def normalize_phone(value):
    """Canonical ten-digit Indian mobile, or '' when it is not one."""
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    elif len(digits) == 13 and digits.startswith("091"):
        digits = digits[3:]
    elif len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    if len(digits) == 10 and digits[0] in "6789":
        return digits
    return ""


def _window(text, start, end, size=60):
    return text[max(0, start - size):min(len(text), end + size)]


def _segment_of(text, position):
    """The smallest claim containing `position`.

    A contact line commonly packs several claims together:
        "Contact 9876543210 | WhatsApp 9123456789"
    Judging a label on the whole line marks BOTH numbers WhatsApp, so the
    line is narrowed to its nearest "|", ";" or tab boundaries.
    """
    line_start = text.rfind("\n", 0, position) + 1
    line_end = text.find("\n", position)
    if line_end == -1:
        line_end = len(text)
    start, end = line_start, line_end
    for token in ("|", ";", "\t", "  -  ", " / "):
        found = text.rfind(token, line_start, position)
        if found != -1:
            start = max(start, found + len(token))
        found = text.find(token, position, line_end)
        if found != -1:
            end = min(end, found)
    return text[start:end]




def extract_phones(text, require_context=True):
    """[(number, had_whatsapp_label)] for numbers with real phone context."""
    results, seen = [], set()
    body = text or ""
    for match in PHONE_RE.finditer(body):
        number = normalize_phone(match.group(1))
        if not number or number in seen:
            continue
        context = _window(body, match.start(), match.end())
        if PHONE_ANTI_CONTEXT_RE.search(context) and \
                not PHONE_CONTEXT_RE.search(context):
            continue
        has_label = bool(PHONE_CONTEXT_RE.search(context))
        explicit_plus91 = "+91" in body[max(0, match.start() - 5):match.end()]
        if require_context and not (has_label or explicit_plus91):
            continue
        # The WhatsApp label is judged on the number's OWN claim.  A 60-
        # character window around "Phone: 9876543210" reaches into a
        # neighbouring "WhatsApp: ...", which would mark both numbers
        # WhatsApp and leave the phone column empty.
        segment = _segment_of(body, match.start())
        seen.add(number)
        results.append((number, bool(WHATSAPP_CONTEXT_RE.search(segment))))
    return results


def extract_whatsapp(text):
    """Only a number explicitly labelled WhatsApp (spec 34)."""
    for number, is_whatsapp in extract_phones(text):
        if is_whatsapp:
            return number
    return ""


# ------------------------------------------------------------------- pin ---
PIN_RE = re.compile(r'(?<!\d)([1-9][0-9]{5})(?!\d)')
PIN_CONTEXT_RE = re.compile(
    r'\b(pin\s?-?\s?code|pincode|pin|postal\s*code|postal|zip\s*code|zip|'
    r'address|addr|street|road|nagar|colony|sector|district|dist|city|'
    r'state|india)\b', re.I)


def extract_pin(text):
    body = text or ""
    for match in PIN_RE.finditer(body):
        value = match.group(1)
        context = _window(body, match.start(), match.end(), size=50)
        # Six digits are also years-in-a-row, ids and counts; an address or
        # PIN word nearby is what makes it a postal code.
        if not PIN_CONTEXT_RE.search(context):
            continue
        if re.search(r"\b(19|20)\d{2}\b", value):
            continue
        return value
    return ""


# ------------------------------------------------------------------ urls ---
URL_RE = re.compile(r'https?://[^\s<>"\'\)\]]+', re.I)

LINKEDIN_RE = re.compile(
    r'https?://(?:[a-z]{2,3}\.)?(?:www\.)?linkedin\.com/'
    r'(?:in|pub)/'
    r'[A-Za-z0-9._%\-]+/?',
    re.I)
LINKEDIN_COMPANY_RE = re.compile(
    r'https?://(?:www\.)?linkedin\.com/(company|school|showcase)/', re.I)

GITHUB_PROFILE_RE = re.compile(
    r'https?://(?:www\.)?github\.com/'
    r'([A-Za-z0-9](?:[A-Za-z0-9-]{0,38}))'
    r'(?:/)?(?:[?#].*)?$',
    re.I)

# Coding-practice, course and credential platforms.  A profile on one of
# these is a profile link, not the person's own website (spec 26).
PLATFORM_HOSTS = (
    "dicoding.com", "leetcode.com", "hackerrank.com", "codechef.com",
    "codeforces.com", "kaggle.com", "coursera.org", "udemy.com",
    "udacity.com", "edx.org", "credly.com", "youracclaim.com",
    "freecodecamp.org", "geeksforgeeks.org", "hackerearth.com",
    "topcoder.com", "codewars.com", "exercism.org", "sololearn.com",
    "datacamp.com", "pluralsight.com", "skillshare.com", "simplilearn.com",
    "greatlearning.in", "internshala.com", "naukri.com", "indeed.com",
    "stackoverflow.com", "stackexchange.com", "replit.com", "codepen.io",
    "gist.github.com", "linktr.ee", "bento.me", "beacons.ai",
)

SOCIAL_HOSTS = (
    "github.com", "linkedin.com", "twitter.com", "x.com", "facebook.com",
    "instagram.com", "youtube.com", "youtu.be", "t.me", "telegram.me",
    "medium.com", "dev.to", "hashnode.com", "stackoverflow.com",
    "leetcode.com", "hackerrank.com", "codechef.com", "codeforces.com",
    "kaggle.com", "gitlab.com", "bitbucket.org", "discord.gg", "reddit.com",
    "pinterest.com", "tiktok.com", "threads.net", "mastodon.social",
    "githubusercontent.com", "gravatar.com", "shields.io", "badgen.net",
    "npmjs.com", "pypi.org", "docker.com", "vercel.com/button",
)


def clean_url(url):
    url = (url or "").strip().rstrip(".,;:!?)\"'")
    # Strip campaign parameters; keep everything meaningful.
    url = re.sub(r"[?&](utm_[a-z]+|fbclid|gclid|ref|source|trk|"
                 r"originalSubdomain)=[^&#]*", "", url, flags=re.I)
    return url.rstrip("?&")


def extract_urls(text):
    seen, output = set(), []
    for match in URL_RE.finditer(text or ""):
        url = clean_url(match.group(0))
        key = url.lower().rstrip("/")
        if not url or key in seen:
            continue
        seen.add(key)
        output.append(url)
    return output


def extract_linkedin(text_or_urls):
    """Canonical https://www.linkedin.com/in/<slug>, personal profiles only."""
    body = text_or_urls if isinstance(text_or_urls, str) \
        else "\n".join(text_or_urls)
    for match in LINKEDIN_RE.finditer(body or ""):
        url = clean_url(match.group(0))
        if LINKEDIN_COMPANY_RE.match(url):
            continue
        slug = url.rstrip("/").rsplit("/", 1)[-1]
        if not slug or slug.lower() in ("in", "pub"):
            continue
        return "https://www.linkedin.com/in/%s" % slug
    # GitHub renders the bare form "in/anil-palagani-a253b1254".
    bare = re.search(r'(?<![\w/])in/([A-Za-z0-9][A-Za-z0-9._-]{2,})', body or "")
    if bare:
        return "https://www.linkedin.com/in/%s" % bare.group(1)
    return ""


def extract_github_profile(text, reserved):
    for match in GITHUB_PROFILE_RE.finditer(text or ""):
        name = match.group(1)
        if name.lower() not in reserved:
            return "https://github.com/%s" % name
    return ""


def is_social(url):
    low = (url or "").lower()
    return any(host in low for host in SOCIAL_HOSTS)


def is_platform(url):
    """A third-party profile/credential page, not a personal site."""
    low = (url or "").lower()
    return any(host in low for host in PLATFORM_HOSTS)


def extract_portfolio(urls, github_website=""):
    """Best personal-site URL: GitHub's own website field wins (spec 64)."""
    if github_website and not is_social(github_website) \
            and not is_platform(github_website):
        return clean_url(github_website)
    for url in urls:
        if is_social(url) or is_platform(url):
            continue
        if re.search(r"\.(png|jpe?g|gif|svg|css|js|ico|woff2?)(\?|$)", url,
                     re.I):
            continue
        return clean_url(url)
    return ""


# ---------------------------------------------------------------- resume ---
RESUME_HINT_RE = re.compile(r'\b(resume|cv|curriculum[\s_-]*vitae)\b', re.I)
RESUME_FILE_RE = re.compile(
    r'https?://[^\s<>"\']+\.(?:pdf|docx?|txt|rtf|odt)(?:\?[^\s<>"\']*)?',
    re.I)


def extract_resume_urls(text, links=()):
    """Resume/CV links, most confident first."""
    body = text or ""
    strong, weak = [], []
    for url in list(links) + extract_urls(body):
        url = clean_url(url)
        low = url.lower()
        if RESUME_FILE_RE.match(url) and RESUME_HINT_RE.search(low):
            strong.append(url)
        elif RESUME_HINT_RE.search(low):
            weak.append(url)
        elif RESUME_FILE_RE.match(url) and low.endswith(".pdf"):
            weak.append(url)
    ordered, seen = [], set()
    for url in strong + weak:
        key = url.lower()
        if key not in seen:
            seen.add(key)
            ordered.append(url)
    return ordered


# ---------------------------------------------------------- delegation ----
# Education and employment extraction moved to education_extractor and
# employment_extractor, where section context is available.  The old names
# forward so existing callers and tests keep working against one
# implementation.






def extract_college(text):
    return extract_institution(text)




# ==========================================================================
# PORTFOLIO PARSER
# ==========================================================================
# Public portfolio / personal-site extraction.
#
# Scope is deliberately tight (spec 25): the homepage plus at most a few
# clearly-labelled About / Contact / Resume pages on the same site.  No
# recursive crawl, no following hundreds of links.
#
# Text is read with `page.locator("body").inner_text()` and the HTML with
# `page.content()` -- never Ctrl+U / Ctrl+A / Ctrl+C.
# --------------------------------------------------------------------------
_SUBPAGE_RE = re.compile(r"\b(about|contact|resume|cv)\b", re.I)


class PortfolioParser:
    def __init__(self, say=print):
        self.say = say

    async def fetch(self, context, url):
        """Return (visible_text, links, pages_visited)."""
        if not url:
            return "", [], []
        url = clean_url(url)
        if not url.lower().startswith(("http://", "https://")):
            url = "https://" + url.lstrip("/")

        page = None
        text_parts, links, visited = [], [], []
        try:
            page = await context.new_page()
            page.set_default_navigation_timeout(C.PORTFOLIO_TIMEOUT)
            home_text, home_links = await self._read(page, url)
            if home_text is None:
                return "", [], []
            visited.append(url)
            text_parts.append("# %s\n%s" % (url, home_text))
            links.extend(home_links)

            for sub in self._subpages(url, home_links):
                if len(visited) > C.PORTFOLIO_MAX_SUBPAGES:
                    break
                sub_text, sub_links = await self._read(page, sub)
                if not sub_text:
                    continue
                visited.append(sub)
                text_parts.append("# %s\n%s" % (sub, sub_text))
                links.extend(sub_links)
        except Exception as exc:
            if browser_died(exc):
                raise BrowserLost(str(exc))
            log.warning("portfolio %s failed: %s", url, exc)
        finally:
            if page is not None:
                try:
                    await page.close()
                except Exception:
                    pass

        unique_links, seen = [], set()
        for link in links:
            key = link.lower().rstrip("/")
            if key and key not in seen:
                seen.add(key)
                unique_links.append(link)
        return "\n\n".join(text_parts), unique_links, visited

    async def _read(self, page, url):
        try:
            response = await page.goto(url, wait_until="domcontentloaded",
                                       timeout=C.PORTFOLIO_TIMEOUT)
        except Exception as exc:
            if browser_died(exc):
                raise BrowserLost(str(exc))
            log.info("portfolio page unreachable %s: %s", url, exc)
            return None, []
        if response is not None and response.status >= 400:
            log.info("portfolio page %s returned %d", url, response.status)
            return None, []
        try:
            # A single-page app may still be painting; a short settle makes
            # the text complete without a fixed long sleep.
            await page.wait_for_load_state("networkidle", timeout=6000)
        except Exception:
            pass
        try:
            text = await page.locator("body").inner_text(timeout=10000)
        except Exception:
            text = ""
        try:
            hrefs = await page.eval_on_selector_all(
                "a[href]", "els => els.map(e => e.href)")
        except Exception:
            hrefs = []
        return text, [clean_url(h) for h in hrefs if h]

    @staticmethod
    def _subpages(base, links):
        """Same-site About / Contact / Resume pages only."""
        host = urllib.parse.urlparse(base).netloc.lower()
        chosen, seen = [], {base.rstrip("/").lower()}
        for link in links:
            parsed = urllib.parse.urlparse(link)
            if parsed.netloc.lower() != host:
                continue
            key = link.rstrip("/").lower()
            if key in seen:
                continue
            target = (parsed.path or "") + "#" + (parsed.fragment or "")
            if not _SUBPAGE_RE.search(target):
                continue
            seen.add(key)
            chosen.append(link)
            if len(chosen) >= C.PORTFOLIO_MAX_SUBPAGES:
                break
        return chosen


# ==========================================================================
# RESULT STORE
# ==========================================================================
# Where collected students live at runtime, behind one interface.
#
#     ResultStore          the contract the extraction pipeline talks to
#     SqliteStore          rank-local file, the only backend implemented now
#     IdentityIndex        read-only cumulative history of past campaigns
#
# WHY AN INTERFACE FOR ONE IMPLEMENTATION
# ---------------------------------------
# A central PostgreSQL store is the eventual destination, and when it arrives
# the extraction pipeline must not change: student_scraper.py asks
# "is this person known?" and says "save this person", and it should not care
# whether the answer comes from a local file or a server across the LAN.  So
# the contract is written down now and the backend is chosen by open_store().
#
# WHAT SQLITE DOES AND DOES NOT GIVE YOU
# --------------------------------------
# Each rank has its own runtime/rank_NNN/students.db.  Inside one rank the
# guarantees are real:
#
#     the database enforces uniqueness, not the code
#         Saving is INSERT ... ON CONFLICT DO NOTHING against UNIQUE indexes.
#         A SELECT-then-INSERT would be a race even in one process, and it is
#         exactly the pattern this replaces.
#
#     claims are leases, never permanent locks
#         A crashed rank's claim expires and the candidate becomes eligible
#         again.  Nothing about a failure blacklists a person -- that policy
#         lives in py and is unchanged.
#
# Across laptops it gives you NOTHING, and this must not be misread:
#
#     Rank 4 on laptop 1 and rank 63 on laptop 7 can both fully extract
#     github.com/student123.  They cannot see each other.  There is no
#     cross-laptop lock, no shared index, and no real-time global duplicate
#     prevention.  Duplicates between laptops are resolved ONLY when
#     export_students.py merges the stores.
#
#     history/identities.db narrows the window between CAMPAIGNS, not within
#     one: it holds identities from campaigns already closed, and it is a
#     LOCAL file.  Giving 8 laptops the same history means merging every
#     laptop's stores, building one master identities.db and copying that same
#     file back to each laptop before the next campaign starts.  That is a
#     manual step.  PostgreSQL would remove it.
#
# CAMPAIGN METADATA NEVER TOUCHES THE 65 COLUMNS
# ----------------------------------------------
# campaign_id, rank, search_country/state/city/query and source_xlsx are
# columns of this table.  They are NOT written into Data Source, Remarks /
# Notes or any other final student field: those 65 columns keep exactly the
# meaning they had before the redesign.  The search city in particular is
# where we LOOKED, and is never the student's Current City.
# --------------------------------------------------------------------------
# How long a claim is honoured before another rank may retry the candidate.
# Long enough for a slow portfolio plus resume, short enough that a crash
# does not park somebody for an hour.
CLAIM_LEASE_SECONDS = 20 * 60

SCHEMA_VERSION = 1


def _now_datetime():
    return datetime.datetime.now().replace(microsecond=0)


def _stamp(moment=None):
    return (moment or _now_datetime()).isoformat(sep=" ")


# ================================================================ contract ==
class ResultStore:
    """The contract.  A PostgreSQL backend implements exactly this."""

    backend = "abstract"
    # False for every local backend.  Only a genuinely shared central store
    # may set it, and only then may anything claim real-time global dedupe.
    provides_global_dedupe = False

    def is_known(self, username="", profile_url="", email=""):
        """Reason this person is already collected, or ""."""
        raise NotImplementedError

    def save_student(self, record, search_unit=None, campaign_id="",
                     rank=0, total_ranks=0, source_xlsx=""):
        """Save one finished 65-column record.

        Returns (saved, reason): (True, "") when the row landed, or
        (False, "<reason>") when the database rejected it as a duplicate.
        """
        raise NotImplementedError

    def claim(self, username, rank=0, lease_seconds=CLAIM_LEASE_SECONDS):
        raise NotImplementedError

    def release_claim(self, username):
        raise NotImplementedError

    def count(self):
        raise NotImplementedError

    def iter_records(self):
        raise NotImplementedError

    def close(self):
        pass


# ================================================================== sqlite ==
_STUDENTS_DDL = """
CREATE TABLE IF NOT EXISTS students (
    id                  INTEGER PRIMARY KEY,
    github_username_key TEXT    NOT NULL,
    profile_url_key     TEXT    NOT NULL DEFAULT '',
    email_key           TEXT    NOT NULL DEFAULT '',
    record_json         TEXT    NOT NULL,
    populated_fields    INTEGER NOT NULL DEFAULT 0,
    collected_at        TEXT    NOT NULL,
    campaign_id         TEXT    NOT NULL DEFAULT '',
    rank                INTEGER NOT NULL DEFAULT 0,
    total_ranks         INTEGER NOT NULL DEFAULT 0,
    search_country      TEXT    NOT NULL DEFAULT '',
    search_state        TEXT    NOT NULL DEFAULT '',
    search_city         TEXT    NOT NULL DEFAULT '',
    search_query        TEXT    NOT NULL DEFAULT '',
    source_xlsx         TEXT    NOT NULL DEFAULT ''
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_students_username
    ON students(github_username_key);
CREATE UNIQUE INDEX IF NOT EXISTS ux_students_profile
    ON students(profile_url_key) WHERE profile_url_key <> '';
CREATE UNIQUE INDEX IF NOT EXISTS ux_students_email
    ON students(email_key) WHERE email_key <> '';

CREATE TABLE IF NOT EXISTS claims (
    github_username_key TEXT    PRIMARY KEY,
    rank                INTEGER NOT NULL,
    claimed_at          TEXT    NOT NULL,
    lease_expires_at    TEXT    NOT NULL,
    attempts            INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""

_IDENTITIES_DDL = """
CREATE TABLE IF NOT EXISTS identities (
    github_username_key TEXT PRIMARY KEY,
    profile_url_key     TEXT NOT NULL DEFAULT '',
    email_key           TEXT NOT NULL DEFAULT '',
    first_campaign_id   TEXT NOT NULL DEFAULT '',
    first_seen_at       TEXT NOT NULL DEFAULT ''
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_identities_profile
    ON identities(profile_url_key) WHERE profile_url_key <> '';
CREATE UNIQUE INDEX IF NOT EXISTS ux_identities_email
    ON identities(email_key) WHERE email_key <> '';
CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""


def _connect(path, read_only=False, create=True):
    if read_only:
        if not os.path.exists(path):
            return None
        uri = "file:%s?mode=ro" % path.replace("?", "%3f").replace("#", "%23")
        connection = sqlite3.connect(uri, uri=True, timeout=30)
    else:
        if not create and not os.path.exists(path):
            return None
        directory = os.path.dirname(os.path.abspath(path))
        if directory:
            os.makedirs(directory, exist_ok=True)
        connection = sqlite3.connect(path, timeout=30)
        # WAL so a reader (the export) never blocks the writer, and FULL so a
        # power cut cannot lose an acknowledged student.
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA synchronous=FULL")
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys=ON")
    return connection


class SqliteStore(ResultStore):
    """One rank's local store.  Rank-local duplicate protection only."""

    backend = "sqlite"
    provides_global_dedupe = False

    def __init__(self, path, history=None, read_only=False):
        self.path = path
        self.read_only = read_only
        self.connection = _connect(path, read_only=read_only)
        if self.connection is None:
            raise RuntimeError("cannot open store %s" % path)
        if not read_only:
            self.connection.executescript(_STUDENTS_DDL)
            self.connection.execute(
                "INSERT OR IGNORE INTO meta(key, value) VALUES(?, ?)",
                ("schema_version", str(SCHEMA_VERSION)))
            self.connection.commit()
        # Cumulative history of CLOSED campaigns.  Read-only, optional.
        self.history = history

    # -- duplicates ---------------------------------------------------------
    def is_known(self, username="", profile_url="", email=""):
        keys = identity_keys(username, profile_url, email)
        if self.history is not None:
            reason = self.history.is_known_keys(keys)
            if reason:
                return "%s (earlier campaign)" % reason
        return self._is_known_keys(keys)

    def _is_known_keys(self, keys):
        cursor = self.connection.cursor()
        if keys["github_username_key"]:
            cursor.execute(
                "SELECT 1 FROM students WHERE github_username_key = ?",
                (keys["github_username_key"],))
            if cursor.fetchone():
                return "username"
        if keys["profile_url_key"]:
            cursor.execute(
                "SELECT 1 FROM students WHERE profile_url_key = ?",
                (keys["profile_url_key"],))
            if cursor.fetchone():
                return "profile url"
        if keys["email_key"]:
            cursor.execute("SELECT 1 FROM students WHERE email_key = ?",
                           (keys["email_key"],))
            if cursor.fetchone():
                return "email"
        return ""

    # -- saving -------------------------------------------------------------
    def save_student(self, record, search_unit=None, campaign_id="",
                     rank=0, total_ranks=0, source_xlsx=""):
        """Insert one record.  The DATABASE decides whether it is a duplicate.

        ON CONFLICT DO NOTHING against the UNIQUE indexes, so two writers
        racing on the same person cannot both succeed.  No SELECT-then-INSERT
        anywhere in this path.
        """
        if self.read_only:
            raise RuntimeError("store opened read-only: %s" % self.path)
        keys = identity_keys(
            username=record.get("GitHub Profile", ""),
            profile_url=record.get("GitHub Profile", ""),
            email=record.get("Email Address", ""))
        if not keys["github_username_key"]:
            return False, "no usable GitHub identity"

        populated = sum(1 for value in record.values() if str(value).strip())
        unit = search_unit
        cursor = self.connection.cursor()
        cursor.execute(
            """INSERT INTO students(
                   github_username_key, profile_url_key, email_key,
                   record_json, populated_fields, collected_at,
                   campaign_id, rank, total_ranks,
                   search_country, search_state, search_city, search_query,
                   source_xlsx)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT DO NOTHING""",
            (keys["github_username_key"], keys["profile_url_key"],
             keys["email_key"],
             json.dumps(record, ensure_ascii=False), populated, _stamp(),
             campaign_id, int(rank), int(total_ranks),
             getattr(unit, "country", ""), getattr(unit, "state", ""),
             getattr(unit, "search_city", ""),
             getattr(unit, "search_query", ""),
             source_xlsx))
        self.connection.commit()
        if cursor.rowcount == 1:
            return True, ""
        # The insert was rejected, so somebody already holds one of the keys.
        return False, self._is_known_keys(keys) or "duplicate"

    # -- claims -------------------------------------------------------------
    def claim(self, username, rank=0, lease_seconds=CLAIM_LEASE_SECONDS):
        """Take a lease on one candidate.  (claimed, holder_rank).

        A lease, never a lock: an expired claim is taken over, so a crash
        cannot park a candidate permanently.
        """
        if self.read_only:
            raise RuntimeError("store opened read-only: %s" % self.path)
        key = normalize_username(username)
        if not key:
            return False, 0
        now = _now_datetime()
        expires = now + datetime.timedelta(seconds=int(lease_seconds))
        cursor = self.connection.cursor()
        cursor.execute(
            """INSERT INTO claims(github_username_key, rank, claimed_at,
                                  lease_expires_at, attempts)
               VALUES(?,?,?,?,1)
               ON CONFLICT(github_username_key) DO UPDATE SET
                   rank             = excluded.rank,
                   claimed_at       = excluded.claimed_at,
                   lease_expires_at = excluded.lease_expires_at,
                   attempts         = claims.attempts + 1
               WHERE claims.lease_expires_at <= ?""",
            (key, int(rank), _stamp(now), _stamp(expires), _stamp(now)))
        self.connection.commit()
        if cursor.rowcount == 1:
            return True, int(rank)
        cursor.execute("SELECT rank FROM claims WHERE github_username_key = ?",
                       (key,))
        row = cursor.fetchone()
        return False, int(row["rank"]) if row else 0

    def release_claim(self, username):
        if self.read_only:
            return
        key = normalize_username(username)
        if not key:
            return
        self.connection.execute(
            "DELETE FROM claims WHERE github_username_key = ?", (key,))
        self.connection.commit()

    def claim_attempts(self, username):
        key = normalize_username(username)
        if not key:
            return 0
        cursor = self.connection.execute(
            "SELECT attempts FROM claims WHERE github_username_key = ?",
            (key,))
        row = cursor.fetchone()
        return int(row["attempts"]) if row else 0

    # -- reading ------------------------------------------------------------
    def count(self):
        return int(self.connection.execute(
            "SELECT COUNT(*) FROM students").fetchone()[0])

    def iter_records(self):
        """Every stored row, oldest first, as a dict.

        record is the 65-column payload; everything else is store metadata
        and is kept separate from it on purpose.
        """
        cursor = self.connection.execute(
            """SELECT github_username_key, profile_url_key, email_key,
                      record_json, populated_fields, collected_at,
                      campaign_id, rank, search_city, search_query,
                      source_xlsx
               FROM students ORDER BY collected_at, id""")
        for row in cursor:
            yield {
                "github_username_key": row["github_username_key"],
                "profile_url_key": row["profile_url_key"],
                "email_key": row["email_key"],
                "record": json.loads(row["record_json"]),
                "populated_fields": int(row["populated_fields"]),
                "collected_at": row["collected_at"],
                "campaign_id": row["campaign_id"],
                "rank": int(row["rank"]),
                "search_city": row["search_city"],
                "search_query": row["search_query"],
                "source_xlsx": row["source_xlsx"],
            }

    def known_keys(self):
        """Every identity in this store, for building a history index."""
        cursor = self.connection.execute(
            """SELECT github_username_key, profile_url_key, email_key,
                      campaign_id, collected_at FROM students""")
        return [dict(row) for row in cursor]

    def close(self):
        try:
            self.connection.close()
        except Exception:
            pass


# ================================================== cumulative history ====
class IdentityIndex:
    """Identities collected in CLOSED campaigns.  Opened READ-ONLY by ranks.

    Identity keys only -- no names, no emails beyond the key, no records --
    so it stays small and fast at a million rows while still answering the
    only question a new campaign needs to ask: have we already done this
    person?

    LOCAL to one laptop when the backend is SQLite.  See the module
    docstring: sharing it across 8 laptops is a manual merge-and-copy step
    between campaigns, not something this file does by itself.
    """

    def __init__(self, path=None, writable=False):
        self.path = path or C.HISTORY_DB
        self.writable = writable
        if writable:
            self.connection = _connect(self.path)
            self.connection.executescript(_IDENTITIES_DDL)
            self.connection.commit()
        else:
            self.connection = _connect(self.path, read_only=True)

    @property
    def available(self):
        return self.connection is not None

    def is_known_keys(self, keys):
        if self.connection is None:
            return ""
        cursor = self.connection.cursor()
        if keys.get("github_username_key"):
            cursor.execute(
                "SELECT 1 FROM identities WHERE github_username_key = ?",
                (keys["github_username_key"],))
            if cursor.fetchone():
                return "username"
        if keys.get("profile_url_key"):
            cursor.execute(
                "SELECT 1 FROM identities WHERE profile_url_key = ?",
                (keys["profile_url_key"],))
            if cursor.fetchone():
                return "profile url"
        if keys.get("email_key"):
            cursor.execute("SELECT 1 FROM identities WHERE email_key = ?",
                           (keys["email_key"],))
            if cursor.fetchone():
                return "email"
        return ""

    def is_known(self, username="", profile_url="", email=""):
        return self.is_known_keys(
            identity_keys(username, profile_url, email))

    def remember(self, keys, campaign_id="", seen_at=""):
        """Add one identity.  Only ever called by an explicit command."""
        if not self.writable:
            raise RuntimeError("identity index opened read-only: %s"
                               % self.path)
        if not keys.get("github_username_key"):
            return False
        cursor = self.connection.execute(
            """INSERT INTO identities(github_username_key, profile_url_key,
                                      email_key, first_campaign_id,
                                      first_seen_at)
               VALUES(?,?,?,?,?) ON CONFLICT DO NOTHING""",
            (keys["github_username_key"], keys.get("profile_url_key", ""),
             keys.get("email_key", ""), campaign_id, seen_at or _stamp()))
        return cursor.rowcount == 1

    def commit(self):
        if self.writable:
            self.connection.commit()

    def count(self):
        if self.connection is None:
            return 0
        return int(self.connection.execute(
            "SELECT COUNT(*) FROM identities").fetchone()[0])

    def close(self):
        if self.connection is not None:
            try:
                self.connection.close()
            except Exception:
                pass


# ================================================================= factory ==
def open_store(target=None, history=None, read_only=False):
    """Open the configured backend.

    A path or a sqlite:// URL gives SqliteStore.  A postgres:// URL is
    recognised and refused clearly rather than half-working: there is no
    PostgreSQL backend yet, and inventing connection details would be worse
    than saying so.
    """
    text = str(target or "")
    lowered = text.lower()
    if lowered.startswith(("postgres://", "postgresql://", "mysql://")):
        raise NotImplementedError(
            "central database backends are not implemented yet.\n"
            "Requested: %s\n"
            "The store interface is ready for one, but it needs connection\n"
            "details and credentials that this project does not have.  Use a\n"
            "rank-local sqlite store for now; duplicates between laptops are\n"
            "resolved by export_students.py." % text)
    if lowered.startswith("sqlite://"):
        text = text[len("sqlite://"):]
    return SqliteStore(text, history=history, read_only=read_only)


def open_history(path=None, writable=False):
    return IdentityIndex(path, writable=writable)


# ==========================================================================
# RESUME PARSER
# ==========================================================================
# Public resume/CV download and text extraction.
#
# Safety first (spec 29): a file is validated on content type, extension and
# size BEFORE it is parsed, archives and executables are refused outright, and
# nothing downloaded is ever executed.
#
# Parsers use what this machine actually has:
#   PDF   pypdf/pdfminer when installed, otherwise poppler's `pdftotext`
#   DOCX  python-docx when installed, otherwise the stdlib (a .docx is a zip
#         holding word/document.xml)
#   ODT   stdlib zip + content.xml
#   TXT   read directly
#   HTML  rendered through Playwright, or tag-stripped
# No third-party install is required for any of them.
# --------------------------------------------------------------------------
_XML_TAG_RE = re.compile(rb"<[^>]+>")
_PARA_BREAK_RE = re.compile(rb"</w:p>|</text:p>")


def _extension_of(url, content_type=""):
    path = urllib.parse.urlparse(url).path
    ext = os.path.splitext(path)[1].lower()
    if ext:
        return ext
    mapping = {
        "application/pdf": ".pdf",
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document": ".docx",
        "application/msword": ".doc",
        "application/vnd.oasis.opendocument.text": ".odt",
        "text/plain": ".txt",
        "text/html": ".html",
    }
    return mapping.get((content_type or "").split(";")[0].strip().lower(), "")


def is_safe_resume(url, content_type, size):
    """(ok, reason) -- checked before anything is written or parsed."""
    ext = _extension_of(url, content_type)
    if ext in C.RESUME_FORBIDDEN_EXT:
        return False, "refused file type %s" % ext
    if size is not None and size > C.RESUME_MAX_BYTES:
        return False, "file too large (%d bytes)" % size
    if size == 0:
        return False, "empty file"
    base = (content_type or "").split(";")[0].strip().lower()
    if base and base not in C.RESUME_ALLOWED_MIME:
        return False, "unsupported content type %s" % base
    if ext and ext not in C.RESUME_ALLOWED_EXT and ext not in (".html", ".htm"):
        return False, "unsupported extension %s" % ext
    return True, ""


class ResumeParser:
    def __init__(self, say=print, download_dir=None):
        self.say = say
        # Where a downloaded resume is written before parsing.  In the
        # distributed architecture this is the CURRENT rank's own
        # runtime/rank_NNN/temp_downloads -- never a shared global path, so two
        # ranks can never collide.  When no rank dir is supplied (the debug
        # mode, which writes no output), a per-instance scratch dir is used,
        # which is likewise private.
        self._download_dir = download_dir
        self._scratch = ""

    def _temp_dir(self):
        if self._download_dir:
            return self._download_dir
        if not self._scratch:
            import tempfile
            self._scratch = tempfile.mkdtemp(prefix="resume_")
        return self._scratch

    # ------------------------------------------------------------ fetch --
    async def fetch(self, context, url, username):
        """Return (text, saved_path, remark)."""
        url = clean_url(url)
        if not url:
            return "", "", ""

        # A Google Docs/Drive viewer link is a web page, not a file.
        if re.search(r"(docs|drive)\.google\.com", url, re.I):
            text, remark = await self._read_web(context, url)
            return text, "", remark

        request = context.request
        try:
            response = await request.get(url, timeout=C.RESUME_TIMEOUT)
        except Exception as exc:
            return "", "", "resume download failed: %s" % exc
        if not response.ok:
            if response.status in (401, 403):
                return "", "", "resume is private (HTTP %d)" % response.status
            return "", "", "resume unavailable (HTTP %d)" % response.status

        headers = {k.lower(): v for k, v in response.headers.items()}
        content_type = headers.get("content-type", "")
        try:
            body = await response.body()
        except Exception as exc:
            return "", "", "resume body unreadable: %s" % exc

        ok, reason = is_safe_resume(url, content_type, len(body))
        if not ok:
            return "", "", reason

        ext = _extension_of(url, content_type) or ".bin"
        if ext in (".html", ".htm") or content_type.startswith("text/html"):
            text, remark = await self._read_web(context, url)
            return text, "", remark

        download_dir = self._temp_dir()
        os.makedirs(download_dir, exist_ok=True)
        safe = re.sub(r"[^A-Za-z0-9._-]", "_", "%s_resume%s" % (username, ext))
        path = os.path.join(download_dir, safe)
        with open(path, "wb") as handle:
            handle.write(body)

        if not self._magic_matches(path, ext):
            os.remove(path)
            return "", "", "resume content did not match its extension"

        text = self.extract_text(path)
        if not text.strip():
            return "", path, "resume produced no text (scanned or empty)"
        return text, path, ""

    async def _read_web(self, context, url):
        page = None
        try:
            page = await context.new_page()
            response = await page.goto(url, wait_until="domcontentloaded",
                                       timeout=C.RESUME_TIMEOUT)
            if response is not None and response.status in (401, 403):
                return "", "resume needs a login; skipped"
            try:
                await page.wait_for_load_state("networkidle", timeout=6000)
            except Exception:
                pass
            text = await page.locator("body").inner_text(timeout=10000)
            if re.search(r"request access|sign in to continue|you need access",
                         text or "", re.I):
                return "", "resume is private; skipped"
            return text or "", ""
        except Exception as exc:
            return "", "resume page failed: %s" % exc
        finally:
            if page is not None:
                try:
                    await page.close()
                except Exception:
                    pass

    @staticmethod
    def _magic_matches(path, ext):
        """Cheap sanity check that the bytes are what the name claims."""
        try:
            with open(path, "rb") as handle:
                head = handle.read(8)
        except OSError:
            return False
        if ext == ".pdf":
            return head.startswith(b"%PDF")
        if ext in (".docx", ".odt"):
            return head.startswith(b"PK")
        if ext == ".doc":
            return head.startswith(b"\xd0\xcf\x11\xe0") or head.startswith(
                b"PK")
        return True

    # ------------------------------------------------------------ parse --
    def extract_text(self, path):
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".pdf":
                return self._pdf(path)
            if ext == ".docx":
                return self._docx(path)
            if ext == ".odt":
                return self._odt(path)
            if ext in (".txt", ".rtf"):
                with open(path, encoding="utf-8", errors="replace") as handle:
                    text = handle.read()
                return self._dertf(text) if ext == ".rtf" else text
            if ext == ".doc":
                return self._legacy_doc(path)
        except Exception as exc:
            log.warning("resume parse failed for %s: %s", path, exc)
        return ""

    @staticmethod
    def _pdf(path):
        for module in ("pypdf", "PyPDF2"):
            try:
                reader = __import__(module).PdfReader(path)
                text = "\n".join((page.extract_text() or "")
                                 for page in reader.pages)
                if text.strip():
                    return text
            except Exception:
                pass
        try:
            from pdfminer.high_level import extract_text
            text = extract_text(path)
            if text.strip():
                return text
        except Exception:
            pass
        # poppler is present on this machine and handles the rest.
        if shutil.which("pdftotext"):
            try:
                done = subprocess.run(
                    ["pdftotext", "-layout", "-enc", "UTF-8", path, "-"],
                    capture_output=True, timeout=60, check=False)
                return done.stdout.decode("utf-8", "replace")
            except Exception as exc:
                log.warning("pdftotext failed: %s", exc)
        return ""

    @staticmethod
    def _docx(path):
        try:
            import docx
            return "\n".join(par.text for par in docx.Document(path).paragraphs)
        except Exception:
            pass
        # A .docx is a zip; word/document.xml holds the text.
        with zipfile.ZipFile(path) as archive:
            raw = archive.read("word/document.xml")
        raw = _PARA_BREAK_RE.sub(b"\n", raw)
        return html.unescape(
            _XML_TAG_RE.sub(b" ", raw).decode("utf-8", "replace"))

    @staticmethod
    def _odt(path):
        with zipfile.ZipFile(path) as archive:
            raw = archive.read("content.xml")
        raw = _PARA_BREAK_RE.sub(b"\n", raw)
        return html.unescape(
            _XML_TAG_RE.sub(b" ", raw).decode("utf-8", "replace"))

    @staticmethod
    def _legacy_doc(path):
        for tool in ("antiword", "catdoc"):
            if shutil.which(tool):
                try:
                    done = subprocess.run([tool, path], capture_output=True,
                                          timeout=60, check=False)
                    text = done.stdout.decode("utf-8", "replace")
                    if text.strip():
                        return text
                except Exception:
                    pass
        if shutil.which("soffice"):
            try:
                out = os.path.dirname(path)
                subprocess.run(
                    ["soffice", "--headless", "--convert-to", "txt:Text",
                     "--outdir", out, path],
                    capture_output=True, timeout=180, check=False)
                converted = os.path.splitext(path)[0] + ".txt"
                if os.path.exists(converted):
                    with open(converted, encoding="utf-8",
                              errors="replace") as handle:
                        text = handle.read()
                    os.remove(converted)
                    return text
            except Exception as exc:
                log.warning("soffice conversion failed: %s", exc)
        return ""

    @staticmethod
    def _dertf(text):
        text = re.sub(r"\\'[0-9a-f]{2}", " ", text)
        text = re.sub(r"\\[a-z]+-?\d* ?", " ", text)
        return re.sub(r"[{}]", " ", text)

    # ---------------------------------------------------------- cleanup --
    @staticmethod
    def discard(path):
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError as exc:
                log.warning("could not remove %s: %s", path, exc)


# ==========================================================================
# SKILLS EXTRACTOR
# ==========================================================================
# Skills, weighted by where they were found.
#
# Searching a whole page for technology words is how "I can C the problem"
# becomes the C programming language and how a footer link turns into a tool.
# A mention only counts as a skill when the place it appears supports that
# reading, so every hit carries the weight of its section:
#
#     skills / tech-stack section   1.00   an explicit declaration
#     resume summary, bio, about    0.70   the person describing themselves
#     profile README body           0.60   prose, often about projects
#     repository descriptions,
#     footers, navigation           0.30   incidental mentions
#
# Primary skills are the ones with the strongest, most repeated evidence --
# not simply the first five found.
# --------------------------------------------------------------------------
# --------------------------------------------------------- vocabularies --
PROGRAMMING_LANGUAGES = [
    "Python", "Java", "JavaScript", "TypeScript", "C++", "C#", "C", "Go",
    "Rust", "PHP", "Ruby", "Kotlin", "Swift", "Dart", "SQL", "R", "MATLAB",
    "Bash", "Shell", "Perl", "Scala", "Haskell", "Elixir", "Lua", "Julia",
    "Objective-C", "Assembly", "Solidity", "Groovy", "VB.NET", "COBOL",
    "PL/SQL", "Racket", "Clojure", "F#", "Erlang", "Fortran",
]

FRAMEWORKS = [
    "HTML5", "CSS3", "HTML", "CSS", "Sass", "SCSS", "Tailwind CSS",
    "React Native", "React", "Next.js", "Node.js", "Express", "Django",
    "Flask", "FastAPI", "Spring Boot", "Spring", ".NET", "ASP.NET",
    "Angular", "Vue", "Svelte", "Flutter", "Laravel", "Rails",
    "TensorFlow", "PyTorch", "Keras", "scikit-learn", "Pandas", "NumPy",
    "Matplotlib", "Seaborn", "OpenCV", "Hugging Face", "LangChain",
    "Streamlit", "Gradio", "AWS", "Azure", "GCP",
    "Firebase", "Supabase", "PostgreSQL", "MySQL", "MongoDB", "Redis",
    "SQLite", "Oracle", "Elasticsearch", "Kafka", "GraphQL", "REST API",
    "Bootstrap", "Tailwind", "jQuery", "Hibernate", "Spark", "Hadoop",
    "Airflow", "dbt", "Power BI", "Tableau", "Selenium", "Cypress",
]

TOOLS = [
    "Docker", "Kubernetes",
    "Git", "GitHub", "GitLab", "Bitbucket", "VS Code", "IntelliJ", "Eclipse",
    "PyCharm", "Jupyter", "Colab", "Linux", "Ubuntu", "Windows Server",
    "Postman", "Figma", "Jenkins", "Terraform", "Ansible", "Jira",
    "Confluence", "Wireshark", "Nmap", "Splunk", "Excel", "Maven", "Gradle",
    "Nginx", "Apache", "Active Directory", "VMware", "Cisco", "ServiceNow",
    "Anaconda", "Slack", "Notion", "Trello",
]

# Competencies that are neither a language, a framework nor a tool.  They
# belong in Primary/Secondary Skills but must not be forced into one of
# the three specific columns.
CONCEPTS = [
    "DSA", "Object Oriented Programming", "Machine Learning",
    "Deep Learning", "Data Structures", "Algorithms", "Operating Systems",
    "Computer Networks", "DBMS", "Competitive Programming",
]

# Spellings that mean the same competency.
CONCEPT_ALIASES = {
    "DSA": ["dsa", "data structures and algorithms",
            "data structures & algorithms", "data structure and algorithm",
            "ds & algo", "ds and algo", "dsa using java"],
    "Object Oriented Programming": ["oops", "oop",
                                    "object oriented programming"],
    "Machine Learning": ["machine learning", "ml"],
    "Deep Learning": ["deep learning"],
    "DBMS": ["dbms", "database management system"],
    "Competitive Programming": ["competitive programming", "cp"],
    "Operating Systems": ["operating systems", "operating system"],
    "Computer Networks": ["computer networks", "computer networking"],
}

ALL_SKILLS = PROGRAMMING_LANGUAGES + FRAMEWORKS + TOOLS + CONCEPTS

# Badge alt text is matched case-insensitively against this map.
_CANONICAL = {}
for _skill in ALL_SKILLS:
    _CANONICAL[_skill.lower()] = _skill
for _canonical, _aliases in CONCEPT_ALIASES.items():
    for _alias in _aliases:
        _CANONICAL[_alias] = _canonical
# Badge labels and ?logo= slugs use their own spellings.
_CANONICAL.update({
    "openjdk": "Java", "java17": "Java", "js": "JavaScript",
    "nodedotjs": "Node.js", "node": "Node.js", "nextdotjs": "Next.js",
    "html": "HTML5", "html-5": "HTML5", "css": "CSS3", "css-3": "CSS3",
    "c++": "C++", "cplusplus": "C++", "csharp": "C#", "c-sharp": "C#",
    "c%2b%2b": "C++", "dotnet": ".NET", "tailwindcss": "Tailwind CSS",
    "mysql": "MySQL", "postgresql": "PostgreSQL", "mongodb": "MongoDB",
    "visualstudiocode": "VS Code", "vscode": "VS Code",
    "github": "GitHub", "git": "Git", "gnubash": "Bash",
    "scikitlearn": "scikit-learn", "sklearn": "scikit-learn",
    "amazonaws": "AWS", "googlecloud": "GCP",
    "microsoftazure": "Azure", "springboot": "Spring Boot",
})


def canonical_skill(name):
    """Map a badge alt/label/logo onto a taxonomy entry, or ""."""
    key = re.sub(r"\s+", " ", (name or "").strip().lower())
    if not key:
        return ""
    if key in _CANONICAL:
        return _CANONICAL[key]
    compact = key.replace(" ", "").replace("-", "").replace(".", "")
    for candidate, canonical in _CANONICAL.items():
        if candidate.replace(" ", "").replace("-", "").replace(".", "") \
                == compact:
            return canonical
    return ""

# Single letters are words far more often than languages.
SINGLE_LETTER = {"C", "R"}

# Names that appear on every page of the site being scraped.  Every GitHub
# profile says "GitHub" in its navigation and footer, so the word alone is
# no evidence that the person lists it as a tool.  It counts only when
# declared: in a skills section, as a badge, or in a self-statement.
AMBIENT_SKILLS = {"GitHub", "Git"}

# Weight per section, keyed by the canonical names py produces.
SECTION_WEIGHT = {
    "_other": 0.2,
    "skills": 1.0,
    "projects": 0.5,
    "about": 0.7,
    "experience": 0.7,
    "education": 0.5,
    "certifications": 0.6,
    "achievements": 0.4,
    "contact": 0.0,
    "_head": 0.6,
}

# Source weight, applied on top of the section weight.
SOURCE_WEIGHT = {
    "resume": 1.0,
    "readme_raw": 0.95,   # the Markdown source: badges live only here
    "portfolio": 0.85,
    "readme": 0.8,
    "bio": 0.9,
    "profile": 0.35,      # the raw GitHub page: repo names, nav, footer
    "repo_language": 0.3,  # supporting evidence only (spec 18)
    "xlsx": 0.5,
}

# A badge is an explicit declaration, worth more than the same word
# appearing in a sentence.
BADGE_WEIGHT = 1.1

PRIMARY_THRESHOLD = 1.2
PRIMARY_LIMIT = 8
SECONDARY_LIMIT = 15


def _pattern(skill):
    escaped = re.escape(skill)
    return re.compile(r"(?<![A-Za-z0-9+#.])" + escaped +
                      r"(?![A-Za-z0-9+#]|\.[A-Za-z])", re.I)


_PATTERNS = {skill: _pattern(skill) for skill in ALL_SKILLS}

_TECH_CUE = re.compile(
    r"\b(languages?|libraries|skills|stack|technologies|tools|frameworks|"
    r"programming|proficient|experienced\s+in|familiar\s+with|know|using|"
    r"built\s+with|written\s+in)\b", re.I)


def _single_letter_allowed(skill, text, in_skills_section):
    """C and R only count where a technical reading is unambiguous."""
    if in_skills_section:
        return True
    pattern = _PATTERNS[skill]
    for match in pattern.finditer(text):
        line = text[text.rfind("\n", 0, match.start()) + 1:]
        line = line[:line.find("\n")] if "\n" in line else line
        if _TECH_CUE.search(line):
            return True
        # "C, C++, Java" or "Python | R | SQL": a delimited technical list.
        around = text[max(0, match.start() - 30):match.end() + 30]
        siblings = sum(1 for other in ("C++", "C#", "Python", "Java", "SQL",
                                       "MATLAB", "JavaScript", "Go")
                       if _PATTERNS.get(other) and
                       _PATTERNS[other].search(around))
        # "C, C++, Java", "Python | R | SQL" and "C and C++" all read as
        # technical lists.
        if siblings >= 1 and re.search(r"[,|/•]|\band\b|&", around, re.I):
            return True
    return False


# "I no longer use Java" is not a claim to Java.
_NEGATION = re.compile(
    r"\b(no\s+longer|not\s+using|don'?t\s+use|never\s+used|stopped\s+using|"
    r"instead\s+of|migrated\s+away\s+from|moving\s+away\s+from|"
    r"used\s+to\s+use)\b", re.I)


def _is_negated(pattern, text):
    """True when every mention of the skill sits in a negative clause."""
    found = list(pattern.finditer(text))
    if not found:
        return False
    for match in found:
        before = text[max(0, match.start() - 45):match.start()]
        if not _NEGATION.search(before):
            return False
    return True


# A URL is not prose.  "github-readme-stats.vercel.app" contains the word
# "github", which is how a stats image turned into a claimed tool.
_URL_RE = re.compile(r"(?:https?://|www\.)\S+", re.I)
_MD_IMAGE_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")


def strip_non_prose(text):
    """Remove URLs and image syntax before scanning text for skill words."""
    body = _MD_IMAGE_RE.sub(" ", text or "")
    return _URL_RE.sub(" ", body)


# An explicit statement about oneself is stronger evidence than a name in
# a badge list: "I'm currently learning DSA using java" says more than a
# row of logos does.
_DECLARATION = re.compile(
    r"\b(currently\s+learning|learning|studying|working\s+with|"
    r"experienced\s+in|proficient\s+in|skilled\s+in|specializ\w+\s+in|"
    r"i\s+use|i\s+work\s+with|i\s+code\s+in|building\s+with|"
    r"hands[- ]on\s+with|familiar\s+with)\b", re.I)
DECLARATION_BOOST = 0.6


def _declared_near(text, pattern):
    """True when a mention sits inside an explicit self-statement."""
    for match in pattern.finditer(text or ""):
        window = text[max(0, match.start() - 90):match.end() + 40]
        if _DECLARATION.search(window):
            return True
    return False


def _badge_skill_names(markdown, in_skills_section):
    """Technologies declared by badges in this block of Markdown.

    A badge only counts when it names something in the taxonomy, so the
    Instagram/LinkedIn/X badges a "Socials" section uses -- written in the
    very same shields.io style -- can never become skills.  Outside a
    skills section a badge additionally has to be a real technology badge.
    """
    names = []
    for badge in extract_badges(markdown):
        if badge["decoration"]:
            continue
        if not in_skills_section and not badge["shields"]:
            continue
        for candidate in (badge["alt"], badge["label"], badge["logo"]):
            skill = canonical_skill(candidate)
            if skill:
                if skill not in names:
                    names.append(skill)
                break
    return names


def concept_names(text):
    """Competencies stated in prose: "currently learning DSA using java"."""
    body = " %s " % re.sub(r"\s+", " ", (text or "").lower())
    found = []
    for canonical, aliases in CONCEPT_ALIASES.items():
        for alias in aliases:
            if re.search(r"(?<![a-z])%s(?![a-z])" % re.escape(alias), body):
                if canonical not in found:
                    found.append(canonical)
                break
    return found


def collect_skills(documents):
    """Score every skill across weighted documents.

    `documents` is a list of (text, source_name) pairs.  Returns
    {skill: score} where a higher score means stronger evidence.
    """
    scores = defaultdict(float)
    seen_in = defaultdict(set)

    for text, source in documents:
        if not text:
            continue
        source_weight = SOURCE_WEIGHT.get(source, 0.5)
        parsed = parse_sections(text)
        blocks = [(name, body) for name, body in parsed.items() if body]
        if not blocks:
            blocks = [("_head", text)]

        for section, body in blocks:
            weight = SECTION_WEIGHT.get(section, 0.4) * source_weight
            if weight <= 0:
                continue
            in_skills = section == "skills"

            # Markdown badges: the only place a README's tech stack
            # actually exists, because <img> contributes nothing to the
            # rendered text.
            for name in _badge_skill_names(body, in_skills):
                scores[name] += BADGE_WEIGHT * source_weight
                seen_in[name].add((source, "badge"))

            prose = strip_non_prose(body)

            for name in concept_names(prose):
                scores[name] += weight
                if _declared_near(prose, re.compile(
                        r"(?i)\b%s\b" % re.escape(name))) or \
                        any(_declared_near(prose, re.compile(
                            r"(?i)(?<![a-z])%s(?![a-z])" % re.escape(alias)))
                            for alias in CONCEPT_ALIASES.get(name, [])):
                    scores[name] += DECLARATION_BOOST
                seen_in[name].add((source, section))

            for skill in ALL_SKILLS:
                if skill in CONCEPTS:
                    continue          # handled by concept_names()
                if skill in SINGLE_LETTER:
                    if not _single_letter_allowed(skill, prose, in_skills):
                        continue
                if not _PATTERNS[skill].search(prose):
                    continue
                if _is_negated(_PATTERNS[skill], prose):
                    continue
                declared = _declared_near(prose, _PATTERNS[skill])
                if skill in AMBIENT_SKILLS and not in_skills and not declared:
                    continue
                scores[skill] += weight
                if declared:
                    scores[skill] += DECLARATION_BOOST
                seen_in[skill].add((source, section))

    # Evidence in more than one place is stronger than repetition in one.
    for skill, places in seen_in.items():
        if len(places) > 1:
            scores[skill] += 0.25 * (len(places) - 1)
    return dict(scores)


def classify(scores):
    """Split scored skills into the five output columns."""
    ranked = sorted(scores.items(), key=lambda item: (-item[1], item[0]))
    languages, frameworks, tools = [], [], []
    for skill, _score in ranked:
        if skill in PROGRAMMING_LANGUAGES:
            languages.append(skill)
        elif skill in FRAMEWORKS:
            frameworks.append(skill)
        elif skill in TOOLS:
            tools.append(skill)

    primary = [skill for skill, score in ranked if score >= PRIMARY_THRESHOLD]
    if not primary:
        # Nothing cleared the bar: take the strongest few rather than none.
        primary = [skill for skill, _score in ranked[:5]]
    primary = primary[:PRIMARY_LIMIT]
    secondary = [skill for skill, _score in ranked
                 if skill not in primary][:SECONDARY_LIMIT]

    return {
        "primary": primary,
        "secondary": secondary,
        "languages": languages,
        "frameworks": frameworks,
        "tools": tools,
    }


def extract_skills(documents):
    return classify(collect_skills(documents))


def as_text(values):
    return join_list(values)


# ==========================================================================
# STATE MANAGER
# ==========================================================================
# Crash-safe run state for this folder.
#
# Every write is atomic (tmp file -> fsync -> os.replace) so a kill at any
# instant leaves either the previous state or the new one, never a truncated
# file.
#
# WHAT THIS FILE IS AND IS NOT
# ----------------------------
# It tracks ONE search at a time: which search is being worked on, whether its
# XLSX has been collected, how far down that XLSX the processing has got, and
# which person is being handled right now.  Those blocks are what govern the
#
#     SUCCESS                 -> clean up
#     ANY UNRESOLVED FAILURE  -> preserve and resume later
#
# rule, and nothing about when a keyword XLSX may be deleted lives anywhere
# else.
#
# It does NOT track which search comes next.  That is current_row.json, written
# by student_scraper after a search finishes.  Keeping the two apart is what
# makes resuming simple: the position is one integer in its own file, and this
# file only ever describes the search currently in hand.
#
# NOTHING HERE CAN REFUSE TO RESUME
# ---------------------------------
# An earlier version compared campaign fingerprints, the rank, the rank count
# and an assignment-algorithm version, and raised rather than resume when any
# of them had moved.  All four are gone, along with the campaign they belonged
# to.  A state file that was written by an older version is read for the keys
# that still mean something and given defaults for the rest, so changing the
# keyword file or num_sys can no longer stop a run -- it just changes which
# searches this copy walks.
# --------------------------------------------------------------------------
class NoCurrentUnit(RuntimeError):
    """There is no search unit recorded, so there is nothing to retry."""


def _blank():
    return {
        "version": C.STATE_VERSION,
        # Where in this copy's search list we are.  A hint for reporting:
        # current_row.json is what the run actually resumes from.
        "position": {"unit_index": 0, "units_done": 0},
        # The identity of the search currently being worked on.
        "search_unit": {
            "country": "", "state": "", "canonical_city": "",
            "search_city": "", "base_keyword": "", "search_query": "",
            "slug": "",
        },
        "phase": "search",
        "collection": {"complete": False, "xlsx_file": "", "rows": 0},
        "processing": {"next_row_index": 0, "total_rows": 0,
                       "complete": False},
        "current_candidate": {"username": "", "source_row": -1},
        "cleanup_complete": False,
        # Page-by-page collection progress for the CURRENT search.  Reset
        # with the rest of the per-search blocks, because it describes one
        # search's crawl and nothing else.
        "crawl": {},
        # Mirror of the extension's learned pagination selector, keyed by
        # host.  Survives a wiped browser profile.
        "next_selectors": {},
    }


class StateManager:
    """This folder's state file.

    path exists so the test suite can point a manager at a temporary file;
    production always uses config.STATE_FILE.
    """

    def __init__(self, path=None):
        self.path = path or C.STATE_FILE
        self.data = self._load()

    # -- loading ------------------------------------------------------------
    def _load(self):
        blank = _blank()
        if not os.path.exists(self.path):
            return blank
        try:
            with open(self.path, encoding="utf-8") as handle:
                loaded = json.load(handle)
        except Exception as exc:
            log.warning("state file unreadable (%s); starting fresh", exc)
            return blank
        if not isinstance(loaded, dict):
            log.warning("state file is not an object; starting fresh")
            return blank

        # Keys this version does not know about are dropped, and keys it
        # expects but the file lacks get their defaults.  A state written by
        # an older version is therefore usable rather than fatal.
        for key, value in blank.items():
            loaded.setdefault(key, value)
        for section in ("position", "search_unit", "collection",
                        "processing", "current_candidate", "crawl"):
            merged = dict(blank[section])
            merged.update(loaded.get(section) or {})
            loaded[section] = merged
        loaded["version"] = C.STATE_VERSION
        return loaded

    @property
    def is_fresh(self):
        """True when nothing has been recorded yet."""
        return not os.path.exists(self.path)

    # -- writing ------------------------------------------------------------
    def save(self):
        directory = os.path.dirname(self.path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, self.path)

    def set(self, **values):
        self.data.update(values)
        self.save()

    def set_in(self, section, **values):
        self.data.setdefault(section, {}).update(values)
        self.save()

    # -- search-unit lifecycle ---------------------------------------------
    def begin_unit(self, index, unit):
        """Start one search.  Resets the per-search blocks unless resuming it.

        Same logic v3's begin_keyword() used, with the identity comparison
        widened from a keyword string to the full unit identity: two cities
        can share a base keyword, so the keyword alone no longer identifies
        the work.
        """
        stored = self.data.get("search_unit") or {}
        same = (stored.get("search_query") == unit.search_query
                and stored.get("canonical_city") == unit.canonical_city
                and stored.get("search_city") == unit.search_city)
        if not same:
            self.data["collection"] = {"complete": False, "xlsx_file": "",
                                       "rows": 0}
            self.data["processing"] = {"next_row_index": 0, "total_rows": 0,
                                       "complete": False}
            self.data["current_candidate"] = {"username": "",
                                              "source_row": -1}
            self.data["cleanup_complete"] = False
            self.data["crawl"] = {}
        self.data["search_unit"] = unit.as_state()
        self.data.setdefault("position", {})["unit_index"] = int(index)
        self.data["phase"] = "search"
        self.save()

    def finish_unit(self):
        self.data["cleanup_complete"] = True
        self.data["processing"]["complete"] = True
        self.data["phase"] = "done"
        self.save()

    def advance_unit(self):
        position = self.data.setdefault("position", {})
        position["unit_index"] = int(position.get("unit_index", 0)) + 1
        position["units_done"] = int(position.get("units_done", 0)) + 1
        self.data["search_unit"] = _blank()["search_unit"]
        self.data["phase"] = "search"
        self.data["collection"] = {"complete": False, "xlsx_file": "",
                                   "rows": 0}
        self.data["processing"] = {"next_row_index": 0, "total_rows": 0,
                                   "complete": False}
        self.data["current_candidate"] = {"username": "", "source_row": -1}
        self.data["crawl"] = {}
        self.data["cleanup_complete"] = False
        self.save()

    # -- targeted recovery --------------------------------------------------
    def reset_for_recollection(self):
        """Clear THIS unit's collection and processing so it is re-collected.

        ONE atomic write.  Two set_in() calls would put a state on disk in
        between where collection had been cleared but processing still claimed
        rows from the file that is about to be archived -- a crash in that
        window would leave the rank resuming into a row count with no source.

        Deliberately untouched:

            position        unit_index and units_done -- this is a RETRY of
                            the current search, not a rewind of the list
            search_unit     the identity being retried
            next_selectors  a learned pagination selector belongs to the host,
                            not to one search

        Returns the search_unit being retried, so a caller can report it.
        """
        unit = dict(self.data.get("search_unit") or {})
        if not unit.get("search_query"):
            raise NoCurrentUnit(
                "%s records no current search unit, so there is nothing to "
                "retry.\nThis rank has not started a search yet; just run it "
                "normally." % self.path)
        blank = _blank()
        self.data["collection"] = dict(blank["collection"])
        self.data["processing"] = dict(blank["processing"])
        self.data["current_candidate"] = dict(blank["current_candidate"])
        self.data["crawl"] = {}
        self.data["cleanup_complete"] = False
        self.data["phase"] = "search"
        self.save()
        return unit

    def describe_current(self):
        """Everything --retry-current-search reports.  Reads only."""
        position = self.data.get("position") or {}
        unit = self.data.get("search_unit") or {}
        collection = self.data.get("collection") or {}
        processing = self.data.get("processing") or {}
        return {
            "state_file": self.path,
            "unit_index": int(position.get("unit_index", 0) or 0),
            "units_done": int(position.get("units_done", 0) or 0),
            "search_query": unit.get("search_query", ""),
            "canonical_city": unit.get("canonical_city", ""),
            "search_city": unit.get("search_city", ""),
            "state": unit.get("state", ""),
            "base_keyword": unit.get("base_keyword", ""),
            "slug": unit.get("slug", ""),
            "phase": self.data.get("phase", ""),
            "collection_complete": bool(collection.get("complete")),
            "collection_file": collection.get("xlsx_file", "") or "",
            "collection_rows": int(collection.get("rows", 0) or 0),
            "next_row_index": int(processing.get("next_row_index", 0) or 0),
            "total_rows": int(processing.get("total_rows", 0) or 0),
            "processing_complete": bool(processing.get("complete")),
            "cleanup_complete": bool(self.data.get("cleanup_complete")),
        }

    # -- convenience --------------------------------------------------------
    @property
    def unit_index(self):
        return int((self.data.get("position") or {}).get("unit_index", 0) or 0)

    @property
    def units_done(self):
        return int((self.data.get("position") or {}).get("units_done", 0) or 0)

    @property
    def search_query(self):
        return (self.data.get("search_unit") or {}).get("search_query", "")

    @property
    def collection_complete(self):
        return bool(self.data.get("collection", {}).get("complete"))

    @property
    def collection_file(self):
        return self.data.get("collection", {}).get("xlsx_file", "") or ""

    @property
    def next_row_index(self):
        return int(self.data.get("processing", {}).get("next_row_index", 0)
                   or 0)

    def next_selector_for(self, host):
        return (self.data.get("next_selectors") or {}).get(host, "")

    def remember_next_selector(self, host, selector):
        if not selector:
            return
        table = self.data.setdefault("next_selectors", {})
        if table.get(host) == selector:
            return
        table[host] = selector
        self.save()


# ==========================================================================
# STUDENT MERGER
# ==========================================================================
# Turns a SourceBundle into one final-schema row.
#
# Each field has its own priority order (spec 64) rather than a blind
# concatenation, and a disagreement between sources is recorded in
# Remarks / Notes rather than silently resolved (spec 65).
#
# Pipeline (spec 66):
#     raw -> clean -> detect -> context-validate -> standardise -> merge ->
#     final validation -> Excel row
# --------------------------------------------------------------------------
class _Sourced:
    """A value plus the source it came from, so conflicts can be explained."""

    __slots__ = ("value", "source")

    def __init__(self, value="", source=""):
        self.value = value
        self.source = source

    def __bool__(self):
        return bool(self.value)


def _first(*pairs):
    """First non-empty (value, source) pair."""
    for value, source in pairs:
        if value:
            return _Sourced(value, source)
    return _Sourced()


class StudentMerger:
    def __init__(self, say=print):
        self.say = say

    def build(self, bundle):
        candidate = bundle.candidate
        fields = bundle.github_fields or {}
        remarks = list(bundle.remarks)

        # ---- raw sources, each tagged with where it came from ------------
        github_text = clean_text(bundle.github_text)
        readme = clean_text(fields.get("readme", ""))
        # The Markdown source of the profile README.  Badges only exist
        # here: rendered, a tech stack is a row of <img> with no text.
        readme_raw = fields.get("readme_raw", "") or ""
        portfolio_text = clean_text(bundle.portfolio_text)
        resume_text = clean_text(bundle.resume_text)
        xlsx_text = clean_text(
            "%s %s" % (candidate.xlsx_bio, " ".join(candidate.raw_cells)))
        bio = clean_text(fields.get("bio", "") or candidate.xlsx_bio)

        documents = [(resume_text, "resume"), (portfolio_text, "portfolio"),
                     (readme_raw, "readme_raw"), (readme, "readme"),
                     (bio, "bio"),
                     (github_text, "profile"), (xlsx_text, "xlsx")]
        everything = "\n".join(text for text, _name in documents if text)
        # Prose the person wrote about themselves, in priority order.
        prose = "\n".join(part for part in (resume_text, portfolio_text,
                                            readme, bio) if part)

        resume_sections = parse_sections(resume_text)
        portfolio_sections = parse_sections(portfolio_text)
        readme_sections = parse_sections(readme_raw or readme)

        record = {column: "" for column in C.FINAL_COLUMNS}

        # ---------------------------------------------------------- name --
        name = _first(
            (flatten(fields.get("fullname", "")), "GitHub"),
            (self._resume_name(resume_text), "Resume"),
            (self._portfolio_name(portfolio_text), "Portfolio"),
            (flatten(candidate.xlsx_name), "Instant Data Scraper"),
            (candidate.username, "username"),
        )
        record["Student Name"] = name.value

        # --------------------------------------------------------- email --
        github_email = fields.get("email", "").strip().lower()
        if github_email and is_placeholder_email(github_email):
            github_email = ""
        email = _first(
            (github_email, "GitHub"),
            (self._one(extract_emails(
                section_text(portfolio_sections, "contact")
                or portfolio_text)), "Portfolio"),
            (self._one(extract_emails(
                section_text(resume_sections, "contact")
                or resume_text)), "Resume"),
            (self._one(extract_emails(readme)), "README"),
            (self._one(extract_emails(xlsx_text)), "Instant Data Scraper"),
        )
        record["Email Address"] = email.value
        record["Email Validation Status"] = (
            "Valid Format" if email.value else "Not Found")

        # --------------------------------------------------------- phone --
        phone = _first(
            (self._one_phone(portfolio_text), "Portfolio"),
            (self._one_phone(resume_text), "Resume"),
            (self._one_phone(github_text), "GitHub"),
            (self._one_phone(xlsx_text), "Instant Data Scraper"),
        )
        record["Phone Number"] = phone.value
        record["Phone Validation Status"] = (
            "Valid Format" if phone.value else "Not Found")
        record["WhatsApp Number"] = (
            extract_whatsapp(portfolio_text)
            or extract_whatsapp(resume_text)
            or extract_whatsapp(github_text))

        # ------------------------------------------------------ location --
        location = best_location([
            (flatten(fields.get("location", "")), 1.00),
            (self._contact_location(resume_sections, resume_text), 0.95),
            (self._contact_location(portfolio_sections, portfolio_text),
             0.90),
            (flatten(candidate.xlsx_location), 0.80),
        ])
        record["Current City"] = location.city
        record["State"] = location.state
        record["Country"] = location.country
        # An unresolvable place is worth saying out loud rather than
        # leaving the reader to wonder why Country is empty.
        for note in location.notes:
            if note.startswith("ambiguous place"):
                remarks.append(note[0].upper() + note[1:])
        record["PIN Code"] = (extract_pin(portfolio_text)
                              or extract_pin(resume_text))

        # ----------------------------------------------------- education --
        # An Education section is consulted first, but it never REPLACES
        # the rest of the README: "3rd Data Science honours student" can
        # sit in the opening paragraph of a profile that also has an
        # About section.
        education = extract_education([
            (resume_text, 0.95),
            (bio, 0.92),
            (section_text(readme_sections, "education"), 0.88),
            (readme_raw or readme, 0.85),
            (portfolio_text, 0.90),
            (xlsx_text, 0.80),
        ])
        degree = education["degree"]
        specialization = education["specialization"]
        graduation = education["graduation"]

        record["Current Course"] = degree or specialization
        record["Current Specialization"] = specialization
        record["Current Education Level"] = education["level"]
        record["Current Year / Semester"] = education["year"]
        record["Expected Graduation Year"] = graduation
        record["College Name"] = education["college"]
        record["University / Affiliation University"] = education["university"]

        conflict = self._graduation_conflict(resume_text, readme, bio,
                                             portfolio_text, xlsx_text,
                                             graduation)
        if conflict:
            remarks.append(conflict)

        # "Student at malla reddy engineering college, hyderabad,
        # telangana." names the institution AND where it is; the tail
        # after the institution name is that place.
        college_location = parse_location(
            self._college_location(education["college"], everything))
        if not college_location.city and education["college"]:
            # The person's own location is the best remaining evidence for
            # where they study, but only when a college is actually known.
            college_location = location
        record["College City"] = college_location.city
        record["College State"] = college_location.state

        if degree in UG_DEGREES or (not degree
                                       and education["level"]
                                       == "Undergraduate"):
            record["UG Degree"] = degree
            record["UG Specialization"] = specialization
            record["UG College / Institute Name"] = education["college"]
            record["UG University"] = education["university"]
            record["UG Graduation Year"] = graduation
        elif degree in PG_DEGREES or education["level"] == "Postgraduate":
            record["PG Degree"] = degree
            record["PG Specialization"] = specialization
            record["PG College / Institute Name"] = education["college"]
            record["PG University"] = education["university"]
            record["PG Graduation Year"] = graduation

        # -------------------------------------------------------- skills --
        skills = extract_skills(documents)
        record["Programming Languages"] = as_text(skills["languages"])
        record["Technologies / Frameworks"] = as_text(skills["frameworks"])
        record["Tools / Software Known"] = as_text(skills["tools"])
        record["Primary Skills"] = as_text(skills["primary"])
        record["Secondary Skills"] = as_text(skills["secondary"])

        # ---------------------------------------------------- employment --
        role, company = extract_company_role(bio)
        if not company:
            role, company = extract_company_role(
                section_text(resume_sections, "experience") or resume_text)
        if not company:
            role, company = extract_company_role(prose)
        record["Current Company"] = company
        record["Current Job Role"] = role
        record["Current Employment Status"] = extract_employment_status(
            prose or everything, has_company=bool(company))

        months = extract_experience(prose or everything)
        record["Total Experience"] = format_experience(months)
        record["Internship Experience"] = join_list(
            extract_internships(prose or everything))

        # ---------------------------------------------------------- urls --
        all_links = list(bundle.github_links) + list(bundle.portfolio_links)
        link_text = "\n".join(all_links + [github_text, portfolio_text,
                                           resume_text])
        record["LinkedIn Profile"] = extract_linkedin(link_text)
        record["GitHub Profile"] = candidate.profile_url
        record["Portfolio Website"] = bundle.portfolio_url
        record["Resume URL"] = bundle.resume_url

        # ---------------------------------------------------- aspiration --
        # Each source is searched in turn: `bio or prose` would stop at a
        # non-empty bio and never look at the README that states the goal.
        interest = self._by_priority(extract_career_interest,
                                     [bio, prose, everything])
        record["Career Interest"] = interest
        record["Preferred Job Role"] = self._by_priority(
            extract_preferred_role, [bio, prose, everything])
        record["Preferred Industry"] = extract_preferred_industry(interest)

        # ---------------------------------------------------- provenance --
        record["Data Source"] = " + ".join(bundle.sources_used())
        record["Source URL"] = candidate.profile_url
        record["Source Platform"] = "GitHub"
        record["Data Collection Date"] = datetime.date.today().isoformat()
        record["source_file"] = candidate.source_file

        record["Data Verified Status"] = self._verified_status(
            bundle, record, name)
        record["Remarks / Notes"] = "; ".join(
            dict.fromkeys(r for r in remarks if r))[:900]

        self._fill_aliases(record)
        return record

    # ------------------------------------------------------------ helpers
    @staticmethod
    def _fill_aliases(record):
        """Alias columns always mirror their canonical column (spec 57)."""
        for alias, canonical in C.ALIAS_COLUMNS.items():
            record[alias] = record.get(canonical, "")

    @staticmethod
    def _one(values):
        return values[0] if values else ""

    @staticmethod
    def _by_priority(function, sources):
        """First non-empty result, consulting sources in priority order."""
        for source in sources:
            value = function(source)
            if value:
                return value
        return ""

    @staticmethod
    def _one_phone(text):
        """A number labelled only WhatsApp is not this source's phone.

        It still reaches the WhatsApp column; letting it also claim
        "Phone Number" would hide a real phone stated by a lower-priority
        source.
        """
        found = extract_phones(text)
        for number, is_whatsapp in found:
            if not is_whatsapp:
                return number
        return ""

    @staticmethod
    def _resume_name(text):
        """A resume's name is its first real line, above the contact block."""
        for line in (text or "").splitlines()[:6]:
            stripped = line.strip()
            if not (3 <= len(stripped) <= 60):
                continue
            if "@" in stripped or re.search(r"\d{4}", stripped):
                continue
            words = stripped.split()
            if not 2 <= len(words) <= 5:
                continue
            if all(word[0].isupper() for word in words if word[:1].isalpha()):
                return stripped.title() if stripped.isupper() else stripped
        return ""

    @staticmethod
    def _portfolio_name(text):
        for line in (text or "").splitlines():
            stripped = re.sub(r"^(hi|hello|hey)[,!]?\s*(i\s*'?m|i am)\s*", "",
                              line.strip(), flags=re.I).strip(" .!,")
            if not (3 <= len(stripped) <= 50):
                continue
            words = stripped.split()
            if not 2 <= len(words) <= 4:
                continue
            if all(word[0].isupper() for word in words if word[:1].isalpha()):
                return stripped
        return ""

    @staticmethod
    def _contact_location(parsed_sections, fallback_text):
        """A location stated in a Contact/About section, else a labelled line."""
        scope = section_text(parsed_sections, "contact", "about") \
            or fallback_text
        for line in (scope or "").splitlines():
            if re.search(r"\b(location|based\s+in|address|city|lives?\s+in)"
                         r"\b", line, re.I):
                value = re.sub(
                    r"^.*?\b(location|based\s+in|address|city|lives?\s+in)"
                    r"\b\s*[:\-]?\s*", "", line, flags=re.I)
                value = flatten(value)
                if 2 <= len(value) <= 60:
                    return value
        return ""

    @staticmethod
    def _college_location(college, text):
        """A place named on the same line as the institute."""
        if not college:
            return ""
        for line in (text or "").splitlines():
            if college.lower() not in line.lower():
                continue
            tail = line.lower().split(college.lower(), 1)[-1]
            city, _state, _country = parse_location_tuple(tail)
            if city:
                return tail
        return ""

    @staticmethod
    def _graduation_conflict(resume_text, readme, bio, portfolio_text,
                             xlsx_text, chosen):
        """Record a disagreement instead of hiding it (spec 29)."""
        sources = [
            ("Resume", extract_graduation_year(resume_text)),
            ("GitHub", extract_graduation_year(
                "\n".join([bio, readme]))),
            ("Portfolio", extract_graduation_year(portfolio_text)),
            ("Instant Data Scraper",
             extract_graduation_year(xlsx_text)),
        ]
        present = [(source, year) for source, year in sources if year]
        if len({year for _source, year in present}) <= 1:
            return ""
        detail = "; ".join("%s %s" % (source, year)
                           for source, year in present)
        winner = next((source for source, year in present
                       if year == chosen), present[0][0])
        return ("Graduation year conflict: %s. %s value used."
                % (detail, winner))

    @staticmethod
    def _verified_status(bundle, record, name):
        """Deterministic rule, stated plainly (spec 62).

        The GitHub account is confirmed whenever the live profile loaded.
        One further identity field agreeing across two different sources
        raises it to 'Verified from public profile'.
        """
        if not bundle.github_ok:
            return "Insufficient Data"
        corroborated = 0
        if bundle.portfolio_text.strip():
            corroborated += 1
        if bundle.resume_text.strip():
            corroborated += 1
        identity = sum(1 for key in ("Email Address", "Student Name",
                                     "Current City", "LinkedIn Profile")
                       if record.get(key))
        if corroborated and identity >= 2 and name.source != "username":
            return "Verified from public profile"
        if identity >= 2 or corroborated:
            return "Partially Verified"
        return "Insufficient Data"


# ==========================================================================
# TEMP STORE
# ==========================================================================
# temporary.txt -- the combined raw record for the ONE current person.
#
# Written before extraction, cleared only after that person's row is safely in
# the store (spec 30/71).  If anything fails first the file stays, so the next
# run can see exactly what was collected.
#
# RANK-LOCAL PATHS
# ----------------
# These files used to be fixed at config.TEMP_STUDENT_FILE and friends.  Under
# they live next to this file, because one folder is one rank and "the ONE
# person currently being processed" is one person per folder.  Two copies have
# two folders, so they cannot overwrite each other's raw dump and a failure
# always preserves evidence about the right person.
#
# configure() still exists so a test can point the module at a temporary
# directory; production never calls it.
# --------------------------------------------------------------------------
RULE = "=" * 50

_PATHS = {
    "temp_student_file": C.TEMP_STUDENT_FILE,
    "temp_student_json": C.TEMP_STUDENT_JSON,
    "temp_download_dir": C.TEMP_DOWNLOAD_DIR,
}


def configure(temp_student_file=None, temp_student_json=None,
              temp_download_dir=None):
    """Point this module somewhere else.  For tests only."""
    if temp_student_file:
        _PATHS["temp_student_file"] = temp_student_file
    if temp_student_json:
        _PATHS["temp_student_json"] = temp_student_json
    if temp_download_dir:
        _PATHS["temp_download_dir"] = temp_download_dir
    return _PATHS


def _path(name):
    return _PATHS[name]


def student_file():
    return _path("temp_student_file")


def student_json():
    return _path("temp_student_json")


def download_dir():
    return _path("temp_download_dir")


def _section(title, pairs):
    lines = [RULE, "SOURCE: %s" % title, RULE, ""]
    for label, value in pairs:
        lines.append("%s:" % label)
        lines.append(str(value if value not in (None, "") else "-"))
        lines.append("")
    return "\n".join(lines)


def build(bundle):
    """Render the full raw dump for one candidate."""
    candidate = bundle.candidate
    fields = bundle.github_fields or {}

    text = _section("INSTANT DATA SCRAPER", [
        ("Keyword", candidate.keyword),
        ("Source File", candidate.source_file),
        ("Source Row", candidate.source_row_number),
        ("Raw XLSX Cells", "\n".join(candidate.raw_cells)),
        ("Detected Username", candidate.username),
        ("Detected GitHub URL", candidate.profile_url),
        ("Name", candidate.xlsx_name),
        ("Bio", candidate.xlsx_bio),
        ("Location", candidate.xlsx_location),
    ])
    text += "\n" + _section("GITHUB PROFILE", [
        ("Structured Name", fields.get("fullname", "")),
        ("Structured Username", fields.get("nickname", "")),
        ("Structured Location", fields.get("location", "")),
        ("Structured Bio", fields.get("bio", "")),
        ("Public Email", fields.get("email", "")),
        ("Public Website", fields.get("website", "")),
        ("Public Links", "\n".join(bundle.github_links)),
        ("PROFILE README (RENDERED)", fields.get("readme", "")),
        ("FULL VISIBLE PROFILE TEXT", bundle.github_text),
    ])
    text += "\n" + _section("GITHUB README RAW", [
        ("Repository", "%s/%s" % (candidate.username, candidate.username)),
        ("RAW MARKDOWN", fields.get("readme_raw", "")),
    ])
    text += "\n" + _section("PORTFOLIO", [
        ("Portfolio URL", bundle.portfolio_url),
        ("Pages Visited", "\n".join(bundle.portfolio_pages)),
        ("FULL VISIBLE PORTFOLIO TEXT", bundle.portfolio_text),
        ("Relevant links", "\n".join(bundle.portfolio_links[:60])),
    ])
    text += "\n" + _section("RESUME", [
        ("Resume URL", bundle.resume_url),
        ("Resume File", bundle.resume_file),
        ("EXTRACTED RESUME TEXT", bundle.resume_text),
    ])
    return text


def write(bundle):
    text = build(bundle)
    _atomic_write(student_file(), text)
    payload = {
        "username": bundle.candidate.username,
        "profile_url": bundle.candidate.profile_url,
        "keyword": bundle.candidate.keyword,
        "source_file": bundle.candidate.source_file,
        "source_row": bundle.candidate.source_row_number,
        "sources_used": bundle.sources_used(),
        "portfolio_url": bundle.portfolio_url,
        "resume_url": bundle.resume_url,
        "resume_file": bundle.resume_file,
        "remarks": bundle.remarks,
    }
    _atomic_write(student_json(),
                  json.dumps(payload, indent=2, ensure_ascii=False))
    return text


def clear():
    """Only ever called after the student's row is safely in the store."""
    for path in (student_file(), student_json()):
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError as exc:
            log.warning("could not clear %s: %s", path, exc)


def stale_student_files():
    """Per-student temp data that should not exist between students."""
    stale = []
    for path in (student_file(), student_json()):
        if os.path.exists(path):
            stale.append(path)
    directory = download_dir()
    if os.path.isdir(directory):
        for name in sorted(os.listdir(directory)):
            stale.append(os.path.join(directory, name))
    return stale


def sweep_completed_keyword():
    """Remove leftovers ONLY after a keyword finished successfully.

    Every row is resolved at this point, so nothing here can still be
    needed.  This is never reached from an error path and is deliberately
    not in a finally: block -- an unresolved student's temp data must
    survive for the next run.
    """
    removed = []
    for path in stale_student_files():
        try:
            os.remove(path)
            removed.append(os.path.basename(path))
        except OSError as exc:
            log.warning("could not remove %s: %s", path, exc)
    return removed


def _atomic_write(path, text):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        handle.write(text)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(tmp, path)


# ==========================================================================
# TRANSIENT FAILURES
# ==========================================================================
# Structured record of profiles that failed for a TEMPORARY reason.
#
# A timeout, a 5xx or a dropped connection is not the person's fault and must
# never become a permanent blacklist (spec 74).  Each one is written here with
# enough detail to see what is blocking progress, and the entry is removed the
# moment that profile succeeds.
#
# Separate from skipped_accounts.txt, which records PERMANENT skips only
# (confirmed 404, deleted account, organisation).
# --------------------------------------------------------------------------
def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


class TransientFailureLog:
    """Atomic, crash-safe JSON keyed by lower-cased username."""

    def __init__(self, path=None):
        # One folder is one rank, so this folder's retry budget is its own.
        self.path = path or C.TRANSIENT_FILE
        self.data = self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return {"version": 1, "failures": {}}
        try:
            with open(self.path, encoding="utf-8") as handle:
                loaded = json.load(handle)
        except Exception as exc:
            log.warning("transient failure log unreadable (%s); starting "
                        "fresh", exc)
            return {"version": 1, "failures": {}}
        loaded.setdefault("version", 1)
        loaded.setdefault("failures", {})
        return loaded

    def save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, self.path)

    # ------------------------------------------------------------------
    def attempts(self, username):
        entry = self.data["failures"].get((username or "").lower())
        return int(entry.get("attempts", 0)) if entry else 0

    def record(self, candidate, reason):
        """Add or update the entry for one candidate.  Returns it."""
        key = (candidate.username or "").lower()
        entry = self.data["failures"].get(key, {})
        entry.update({
            "username": candidate.username,
            "github_url": candidate.profile_url,
            "keyword": candidate.keyword,
            "source_xlsx": candidate.source_file,
            "source_row": candidate.source_row_number,
            "reason": str(reason)[:300],
            "attempts": int(entry.get("attempts", 0)) + 1,
            "first_seen": entry.get("first_seen") or _now(),
            "last_attempt": _now(),
        })
        self.data["failures"][key] = entry
        self.save()
        log.warning("transient failure %s (attempt %d): %s",
                    candidate.username, entry["attempts"], reason)
        return entry

    def clear(self, username):
        """Drop the entry once that profile finally works."""
        key = (username or "").lower()
        if key in self.data["failures"]:
            del self.data["failures"][key]
            self.save()

    def blocking(self):
        """Entries that have used up their retries."""
        return [entry for entry in self.data["failures"].values()
                if int(entry.get("attempts", 0)) >= C.PROFILE_ATTEMPTS_MAX]

    def all_entries(self):
        return sorted(self.data["failures"].values(),
                      key=lambda entry: entry.get("last_attempt", ""))


def describe_transient_failure(entry):
    return "\n".join([
        "TRANSIENT PROFILE FAILURE",
        "Username: %s" % entry.get("username", ""),
        "GitHub URL: %s" % entry.get("github_url", ""),
        "Keyword: %s" % entry.get("keyword", ""),
        "Source XLSX: %s" % entry.get("source_xlsx", ""),
        "Source row: %s" % entry.get("source_row", ""),
        "Attempts: %s" % entry.get("attempts", ""),
        "Reason: %s" % entry.get("reason", ""),
        "Last attempted: %s" % entry.get("last_attempt", ""),
    ])




class SearchUnit:
    """One search to run: one city spelling crossed with one keyword.

    search_city is the spelling actually typed into GitHub, which may be an
    alias ("Bangalore"); canonical_city is the city's one true name
    ("Bengaluru") and is what decides which rank owns it, so every spelling of
    a city lands on the same rank.

    NONE of these fields ever reaches a person's Current City / State /
    Country.  They record where we LOOKED, not where anybody lives.
    """

    __slots__ = ("country", "state", "canonical_city", "search_city",
                 "base_keyword", "search_query", "slug")

    def __init__(self, country, state, canonical_city, search_city,
                 base_keyword, search_query, slug):
        self.country = country
        self.state = state
        self.canonical_city = canonical_city
        self.search_city = search_city
        self.base_keyword = base_keyword
        self.search_query = search_query
        self.slug = slug

    @property
    def identity(self):
        return (self.canonical_city, self.search_city, self.base_keyword)

    def as_state(self):
        """The search_unit block written into scraper_state.json."""
        return {
            "country": self.country,
            "state": self.state,
            "canonical_city": self.canonical_city,
            "search_city": self.search_city,
            "base_keyword": self.base_keyword,
            "search_query": self.search_query,
            "slug": self.slug,
        }

    def __repr__(self):
        return "SearchUnit(%r, city=%r as %r)" % (
            self.search_query, self.canonical_city, self.search_city)


def make_unit(city, spelling, keyword):
    """Build one SearchUnit.

    A multi-word city must be quoted or GitHub reads only its first word as
    the location, so "New Delhi" is sent as location:"New Delhi".
    """
    city_term = '"%s"' % spelling if " " in spelling else spelling
    return SearchUnit(
        country=city.country,
        state=city.state,
        canonical_city=city.canonical,
        search_city=spelling,
        base_keyword=keyword,
        search_query=C.SEARCH_QUERY_TEMPLATE.format(keyword=keyword,
                                                    city=city_term),
        slug=slugify("%s %s" % (keyword, spelling)))


# ==========================================================================
# KEYWORD SHUFFLE
# ==========================================================================
# The keyword order for one lap, seeded from (rank, lap) so a restart does
# not reshuffle and two ranks never run a lap in the same order.
KEYWORD_CYCLE_START = 1


def keyword_cycle_seed(cycle_number, this_rank=None):
    """The seed for one (rank, lap) pair -- sha1, because hash() is salted
    per process and would reshuffle after a restart."""
    this_rank = rank if this_rank is None else this_rank
    source = "%d|%d" % (int(this_rank), int(cycle_number))
    return int(hashlib.sha1(source.encode("utf-8")).hexdigest()[:12], 16)


def shuffle_keywords(keywords, cycle_number, this_rank=None):
    """`keywords` in this lap's order -- a permutation, checked rather than
    assumed, because a dropped keyword would silently lose an audience."""
    ordered = list(keywords)
    random.Random(keyword_cycle_seed(cycle_number, this_rank)).shuffle(ordered)
    if len(ordered) != len(keywords) or set(ordered) != set(keywords):
        raise RuntimeError(
            "the shuffle corrupted the keyword plan for cycle %s "
            "(%d in, %d out)" % (cycle_number, len(keywords), len(ordered)))
    return ordered


def keyword_order(cycle_number=None, keywords=None):
    """The keyword list in this lap's order; the file order when the shuffle
    is switched off."""
    resolved = load_base_keywords() if keywords is None else list(keywords)
    if not KEYWORD_SHUFFLE_ENABLED:
        return resolved
    if cycle_number is None:
        cycle_number = read_keyword_cycle()
    return shuffle_keywords(resolved, cycle_number)


def read_keyword_cycle():
    """Which lap this copy is on; lap 1 when the file is missing or
    unreadable, because a bad lap number costs an order, never data."""
    if not os.path.exists(C.KEYWORD_CYCLE_FILE):
        # Deliberately does NOT create the file.  banner() and the report
        # modes call this, and a read that writes is how a read-only path
        # ends up leaving state behind.  _run() writes it at startup.
        return KEYWORD_CYCLE_START
    try:
        with open(C.KEYWORD_CYCLE_FILE, encoding="utf-8") as handle:
            data = json.load(handle)
        number = int(data.get("keyword_cycle", KEYWORD_CYCLE_START))
        if number < KEYWORD_CYCLE_START:
            return KEYWORD_CYCLE_START
        return number
    except Exception as exc:
        log.warning("%s unreadable (%s); starting at cycle %d",
                    os.path.basename(C.KEYWORD_CYCLE_FILE), exc,
                    KEYWORD_CYCLE_START)
        return KEYWORD_CYCLE_START


def save_keyword_cycle(number, keywords=None):
    """Record the lap, atomically.  The keywords written beside it are a
    preview for reading by eye; the order is re-derived from the number."""
    number = int(number)
    try:
        resolved = load_base_keywords() if keywords is None else list(keywords)
    except Exception:
        resolved = []
    order = (shuffle_keywords(resolved, number)
             if resolved and KEYWORD_SHUFFLE_ENABLED else resolved)
    payload = {
        "keyword_cycle": number,
        "rank": rank,
        "num_sys": num_sys,
        "shuffled": bool(KEYWORD_SHUFFLE_ENABLED),
        "shuffle_seed": keyword_cycle_seed(number),
        "keywords": len(order),
        "keywords_fingerprint": (keywords_fingerprint(resolved)
                                 if resolved else ""),
        "first_keywords": order[:10],
        "opened_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    tmp = C.KEYWORD_CYCLE_FILE + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=4)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, C.KEYWORD_CYCLE_FILE)
    except Exception as exc:
        # The cycle number is a pacing detail, not data.  A folder that
        # cannot write it keeps collecting; it just re-runs this lap's order
        # after a restart instead of moving to the next one.
        log.warning("could not write %s (%s)",
                    os.path.basename(C.KEYWORD_CYCLE_FILE), exc)
    return number


def open_next_keyword_cycle(keywords=None):
    """Move to the next lap, persisted before its first search so a crash
    resumes the new lap rather than re-running the old order."""
    number = read_keyword_cycle() + 1
    save_keyword_cycle(number, keywords)
    return number


def describe_keyword_cycle(cycle_number, order):
    """The startup/lap line: which lap, and the first few keywords of it."""
    if not KEYWORD_SHUFFLE_ENABLED:
        return "Keyword order       : new_mylib.py file order (shuffle off)"
    preview = ", ".join(order[:4])
    return ("Keyword order       : shuffled for cycle %d (%s%s)"
            % (cycle_number, preview, " ..." if len(order) > 4 else ""))


# get_cities_by_rank() used to be defined here.  It now sits at the TOP of
# this file, directly under `rank` and `num_sys`, because those three things
# are one idea and splitting them across 11,900 lines made the work division
# hard to find.
def build_search_list(this_rank=None, systems=None, say_total=True,
                      keywords=None):
    """Every search this rank must run, in a fixed order.

        this rank's cities
            -> each spelling of the city, canonical spelling first
                -> each keyword, in `keywords` order

    The order never changes for a given (cities, keywords, rank, num_sys),
    which is what lets current_row.json be a plain index into this list.

    `keywords` is the keyword order for the current lap -- see the KEYWORD
    SHUFFLE block above.  Left out, it is new_mylib.py's file order, which
    is what every caller got before the shuffle existed, so --dry-run and
    the report modes are unaffected.

    Returns a list of (city, spelling, keyword) tuples rather than finished
    SearchUnits: this rank owns hundreds of thousands of searches, and the
    tuples share their strings whereas the objects would not.  Call
    make_unit() on one tuple when it is that search's turn.
    """
    this_rank = rank if this_rank is None else this_rank
    systems = num_sys if systems is None else systems

    all_cities = list(cities())
    mine = get_cities_by_rank(all_cities, this_rank, systems)
    keywords = load_base_keywords() if keywords is None else list(keywords)

    searches = []
    for city in mine:
        for spelling in city.search_aliases:
            for keyword in keywords:
                searches.append((city, spelling, keyword))

    if say_total:
        print("Rank %d of %d: %d cities x %d keywords = %s searches"
              % (this_rank, systems, len(mine), len(keywords),
                 format(len(searches), ",")))
    return searches


# ----------------------------------------------------- current_row.json ----
# Where this rank is in its own search list.  Written AFTER a search
# finishes, so a crash resumes on the NEXT search instead of repeating the
# one that was already done.
START = 0


def read_current_row(total_rows):
    """The index of the next search to run.

    Missing file -> create it at START and begin there.  An index outside the
    list -- the list shrank because a city or keyword was removed -- is reset
    to START rather than used, because an index past the end would otherwise
    skip the whole run.
    """
    if not os.path.exists(C.CURRENT_ROW_FILE):
        save_current_row(START)
        return START
    try:
        with open(C.CURRENT_ROW_FILE, encoding="utf-8") as handle:
            data = json.load(handle)
        current_row = int(data.get("current_row", START))
        if current_row < START or current_row >= total_rows:
            return START
        return current_row
    except Exception:
        save_current_row(START)
        return START


def save_current_row(row_number):
    """Record which search to run next.  Atomic: a crash leaves one or other.

    os.replace() rather than a plain write, because a torn current_row.json
    would be unreadable and send the whole rank back to search 0.
    """
    tmp = C.CURRENT_ROW_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump({"current_row": int(row_number)}, handle, indent=4)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(tmp, C.CURRENT_ROW_FILE)


class Stopped(RuntimeError):
    """Ctrl+C, or a condition that must halt the run without cleanup."""


class GlobalThrottle(Stopped):
    """Several different profiles hit a GitHub rate-limit/challenge in a short
    window: an account/session-wide block.  Halts the rank (and the device)
    without deferring those profiles -- we respect GitHub's restriction rather
    than keep probing.  A Stopped, so all the safe-stop machinery applies."""


class RateLimited(RuntimeError):
    """A GitHub rate-limit/challenge on the SEARCH page or during PAGINATION.

    Not a Stopped: the run loop cools down and retries the SAME SearchUnit, so
    the unit stays unfinished and the next unit/rank never starts.  Carries the
    source and any server Retry-After (seconds)."""

    def __init__(self, message, source="search", retry_after=None,
                 kind="rate-limit"):
        super().__init__(message)
        self.source = source
        self.retry_after = retry_after
        # "rate-limit" escalates and eventually stops the device; "outage"
        # (a GitHub 5xx) waits at the top of the ladder and retries forever.
        self.kind = kind


class TransientProfileFailure(RuntimeError):
    """A profile failed for a temporary reason; retry, never blacklist."""


class CollectResult:
    """What collect() decided.  THREE outcomes, not two.

        RESULTS       rows > 0 and the XLSX is on disk -- process it
        ZERO_RESULTS  a valid Users search with no users.  The search unit is
                      COMPLETE: no XLSX exists, no candidate exists, nothing
                      is unresolved.  Advance to the next search.
        FAILED        preserve everything and stop

    ZERO_RESULTS used to be indistinguishable from FAILED, so a legitimately
    empty query blocked its rank forever: every run retried the same search
    and failed the same way.  It is now a success with zero rows.
    """

    RESULTS = "results"
    ZERO_RESULTS = "zero-results"
    FAILED = "failed"


def setup_logging(verbose=False):
    """Log to this folder's own logs/student_scraper.log.

    One folder is one rank, so the log is already per-rank: a shared log from
    many copies would be unreadable and useless for troubleshooting any one
    of them.

    The file ROLLS AT MIDNIGHT and C.LOG_RETENTION_DAYS dated files are kept
    behind the live one; the oldest is deleted at each roll.  This used to be
    a plain FileHandler, which on a loop that never ends meant one file that
    grew for the life of the process.  Rotation happens in-process, so a run
    that spans weeks needs nothing scheduled outside it.
    """
    os.makedirs(C.LOG_DIR, exist_ok=True)
    log_file = C.LOG_FILE
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    log.handlers.clear()
    file_handler = logging.handlers.TimedRotatingFileHandler(
        log_file, when="midnight", backupCount=C.LOG_RETENTION_DAYS,
        encoding="utf-8")
    # Dated suffixes (student_scraper.log.2026-10-07) rather than the default
    # .1/.2/.3, so which day a file holds is readable without opening it.
    file_handler.suffix = "%Y-%m-%d"
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S"))
    log.addHandler(file_handler)
    if verbose:
        log.addHandler(logging.StreamHandler(sys.stderr))


def say(message=""):
    print(message, flush=True)
    log.info(message.strip() or "-")


def _request_stop(_signum, _frame):
    global STOP_REQUESTED
    if STOP_REQUESTED:
        return
    STOP_REQUESTED = True
    print("\nStopping after the current safe point...", flush=True)


def check_stop():
    if STOP_REQUESTED:
        raise Stopped("interrupted")




# guard_output_locks() used to run at scraper startup, because every student
# was appended straight into student_data.xlsx and an open workbook meant a
# lost save.  Ranks do not write the workbook any more, so a colleague having
# student_data.xlsx open in LibreOffice no longer blocks collection.  The check
# moved to where it is actually needed -- guard_locks(), called by
# the export -- rather than being left here doing nothing.


# ==========================================================================
# SCHEDULE
# ==========================================================================
# When this copy stops: the night window, and the break after a full work
# period.  Both are taken between searches only, and the work clock is
# stopped around every long wait so a period means hours of real collecting.
_WORK_STARTED = None        # has the clock ever been started?
_WORK_BANKED = 0.0          # seconds of work banked before the last pause
_WORK_MARK = None           # monotonic at the last resume, or None if paused
_WORK_PAUSE_DEPTH = 0       # so nested pauses stop the clock exactly once
_WORK_PERIOD_HOURS = 0.0    # this period's length, re-drawn at every reset


def pick_work_period_hours():
    """A fresh period length in hours, to one decimal place; 0.0 when the
    work/break cycle is switched off."""
    low = max(0.0, float(WORK_PERIOD_MIN_HOURS))
    high = max(low, float(WORK_PERIOD_MAX_HOURS))
    if high <= 0:
        return 0.0
    if high == low:
        return high
    return random.randint(int(low * 10), int(high * 10)) / 10.0


def pick_break_seconds():
    """A fresh break length, in seconds, from the configured range."""
    low = max(0, int(BREAK_MIN_MINUTES)) * 60
    high = max(low, int(BREAK_MAX_MINUTES) * 60)
    if high <= low:
        return low
    return random.randint(low, high)


def reset_work_timer():
    """Begin a fresh work period.  Called at startup and after every break."""
    global _WORK_STARTED, _WORK_BANKED, _WORK_MARK, _WORK_PAUSE_DEPTH
    global _WORK_PERIOD_HOURS
    _WORK_STARTED = time.monotonic()
    _WORK_BANKED = 0.0
    _WORK_MARK = time.monotonic()
    _WORK_PAUSE_DEPTH = 0
    _WORK_PERIOD_HOURS = pick_work_period_hours()
    return _WORK_PERIOD_HOURS


def work_clock_pause():
    """Stop counting working time.  Safe to nest."""
    global _WORK_BANKED, _WORK_MARK, _WORK_PAUSE_DEPTH
    _WORK_PAUSE_DEPTH += 1
    if _WORK_PAUSE_DEPTH == 1 and _WORK_MARK is not None:
        _WORK_BANKED += time.monotonic() - _WORK_MARK
        _WORK_MARK = None


def work_clock_resume():
    """Start counting working time again.  Safe to nest."""
    global _WORK_MARK, _WORK_PAUSE_DEPTH
    if _WORK_PAUSE_DEPTH > 0:
        _WORK_PAUSE_DEPTH -= 1
    if _WORK_PAUSE_DEPTH == 0 and _WORK_MARK is None:
        _WORK_MARK = time.monotonic()


def work_seconds():
    """Seconds actually spent working since the last break."""
    if _WORK_STARTED is None:
        return 0.0
    total = _WORK_BANKED
    if _WORK_MARK is not None:
        total += time.monotonic() - _WORK_MARK
    return total


def work_period_hours():
    """This period's length.  0.0 when the work/break cycle is off."""
    return _WORK_PERIOD_HOURS


def night_seconds_remaining(now=None):
    """Seconds until the night window ends, or 0 when outside it -- the only
    place the window is decided.  It may cross midnight; START == END is off."""
    if not NIGHT_PAUSE_ENABLED:
        return 0
    now = now or datetime.datetime.now()
    start = int(NIGHT_START_HOUR) % 24
    end = int(NIGHT_END_HOUR) % 24
    if start == end:
        return 0
    if start < end:
        inside = start <= now.hour < end
    else:
        inside = now.hour >= start or now.hour < end
    if not inside:
        return 0
    target = now.replace(hour=end, minute=0, second=0, microsecond=0)
    if target <= now:
        target = target + datetime.timedelta(days=1)
    return max(0, int((target - now).total_seconds()))


def due_pause(now=None):
    """(reason, seconds) when a scheduled stop is due, else None.  Night is
    checked first, so a period ending at 23:58 does not take both."""
    night = night_seconds_remaining(now)
    if night > 0:
        return ("night stop %02d:00-%02d:00"
                % (int(NIGHT_START_HOUR) % 24, int(NIGHT_END_HOUR) % 24),
                night)

    period = work_period_hours()
    if _WORK_STARTED is None or period <= 0:
        return None
    worked = work_seconds()
    if worked < period * 3600.0:
        return None
    # Drawn fresh here rather than once per run, so repeated breaks are not
    # identical lengths.
    seconds = pick_break_seconds()
    if seconds <= 0:
        return None
    return ("%.1f-hour work period reached (%.2f h actually worked) -- "
            "taking a %s break (random %d-%d min)"
            % (period, worked / 3600.0, fmt_hms(seconds),
               int(BREAK_MIN_MINUTES), int(BREAK_MAX_MINUTES)), seconds)


def night_pause_status():
    """One line about the night stop, for the startup banner."""
    if not NIGHT_PAUSE_ENABLED:
        return "disabled (24-hour collecting permitted)"
    start = int(NIGHT_START_HOUR) % 24
    end = int(NIGHT_END_HOUR) % 24
    if start == end:
        return "disabled (start hour == end hour)"
    return "%02d:00 - %02d:00 local" % (start, end)


def work_break_status():
    """One line for the banner: the drawn period once the clock is running,
    the configured range before that."""
    low = max(0.0, float(WORK_PERIOD_MIN_HOURS))
    high = max(low, float(WORK_PERIOD_MAX_HOURS))
    if high <= 0:
        return "disabled (no scheduled breaks)"
    breaks = "%d-%d min off" % (int(BREAK_MIN_MINUTES),
                                int(BREAK_MAX_MINUTES))
    period = work_period_hours()
    if period <= 0:                      # the clock has not been started yet
        return ("%.1f-%.1f h of work, then %s (both drawn fresh each time)"
                % (low, high, breaks))
    return ("%.1f h of work this period, then %s (both drawn fresh each time)"
            % (period, breaks))


# ============================================================== workflow ===
class Workflow:
    # Class-level defaults so a subclass that builds its own state (the
    # lifecycle test stubs do) still has them without needing to track
    # every attribute this class gains.
    keyword_number = 0
    keyword_total = 0
    # Set by __init__ to a real BrowserManager.  Declared here for the same
    # reason as the rest: process_rows() credits the browser's restart budget
    # after every student, and the stubs built with __new__ set only the
    # attributes their own test needs -- sixteen test files do this, so the
    # default belongs on the class, not in sixteen fixtures.
    browser = None
    # A Workflow built with __new__ (every test stub does this) must still
    # have a usable master index.  _existing stays None on the class and the
    # property creates a dict on first use, which binds it to the INSTANCE --
    # so stubs never share one dict, the way a plain `existing = {}` class
    # attribute would make them.
    _existing = None
    _master_dirty = False

    @property
    def existing(self):
        if self._existing is None:
            self._existing = {}
        return self._existing

    @existing.setter
    def existing(self, value):
        self._existing = value

    def __init__(self, options, searches, existing=None):
        self.options = options
        # Every search this copy must run, as (city, spelling, keyword)
        # tuples.  current_row.json is an index into exactly this list.
        self.searches = searches
        # naukri's `existing_data`: the master spreadsheet, in memory.
        self.existing = load_existing_master() if existing is None else existing
        self._master_dirty = False
        self.state = StateManager()
        self.browser = BrowserManager(headless=not options.show)
        self.scraper = None
        self.profiles = None
        self.portfolios = PortfolioParser(say)
        self.resumes = ResumeParser(say, download_dir=C.TEMP_DOWNLOAD_DIR)
        self.merger = StudentMerger(say)
        self.reader = XlsxReader()
        self.transient = TransientFailureLog()
        self.deferred = DeferredQueue()
        self.throttle = ThrottleCircuit()
        self.cooldown = RateLimitCooldown()
        # Usernames deferred during THIS process so that, if a global throttle
        # then trips, we can undo those deferrals (they were global victims,
        # not profile-specific).  Reset per search.
        self._deferred_this_run = set()
        self._throttle_min_pos = None
        self.processed = 0
        self.keyword_number = 0
        self.keyword_total = 0
        self.unit = None
        # How many times the whole search list has been walked end to end.
        self.laps = 0
        # The lap the keyword order belongs to; set for real in _run().
        self.keyword_cycle = KEYWORD_CYCLE_START

    # ------------------------------------------------------------- start --
    async def run(self):
        """Run this copy, forever.

        A copy owns hundreds of thousands of searches, so a run is far more
        likely to end by Ctrl+C, a rate-limit cooldown or a low battery than
        by anything else.  Every one of those paths still leaves
        current_data.xlsx holding what was collected, which is what the
        finally below is for.
        """
        try:
            return await self._run()
        finally:
            say("[CYCLE] %d row(s) waiting in %s"
                % (row_count(),
                   os.path.basename(C.DETAIL_OUTPUT_FILE)))

    async def _run(self):
        searches = self.searches
        if not searches:
            say("Rank %d of %d has no cities to search." % (rank, num_sys))
            say("num_sys (%d) is large enough that this rank's slice of the "
                "city list is empty." % num_sys)
            say("Lower num_sys, or give this folder a lower rank.")
            say("Run --dry-run to see what every copy would get.")
            return 1

        # Drawn before the banner, so it reports the numbers the run uses.
        reset_work_timer()
        self.keyword_cycle = read_keyword_cycle()
        banner(searches, self.existing)
        save_keyword_cycle(self.keyword_cycle)
        blocking = self.transient.blocking()
        if blocking:
            say("Transient  : %d profile(s) previously exhausted retries"
                % len(blocking))
            for entry in blocking[:3]:
                say("             %s (%s)" % (entry["username"],
                                              entry["reason"][:60]))
            say("=" * 60)

        await self.browser.start(with_extension=True)
        self.scraper = InstantDataScraper(self.browser, self.state, say)
        self.profiles = GitHubProfile(self.browser, say)

        page = await self.browser.page()
        # The extension opens its own options page a second or two after
        # it loads and leaves it open for the whole session, costing a
        # renderer for nothing.  Sweeping immediately would run too early.
        await asyncio.sleep(2)
        surplus = await self.browser.close_extra_pages(keep=(page,))
        if surplus:
            say("Memory     : closed %d surplus browser page(s)" % surplus)
        login = await self.browser.github_login_name(page)
        if login:
            say("Logged in as: %s" % login)
        else:
            say("GitHub login not detected.")
            say("Run: python3 student_scraper.py --login")

        # A cooldown persisted by a previous run must be respected before we
        # touch GitHub again -- restarting the command cannot bypass it.
        if not await self._respect_active_cooldown():
            return 4

        total = len(searches)
        index = read_current_row(total)
        if index:
            say("")
            say("[RESUME] continuing at search %s of %s (current_row.json)"
                % (format(index + 1, ","), format(total, ",")))

        # ------------------------------------------------------------------
        # THE LOOP.  It does not end.
        #
        #     index = read_current_row()
        #     ... run searches[index] ...
        #     save_current_row(index + 1)
        #
        # current_row.json is written AFTER the search finishes, so a crash
        # resumes on the NEXT search rather than repeating the one that was
        # already done.  Reaching the last search wraps back to 0 and the
        # whole list runs again -- a city crawled months ago has new people
        # in it by the time the loop returns to it.
        # ------------------------------------------------------------------
        while True:
            check_stop()
            # Battery/network gate at a SAFE POINT: between two searches,
            # with nothing half-collected.  Pauses for a charger or a dropped
            # link; stops cleanly (state preserved) when the battery is
            # critical or the network stays down.
            if not await self._health_gate():
                return 4
            # The night window, or the break after a full work period.
            page = await self._schedule_gate(page)

            city, spelling, keyword = searches[index]
            unit = make_unit(city, spelling, keyword)
            self.unit = unit
            say("")
            say("=" * 60)
            say("[Search %s/%s]  [Rank %d/%d]%s"
                % (format(index + 1, ","), format(total, ","), rank, num_sys,
                   "" if not self.laps else "  [lap %d]" % (self.laps + 1)))
            say("State        : %s" % unit.state)
            say("City         : %s%s"
                % (unit.canonical_city,
                   "" if unit.search_city == unit.canonical_city
                   else ' (searching as "%s")' % unit.search_city))
            say("Base keyword : %s" % unit.base_keyword)
            say("Search query : %s" % unit.search_query)
            say("=" * 60)
            # Compact rank card at this (phase) transition -- not per log line.
            say(rank_card(
                rank, num_sys, index + 1, total, unit.search_query,
                len(self.existing), self.deferred.count()))
            self.keyword_number = index + 1
            self.keyword_total = total
            self.state.begin_unit(index, unit)

            # The page from the previous keyword may have died with a
            # browser that was restarted mid-run, so it is never reused.
            alive = (self.browser is not None
                     and await self.browser.is_alive())
            stale = page is None or page.is_closed()
            say("[NEXT KEYWORD BROWSER] context alive: %s"
                % ("yes" if alive else "no"))
            say("[NEXT KEYWORD BROWSER] page closed: %s"
                % ("yes" if stale else "no"))
            if not alive or stale:
                say("[BROWSER] refreshing page reference")
            page = await self.live_page(page)

            try:
                await self.do_search_unit(page, unit)
            except RateLimited as limit:
                # A GitHub rate-limit/challenge on the search page or during
                # pagination.  Cool down and retry the SAME search; never
                # advance.  current_row.json is NOT written, so a crash during
                # the cooldown also resumes on this same search.
                if not await self._cooldown_and_wait(limit):
                    raise Stopped("rate-limit cooldown exhausted")
                page = await self.live_page(page)
                continue
            except GlobalThrottle:
                # A profile-level account-wide block already stops the device;
                # also record a cooldown so a restart respects it.
                self.cooldown.escalate("profile", "global throttle")
                raise
            except Stopped:
                raise
            except BrowserLost as exc:
                say("[STOP] the browser process died: %s" % exc)
                raise Stopped(str(exc))

            # The search completed: any rate-limit escalation is over.
            if self.cooldown.clear():
                say("[RATE LIMIT] Cleared; normal pacing restored.")

            # ---- the search is DONE; only now does the position move ----
            # naukri writes its master after each page.  Here it is after each
            # completed SEARCH: the master is the whole collection, so writing
            # it per person would rewrite every row for every row.
            self.save_master_now("search complete")
            index += 1
            self.state.advance_unit()

            if index >= total:
                # End of the list.  Give every parked profile one more go,
                # then wrap around and start the whole list again.
                self.laps += 1
                say("")
                say("[LAP %d] finished all %s search(es) for rank %d."
                    % (self.laps, format(total, ","), rank))
                say("        %d person/people in %s"
                    % (len(self.existing), os.path.basename(master_path())))
                if self.deferred.count():
                    page = await self._retry_deferred(page)
                index = START
                # The longest pause belongs here, where nothing is in
                # flight -- but not when --limit (checked below) ends the run
                # anyway, or a `--limit 5` run would wait 45 minutes to exit.
                spent = (getattr(self.options, "limit", 0)
                         and self.processed >= self.options.limit)
                if int(LAP_COOLDOWN_MINUTES) > 0 and not spent:
                    page = await self._long_pause(
                        page, int(LAP_COOLDOWN_MINUTES) * 60,
                        "cooldown after lap %d (keywords re-shuffle for "
                        "lap %d)" % (self.laps, self.laps + 1))
                # A new lap re-draws the keyword order; same searches, so
                # `total` is unchanged and index 0 is still index 0.
                searches = self.next_keyword_cycle()
                self.searches = searches
                total = len(searches)
                say("[LAP %d] starting again from search 1." % (self.laps + 1))

            save_current_row(index)

            if self.options.limit and self.processed >= self.options.limit:
                say("")
                say("[LIMIT] processed %d record(s); stopping as requested."
                    % self.processed)
                return 0

            city, spelling, keyword = searches[index]
            say("")
            say("[NEXT] %s" % make_unit(city, spelling, keyword).search_query)
            # Conservative delay between two DIFFERENT completed searches.
            await self._search_delay()

    def save_master_now(self, why=""):
        """Write the master back, as naukri does between searches.

        Best-effort: current_data.xlsx already holds the people, so a failed
        master write costs the duplicate index, not the data, and must never
        stop collection.
        """
        if not self._master_dirty:
            return True
        try:
            rows = save_master(self.existing)
            self._master_dirty = False
            say("[MASTER] %d person/people written%s"
                % (rows, " (%s)" % why if why else ""))
            return True
        except Exception as exc:
            say("[MASTER] write failed (%s); the people are still in %s"
                % (exc, os.path.basename(C.DETAIL_OUTPUT_FILE)))
            return False

    def write_to_cycle_file(self, record):
        """Append one person to current_data.xlsx, rolling over at CYCLE_SIZE.

        Best-effort by design: the master index has already accepted the key,
        so a failed spreadsheet write must never break collection.  It is
        reported loudly instead -- a silent failure here would look like the
        scraper was finding nobody.
        """
        try:
            return append_detail_to_excel(record, say=say)
        except Exception as exc:          # writing xlsx must never stop a run
            say("[CYCLE] FAILED to write %s (%s)"
                % (os.path.basename(C.DETAIL_OUTPUT_FILE), exc))
            say("[CYCLE] the person IS saved in %s (all 65 columns), so they "
                "are not lost" % os.path.basename(master_path()))
            return ""

    # ------------------------------------------------ pacing & cooldown ---
    async def _health_gate(self, check_network=True):
        """True to continue; False to stop the device cleanly.

        Called at two safe points.  Between searches the FULL gate runs.
        Inside the profile loop it runs with check_network=False: one search
        can hold the run for tens of minutes, so deferring the battery
        decision to the next search would let a 20% battery reach empty
        mid-profile -- the very thing this gate exists to prevent.  The
        network is deliberately left to the between-search gate: a link that
        drops mid-search is already handled as a transient profile failure,
        and a TCP connect per profile would buy nothing.

        A pause waits in HEALTH_RECHECK_SECONDS steps through the
        interruptible sleep, so Ctrl+C stays responsive while waiting, and
        continues BY ITSELF as soon as the condition clears -- no rerun, no
        operator.  With HEALTH_MAX_WAIT_MINUTES = None it waits indefinitely:
        an outage at 2am should find the run still going in the morning.
        That cannot strand a dying laptop, because the battery keeps draining
        while paused and BATTERY_CRITICAL_PERCENT still stops it cleanly.

        Once paused for the battery, continuing requires
        BATTERY_RESUME_PERCENT (battery_resumed), or the
        charger back -- mains power resumes immediately at any percent.
        """
        waited = 0
        limit = (C.HEALTH_MAX_WAIT_MINUTES * 60
                 if C.HEALTH_MAX_WAIT_MINUTES else None)
        announced = False
        battery_paused = False
        while True:
            check_stop()
            percent, plugged = battery_state()
            verdict = battery_verdict(percent, plugged)
            online = network_ok() if check_network else None

            if verdict == "stop":
                say("")
                say("[BATTERY CRITICAL] %d%% on battery power." % percent)
                say("[DEVICE STOPPED] state preserved; charge and rerun.")
                return False
            if verdict != "ok":
                battery_paused = True
            # Hysteresis: having paused for the battery, "ok" is not enough.
            battery_clear = verdict == "ok" and (
                not battery_paused
                or battery_resumed(percent, plugged))

            if battery_clear and online is not False:
                if announced:
                    say("[HEALTH] recovered; continuing.")
                return True

            if limit is not None and waited >= limit:
                say("")
                reason = ("battery still low" if not battery_clear
                          else "network still down")
                say("[DEVICE STOPPED] %s after %d minute(s)."
                    % (reason, C.HEALTH_MAX_WAIT_MINUTES))
                say("State preserved. Rerun to resume.")
                return False

            if not announced:
                say("")
                say(describe_health(percent, plugged, online))
                if not battery_clear:
                    say("[PAUSED] battery %d%% on battery power; waiting for "
                        "a charger." % percent)
                else:
                    say("[PAUSED] %s is unreachable; waiting for the network."
                        % C.HEALTH_HOST)
                if limit is None:
                    say("         re-checking every %ds; resumes by itself."
                        % C.HEALTH_RECHECK_SECONDS)
                else:
                    say("         re-checking every %ds, up to %d minute(s)."
                        % (C.HEALTH_RECHECK_SECONDS,
                           C.HEALTH_MAX_WAIT_MINUTES))
                announced = True

            cd = Countdown(C.HEALTH_RECHECK_SECONDS, "WAIT",
                                    prefix="[HEALTH]")
            cd.start()
            # Waiting for a charger or the network is not working time.
            work_clock_pause()
            try:
                await sleep_interruptible(C.HEALTH_RECHECK_SECONDS,
                                                    check_stop,
                                                    on_tick=cd.tick)
            finally:
                work_clock_resume()
            cd.finish()
            waited += C.HEALTH_RECHECK_SECONDS

    async def _crawl_health_gate(self):
        """The health gate as the page crawler wants it: pause in place, or
        raise so the device stops cleanly with every staged page preserved.

        Network is not re-checked between pages for the same reason it is not
        re-checked between profiles -- a link that drops mid-crawl already
        surfaces as a page failure, and the between-search gate covers it.
        """
        if not await self._health_gate(check_network=False):
            raise Stopped("battery critical during collection")

    def next_keyword_cycle(self):
        """Open the next lap and return its search list; best-effort, because
        losing the ORDER must never stop a lap from running."""
        if not KEYWORD_SHUFFLE_ENABLED:
            return self.searches
        try:
            # Read once and pass on: new_mylib.py is parsed as text.
            keywords = load_base_keywords()
            self.keyword_cycle = open_next_keyword_cycle(keywords)
            order = keyword_order(self.keyword_cycle, keywords)
            searches = build_search_list(say_total=False, keywords=order)
        except Exception as exc:
            say("[CYCLE] could not shuffle the keywords (%s); this lap keeps "
                "the previous order" % exc)
            return self.searches
        say("")
        say("[CYCLE] crawl cycle %d opened -- keywords shuffled once for "
            "this lap" % self.keyword_cycle)
        say("        order: %s%s"
            % (", ".join(order[:6]), " ..." if len(order) > 6 else ""))
        return searches

    async def _schedule_gate(self, page):
        """Take whichever scheduled stop is due, between searches only, and
        hand back the live page."""
        pause = due_pause()
        if pause is None:
            return page
        reason, seconds = pause
        return await self._long_pause(page, seconds, reason)

    async def _long_pause(self, page, seconds, reason):
        """Wait out a scheduled stop and hand back a usable page.

        A pause of at least BROWSER_RESTART_PAUSE_SECONDS closes Chromium and
        relaunches it after; the position is already saved, so being killed
        mid-pause costs nothing, and the work clock is stopped throughout.
        """
        seconds = int(max(0, seconds))
        if seconds <= 0:
            return page
        resume_at = (datetime.datetime.now()
                     + datetime.timedelta(seconds=seconds))
        say("")
        say("[PAUSE] %s" % reason)
        say("        waiting %s  ->  resuming at %s"
            % (fmt_hms(seconds), resume_at.strftime("%Y-%m-%d %H:%M:%S")))
        self.save_master_now("scheduled pause")
        closed = seconds >= BROWSER_RESTART_PAUSE_SECONDS
        if closed:
            say("        closing Chromium for the wait")
            await self._close_browser_for_pause()
            page = None
        cd = Countdown(seconds, "BREAK", prefix="[PAUSE]")
        cd.start()
        work_clock_pause()
        try:
            await sleep_interruptible(seconds, check_stop, on_tick=cd.tick)
        finally:
            work_clock_resume()
            cd.finish()
        # Any stop ends the work period, the night stop included.
        period = reset_work_timer()
        say("[RESUME] %s" % reason)
        if period > 0:
            say("         next break after %.1f h of work" % period)
        if closed:
            return await self.restart_browser()
        return await self.live_page(page)

    async def _close_browser_for_pause(self):
        """Shut Chromium down for a long pause, leaving the manager in the
        state restart() expects of a browser that is already gone."""
        if self.browser is None:
            return
        try:
            await self.browser.close()
        except Exception as exc:
            log.info("browser already gone before the pause: %s", exc)
        self.browser.playwright = None
        self.browser.context = None
        self.browser._context_closed = True
        self.browser._page = None
        # Rebuilt by restart_browser(); holding them would keep dead handles.
        self.scraper = None
        self.profiles = None

    async def _search_delay(self):
        """A random, interruptible delay between two DIFFERENT completed
        searches (never before resuming an unfinished one).

        The pause is proportional to the load the search that just finished
        actually put on GitHub.  A ZERO-RESULT search is a single results-page
        load -- no pagination, no profile fetches, no download -- so it gets
        the shorter ZERO_RESULT range; anything that really collected gets the
        full SEARCH range.  Most keyword-x-city pairs match nobody, so this is
        where the bulk of a run's waiting is spent.
        """
        if getattr(self, "_last_search_empty", False):
            low, high = ZERO_RESULT_DELAY_MIN, ZERO_RESULT_DELAY_MAX
        else:
            low, high = SEARCH_DELAY_MIN, SEARCH_DELAY_MAX
        seconds = random.uniform(low, high)
        cd = Countdown(seconds, "SEARCH")
        cd.start()
        await sleep_interruptible(seconds, check_stop,
                                            on_tick=cd.tick)
        cd.finish()

    async def _profile_delay(self):
        """A random, interruptible delay between two profile-page loads."""
        seconds = random.uniform(PROFILE_DELAY_MIN, PROFILE_DELAY_MAX)
        cd = Countdown(seconds, "PROFILE")
        cd.start()
        await sleep_interruptible(seconds, check_stop,
                                            on_tick=cd.tick)
        cd.finish()

    async def _respect_active_cooldown(self):
        """Honour a cooldown persisted by a previous run before touching GitHub.

        Returns True when it is safe to proceed (no cooldown, or it has been
        waited out) and False when the escalation is exhausted (stop the
        device).  Restarting the command therefore cannot bypass a cooldown."""
        if self.cooldown.exhausted():
            say("")
            say("[RATE LIMIT] Cooldown level %d exhausted" % self.cooldown.level)
            say("[DEVICE STOPPED]")
            say("State preserved. Resume later.")
            return False
        if not self.cooldown.active():
            return True
        remaining = self.cooldown.remaining()
        until = self.cooldown.data.get("cooldown_until", "")
        cd = CooldownDisplay(
            self.cooldown.level,
            self.cooldown.reason or "GitHub rate limit", until, remaining)
        cd.start()
        # A cooldown is not working time, so it cannot earn a work break.
        work_clock_pause()
        try:
            await sleep_interruptible(remaining, check_stop,
                                                on_tick=cd.tick)
        finally:
            work_clock_resume()
        cd.finish()
        return True

    async def _cooldown_and_wait(self, limit):
        """Escalate the cooldown for a search/pagination rate-limit and wait it
        out, so the SAME unit can be retried.  Returns True to retry, or False
        when the escalation is exhausted (the caller stops the device)."""
        level = self.cooldown.escalate(limit.source, str(limit),
                                       retry_after=limit.retry_after,
                                       kind=limit.kind)
        say("")
        if limit.kind == "outage":
            say("[OUTAGE] GitHub %s returned a server error; waiting it out."
                % limit.source)
            say("[OUTAGE] The job is not stopped -- this retries until "
                "GitHub recovers.")
        else:
            say("[RATE LIMIT] GitHub %s rate limit detected" % limit.source)
        if self.cooldown.exhausted():
            say("[RATE LIMIT] Cooldown level %d exhausted" % level)
            say("[DEVICE STOPPED]")
            say("State preserved. Resume later.")
            return False
        seconds = self.cooldown.remaining()
        until = self.cooldown.data.get("cooldown_until", "")
        cd = CooldownDisplay(
            level, self.cooldown.reason or ("GitHub %s rate limit"
                                            % limit.source), until, seconds)
        cd.start()
        # As above: cooling down is not collecting.
        work_clock_pause()
        try:
            await sleep_interruptible(seconds, check_stop,
                                                on_tick=cd.tick)
        finally:
            work_clock_resume()
        cd.finish()
        # Say exactly WHAT is being retried (req 6).
        say(retry_banner(
            limit.source, query=self._unit_query(),
            page=self._crawl_page(),
            username=getattr(self, "_rate_limited_profile", "")))
        return True

    def _unit_query(self):
        unit = getattr(self, "unit", None)
        return getattr(unit, "search_query", "") if unit else ""

    def _crawl_page(self):
        """The results page a pagination retry resumes at, if known."""
        state = getattr(self, "state", None)
        if state is None:
            return None
        current = int((state.data.get("crawl") or {}).get(
            "current_page", 0) or 0)
        return current + 1 if current else None

    async def _retry_deferred(self, page):
        """One conservative retry of every deferred profile.  Returns page.

        A success is saved by the normal do_candidate path and removed from the
        queue; a profile that is still rate-limited/unavailable stays deferred.
        Never stops the rank and never raises: this is best-effort cleanup.
        """
        entries = self.deferred.all_entries()
        if not entries:
            return page
        say("")
        say("=" * 60)
        say("[DEFERRED RETRY] rank %d: retrying %d parked profile(s)"
            % (rank, len(entries)))
        say("=" * 60)
        for entry in entries:
            check_stop()
            # Same gate as the main profile loop: the deferred queue can hold
            # hundreds of profiles, which is just as long a stretch of work.
            if not await self._health_gate(check_network=False):
                return page
            candidate = candidate_from(entry)
            self.unit = None
            say("")
            say("[DEFERRED RETRY] %s" % candidate.username)
            page = await self.ensure_browser(page)
            try:
                await self.do_candidate(page, candidate)
            except TransientProfileFailure as exc:
                self.deferred.touch(candidate.username, str(exc))
                say("  still unavailable (%s); kept deferred." % exc)
                # A global block during the deferred-retry pass must also stop:
                # do not keep hammering GitHub across every parked profile.
                if is_throttle_reason(str(exc)):
                    distinct = self.throttle.record_throttle(candidate.username)
                    if self.throttle.tripped():
                        say("")
                        say("[GLOBAL THROTTLE]")
                        say("GitHub rate-limit/challenge detected across %d "
                            "different profiles during deferred retry." % distinct)
                        say("Rank state saved. No profiles were discarded.")
                        say("Device stopped to respect GitHub's restriction.")
                        raise GlobalThrottle(
                            "global throttle during deferred retry")
            except Stopped:
                raise
            except Exception as exc:              # keep deferred, never crash
                self.deferred.touch(candidate.username, str(exc))
                say("  retry error (%s); kept deferred." % exc)
            else:
                self.deferred.remove(candidate.username)
                say("  recovered; removed from the deferred queue.")
            await pace(C.PACING_PROFILE)
        return page

    async def retry_deferred_only(self):
        """--retry-deferred: retry ONLY this rank's deferred profiles.

        Opens a browser, gives each parked profile one conservative attempt,
        saves and removes the ones that recover, keeps the rest deferred, and
        refreshes the workbook if anything was recovered.  Touches no search
        state and never resets the rank.
        """
        banner(self.searches, self.existing)
        before = self.deferred.count()
        if not before:
            say("Rank %d has no deferred profiles." % rank)
            return 0
        say("Deferred profiles to retry: %d" % before)
        await self.browser.start(with_extension=False)
        self.scraper = InstantDataScraper(self.browser, self.state, say)
        self.profiles = GitHubProfile(self.browser, say)
        page = await self.browser.page()
        await asyncio.sleep(1)
        await self.browser.close_extra_pages(keep=(page,))
        login = await self.browser.github_login_name(page)
        say("Logged in as: %s" % login if login
            else "GitHub login not detected.")
        await self._retry_deferred(page)
        remaining = self.deferred.count()
        recovered = before - remaining
        say("")
        say("[DEFERRED] rank %d: %d recovered, %d still deferred."
            % (rank, recovered, remaining))
        return 0

    # ------------------------------------------------------- search unit --
    async def do_search_unit(self, page, unit):
        page = await self.live_page(page)
        keyword = unit.search_query
        # Assume this search will do real work; the zero-result path below
        # says otherwise.  Read by _search_delay to size the pause that
        # follows THIS unit.
        self._last_search_empty = False
        # Rank-local download directory, so the export-recovery scan can never
        # see another rank's file.  The name stays readable:
        #     desktop-support-engineer-hyderabad.xlsx
        target = os.path.join(C.KEYWORD_DOWNLOAD_DIR,
                              "%s.xlsx" % unit.slug)

        existing = self.state.collection_file or target
        rows = validate_export(existing) if os.path.exists(existing) else -1

        if rows >= 0 and self.state.collection_complete:
            say("[RESUME] reusing %s (%d rows)"
                % (os.path.basename(existing), rows))
            target = existing
        elif self.options.process_existing and rows >= 0:
            say("[EXISTING] using %s (%d rows)"
                % (os.path.basename(existing), rows))
            target = existing
            self.state.set_in("collection", complete=True, xlsx_file=target,
                              rows=rows)
        else:
            if self.options.process_existing:
                say("[SKIP] no downloaded XLSX for this keyword")
                return
            status, rows = await self.collect(page, keyword, target)

            if status == CollectResult.FAILED:
                say("[KEEP] collection did not complete; nothing removed")
                raise Stopped("collection failed for %r" % keyword)

            if status == CollectResult.ZERO_RESULTS:
                # A COMPLETE search unit with no candidates.  There is no
                # XLSX to read, delete or preserve and no student to save, so
                # the unit is closed here rather than sent into the profile
                # phase with a filename that was never created.
                #
                # self.processed is untouched, so a zero-result search does
                # not consume --limit: the limit counts students saved, not
                # searches visited.
                say("[ZERO RESULTS] valid search completed with no "
                    "candidates")
                say("[COMPLETE] no XLSX required")
                # One page load only: the shorter post-search pause applies.
                self._last_search_empty = True
                self.state.set_in("collection", complete=True, xlsx_file="",
                                  rows=0)
                self.state.set_in("processing", next_row_index=0,
                                  total_rows=0, complete=True)
                self.state.finish_unit()
                # Nothing was downloaded and no student was started, so the
                # only thing that could be lying around is debris from an
                # earlier attempt at this same unit.
                leftovers = sweep_completed_keyword()
                if leftovers:
                    say("[CLEANUP] stale student temp files removed: %s"
                        % ", ".join(leftovers))
                return

            self.state.set_in("collection", complete=True, xlsx_file=target,
                              rows=rows)

        self.state.set(phase="profiles")
        page = await self.enter_profile_phase(page)
        resolved = await self.process_rows(
            page, target, unit,
            keyword_number=self.keyword_number,
            keyword_total=self.keyword_total)

        if not resolved:
            say("[KEEP] %s preserved (unresolved rows remain)"
                % os.path.basename(target))
            raise Stopped("unresolved rows for %r" % keyword)

        # Only here -- every source row resolved -- may anything be deleted.
        self.state.finish_unit()
        try:
            os.remove(target)
            say("[CLEANUP] temporary XLSX removed")
        except OSError as exc:
            say("[WARN] could not remove %s: %s" % (target, exc))
        # Rank-local by construction: every path temp_store knows came from
        # this rank's context, so rank 12 cannot reach runtime/rank_013.
        leftovers = sweep_completed_keyword()
        if leftovers:
            say("[CLEANUP] stale student temp files removed: %s"
                % ", ".join(leftovers))

    # -------------------------------------------------------- collection --
    async def collect(self, page, keyword, target):
        """Run one full extension crawl.  Returns row count, or -1."""
        # Collection gets the same recovery the profile phase has: a dead
        # page here used to end the run with TargetClosedError on the very
        # first navigation of the next keyword.
        if self.browser is not None:
            alive = await self.browser.is_alive()
            if not alive or page is None or page.is_closed():
                say("[BROWSER] collection page unavailable")
                page = await self.restart_browser()
                say("[RESUME] %s" % keyword)

        url = C.SEARCH_URL.format(query=urllib.parse.quote_plus(keyword))
        # A SLOW SEARCH PAGE MUST NOT KILL THE RUN.  This navigation was the
        # one unguarded await in collect(): a Playwright TimeoutError here
        # unwound all the way out of main() as a traceback.  See
        # C.SEARCH_NAV_ATTEMPTS for the run that proved it.
        for attempt in range(1, C.SEARCH_NAV_ATTEMPTS + 1):
            try:
                await page.goto(url, wait_until="domcontentloaded",
                                timeout=C.NAVIGATION_TIMEOUT)
                break
            except Exception as exc:
                log.warning("search page did not open (attempt %d of %d): %s",
                            attempt, C.SEARCH_NAV_ATTEMPTS, exc)
                if attempt >= C.SEARCH_NAV_ATTEMPTS:
                    say("[FAIL] the search page did not open after %d "
                        "attempt(s): %s" % (attempt, exc))
                    return CollectResult.FAILED, -1
                if browser_died(exc) and self.browser is not None:
                    say("[BROWSER] the browser died opening the search page")
                    page = await self.restart_browser()
                else:
                    say("[RETRY] the search page did not open "
                        "(attempt %d of %d); retrying"
                        % (attempt, C.SEARCH_NAV_ATTEMPTS))

        query = urllib.parse.parse_qs(urllib.parse.urlparse(page.url).query)
        if query.get("type", [""])[0] != "users":
            say("[FAIL] the search page is not type=users; refusing to crawl")
            return CollectResult.FAILED, -1
        say("[SEARCH] GitHub Users page opened")

        # WHAT IS THIS PAGE?  Asked once, from the page itself, BEFORE the
        # extension is opened: a zero-result search should not pay for Instant
        # Data Scraper, and a blocked page should not be poked at further.
        #
        # This replaces a hand-rolled substring test for
        #     "We couldn't find any users"
        # which used an ASCII apostrophe and so never matched GitHub's
        # typographic one.  That check failed silently, the flow continued
        # into the extension, and a second, separately broken test there
        # reported the empty page as a collection FAILURE -- which preserved
        # the search unit as unresolved and blocked the rank on every rerun.
        state = await self.scraper.capture_page_state(page, reason="collect")
        page_state, why = classify_search_page(state)
        log.info("page state: %s (%s)", page_state, why)

        if page_state == PageState.ZERO_RESULTS:
            say("[SEARCH] GitHub returned 0 users")
            return CollectResult.ZERO_RESULTS, 0

        if page_state in (PageState.BLOCKED, PageState.HTTP_ERROR):
            # A rate-limit/challenge on the SEARCH page is not a hard stop:
            # cool down and retry the SAME search.  Neither is a GitHub 5xx --
            # that used to stop the device on the very FIRST 503, ending an
            # overnight run over a hiccup that cleared in minutes.  Checked
            # first, because GitHub wraps a 503 in the screen it uses for 429.
            if is_outage_reason(why):
                say("[OUTAGE] GitHub returned a server error: %s" % why)
                say("         GitHub is broken, not limiting us; waiting and "
                    "retrying. The job is NOT stopped.")
                raise RateLimited(why, source="search", kind="outage")
            if is_throttle_reason(why):
                say("[BLOCKED] GitHub search rate limit: %s" % why)
                raise RateLimited(why, source="search")
            # Anything else -- a 404, an unexpected redirect -- is a real
            # error that waiting cannot fix, so it still stops.
            say("[BLOCKED] %s" % why)
            say("          Stopping safely. Nothing was deleted.")
            raise Stopped(why)

        if page_state == PageState.LOGIN_REQUIRED:
            say("[FAIL] %s" % why)
            say("       Run: python3 student_scraper.py --login")
            return CollectResult.FAILED, -1

        if page_state not in PageState.COLLECTABLE:
            say("[FAIL] %s (%s)" % (why, page_state))
            self.scraper.log_state(state)
            return CollectResult.FAILED, -1

        say("[SEARCH] %s" % why)

        # ---- multi-page: one fresh extension popup per results page -------
        # IDS's own cross-page crawl loses its table the moment GitHub
        # replaces the results DOM, which the diagnostic proved it does.  A
        # fresh popup per page sees every page; the pages are staged and
        # merged into ONE keyword XLSX only once all of them succeed.
        if page_state == PageState.RESULTS_MULTI_PAGE:
            return await self._collect_page_by_page(page, target, state)

        popup = None
        try:
            # Nothing but the search tab should be resident while the
            # extension crawls dozens of pages.
            await self.browser.close_extra_pages(keep=(page,))
            popup = await self.scraper.open_instant_data_scraper(page)

            # Three answers, three different jobs.  A result set with no
            # next control is COMPLETE, not broken: there is simply nothing
            # to paginate, so it is exported as it stands.
            pagination = await self.scraper.ensure_next_selector(popup, page)

            if pagination == NextControl.FAILED:
                say("[FAIL] this search could not be collected; see the "
                    "reason above")
                return CollectResult.FAILED, -1

            if pagination == NextControl.SINGLE_PAGE:
                say("[SCRAPER] single-page result set: no crawl needed")
                outcome, pages, rows = \
                    await self.scraper.finish_single_page(popup, page)
            else:
                if not await self.scraper.start_crawling(popup):
                    say("[FAIL] Start Crawling was not available")
                    return CollectResult.FAILED, -1
                # had_next carries the PRE-CRAWL fact that GitHub offered
                # a live next page.  Without it the crawl cannot tell
                # "finished" from "stopped before it paginated", which is how
                # a 2-page search was accepted after 1 page.
                outcome, pages, rows = \
                    await self.scraper.wait_for_crawl_complete(
                        popup, page, max_pages=self.options.max_pages,
                        had_next=(page_state
                                  == PageState.RESULTS_MULTI_PAGE))

            if outcome in CrawlOutcome.EXPORTABLE:
                say("[SCRAPER] collected %d row(s) over %d page(s)%s"
                    % (rows, pages,
                       " (single page)"
                       if outcome == CrawlOutcome.SINGLE_PAGE else ""))
                # BEFORE export: make sure IDS selected the right table.  On a
                # small results page it can auto-pick a table that is mostly
                # page chrome (one real user + garbage rows); cycling 'Try
                # another table' fixes it, and this happens pre-export because
                # export can kill Chromium in this environment.
                expected_ids = self.scraper.dom_result_identities(state)
                # Establish the correct table AND require it to stay stable:
                # the extension mutates its model (mn.data) asynchronously
                # after "Try another table", so a table validated a moment ago
                # can shift before the XLSX click.  stabilize_results_table
                # requires several consecutive matching reads and a final
                # pre-click re-read, bounded to a couple of attempts.
                if not await self.scraper.stabilize_results_table(
                        popup, page, expected_ids):
                    # EXACT SINGLE-PAGE DOM FALLBACK: a small result set (1..10
                    # on one page) has too little repetition for IDS to infer a
                    # table.  When GitHub's exact count matches exactly that
                    # many valid unique DOM identities, use those ALREADY-
                    # VERIFIED identities directly rather than losing them.
                    fallback = self._exact_single_page_dom_fallback(
                        state, outcome, target, keyword, expected_ids)
                    if fallback is not None:
                        return fallback
                    say("[INCOMPLETE] IDS result table did not remain stable")
                    say("             Nothing was exported; the search will "
                        "re-collect on the next run.")
                    return CollectResult.FAILED, -1
                # PROTECTION 1 (pre-export): one last read the instant before
                # writing; require it STILL matches the DOM.
                if expected_ids:
                    now = await self.scraper.preview_identities(
                        popup, expected_ids)
                    if now != expected_ids:
                        say("[INCOMPLETE] the IDS table changed after it was "
                            "validated; refusing to export a shifted table")
                        return CollectResult.FAILED, -1
                # Export from the VERIFIED PREVIEW, not the #xlsx button: the
                # button can write a table that diverged from the validated
                # preview and can kill Chromium on click.  Python writes the
                # page XLSX straight from the stabilised grid instead.
                if expected_ids:
                    exported = await self.scraper.export_page_from_preview(
                        popup, target, expected_ids)
                else:
                    # No DOM identities to verify against (not a Users results
                    # page); fall back to the extension export.
                    exported = await self.scraper.export_xlsx(popup, target)
                if exported < 0:
                    return CollectResult.FAILED, -1
                # PROTECTION 2 (post-export): the written file's identities
                # must EXACTLY equal the page's DOM identities -- not merely
                # match a count.  This is what catches an export whose model
                # diverged from the validated preview (the Engineering p2
                # case: preview 4/4 but XLSX held 1 user + page chrome).
                usable = len(self.reader.read(target, keyword))
                ok, why = self._verify_export_identities(
                    state, target, usable, exported, pages)
                if not ok:
                    preserved = self._preserve_incomplete(target)
                    say("[INCOMPLETE] %s" % why)
                    if preserved:
                        say("             Preserved for diagnosis: %s"
                            % os.path.basename(preserved))
                    say("             collection not marked complete; the "
                        "search will re-collect on the next run.")
                    return CollectResult.FAILED, -1
                say("[DOWNLOAD] %s" % os.path.basename(target))
                say("[OK] %d rows (%d usable)" % (exported, usable))
                return CollectResult.RESULTS, usable

            # The crawl stopped early.  Whatever the reason, rows already
            # sitting in the extension are real data from pages GitHub
            # served successfully -- they are saved beside the keyword file
            # rather than thrown away.  A partial is never processed
            # automatically, because doing so would delete the keyword and
            # silently lose the pages we never reached.
            partial = ""
            if rows and rows > 0:
                partial = target.replace(".xlsx", ".partial.xlsx")
                saved = await self.scraper.export_xlsx(popup, partial)
                if saved >= 0:
                    say("[SAVED] %d row(s) kept in %s"
                        % (saved, os.path.basename(partial)))
                else:
                    partial = ""

            if outcome == CrawlOutcome.RATE_LIMITED:
                say("[BLOCKED] GitHub is throttling this crawl.")
                say("          Nothing was deleted. Wait a while, then run")
                say("          the same command to retry this keyword.")
                if partial:
                    say("          Partial data: %s"
                        % os.path.basename(partial))
                raise Stopped("rate limited during collection")

            if outcome == CrawlOutcome.INCOMPLETE:
                say("[INCOMPLETE] pagination did not finish; this search is "
                    "NOT complete")
                say("             %d row(s) from %d page(s) were collected "
                    "and kept." % (max(rows, 0), max(pages, 0)))
                say("             Nothing was deleted. Run the same command "
                    "to retry this search.")
                if partial:
                    say("             Partial data: %s"
                        % os.path.basename(partial))
                return CollectResult.FAILED, -1

            say("[FAIL] crawling ended as '%s' after %d page(s)"
                % (outcome, max(pages, 0)))
            if partial:
                say("       Partial data preserved: %s"
                    % os.path.basename(partial))
            return CollectResult.FAILED, -1
        finally:
            if popup is not None:
                try:
                    await popup.close()
                except Exception:
                    pass
            alive = (self.browser is not None
                     and await self.browser.is_alive())
            say("[COLLECT] browser alive: %s" % ("yes" if alive else "no"))
            say("[COLLECT] context pages: %d" % self.browser.page_count())
            if alive:
                # A finished crawl leaves the popup's collected rows and
                # every crawled page's React state in memory.  Profile
                # visits start next, so that is released first -- on a
                # machine short of RAM it is the difference between
                # finishing and losing the tab.
                await self.browser.release_page_memory(page)
                try:
                    await self.browser.close_extra_pages(keep=(page,))
                except Exception:
                    pass

    # -------------------------------------------------- browser health ---
    async def enter_profile_phase(self, page):
        """Report health at the phase boundary and recover if necessary."""
        if self.browser is None:
            return page          # no browser to check (offline processing)
        alive = await self.browser.is_alive()
        closed = page is None or page.is_closed()
        say("")
        say("[PROFILE PHASE] browser alive: %s" % ("yes" if alive else "no"))
        say("[PROFILE PHASE] context pages: %d" % self.browser.page_count())
        say("[PROFILE PHASE] profile page closed: %s"
            % ("yes" if closed else "no"))
        if alive and not closed:
            return page
        return await self.restart_browser()

    async def restart_browser(self):
        """Relaunch Chromium and hand back a usable page.

        Collection is NOT repeated: the keyword XLSX is already on disk and
        scraper_state.json still points at it, so the run continues from
        the same unresolved candidate.
        """
        say("[BROWSER] context unavailable")
        say("[BROWSER] restarting Chromium")
        page = await self.browser.restart(with_extension=True)
        self.scraper = InstantDataScraper(self.browser, self.state, say)
        self.profiles = GitHubProfile(self.browser, say)
        login = await self.browser.github_login_name(page)
        await self.browser.close_extra_pages(keep=(page,))
        say("[BROWSER] session restored%s"
            % (" (logged in as %s)" % login if login else ""))
        if self.state.search_query:
            say("[RESUME] %s" % self.state.search_query)
        return page

    async def live_page(self, current=None):
        """The authoritative page for right now.

        A page captured before a restart is dead, and using it raises
        TargetClosedError.  Every phase asks for the page again instead of
        trusting one handed over minutes earlier.
        """
        if self.browser is None:
            return current
        if not await self.browser.is_alive():
            return await self.restart_browser()
        page = await self.browser.get_live_page()
        if page is None or page.is_closed():
            return await self.restart_browser()
        return page

    async def ensure_browser(self, page):
        """A live page, restarting Chromium only when one is needed."""
        if self.browser is None:
            return page
        if page is not None and not page.is_closed() \
                and await self.browser.is_alive():
            return page
        return await self.restart_browser()

    async def _reclaim(self, page, done):
        """Give memory back between students.

        Portfolio and resume pages close themselves, so a surplus page here
        means something opened one unexpectedly.  Every
        RELEASE_PAGE_EVERY students the working tab is emptied too, so one
        heavy profile does not stay resident.
        """
        if not C.LOW_MEMORY_MODE or self.browser is None:
            return
        try:
            surplus = await self.browser.close_extra_pages(keep=(page,))
            if surplus:
                log.info("closed %d surplus page(s) after student %d",
                         surplus, done)
            if C.RELEASE_PAGE_EVERY and done % C.RELEASE_PAGE_EVERY == 0:
                await self.browser.release_page_memory(page)
        except Exception as exc:
            log.debug("reclaim skipped: %s", exc)

    def _verify_export_count(self, state, usable, raw, pages):
        """(ok, why) -- may a single-page / IDS-crawl export be accepted?

        Enforced only when GitHub stated an exact, reachable total: then the
        USABLE identities must equal it (a shortfall means rows were lost, an
        excess means the result set shifted).  When the total is abbreviated
        (9.1k), absent, or above the result cap, there is nothing exact to
        check against and the export stands.  This is the same rule
        verify_collection() applies -- single-page must not be a
        hole in it.
        """
        expected_total, mode = expected_total_users(state)
        if mode != ValidationMode.EXACT_TOTAL or not expected_total:
            return True, ""
        if usable == expected_total:
            return True, ""
        if usable < expected_total:
            return (False,
                    "GitHub page contains %d users but IDS exported %d row(s) "
                    "(%d usable); %d missing"
                    % (expected_total, raw, usable, expected_total - usable))
        return (False,
                "GitHub page states %d users but %d usable identities were "
                "collected; the result set changed during collection"
                % (expected_total, usable))

    def _verify_export_identities(self, state, target, usable, raw, pages):
        """(ok, why): the exported file must match the page's DOM identities.

        Set equality against the DOM is the strong check -- it rejects an
        export that has the right COUNT but wrong identities (e.g. a real user
        plus a cell that parses to "images").  When the DOM identities are not
        available it falls back to the exact-count check, so nothing is
        weakened.
        """
        expected_ids = self.scraper.dom_result_identities(state)
        if expected_ids:
            got = self.scraper.export_identities(target)
            if got == expected_ids:
                return True, ""
            missing = sorted(expected_ids - got)
            extra = sorted(got - expected_ids)
            return (False,
                    "exported identities do not match the %d DOM user(s): "
                    "%d exported, missing %s, unexpected %s"
                    % (len(expected_ids), len(got),
                       missing[:5] or "none", extra[:5] or "none"))
        # No DOM identities to compare -- fall back to the count guard.
        return self._verify_export_count(state, usable, raw, pages)

    def _exact_single_page_dom_fallback(self, state, outcome, target, keyword,
                                        expected_ids):
        """Collect a small exact single-page result set from its verified DOM
        identities when IDS cannot infer a structured grid.

        Allowed ONLY when ALL hold: the page is a single page, GitHub stated an
        EXACT total between 1 and 10, and the DOM shows exactly that many valid
        UNIQUE github identities.  Writes one row per identity (just the
        profile URL -- NOT scraped profile data) and reads the file back with
        XlsxReader, requiring the recovered identity SET to equal exactly the
        expected DOM set.

        Returns (status, count) when it takes over -- success, or a
        verification failure (INCOMPLETE, state preserved) -- and None when the
        conditions are NOT met, so the caller reports the normal INCOMPLETE and
        the SearchUnit does not advance.  This never runs for multi-page
        searches (a different code path), abbreviated/unknown totals, or counts
        above GITHUB_USERS_PER_PAGE (10).
        """
        if outcome != CrawlOutcome.SINGLE_PAGE:
            return None
        total, mode = expected_total_users(state)
        # exact, reachable total only; abbreviated/unknown/oversized -> normal
        if mode != ValidationMode.EXACT_TOTAL or not total:
            return None
        if not 1 <= total <= GITHUB_USERS_PER_PAGE:
            return None            # 11+ never (and can't be one page anyway)
        # Same authoritative-DOM validation + one-identity-per-row serialization
        # used for a single page of a multi-page crawl.
        ok, count, reason = write_dom_fallback(state, target, expected_ids,
                                               total)
        if not ok:
            preserved = self._preserve_incomplete(target)
            say("[INCOMPLETE] exact single-page DOM fallback rejected: %s"
                % reason)
            if preserved:
                say("             Preserved for diagnosis: %s"
                    % os.path.basename(preserved))
            return CollectResult.FAILED, -1
        usable = len(self.reader.read(target, keyword))
        say("[COLLECT] exact single-page DOM fallback: %d verified user(s)"
            % count)
        say("[OK] %d row(s)" % usable)
        return CollectResult.RESULTS, usable

    def _preserve_incomplete(self, target):
        """Move an unverified export aside so it is not treated as complete.

        Renamed to <slug>.partial.xlsx, exactly as a multi-page shortfall
        preserves its rows.  Returns the preserved path, or "" if there was
        nothing to move.
        """
        if not os.path.exists(target):
            return ""
        partial = target.replace(".xlsx", ".partial.xlsx")
        try:
            os.replace(target, partial)
            return partial
        except OSError as exc:
            log.warning("could not preserve %s: %s", target, exc)
            return target

    async def _collect_page_by_page(self, page, target, first_state):
        """Multi-page collection.  ONE keyword XLSX, written once, verified.

        SUCCESS -> the staged pages are merged, verified and removed.
        FAILURE -> staging is preserved, no keyword XLSX exists, and the
                   search unit stays incomplete.  Unchanged rule.
        """
        slug = self.unit.slug if self.unit is not None else \
            slugify(os.path.basename(target).replace(".xlsx", ""))
        attempt = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        collector = PageByPageCollector(
            self.browser, self.scraper, say,
            check_stop=check_stop, health_gate=self._crawl_health_gate)
        try:
            await collector.collect(page, slug, first_state=first_state,
                                    max_pages=self.options.max_pages)
            self.state.set_in("crawl", **collector.state_block(attempt))
            unique, dupes, unusable, ok, why = collector.merge_result(target)
            # Compact crawl summary card at this phase transition (req 4).
            say(crawl_card(
                getattr(self, "keyword_number", 0),
                getattr(self, "keyword_total", 0),
                len(collector.pages),
                collector.expected_pages or len(collector.pages),
                unique, collector.expected_total))
        except PaginationRateLimit as limit:
            # A results page was rate-limited.  Staging is preserved; surface a
            # RateLimited so the run loop cools down and retries THIS search.
            self.state.set_in("crawl", **collector.state_block(attempt))
            if collector.staging and os.path.isdir(collector.staging):
                say("[PAGINATION] rate-limited; %d staged page(s) preserved: %s"
                    % (len(collector.pages),
                       os.path.relpath(collector.staging, C.PROJECT_ROOT)))
            raise RateLimited(
                str(limit), source="pagination",
                retry_after=limit.retry_after,
                kind="outage" if is_outage_reason(str(limit))
                else "rate-limit")
        except PageCrawlError as exc:
            self.state.set_in("crawl", **collector.state_block(attempt))
            self._report_crawl_incomplete(collector, str(exc))
            return CollectResult.FAILED, -1
        except BrowserLost:
            self.state.set_in("crawl", **collector.state_block(attempt))
            say("[INCOMPLETE] the browser died during page-by-page "
                "collection")
            say("             Staging preserved; nothing was deleted.")
            raise

        if ok:
            say("[DOWNLOAD] %s" % os.path.basename(target))
            say("[OK] %d rows (%s)" % (unique, why))
            collector.cleanup()          # only now: the keyword XLSX is proven
            return CollectResult.RESULTS, unique

        # The merge failed verification.  Only the RECONCILABLE exact-total
        # case (every page DOM-validated, unique rows, no unusable, natural end,
        # only the COUNT differs) gets the dynamic-result reconciliation; any
        # other failure is the normal INCOMPLETE.
        if not collector.is_exact_total_mismatch(unique, dupes, unusable):
            self._remove(target)
            self._report_crawl_incomplete(collector, why)
            return CollectResult.FAILED, -1

        return await self._reconcile_total(page, collector, target, unique)

    def _remove(self, path):
        try:
            os.remove(path)
        except OSError:
            pass

    def _report_crawl_incomplete(self, collector, why):
        say("")
        say("[INCOMPLETE] %s" % why)
        say("             %d page(s) exported before stopping."
            % len(collector.pages))
        if collector.staging and os.path.isdir(collector.staging):
            say("             Staging preserved: %s"
                % os.path.relpath(collector.staging, C.PROJECT_ROOT))
        say("             No keyword XLSX was created and nothing was deleted.")
        say("             Run the same command to retry this search.")

    async def _reconcile_total(self, page, collector, target, unique):
        """Dynamic-result reconciliation for a reconcilable EXACT_TOTAL
        mismatch.  `target` holds the first crawl's merged set.  Returns a
        (status, count).  Never weakens validation: it accepts only a snapshot
        GitHub itself now confirms, or -- with the explicit flag -- a set two
        fully validated crawls agree on."""
        initial = collector.expected_total
        first_ids = InstantDataScraper.export_identities(target)
        # STRICT_TOTAL_MATCH is the default that stops on a persistent mismatch;
        # the explicit flag is the ONLY thing that permits accepting a stable
        # mismatch (two identical fully-validated crawls).
        accept_stable = getattr(self.options,
                                "accept_stable_pagination_mismatch", False)
        accepted, info = await collector.reconcile(
            page, target, unique, first_ids, accept_stable)

        if info["kind"] == "reconciled":
            say("")
            say("[DYNAMIC RESULT RECONCILED]")
            say("Initial GitHub total : %s" % info["initial"])
            say("Final GitHub total   : %s" % info["final"])
            say("Collected unique     : %d" % info["unique"])
            say("Result set changed during pagination; final snapshot "
                "validated.")
            collector.cleanup()
            return CollectResult.RESULTS, unique

        if info["kind"] == "stable":
            say("")
            say("[STABLE PAGINATION SET]")
            say("GitHub displayed total: %s" % info["initial"])
            say("Observed stable unique pagination set: %d" % info["unique"])
            say("Two independent validated crawls produced the same identity "
                "set.")
            if accepted:
                say("[ACCEPTED] --accept-stable-pagination-mismatch: taking the "
                    "%d verified user(s)." % unique)
                collector.cleanup()
                return CollectResult.RESULTS, unique
            self._remove(target)
            say("[INCOMPLETE] strict total match (default): the count still "
                "does not equal GitHub's stated total.")
            say("             Re-run with --accept-stable-pagination-mismatch "
                "to accept this verified set.")
            self._report_crawl_incomplete(collector, "stable pagination "
                                          "mismatch (%d vs stated %s)"
                                          % (unique, initial))
            return CollectResult.FAILED, -1

        # reconciliation did not confirm the set (final total differs / a
        # second crawl produced a different set / it could not complete)
        self._remove(target)
        detail = {"changed": "a verification re-crawl produced a different "
                  "identity set",
                  "verify_failed": "the verification re-crawl could not "
                  "complete"}.get(info["kind"], "the totals could not be "
                                   "reconciled")
        say("")
        say("[INCOMPLETE] dynamic-result reconciliation failed: %s." % detail)
        say("             Initial %s, final %s, collected %d."
            % (info["initial"], info["final"], info["unique"]))
        self._report_crawl_incomplete(collector, "total mismatch not "
                                      "reconciled")
        return CollectResult.FAILED, -1

    # ---------------------------------------------------------- per row ---
    async def process_rows(self, page, path, unit, keyword_number=0,
                           keyword_total=0):
        keyword = unit.search_query
        candidates = self.reader.read(path, keyword)
        total = len(candidates)
        self.state.set_in("processing", total_rows=total)
        source_rows = int(self.state.data.get("collection", {}).get("rows", 0)
                          or 0) or total

        # EVERY candidate is processed, known or not -- find_naukri_job.py's
        # behaviour.  Its `if key in existing_data:` branch only bumps the
        # counters; scrape_job_detail_async() and append_detail_to_excel()
        # sit at the SAME indentation as that if/else, so they run either
        # way and a job it already knows is fetched and written again.
        #
        # This used to skip the known ones here, before any fetching, which
        # saved the profile + README + portfolio + resume for each of them --
        # 111 profiles on one search of the previous run.  That saving is
        # given up on purpose so the two scripts behave the same way.
        pending = list(range(total))
        to_process = total
        ordinal = {index: number
                   for number, index in enumerate(pending, start=1)}

        say("")
        say("=" * 50)
        if keyword_total:
            say("[Search %d/%d] [Rank %d/%d] %s"
                % (keyword_number, keyword_total, rank,
                   num_sys, keyword))
        else:
            say("Search: %s" % keyword)
        say("Source XLSX: %s" % os.path.basename(path))
        say("Source rows: %d" % source_rows)
        say("Unique candidates: %d" % total)
        say("Profiles to process: %d" % to_process)
        say("=" * 50)

        if total == 0:
            say("[PROCESS] no usable rows in %s" % os.path.basename(path))
            return True

        # Per-search throttle bookkeeping (see _raise_global_throttle).
        self._deferred_this_run = set()
        self._throttle_min_pos = None
        # Whether a profile page has been loaded yet this search: the between-
        # profiles delay applies only from the SECOND load onward.
        self._profile_loaded_once = False

        start = self.state.next_row_index
        if start:
            done = sum(1 for index in pending if index < start)
            say("")
            say("[RESUME]")
            say("Keyword: %s" % keyword)
            say("Row/Candidate: %d/%d" % (start + 1, total))
            say("Profiles processed so far: %d/%d" % (done, to_process))
            say("Source XLSX: %s" % os.path.basename(path))

        # Compact profile-phase card at the transition into fetching (req 4):
        # first candidate this run will actually load, shown once -- the
        # per-candidate lines below carry the running progress.
        first = next((index for index in pending if index >= start), None)
        if first is not None:
            say(profile_card(
                ordinal[first], to_process, candidates[first].username,
                keyword_number or getattr(self, "keyword_number", 0),
                keyword_total or getattr(self, "keyword_total", 0)))

        for position in range(start, total):
            check_stop()
            # Battery gate INSIDE the profile loop.  The between-search gate
            # is not enough on its own: one search can run for tens of
            # minutes, so a battery that reaches the pause threshold here
            # would otherwise keep scraping until the search ended.  Network
            # is not re-checked per profile (see _health_gate).  Returning
            # False preserves the keyword XLSX and stops the device cleanly,
            # exactly as the --limit path below does.
            if not await self._health_gate(check_network=False):
                return False
            if self.options.limit and self.processed >= self.options.limit:
                say("[LIMIT] reached --limit %d" % self.options.limit)
                # Rows remain: the keyword is deliberately incomplete.
                return False
            candidate = candidates[position]
            say("")
            if position in ordinal:
                say("[Candidate %d/%d] [Rank %d/%d] [%s] %s"
                    % (ordinal[position], to_process, rank,
                       num_sys, keyword,
                       candidate.username))
            else:
                # A known person consumes no slot in the denominator.
                say("[SKIP] [%s] %s (already saved)"
                    % (keyword, candidate.username))
            self.state.set_in("current_candidate",
                              username=candidate.username,
                              source_row=candidate.source_row_number)
            # Conservative delay BETWEEN two profile-page loads: only for a
            # candidate that will actually be fetched (a known person is a
            # cheap skip, not a load), and only from the second load onward --
            # never before the first, never after the last, and not stacked on
            # the retry/cooldown waits (those live elsewhere).
            if position in ordinal:
                if self._profile_loaded_once:
                    await self._profile_delay()
                self._profile_loaded_once = True
            # Cheap health check before every navigation.  A dead browser
            # is restarted rather than navigated, so the run continues on
            # THIS candidate instead of failing on it.
            page = await self.ensure_browser(page)
            try:
                await self.do_candidate(page, candidate)
            except TransientProfileFailure as exc:
                entry = self.transient.record(candidate, exc)
                if entry["attempts"] < C.PROFILE_ATTEMPTS_MAX:
                    say("[RETRY] %s failed transiently (attempt %d/%d): %s"
                        % (candidate.username, entry["attempts"],
                           C.PROFILE_ATTEMPTS_MAX, exc))
                    await sleep_jitter(3 * entry["attempts"], check_stop)
                    try:
                        await self.do_candidate(page, candidate)
                    except TransientProfileFailure as again:
                        entry = self.transient.record(candidate, again)
                    except Exception as other:
                        say("[ERROR] %s: %s" % (candidate.username, other))
                        log.exception("row %d failed", position + 1)
                        return False
                    else:
                        self.transient.clear(candidate.username)
                        self.state.set_in("processing",
                                          next_row_index=position + 1)
                        continue
                # GLOBAL THROTTLE CIRCUIT BREAKER.  If this failure is a GitHub
                # rate-limit/challenge, record it; when several DIFFERENT
                # profiles hit that within the window it is an account/session
                # block, so stop the rank and the device rather than keep
                # probing GitHub -- and do NOT defer these profiles.
                if is_throttle_reason(str(exc)):
                    if self._throttle_min_pos is None:
                        self._throttle_min_pos = position
                    distinct = self.throttle.record_throttle(candidate.username)
                    if self.throttle.tripped():
                        self._raise_global_throttle(distinct, str(exc))
                    # ONE profile rate-limited is not an account-wide block.
                    # Rather than leaving the search unresolved -- which stops
                    # the rank until a human reruns it -- cool down and retry
                    # this same search from its saved resume point, exactly as
                    # a search-page or pagination rate limit already does.
                    # The profile keeps its failure count, so if it is still
                    # limited after the wait it reaches MAX_DEFERRED_FAILURES,
                    # gets parked, and the rank moves on by itself.
                    # C.PROFILE_RATE_LIMIT_COOLDOWN directly, not getattr with a False
                    # default: a getattr default turns "the constant was
                    # deleted" into "the feature is off", silently.
                    if (C.PROFILE_RATE_LIMIT_COOLDOWN
                            and entry["attempts"] < C.MAX_DEFERRED_FAILURES):
                        say("")
                        say("[RATE LIMIT] %s was rate-limited (attempt %d/%d)."
                            % (candidate.username, entry["attempts"],
                               C.MAX_DEFERRED_FAILURES))
                        say("[RATE LIMIT] Cooling down and retrying this "
                            "search instead of stopping the rank.")
                        self.state.set_in("processing",
                                          next_row_index=position)
                        self._rate_limited_profile = candidate.username
                        raise RateLimited(
                            str(exc), source="profile",
                            kind="outage" if is_outage_reason(str(exc))
                            else "rate-limit")
                # Still failing after the inline retry.  A profile that has
                # failed MAX_DEFERRED_FAILURES separate times is REPEATEDLY
                # unavailable: park it in the rank's deferred queue and move on,
                # so it can no longer block every other candidate and search.
                # Below that threshold it stays retryable (the search does not
                # complete, so the next run resumes to it).
                if entry["attempts"] >= C.MAX_DEFERRED_FAILURES:
                    self.deferred.add(candidate, str(exc), entry["attempts"])
                    self._deferred_this_run.add(candidate.username.lower())
                    self.transient.clear(candidate.username)
                    say("")
                    say("[DEFERRED] %s" % candidate.username)
                    say("  Reason: %s" % exc)
                    say("  Failures: %d" % entry["attempts"])
                    say("  Continuing with remaining candidates.")
                    self.state.set_in("processing", next_row_index=position + 1)
                    continue
                self._report_blocking(entry)
                return False
            except Stopped:
                raise
            except BrowserLost as exc:
                # The browser died mid-candidate.  Restart and give this
                # one candidate a second attempt before giving up on the
                # keyword; nothing is deleted either way.
                log.warning("browser lost on %s: %s", candidate.username,
                            exc)
                if self.browser.restarts >= C.BROWSER_MAX_RESTARTS:
                    say("[BROWSER] giving up after %d restart(s)"
                        % self.browser.restarts)
                    raise
                page = await self.restart_browser()
                try:
                    await self.do_candidate(page, candidate)
                except Exception as again:
                    say("[ERROR] %s after restart: %s"
                        % (candidate.username, again))
                    return False
            except Exception as exc:
                say("[ERROR] %s: %s" % (candidate.username, exc))
                log.exception("row %d failed", position + 1)
                return False
            # Written only after the row is safely on disk (or skipped).
            self.state.set_in("processing", next_row_index=position + 1)
            # One student survived, so the browser is working: credit it
            # against the restart budget (see BrowserManager.note_healthy).
            if self.browser is not None:
                self.browser.note_healthy()
            await self._reclaim(page, position + 1)

        self.state.set_in("processing", complete=True)
        say("")
        say("[COMPLETE] [%s] %d/%d profiles processed, "
            "%d candidate(s) resolved"
            % (keyword, to_process, to_process, total))
        return True

    def _raise_global_throttle(self, distinct, reason):
        """Undo this run's throttle-driven deferrals, rewind, and halt.

        The profiles that formed the streak are NOT left in the deferred queue
        (they were global victims, not profile-specific), and next_row_index is
        rewound so a resume after the cooldown retries them.  State and the
        keyword XLSX are preserved -- nothing is discarded.  Raises
        GlobalThrottle, which stops the rank and the device.
        """
        streak = set(self.throttle.streak_usernames())
        undone = 0
        for username in list(self._deferred_this_run):
            if username in streak and self.deferred.has(username):
                self.deferred.remove(username)
                undone += 1
        if self._throttle_min_pos is not None:
            # Resume retries from the first throttled profile of this run.
            self.state.set_in("processing",
                              next_row_index=self._throttle_min_pos)
        say("")
        say("[GLOBAL THROTTLE]")
        say("GitHub rate-limit/challenge detected across %d different "
            "profiles." % distinct)
        say("Rank state saved.")
        say("No profiles were discarded.")
        if undone:
            say("%d profile(s) were removed from the deferred queue (they were "
                "throttled by the global block, not profile-specific)." % undone)
        say("Device stopped to respect GitHub's restriction.")
        raise GlobalThrottle(
            "global throttle: %d distinct profiles rate-limited within %d "
            "minute(s)" % (distinct, C.GLOBAL_THROTTLE_WINDOW_MINUTES))

    async def do_candidate(self, page, candidate):
        # Asked BEFORE any expensive work: a known person costs one dict
        # lookup, not a profile fetch plus a portfolio plus a resume.
        key = "github:%s" % normalize_username(candidate.username)
        if key in self.existing:
            say("[DUPLICATE] already collected (%s); skipping"
                % candidate.username)
            log.info("duplicate %s", candidate.username)
            return

        bundle = SourceBundle(candidate=candidate)

        # -- GitHub ---------------------------------------------------------
        try:
            fields, text = await self.profiles.fetch(page, candidate.username)
            bundle.github_ok = True
            bundle.github_fields = fields
            bundle.github_text = text
            bundle.github_links = list(fields.get("links", []))
            say("[GITHUB] OK")
            # A normal profile load proves the account is not blocked right
            # now: clear the global-throttle streak.
            self.throttle.record_success()
        except GitHubProfileError as exc:
            if exc.permanent:
                say("[GITHUB] permanently unavailable (%s); skipping" % exc)
                self._record_skip(candidate.username, str(exc))
                self.transient.clear(candidate.username)
                return
            say("[GITHUB] temporary failure: %s" % exc)
            raise TransientProfileFailure(str(exc))
        await pace(C.PACING_PROFILE)

        # -- portfolio ------------------------------------------------------
        website = bundle.github_fields.get("website", "")
        portfolio_url = extract_portfolio(bundle.github_links, website)
        if portfolio_url:
            bundle.portfolio_url = clean_url(portfolio_url)
            text, links, visited = await self.portfolios.fetch(
                self.browser.context, bundle.portfolio_url)
            bundle.portfolio_text = text
            bundle.portfolio_links = links
            bundle.portfolio_pages = visited
            if text.strip():
                say("[PORTFOLIO] Found (%d page(s))" % len(visited))
            else:
                say("[PORTFOLIO] Link present but unreadable")
                bundle.note("portfolio unavailable")
        else:
            say("[PORTFOLIO] Not found")

        # -- resume ---------------------------------------------------------
        haystack = "\n".join([bundle.github_text, bundle.portfolio_text])
        links = bundle.github_links + bundle.portfolio_links
        resume_urls = extract_resume_urls(haystack, links)
        if resume_urls:
            bundle.resume_url = resume_urls[0]
            text, saved, remark = await self.resumes.fetch(
                self.browser.context, bundle.resume_url, candidate.username)
            bundle.resume_text = text
            bundle.resume_file = saved
            if remark:
                bundle.note(remark)
            if text.strip():
                extension = os.path.splitext(saved)[1].upper().lstrip(".") \
                    or "WEB"
                say("[RESUME] Found %s" % extension)
            else:
                say("[RESUME] Link present but not readable")
        else:
            say("[RESUME] Not found")

        # -- raw dump, extraction, save --------------------------------------
        write(bundle)
        record = self.merger.build(bundle)
        populated = sum(1 for value in record.values() if str(value).strip())
        say("[EXTRACT] %d fields populated" % populated)

        # CONTACTABLE-ONLY COLLECTION (config.REQUIRE_CONTACT).  A person is
        # kept when ANY contact route was found -- an email OR a phone number.
        # The merger has already discarded placeholder and non-personal
        # addresses and non-conforming numbers, so a non-empty value is a real
        # one.  Someone with neither is NOT written to the store or the
        # workbook; we drop them here, after fetching, so every source
        # (GitHub / portfolio / resume / README / source row) had its chance.
        if C.REQUIRE_CONTACT and not self._contact_found(record):
            say("[SKIP] no email or phone for %s; not saved (REQUIRE_CONTACT)"
                % candidate.username)
            self._record_skip(candidate.username, "no email or phone")
            clear()
            self.resumes.discard(bundle.resume_file)
            self.transient.clear(candidate.username)
            return

        # THE MASTER IS A TALLY, NOT A GATE -- find_naukri_job.py's structure:
        #
        #     if key in existing_data:          <- bump seen_count/last_seen_at
        #         ...
        #     else:
        #         existing_data[key] = {...}    <- first sighting
        #
        #     input_row  = {...}                <- OUTSIDE the if/else, so both
        #     detail_row = await scrape_job_detail_async(...)   branches reach it
        #     append_detail_to_excel(detail_row, "current_data.xlsx")
        #
        # The row is therefore written whether or not the key was already
        # known, and current_data.xlsx can hold the same person once per
        # keyword that found them.  master_index.xlsx stays one row per
        # person, with seen_count saying how many times we met them.
        #
        # The SEARCH location never reaches the 65 student columns.  It is
        # where we LOOKED; this person's Current City came from their own
        # profile, resume or portfolio and may be somewhere else entirely.
        key = make_unique_key(record) or ("github:%s"
                                          % normalize_username(candidate.username))
        now = _now()
        if key in self.existing:
            seen = self.existing[key]
            try:
                count = int(seen.get("seen_count", 0) or 0)
            except (TypeError, ValueError):
                count = 0
            seen["seen_count"] = count + 1
            seen["last_seen_at"] = now
            say("[SEEN AGAIN] %s (%d time(s)); row written again"
                % (candidate.username, count + 1))
        else:
            record["first_seen_at"] = now
            record["last_seen_at"] = now
            record["seen_count"] = 1
            self.existing[key] = record
            say("[SAVE] master index (%d person/people)" % len(self.existing))
        self._master_dirty = True
        self.processed += 1
        self.write_to_cycle_file(record)

        clear()
        self.resumes.discard(bundle.resume_file)
        self.transient.clear(candidate.username)

    @staticmethod
    def _report_blocking(entry):
        say("")
        say("=" * 50)
        for line in describe_transient_failure(entry).splitlines():
            say(line)
        say("")
        say("This search has NOT been marked complete.")
        say("XLSX preserved. State saved.")
        say("This profile is NOT blacklisted.")
        say("Run the scraper again to retry.")
        say("=" * 50)

    @staticmethod
    def _contact_found(record):
        """True when the record carries at least one way to reach the person.

        Any ONE of config.CONTACT_FIELDS (email or phone) is enough: a lead
        with only a phone is still a lead.  Values are already validated by
        the merger, so a non-empty field is a usable one.
        """
        for field in C.CONTACT_FIELDS:
            if str(record.get(field, "") or "").strip():
                return True
        return False

    def _record_skip(self, username, reason):
        """PERMANENT skips only, in this rank's own file."""
        try:
            with open(C.SKIPPED_FILE, "a",
                      encoding="utf-8") as handle:
                handle.write("%s\t%s\n" % (username, reason))
        except OSError:
            pass

    async def close(self):
        # Ctrl+C reaches the whole process group, so Playwright's driver may
        # already be gone by the time this runs.  Closing a browser that is
        # no longer there is not a failure, and letting it raise here would
        # replace the real reason the run ended.
        try:
            await self.browser.close()
        except Exception as exc:
            log.info("browser already gone at close: %s", exc)
        if self._master_dirty:
            try:
                save_master(self.existing)
            except Exception as exc:
                say("[MASTER] could not be written (%s); it keeps its previous "
                    "contents and current_data.xlsx is unaffected" % exc)


# ================================================================ modes ====
async def mode_login(show=True):
    """Authenticate this folder's Chromium profile.

    Each clone has its own persistent profile, so each needs its own
    interactive login.  Copying a profile directory from another clone is not
    offered, because copying a live Chromium profile is unsafe and copying a
    stopped one still duplicates session state GitHub has every right to
    treat as suspicious.  Nothing here touches GitHub's authentication beyond
    driving a real login form in a real browser.
    """
    setup_logging()
    browser = BrowserManager(headless=not show)
    say("Rank %d of %d -- profile: %s" % (rank, num_sys, C.PROFILE_DIR))
    say("Opening Chromium. Log into GitHub in the window, then close it.")
    await browser.start(with_extension=True)
    page = await browser.page()
    await page.goto("https://github.com/login",
                    wait_until="domcontentloaded")
    say("Waiting for the login to complete (press Ctrl+C to abort)...")
    try:
        for _ in range(600):
            await asyncio.sleep(1)
            try:
                name = await page.get_attribute('meta[name="user-login"]',
                                                "content", timeout=1000)
            except Exception:
                name = None
            if name:
                say("Logged in as: %s" % name)
                try:
                    count = await save_github_auth(browser.context)
                except Exception as exc:
                    say("WARNING: could not write %s (%s)."
                        % (os.path.basename(C.AUTH_FILE), exc))
                    say("The login still works on THIS machine -- it is in "
                        "%s -- but there is no portable copy."
                        % os.path.basename(C.PROFILE_DIR))
                else:
                    say("")
                    say("Session saved : %s (%d cookie(s))"
                        % (C.AUTH_FILE, count))
                    say("              : permissions 0600 -- this file IS the")
                    say("                login. Anyone holding it is signed in")
                    say("                as you, with no password needed.")
                    say("")
                    say("To set up another laptop, copy just this file into "
                        "its folder.")
                    say("It is ~50 KB; the profile directory is about 1 GB.")
                break
        else:
            say("Timed out waiting for a login.")
    finally:
        await browser.close()
    return 0


async def mode_check_extension(show=False):
    setup_logging()
    ok = True

    def report(label, good, detail=""):
        nonlocal ok
        ok = ok and good
        say("  [%s] %-38s %s" % ("OK" if good else "FAIL", label, detail))

    say("Instant Data Scraper diagnostic")
    say("=" * 50)
    directory = C.find_extension_dir()
    report("extension directory found", bool(directory), directory or
           "not found")
    if not directory:
        return 1

    browser = BrowserManager(headless=not show)
    popup = None
    try:
        await browser.start(with_extension=True)
        report("Chromium started", True)
        report("persistent context loaded", True, C.PROFILE_DIR)
        report("extension service worker", browser.service_worker is not None)
        report("extension id", bool(browser.extension_id),
               browser.extension_id)

        page = await browser.page()
        await page.goto(C.SEARCH_URL.format(query="student"),
                        wait_until="domcontentloaded")
        report("GitHub test page opened", "github.com" in page.url, page.url)

        state = StateManager()
        scraper = InstantDataScraper(browser, state, say)
        popup = await scraper.open_instant_data_scraper(page)
        report("extension UI opened", popup is not None, popup.url[:70])

        has_next = await popup.locator(C.SEL_NEXT_BUTTON).count() > 0
        has_start = await popup.locator(C.SEL_START_CRAWL).count() > 0
        report('pagination control present', has_next or has_start,
               "Start crawling already enabled" if has_start
               else 'Locate "Next" offered')

        has_xlsx = await popup.locator(C.SEL_XLSX).count() > 0
        report("XLSX control", has_xlsx)
        pages, rows = await scraper.read_stats(popup)
        report("stats readable", pages >= 0,
               "pages=%d rows=%d" % (pages, rows))

        # The real test: can the extension be taught to paginate, so that
        # "Start crawling" becomes available without any manual step?
        learned = await scraper.ensure_next_selector(popup, page)
        ready = await popup.locator(C.SEL_START_CRAWL).count() > 0
        host = "github.com"
        selector = state.next_selector_for(host)
        report("Start Crawling control", learned and ready,
               ("pagination selector: %s" % selector) if selector
               else "no Next control on this page")
    except Exception as exc:
        report("diagnostic completed", False, str(exc))
    finally:
        if popup is not None:
            try:
                await popup.close()
            except Exception:
                pass
        await browser.close()

    say("=" * 50)
    say("RESULT: %s" % ("all checks passed" if ok else "problems found"))
    return 0 if ok else 1


def banner(searches, existing=None):
    """The startup block.  Everything needed to tell two copies apart.

    `existing` is the master index this run loaded -- the dict Workflow keeps
    as self.existing.  It used to be a ResultStore, back when students.db was
    the duplicate index; the parameter was never renamed when the master
    workbook replaced it, and `self.store` stopped existing, so every real run
    died here with AttributeError.  Hence the test in test_workflow.py.
    """
    # A search is a (city, spelling, keyword) tuple -- build_search_list
    # keeps tuples and calls make_unit() only when a search's turn
    # comes, so unit[0] is the City.
    cities = len({unit[0].canonical for unit in searches})
    say("=" * 60)
    say("GitHub Student Scraper")
    say("Country             : %s" % COUNTRY)
    say("Rank                : %d of %d" % (rank, num_sys))
    say("Cities this rank    : %d of %d"
        % (cities, canonical_city_count()))
    say("Base keywords       : %d" % len(load_base_keywords()))
    say("Searches this rank  : %s" % format(len(searches), ","))
    try:
        cycle = read_keyword_cycle()
        say(describe_keyword_cycle(cycle, keyword_order(cycle)))
    except Exception as exc:                  # a banner must never stop a run
        say("Keyword order       : unavailable (%s)" % exc)
    say("Night stop          : %s" % night_pause_status())
    say("Work / break        : %s" % work_break_status())
    say("Pacing (random each time)")
    say("  search / empty    : %g-%g s / %g-%g s"
        % (SEARCH_DELAY_MIN, SEARCH_DELAY_MAX,
           ZERO_RESULT_DELAY_MIN, ZERO_RESULT_DELAY_MAX))
    say("  results page      : %g-%g s" % (PAGE_DELAY_MIN, PAGE_DELAY_MAX))
    say("  profile page      : %g-%g s"
        % (PROFILE_DELAY_MIN, PROFILE_DELAY_MAX))
    say("  cooldown per lap  : %s"
        % ("%d min" % int(LAP_COOLDOWN_MINUTES)
           if int(LAP_COOLDOWN_MINUTES) > 0 else "none"))
    say("Folder              : %s" % C.PROJECT_ROOT)
    say("Collected file      : %s" % os.path.basename(C.DETAIL_OUTPUT_FILE))
    say("Cycle size          : %d rows per cycle file" % C.CYCLE_SIZE)
    # NO "Duplicate index" LINE HERE, ON PURPOSE
    # ------------------------------------------
    # The master index is not part of the spec this script was written to --
    # that spec is: append to current_data.xlsx, roll at CYCLE_SIZE, save
    # current_row.json after each row, loop forever.  The index is an addition
    # that keeps the run from re-fetching a person a different keyword already
    # found, and it earns its place silently: a line about it in the startup
    # banner only invited "what is this file and why is it empty?".
    #
    # The count is still reported where it means something -- in the per-search
    # card (rank_card) and at the end of each lap.  The parameter stays so
    # banner() keeps its signature and the tests that call it keep working.
    say("=" * 60)


def mode_validate_keywords():
    """Check the production keyword file and show the resolved list."""
    setup_logging()
    path = os.path.join(C.PROJECT_ROOT, KEYWORD_SOURCE_FILE)
    say("=" * 60)
    say("PRODUCTION BASE KEYWORDS")
    say("=" * 60)
    say("File : %s" % path)
    problems = base_keyword_problems()
    try:
        keywords = load_base_keywords()
    except BaseKeywordError as exc:
        say("")
        say("ERROR: %s" % exc)
        return 1
    say("Keywords : %d" % len(keywords))
    say("Fingerprint : %s" % keywords_fingerprint(keywords))
    say("")
    for position, keyword in enumerate(keywords, 1):
        say("  %3d. %s" % (position, keyword))
    say("")
    if problems:
        say("PROBLEMS (%d):" % len(problems))
        for number, text, reason in problems:
            say("  line %-4d %-40s %s" % (number, text[:40], reason))
        say("RESULT: FAIL")
        return 1
    say("No problems found.")
    say("")
    say("Cities come from py: %d canonical, %d search spellings."
        % (canonical_city_count(), search_spelling_count()))
    say("SearchUnits this produces: %s"
        % format(len(keywords) * search_spelling_count(), ","))
    say("RESULT: PASS")
    return 0


def mode_dry_run(options):
    """Show exactly what a run would do.  Writes nothing at all.

    Opens no page, downloads nothing, touches no state and saves no student.
    Safe to run at any time, including while other copies are working.
    """
    setup_logging()
    try:
        cycle = read_keyword_cycle()
        keywords = keyword_order(cycle)       # the order the next run uses
        searches = build_search_list(say_total=False, keywords=keywords)
    except BaseKeywordError as exc:
        say("ERROR: %s" % exc)
        return 1
    if not searches:
        say("ERROR: rank %d of %d has no cities to search." % (rank, num_sys))
        say("num_sys (%d) is large enough that this rank's slice is empty."
            % num_sys)
        say("Lower num_sys, or give this folder a lower rank.")
        return 2

    all_cities = list(cities())
    mine = get_cities_by_rank(all_cities, rank, num_sys)

    say("=" * 60)
    say("Plan Preview -- nothing will be written")
    say("=" * 60)
    say("Keyword source      : %s (%s)"
        % (KEYWORD_SOURCE_FILE, KEYWORD_SOURCE_DICT))
    problems = base_keyword_problems()
    say("                      %s"
        % ("VALID" if not problems
           else "%d PROBLEM(S) -- run --validate-keywords" % len(problems)))
    say("")
    say("rank      = %d" % rank)
    say("num_sys   = %d" % num_sys)
    # Where the pair came from, so "this laptop is crawling the wrong
    # cities" is one line to diagnose instead of a hunt through two files
    # and the environment.
    say("            (from %s)" % RANK_SOURCE)
    say("")
    say("Cities (all copies) : %d" % len(all_cities))
    say("Cities (this copy)  : %d" % len(mine))
    say("Keywords            : %d" % len(keywords))
    say("Searches this copy  : %s" % format(len(searches), ","))
    say(describe_keyword_cycle(cycle, keywords))
    say("Night stop          : %s" % night_pause_status())
    say("Work / break        : %s" % work_break_status())
    say("")

    current = START
    if os.path.exists(C.CURRENT_ROW_FILE):
        current = read_current_row(len(searches))
    say("Resume point        : search %s of %s"
        % (format(current + 1, ","), format(len(searches), ",")))
    if searches:
        city, spelling, keyword = searches[min(current, len(searches) - 1)]
        say("Next search         : %s" % make_unit(city, spelling,
                                                   keyword).search_query)
    say("")
    say("First 5 cities      : %s"
        % ", ".join(city.canonical for city in mine[:5]))
    say("Last 5 cities       : %s"
        % ", ".join(city.canonical for city in mine[-5:]))
    say("")
    say("-" * 60)
    say("WHAT EVERY COPY GETS")
    say("-" * 60)
    for other in range(1, num_sys + 1):
        share = get_cities_by_rank(all_cities, other, num_sys)
        units = sum(len(city.search_aliases) for city in share) * len(keywords)
        say("  rank %-3d : %4d cities, %s searches%s"
            % (other, len(share), format(units, ","),
               "   <-- this folder" if other == rank else ""))
    covered = sum(len(get_cities_by_rank(all_cities, r, num_sys))
                  for r in range(1, num_sys + 1))
    say("")
    say("  %d of %d cities covered%s"
        % (covered, len(all_cities),
           "" if covered == len(all_cities) else "   *** SEARCH_CITIES ARE MISSING"))
    idle = [r for r in range(1, num_sys + 1)
            if not get_cities_by_rank(all_cities, r, num_sys)]
    if idle:
        say("")
        say("  NOTE: rank(s) %s get no cities at num_sys=%d."
            % (", ".join(str(r) for r in idle), num_sys))
        say("        No city is lost -- the earlier ranks cover them all --")
        say("        but those copies would have nothing to do.  Lower")
        say("        num_sys to put every copy to work.")
    say("=" * 60)
    say("Nothing was written.")
    return 0


def mode_retry_current_search(options):
    """Re-collect the search this folder is currently on.

    For the case where a search produced a wrong or truncated XLSX: the
    collected file is set aside and the state is cleared so the next run
    crawls that same search again from scratch.

    DELIBERATELY UNTOUCHED

        current_row.json          this is a RETRY of the current search,
                                  not a rewind of the whole list
        students.db               already-saved people stay, so re-collection
                                  skips them as duplicates and only genuinely
                                  new people cost anything
        history/identities.db
        the Playwright profile
        next_selectors, transient failures, skipped accounts

    The state is written BEFORE the XLSX is moved.  Either order is safe -- a
    missing file with a "complete" flag also re-collects -- but the state is
    the authoritative record, so it changes first.
    """
    setup_logging(options.verbose)

    say("=" * 60)
    say("RETRY CURRENT SEARCH %s"
        % ("" if options.apply else "(REPORT ONLY)"))
    say("=" * 60)

    if not os.path.exists(C.STATE_FILE):
        say("There is no state at %s" % C.STATE_FILE)
        say("There is nothing to retry; run the scraper normally.")
        return 0

    state = StateManager()
    current = state.describe_current()
    if not current["search_query"]:
        say("No current search is recorded.")
        say("There is nothing to retry; run the scraper normally.")
        return 0

    xlsx = current["collection_file"]
    if not xlsx and current["slug"]:
        xlsx = os.path.join(C.KEYWORD_DOWNLOAD_DIR, "%s.xlsx" % current["slug"])
    partial = (xlsx.replace(".xlsx", ".partial.xlsx")
               if xlsx.endswith(".xlsx") else "")
    files = [path for path in (xlsx, partial) if path and os.path.exists(path)]

    saved_here = 0
    store_total = 0
    if os.path.exists(C.STORE_DB):
        store = open_store(C.STORE_DB, read_only=True)
        store_total = store.count()
        saved_here = sum(1 for row in store.iter_records()
                         if row["search_query"] == current["search_query"])
        store.close()

    # ------------------------------------------------------------- report --
    say("Rank                       : %d/%d" % (rank, num_sys))
    say("Folder                     : %s" % C.PROJECT_ROOT)
    say("")
    say("Search                     : %s" % current["search_query"])
    say("  state                    : %s" % current["state"])
    say("  canonical city           : %s%s"
        % (current["canonical_city"],
           "" if current["search_city"] == current["canonical_city"]
           else ' (searched as "%s")' % current["search_city"]))
    say("  base keyword             : %s" % current["base_keyword"])
    say("  position                 : search %d (current_row.json)"
        % (current["unit_index"] + 1))
    say("")
    say("Current collection status   : %s"
        % ("complete" if current["collection_complete"] else "incomplete"))
    say("Current XLSX                : %s"
        % (os.path.basename(xlsx) if xlsx else "(none recorded)"))
    say("                              %s"
        % ("present on disk" if xlsx and os.path.exists(xlsx)
           else "not on disk"))
    if partial and os.path.exists(partial):
        say("Partial XLSX                : %s (present)"
            % os.path.basename(partial))
    say("Current collected rows      : %d" % current["collection_rows"])
    say("Current processing position : row %d of %d%s"
        % (current["next_row_index"], current["total_rows"],
           " (complete)" if current["processing_complete"] else ""))
    say("Phase                       : %s" % current["phase"])
    say("")
    say("People already saved from this search : %d" % saved_here)
    say("People in this folder\'s store in total : %d" % store_total)
    say("")

    # ------------------------------------------------------------ the plan --
    say("What --apply would do:")
    say("  1. reset ONLY this search\'s collection and processing state,")
    say("     in one atomic write, through StateManager")
    if files:
        say("  2. MOVE %d file(s) aside into %s/"
            % (len(files), os.path.relpath(C.RETRY_BACKUP_DIR,
                                           C.PROJECT_ROOT)))
        for path in files:
            say("       %s" % os.path.basename(path))
    else:
        say("  2. nothing to move; no XLSX is on disk")
    say("  3. leave everything else exactly as it is")
    say("")
    say("Preserved: current_row.json, the search identity,")
    say("           students.db (%d person/people), the history index,"
        % store_total)
    say("           current_data.xlsx and every cycle file,")
    say("           the Playwright profile, learned pagination selectors,")
    say("           transient failures, skipped accounts.")
    say("")
    say("The %d person/people already saved from this search stay in the "
        "store," % saved_here)
    say("so re-collection skips them as duplicates.")

    if not options.apply:
        say("")
        say("This was a report. Nothing has been changed.")
        say("Re-run with --apply to do it.")
        return 0

    # ----------------------------------------------------------- do it -----
    try:
        say("")
        unit = state.reset_for_recollection()
        say("[RESET] collection and processing cleared for %s"
            % unit.get("search_query", ""))

        moved = []
        if files:
            os.makedirs(C.RETRY_BACKUP_DIR, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            for path in files:
                # Second-resolution stamps collide when a retry is repeated
                # quickly, and an archive that overwrites the evidence it was
                # created to keep is worse than no archive.
                target = os.path.join(
                    C.RETRY_BACKUP_DIR,
                    "%s.%s" % (os.path.basename(path), stamp))
                serial = 1
                while os.path.exists(target):
                    target = os.path.join(
                        C.RETRY_BACKUP_DIR,
                        "%s.%s-%d" % (os.path.basename(path), stamp, serial))
                    serial += 1
                try:
                    shutil.move(path, target)
                    moved.append(target)
                    say("[ARCHIVE] %s -> %s/"
                        % (os.path.basename(path),
                           os.path.basename(C.RETRY_BACKUP_DIR)))
                except OSError as exc:
                    # The state already says "re-collect", so a file left
                    # behind is simply overwritten by the next collection.
                    say("[WARN] could not move %s: %s"
                        % (os.path.basename(path), exc))
        after = state.describe_current()
        say("")
        say("Now:")
        say("  collection complete : %s" % after["collection_complete"])
        say("  collection XLSX     : %s"
            % (after["collection_file"] or "(none)"))
        say("  collected rows      : %d" % after["collection_rows"])
        say("  processing position : row %d of %d"
            % (after["next_row_index"], after["total_rows"]))
        say("  phase               : %s" % after["phase"])
        say("  search              : %s (unchanged)" % after["search_query"])
        say("  files archived      : %d" % len(moved))
        say("")
        say("The next run will re-collect %r from GitHub."
            % after["search_query"])
        say("Run: python3 student_scraper.py")
    except NoCurrentUnit as exc:
        say("ERROR: %s" % exc)
        return 1
    return 0




def mode_validate_output():
    setup_logging()
    say("Validating %s" % C.FINAL_WORKBOOK_PATH)
    say("=" * 60)
    ok, problems, summary = validate_output()
    total = 0
    for name, columns, rows, verdict in summary:
        total += rows
        say("  %-24s %3d cols  %5d rows  %s"
            % (name, columns, rows, verdict))
    if not summary:
        say("  (the final workbook does not exist yet)")
        say("  Create it with: python3 student_scraper.py --export-final")
    else:
        say("  %-24s %8s  %5d rows total" % ("", "", total))
    say("")
    for lock in find_lock_files(C.FINAL_OUTPUT_DIR):
        say("  LOCK: %s" % describe_locks([lock])[0].strip())
    for problem in problems:
        say("  PROBLEM: %s" % problem)
    say("")
    say("Checked: 65-column schema and order, alias-column consistency,")
    say("         duplicate usernames, the %s row maximum."
        % format(C.XLSX_MAX_DATA_ROWS, ","))
    say("RESULT: %s" % ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


# ================================================================= main ====
def main():
    parser = argparse.ArgumentParser(
        description="Collect public GitHub student/person data. "
                    "Work division is set by `rank` and `num_sys` at the top "
                    "of this file, not on the command line.")

    reports = parser.add_argument_group("reports (write nothing)")
    reports.add_argument("--dry-run", action="store_true",
                         help="show this copy's share of the cities, what "
                              "every other copy would get, and the resume "
                              "point")
    reports.add_argument("--validate-keywords", action="store_true",
                         help="check the keyword list in %s and show it"
                              % KEYWORD_SOURCE_FILE)
    reports.add_argument("--validate-output", action="store_true",
                         help="check a workbook against the 65-column "
                              "contract")
    reports.add_argument("--check-extension", action="store_true",
                         help="diagnose the Instant Data Scraper setup")

    output = parser.add_argument_group("output")
    output.add_argument("--export-final", action="store_true",
                        help="merge this folder's store, plus any named with "
                             "--store, into one student_data.xlsx")
    output.add_argument("--store", action="append", default=[], metavar="DB",
                        help="another copy's students.db to merge "
                             "(repeatable)")
    output.add_argument("--list-stores", action="store_true",
                        help="show which stores would be merged, then stop")

    diagnostics = parser.add_argument_group(
        "diagnostics (one search, nothing collected)")
    diagnostics.add_argument("--diagnose-page", metavar="QUERY", default="",
                             help="classify ONE search page and print the "
                                  "evidence")
    diagnostics.add_argument("--diagnose-page2", metavar="QUERY", default="",
                             help="diagnose Instant Data Scraper on search "
                                  "page 2")
    diagnostics.add_argument("--diagnose-ids-rows", metavar="QUERY",
                             default="",
                             help="compare the GitHub DOM users to a fresh "
                                  "IDS export on one page")
    diagnostics.add_argument("--page", type=int, default=1,
                             help="which result page --diagnose-ids-rows "
                                  "should examine; default 1")
    diagnostics.add_argument("--fresh-profile", action="store_true",
                             help="use a throwaway browser profile (logged "
                                  "out) for a diagnostic")

    recovery = parser.add_argument_group("recovery")
    recovery.add_argument("--retry-current-search", action="store_true",
                          help="re-collect the search this copy is currently "
                               "on: resets ONLY that search's collection and "
                               "processing state and moves its XLSX aside "
                               "(report unless --apply). Saved people, "
                               "current_row.json and the cycle files are "
                               "untouched.")
    recovery.add_argument("--retry-deferred", action="store_true",
                          help="retry ONLY this copy's deferred profiles "
                               "(deferred_profiles.json), one conservative "
                               "attempt each; a success is saved and removed, "
                               "one still unavailable stays deferred. Never "
                               "blacklists.")
    recovery.add_argument("--apply", action="store_true",
                          help="turn a retry REPORT into an action")

    run = parser.add_argument_group("running")
    run.add_argument("--login", action="store_true",
                     help="open a visible browser once to log THIS folder's "
                          "profile into GitHub")
    run.add_argument("--show", action="store_true",
                     help="run with a visible browser window")
    run.add_argument("--limit", type=int, default=0,
                     help="stop after N saved records (testing)")
    run.add_argument("--process-existing", action="store_true",
                     help="process already-downloaded keyword XLSX files "
                          "without running the extension")
    run.add_argument("--accept-stable-pagination-mismatch", action="store_true",
                     help="accept a validated page-by-page result set whose "
                          "unique count differs from GitHub's stated total, "
                          "ONLY after two independent fully-validated crawls "
                          "produce the identical identity set. Default "
                          "(STRICT) stops and preserves staging instead.")
    run.add_argument("--max-pages", type=int, default=None,
                     help="stop the extension crawl after N pages (testing "
                          "only; production crawls to the end)")
    run.add_argument("--verbose", action="store_true",
                     help="also log to stderr")
    options = parser.parse_args()

    signal.signal(signal.SIGINT, _request_stop)
    signal.signal(signal.SIGTERM, _request_stop)

    # Everything this run creates is owned by this user alone.
    #
    # A collected row is a real student's name, email, phone and WhatsApp
    # number.  Those files were taking the default umask -- students.db and
    # current_data.xlsx came out 0664, readable by every account on the
    # machine, which does not matter much on one laptop and matters on a
    # shared server.  github_auth.json was already chmod 0600 by hand; this
    # is the same rule applied to the data instead of only to the login.
    os.umask(0o077)

    # ---- reports and maintenance: nothing is written ---------------------
    if options.validate_keywords:
        return mode_validate_keywords()
    if options.validate_output:
        return mode_validate_output()
    if options.dry_run:
        return mode_dry_run(options)
    if options.retry_current_search:
        return mode_retry_current_search(options)
    if options.list_stores or options.export_final:
        setup_logging(options.verbose)
        argv = ["--store=%s" % path for path in options.store]
        if options.list_stores:
            argv.append("--list-stores")
        return mode_export_final(argv)
    if options.diagnose_page:
        setup_logging(options.verbose)
        return asyncio.run(mode_diagnose_page(options.diagnose_page))
    if options.diagnose_page2:
        setup_logging(options.verbose)
        return asyncio.run(mode_diagnose_page2(options.diagnose_page2,
                                               options.fresh_profile))
    if options.diagnose_ids_rows:
        setup_logging(options.verbose)
        return asyncio.run(mode_diagnose_ids_rows(options.diagnose_ids_rows,
                                                  options.page,
                                                  options.fresh_profile))

    C.ensure_runtime_dirs()

    if options.login:
        setup_logging(options.verbose)
        return asyncio.run(mode_login(show=True))
    if options.check_extension:
        setup_logging(options.verbose)
        return asyncio.run(mode_check_extension(show=options.show))

    setup_logging(options.verbose)

    # Work out this copy's share before Chromium starts, so a bad rank /
    # num_sys combination fails in a second rather than after a browser
    # launch.
    if rank < 1 or num_sys < 1 or rank > num_sys:
        say("ERROR: rank must be between 1 and num_sys.")
        say("       This file has rank = %d and num_sys = %d." % (rank, num_sys))
        say("       Edit the two lines at the top of student_scraper.py.")
        return 2
    try:
        # Same searches, in this lap's keyword order.
        searches = build_search_list(
            keywords=keyword_order(read_keyword_cycle()))
    except BaseKeywordError as exc:
        say("ERROR: %s" % exc)
        return 2

    try:
        existing = load_existing_master()
    except Exception as exc:
        say("ERROR: cannot read %s: %s" % (master_path(), exc))
        say("       It is the record of everyone already collected; refusing "
            "to start rather than re-collect them all.")
        return 2
    # The banner reports the index a few lines below; printing it here too
    # just said the same thing twice.
    workflow = Workflow(options, searches, existing)

    async def runner():
        try:
            if options.retry_deferred:
                return await workflow.retry_deferred_only()
            return await workflow.run()
        finally:
            await workflow.close()

    try:
        return asyncio.run(runner())
    except GlobalThrottle as exc:
        # Distinct from an ordinary stop: the account/session is throttled.
        log.info("global throttle: %s", exc)
        say("")
        say("Progress saved (%s)." % os.path.basename(C.CURRENT_ROW_FILE))
        say("Keyword XLSX preserved. No other copy was touched.")
        say("Wait for GitHub's limit to clear, then run the same command.")
        return 4
    except Stopped as exc:
        log.info("stopped: %s", exc)
        say("")
        say("Interrupted safely.")
        say("Progress saved (%s)." % os.path.basename(C.CURRENT_ROW_FILE))
        say("Keyword XLSX preserved. No other copy was touched.")
        say("Run the same command to resume.")
        return 130
    except KeyboardInterrupt:
        say("")
        say("Interrupted safely.")
        say("Progress saved. Keyword XLSX preserved.")
        return 130
    except Exception as exc:
        # Ctrl+C goes to the whole process group, so Playwright's driver dies
        # at once while this code is still between two safe points.  The next
        # browser call then fails -- "extension service worker is not
        # available", "Connection closed while reading from the driver" --
        # and that is a CONSEQUENCE of the stop, not a new fault.  Printing a
        # traceback for it hides the thing the operator needs to know: the
        # run stopped cleanly and the position was saved.
        #
        # Only when NO stop was asked for is it a real crash, and then it is
        # re-raised untouched.
        if not STOP_REQUESTED:
            raise
        log.info("stopped while the browser was closing: %s", exc)
        say("")
        say("Interrupted safely.")
        say("Progress saved (%s)." % os.path.basename(C.CURRENT_ROW_FILE))
        say("Keyword XLSX preserved. No other copy was touched.")
        say("Run the same command to resume.")
        return 130



# ==========================================================================
# EXPORT STUDENTS
# ==========================================================================
# Merge every store into ONE student_data.xlsx.
#
#     runtime/rank_001/students.db          this campaign, per rank
#     runtime/rank_002/students.db
#     runtime/archive/<campaign>/rank_*/    campaigns already closed
#     imported/legacy_students.db           the pre-redesign workbooks
#                     |
#                     v
#         global de-duplication by identity
#                     |
#                     v
#     Student Data original/student_data.xlsx
#
# WHY THE EXPORT IS A SEPARATE STEP
# ---------------------------------
# 80 ranks across 8 laptops must not open one workbook.  openpyxl rewrites the
# entire file on every save, so two concurrent writers means one of them loses
# its rows -- and a third reading at the same time can see a half-written file.
# Ranks therefore write to their own database and the workbook is rendered from
# those databases when somebody asks for it.
#
# THIS IS ALSO WHERE CROSS-LAPTOP DUPLICATES ARE RESOLVED
# -------------------------------------------------------
# Rank-local SQLite cannot stop rank 4 on laptop 1 and rank 63 on laptop 7 from
# both collecting github.com/student123 -- they cannot see each other.  Both
# rows exist, and the merge here is what collapses them.  That is the honest
# limit of the current design: duplicate PREVENTION is per rank, duplicate
# RESOLUTION is global and happens now.
#
# To merge stores from other laptops, copy their runtime/rank_NNN/students.db
# files across and pass them with --store.
#
# DETERMINISTIC CONFLICT RESOLUTION
# ---------------------------------
# When two stores hold the same person, the winner is chosen by rule, not by
# whichever was read first, so the same inputs always produce the same output:
#
#     1  more populated fields wins        the more complete record
#     2  earlier collected_at wins         the original observation
#     3  lower rank wins                   a stable, arbitrary tie-break
# --------------------------------------------------------------------------
#!/usr/bin/env python3






# ================================================================ discovery ==
def discover_stores(store_db=None, extra=None, include_imported=True):
    """Every store to merge, labelled.  [(label, path)] in a stable order.

    One folder is one rank, so this folder contributes exactly one store.
    Every OTHER copy's store is named explicitly with --store, which is how
    the stores from several laptops are merged into one workbook:

        python3 export_students.py --store /media/usb/rank2/students.db \
                                   --store /media/usb/rank3/students.db

    Being explicit matters here.  Scanning for stores would silently produce
    a workbook missing whichever laptop had not been copied over yet, and a
    merge that quietly leaves a laptop out looks exactly like a merge that
    worked.
    """
    store_db = store_db or C.STORE_DB
    found = []

    if include_imported and os.path.exists(C.LEGACY_IMPORT_DB):
        found.append(("imported", C.LEGACY_IMPORT_DB))

    if os.path.exists(store_db):
        found.append(("this folder", store_db))

    for path in (extra or []):
        if os.path.exists(path):
            found.append(("extra/%s" % os.path.basename(path), path))
        else:
            raise SystemExit("--store %s does not exist" % path)
    return found


# ==================================================================== merge ==
def _better(candidate, incumbent):
    """True when `candidate` should replace `incumbent`.  Pure ordering."""
    if candidate["populated_fields"] != incumbent["populated_fields"]:
        return candidate["populated_fields"] > incumbent["populated_fields"]
    if candidate["collected_at"] != incumbent["collected_at"]:
        return candidate["collected_at"] < incumbent["collected_at"]
    return candidate["rank"] < incumbent["rank"]


def merge(stores):
    """Globally de-duplicated records, plus a report.

    Keyed by normalised username first, then profile URL, then real email, so
    the same person reached by three different routes collapses to one row.
    """
    by_username, alias = {}, {}
    report = {"per_store": [], "read": 0, "duplicates": 0, "conflicts": 0,
              "unusable": 0}

    for label, path in stores:
        store = open_store(path, read_only=True)
        count = 0
        for row in store.iter_records():
            count += 1
            report["read"] += 1
            keys = identity_keys(
                username=row["github_username_key"],
                profile_url=row["profile_url_key"],
                email=row["email_key"])
            primary = keys["github_username_key"]
            if not primary:
                report["unusable"] += 1
                continue
            # Follow a secondary key to an identity already seen.
            for secondary in (keys["profile_url_key"], keys["email_key"]):
                if secondary and secondary in alias:
                    primary = alias[secondary]
                    break
            row = dict(row, source_store=label)
            existing = by_username.get(primary)
            if existing is None:
                by_username[primary] = row
            else:
                report["duplicates"] += 1
                if _better(row, existing):
                    report["conflicts"] += 1
                    by_username[primary] = row
            for secondary in (keys["profile_url_key"], keys["email_key"]):
                if secondary:
                    alias.setdefault(secondary, primary)
        store.close()
        report["per_store"].append((label, count))

    # Sorted by identity so two exports of the same data are byte-comparable.
    ordered = [by_username[key] for key in sorted(by_username)]
    report["unique"] = len(ordered)
    return ordered, report


# =================================================================== export ==
def export(path=None, store_db=None, extra=None,
           include_imported=True, bundle_stores=None, out=say):
    """Do the whole export.  Returns a report dict.

    bundle_stores, when given, is a pre-validated [(label, path)] list from
    device bundles; it is merged ALONGSIDE this folder's own store.  The
    existing dedupe/conflict/atomic-write path is reused unchanged.
    """
    path = path or C.FINAL_WORKBOOK_PATH
    if bundle_stores is not None:
        stores = list(bundle_stores)
        if include_imported and os.path.exists(C.LEGACY_IMPORT_DB):
            stores.insert(0, ("imported", C.LEGACY_IMPORT_DB))
    else:
        stores = discover_stores(store_db, extra, include_imported)
    if not stores:
        out("No stores found. Nothing to export.")
        out("This folder writes students.db as it collects; name another")
        out("copy's store with --store /path/to/students.db.")
        return {"stores": [], "unique": 0, "written": 0, "path": path}

    out("Stores to merge:")
    for label, store_path in stores:
        out("  %-28s %s" % (label, store_path))
    out("")

    # Refuse before doing any work if the target is open in LibreOffice.
    guard_locks(os.path.dirname(path), say=out)

    records, report = merge(stores)
    out("Rows read            : %d" % report["read"])
    out("Unique students      : %d" % report["unique"])
    out("Duplicates collapsed : %d" % report["duplicates"])
    out("  of which a better record replaced an earlier one: %d"
        % report["conflicts"])
    if report["unusable"]:
        out("Rows with no usable GitHub identity (skipped): %d"
            % report["unusable"])
    out("")

    if report["unique"] > C.XLSX_WARN_DATA_ROWS:
        out("WARNING: %d rows is close to the xlsx worksheet maximum of %d."
            % (report["unique"], C.XLSX_MAX_DATA_ROWS))
        out("         The store remains the complete dataset; plan to export")
        out("         by state or to CSV before the limit is reached.")
        out("")

    written_path, count = write_final(
        [row["record"] for row in records], path)
    out("Wrote %d row(s) to %s" % (count, written_path))

    ok, problems, summary = validate_output(written_path)
    for name, columns, rows, verdict in summary:
        out("  %-24s %3d cols  %5d rows  %s" % (name, columns, rows, verdict))
    for problem in problems:
        out("  PROBLEM: %s" % problem)
    out("RESULT: %s" % ("PASS" if ok else "FAIL"))

    report.update({"stores": stores, "written": count, "path": written_path,
                   "valid": ok, "problems": problems})
    return report


def refresh_final_excel(say_fn=say, **export_kwargs):
    """Regenerate the ONE visible workbook from every LOCAL rank store, safely.

    This is the automatic refresh called when a rank reaches COMPLETE.  It
    reuses export() unchanged -- same store discovery, same global de-dedupe,
    the same atomic 65-column write to Student Data original/student_data.xlsx.
    It NEVER raises: on any failure (a locked workbook, a disk error, ...) the
    previous student_data.xlsx is left exactly as it was and the caller's
    completed rank state is safe.  Prints the short [EXCEL] status lines.
    Returns True only when a fresh workbook was written.
    """
    try:
        report = export(out=lambda *_a, **_k: None, **export_kwargs)
    except WorkbookLocked:
        say_fn("[EXCEL] Export failed; previous student_data.xlsx preserved.")
        say_fn("        (the workbook is open elsewhere; close it and rerun "
               "python3 export_students.py)")
        return False
    except Exception as exc:                     # never lose completed state
        say_fn("[EXCEL] Export failed; previous student_data.xlsx preserved.")
        say_fn("        (%s: %s)" % (type(exc).__name__, exc))
        return False
    if not report.get("stores"):
        say_fn("[EXCEL] No rank stores on this machine yet; nothing to export.")
        return False
    if not report.get("valid", True):
        say_fn("[EXCEL] Export failed validation; previous student_data.xlsx "
               "preserved.")
        return False
    say_fn("[EXCEL] Updated: %s" % report["path"])
    say_fn("[EXCEL] Students: %d" % report.get("written", 0))
    say_fn("[EXCEL] Columns: %d" % len(C.FINAL_COLUMNS))
    return True


def _export_from_bundles(options):
    """Validate the supplied bundles together, then merge them all.

    Refuses to mix incompatible campaigns/plans, a corrupted store, or a rank
    supplied twice -- combine_bundles raises BundleError on any of those.
    """
    try:
        store_paths, report = combine_bundles(options.bundle)
    except BundleError as exc:
        say("ERROR: %s" % exc)
        return 2
    say("=" * 60)
    say("MASTER EXPORT FROM DEVICE BUNDLES")
    say("=" * 60)
    say("Campaign          : %s" % report["campaign_id"])
    say("Plan fingerprint  : %s" % (report["plan_fingerprint"] or "")[:16])
    say("Total ranks       : %s" % report["total_ranks"])
    say("Bundles           : %d" % report["bundles"])
    say("Ranks covered     : %d of %d"
        % (len(report["ranks_covered"]), report["total_ranks"] or 0))
    if report["ranks_missing"]:
        say("Ranks MISSING     : %s"
            % _compact_ranks(report["ranks_missing"]))
        say("  (their students will be absent from this export; supply the "
            "missing device bundle(s) to include them)")
    say("")
    bundle_stores = [("bundle/rank_%03d" % rank, path)
                     for rank, path in sorted(store_paths.items())]
    try:
        result = export(options.out, bundle_stores=bundle_stores,
                        include_imported=not options.no_imported)
    except WorkbookLocked as exc:
        say("")
        say("ERROR: the final workbook is open elsewhere.")
        say(str(exc))
        return 1
    except RowLimitExceeded as exc:
        say("")
        say("ERROR: %s" % exc)
        return 1
    return 0 if result.get("valid", True) else 1


def mode_export_final(argv=None):
    parser = argparse.ArgumentParser(
        description="Merge this folder's store, plus any named with --store, into one student_data.xlsx.")
    parser.add_argument("--out", default=None,
                        help="workbook to write (default: %s)"
                             % C.FINAL_WORKBOOK_PATH)
    parser.add_argument("--store", action="append", default=[],
                        help="an extra students.db to merge, e.g. one copied "
                             "from another laptop (repeatable)")
    parser.add_argument("--no-imported", action="store_true",
                        help="skip imported/legacy_students.db")
    parser.add_argument("--list-stores", action="store_true",
                        help="show which stores would be merged, then stop")
    parser.add_argument("--bundle", action="append", default=[],
                        help="a device result bundle to merge (repeatable); "
                             "checksums, campaign and plan fingerprint are "
                             "verified before merging")
    options = parser.parse_args(argv)

    if options.bundle:
        return _export_from_bundles(options)

    if options.list_stores:
        stores = discover_stores(extra=options.store,
                                 include_imported=not options.no_imported)
        if not stores:
            say("No stores found.")
            return 0
        for label, path in stores:
            store = open_store(path, read_only=True)
            say("  %-28s %6d student(s)  %s" % (label, store.count(), path))
            store.close()
        return 0

    try:
        report = export(options.out, extra=options.store,
                        include_imported=not options.no_imported)
    except WorkbookLocked as exc:
        say("")
        say("ERROR: the final workbook is open elsewhere.")
        say(str(exc))
        say("Close it and run the export again. The lock was NOT removed.")
        say("[EXCEL] Export failed; previous student_data.xlsx preserved.")
        return 1
    except RowLimitExceeded as exc:
        say("")
        say("ERROR: %s" % exc)
        say("[EXCEL] Export failed; previous student_data.xlsx preserved.")
        return 1
    if report.get("stores") and report.get("valid", True):
        say("")
        say("[EXCEL] Updated: %s" % report["path"])
        say("[EXCEL] Students: %d" % report.get("written", 0))
        say("[EXCEL] Columns: %d" % len(C.FINAL_COLUMNS))
    return 0 if report.get("valid", True) else 1


# ==========================================================================
# DIAGNOSE PAGE
# ==========================================================================
# Classify ONE GitHub search page and print the evidence.  Collects nothing.
#
#     python3 diagnose_page.py "College Student Visakhapatnam"
#
# Opens the page in this rank's browser, runs the page probe, prints every
# signal the classifier uses and the verdict it reaches.  It never opens the
# extension, never downloads anything, never touches state and never saves a
# student -- so it settles "is this really empty, or is detection wrong?"
# without running a scrape.
# --------------------------------------------------------------------------
#!/usr/bin/env python3




async def mode_diagnose_page(query):
    browser = BrowserManager(headless=False)
    await browser.start(with_extension=False)
    try:
        page = await browser.page()
        url = C.SEARCH_URL.format(query=urllib.parse.quote_plus(query))
        print("URL: %s" % url)
        await page.goto(url, wait_until="domcontentloaded",
                        timeout=C.NAVIGATION_TIMEOUT)
        login = await browser.github_login_name(page)
        print("Signed in as: %s" % (login or "NOT SIGNED IN"))

        scraper = InstantDataScraper(browser, state=None, say=print)
        state = await scraper.capture_page_state(page, reason="diagnose")
        verdict, why = classify_search_page(state)

        print("")
        print("=" * 60)
        print("VERDICT: %s" % verdict)
        print("WHY    : %s" % why)
        print("=" * 60)
        # How many pages does this search REALLY have?  The header count and
        # the pagination strip answer it directly, and it is the number the
        # crawl must be checked against.
        import re as _re
        body = "%s %s" % (state.get("bodyHead") or "",
                          state.get("bodyTail") or "")
        users = _re.search(r"Users\s+([\d.,k]+)\s*\(", body, _re.I)
        strip = _re.search(r"(Previous\s*[\d\s]+Next|\bNext\b)", body, _re.I)
        print("")
        print("PAGE COUNT EVIDENCE")
        print("  reported total users     : %s"
              % (users.group(1) if users else "not stated on the page"))
        print("  pagination strip         : %s"
              % (strip.group(1).strip() if strip else "none"))
        print("  browser page number      : %s" % state.get("browserPage"))
        if users:
            try:
                total = int(users.group(1).replace(",", "").replace(".", ""))
                print("  expected pages (10/page) : %d" % -(-total // 10))
                print("  -> a complete crawl must collect %d row(s)" % total)
            except ValueError:
                pass
        print("")
        for key in ("url", "title", "isUsersSearch", "loginForm", "loginCopy",
                    "blocked_marker", "saysError", "noResults",
                    "zeroUsersCopy", "zeroUsersCount",
                    "resultsListPresent", "resultsListChildren",
                    "profileLinkCount", "resultProfileLinkCount",
                    "hasNext", "nextByShape", "nextByText",
                    "paginationPresent", "nextDisabled", "browserPage",
                    "bodyChars"):
            print("  %-24s %s" % (key, state.get(key)))
        print("")
        print("  body (first 400 chars):")
        print("  %s" % (state.get("bodyHead") or "")[:400])
        print("")
        print("Interpretation:")
        if verdict == "zero-results":
            print("  GENUINE ZERO RESULTS. The search unit will complete with")
            print("  no candidates and the rank will advance.")
        elif verdict in ("results-single-page", "results-multi-page"):
            print("  USERS ARE PRESENT. If the previous run called this empty,")
            print("  that was the detection bug, now fixed.")
        else:
            print("  NOT A USABLE RESULTS PAGE. Collection would fail and")
            print("  preserve, which is correct for this state.")
        print("")
        print("Full capture as JSON:")
        print(json.dumps(state, indent=2, ensure_ascii=False, default=str))
    finally:
        await browser.close()
    return 0


# ==========================================================================
# DIAGNOSE PAGE2
# ==========================================================================
# Why does Instant Data Scraper lose its table on GitHub search page 2?
#
#     python3 diagnose_page2.py "University Student Visakhapatnam"
#     python3 diagnose_page2.py "University Student Visakhapatnam" --rank 1
#
# WHAT THIS ANSWERS
# -----------------
# The one question that decides the fix:
#
#     if IDS is freshly initialised while the browser is ALREADY on page 2,
#     can it detect those 9 users?
#
#     YES -> IDS loses its table across its own navigation.  A page-by-page
#            architecture (Playwright navigates, IDS is reinitialised per page,
#            page exports are merged) will work.
#     NO  -> page 2's markup or IDS's table configuration is the problem, and
#            page-by-page merging would fail the same way.
#
# It also captures the fifteen specific observations asked for: URL,
# readyState, page-2 result count, container rendering, skeleton/loading state,
# DOM replacement across navigation, popup status text, IDS rows, IDS pages,
# IDS error, "Table not found", popup/tab association, tab id, fresh-IDS
# detection, and how long page 2 needs before its DOM stops changing.
#
# WHAT IT DOES NOT DO
# -------------------
# Collects no students.  Writes no XLSX.  Touches no scraper state, no store,
# no history, no campaign.  It clicks "Next" once and reads the DOM.
#
# It DOES launch Chromium on the rank's persistent profile, because a
# logged-out page has different markup and the failure must be reproduced
# faithfully.  That writes cache and cookies exactly as any browsing session
# does; it changes no login and no scraper file.  Use --fresh-profile to use a
# throwaway profile instead, at the cost of being logged out.
# --------------------------------------------------------------------------
#!/usr/bin/env python3



# Stamp the results DOM, then see what survives navigation.  A surviving
# stamp means GitHub mutated the existing list; a vanished stamp means it
# replaced the nodes, which is what would invalidate a stored table path.
STAMP_JS = r"""
() => {
  const list = document.querySelector('[data-testid="results-list"]')
            || document.querySelector('mode_diagnose_page2');
  if (!list) return {stamped: false};
  list.setAttribute('data-diag-list', 'yes');
  let children = 0;
  Array.from(list.children).forEach((node, index) => {
    node.setAttribute('data-diag-row', String(index));
    children += 1;
  });
  window.__diagMutations = 0;
  const observer = new MutationObserver(records => {
    window.__diagMutations += records.length;
  });
  observer.observe(list.parentNode || document.body,
                  {childList: true, subtree: true});
  window.__diagObserver = observer;
  return {stamped: true, children: children};
}
"""

PROBE_JS = r"""
() => {
  const list = document.querySelector('[data-testid="results-list"]');
  const mode_diagnose_page2 = document.querySelector('mode_diagnose_page2');
  const scope = list || mode_diagnose_page2 || document.body;

  const reserved = new Set(['search','login','settings','topics','orgs',
      'marketplace','features','pricing','about','security','enterprise',
      'notifications','issues','pulls','codespaces','sponsors','explore',
      'trending','collections','events','new','organizations','users']);
  const CHROME = 'header, nav, footer, [role="banner"], [role="navigation"],'
               + ' [role="contentinfo"], .Header, .header, .footer';
  const users = new Set();
  scope.querySelectorAll('a[href]').forEach(a => {
    const m = (a.getAttribute('href') || '')
        .match(/^\/([A-Za-z0-9][A-Za-z0-9-]{0,38})$/);
    if (!m || reserved.has(m[1].toLowerCase())) return;
    if (a.closest(CHROME)) return;
    users.add(m[1]);
  });

  const skeletonSelectors = [
    '[class*="keleton"]', '[aria-busy="true"]', '[role="progressbar"]',
    'progress', '[class*="oading"]', '[data-testid*="loading"]',
  ];
  const skeletons = {};
  skeletonSelectors.forEach(s => {
    try { skeletons[s] = document.querySelectorAll(s).length; }
    catch (e) { skeletons[s] = -1; }
  });

  const text = (document.body ? document.body.innerText : '') || '';
  const total = text.match(/Users\s+([\d.,k]+)\s*\(/i);

  return {
    url: location.href,
    readyState: document.readyState,
    visibility: document.visibilityState,
    resultsListPresent: !!list,
    resultsListChildren: list ? list.children.length : 0,
    mainPresent: !!mode_diagnose_page2,
    userLinkCount: users.size,
    userLinks: Array.from(users).slice(0, 25),
    reportedTotalUsers: total ? total[1] : null,
    // Did the stamped nodes survive?  0 surviving rows with results on screen
    // means GitHub replaced the list wholesale.
    stampedListSurvived: !!document.querySelector('[data-diag-list]'),
    stampedRowsSurviving: document.querySelectorAll('[data-diag-row]').length,
    mutationsObserved: (typeof window.__diagMutations === 'number')
        ? window.__diagMutations : null,
    domSignature: (list || mode_diagnose_page2)
        ? (list || mode_diagnose_page2).innerHTML.length : 0,
    skeletons: skeletons,
    bodyChars: text.length,
  };
}
"""


def search_url_for(query):
    """The Users-search URL, built from the SAME template production uses.

    config.SEARCH_URL is the single source of the template; this formats it
    exactly as collect() does, so the diagnostic can never
    open a different URL from the one the scraper would.
    """
    return C.SEARCH_URL.format(query=urllib.parse.quote_plus(query))


def verify_search_landing(requested_url, final_url, page_state):
    """Did we actually land on the requested Users search?  (ok, reason).

    Pure, so it is unit-testable without a browser.  Rejects the homepage, a
    login wall, an error or rate-limit page, and any redirect away from
    /search?type=users -- all of which would otherwise be "diagnosed" as if
    they were the search results.
    """
    from student_scraper import PageState

    parsed = urllib.parse.urlparse(final_url or "")
    query = urllib.parse.parse_qs(parsed.query)

    if not (parsed.netloc or "").lower().endswith("github.com"):
        return (False, "left GitHub entirely: %s" % (final_url or "(none)"))
    if parsed.path.rstrip("/") in ("", "/"):
        return (False,
                "landed on the GitHub HOMEPAGE (%s), not the search"
                % (final_url or ""))
    if parsed.path.rstrip("/") in ("/login", "/session"):
        return (False, "redirected to the GitHub LOGIN page (%s)"
                % final_url)
    if parsed.path.rstrip("/") != "/search":
        return (False, "redirected away from /search to %s" % final_url)
    if query.get("type", [""])[0] != "users":
        return (False, "the search is not type=users: %s" % final_url)

    verdict, why = page_state
    if verdict in (PageState.LOGIN_REQUIRED, PageState.BLOCKED,
                   PageState.HTTP_ERROR, PageState.GITHUB_ERROR):
        return (False, "%s: %s" % (verdict, why))
    if verdict == PageState.UNKNOWN_INVALID:
        return (False, "the page did not render as a Users search (%s)" % why)
    if verdict == PageState.ZERO_RESULTS:
        return (False,
                "this search has ZERO users, so there is no page 2 to "
                "diagnose (%s)" % why)
    return (True, why)




class Diagnostic:
    def __init__(self, browser, scraper):
        self.browser = browser
        self.scraper = scraper
        self.report = {"captured_at": datetime.datetime.now().isoformat(
            timespec="seconds"), "phases": {}}

    # ------------------------------------------------------------ helpers --
    async def probe(self, page):
        try:
            return await page.evaluate(PROBE_JS)
        except Exception as exc:
            return {"probe_error": str(exc)}

    async def popup_state(self, popup, label=""):
        """Everything the extension popup is willing to tell us."""
        if popup is None or popup.is_closed():
            return {"popup": "closed"}
        pages, rows = await self.scraper.read_stats(popup)
        info = {
            "popup_url": popup.url,
            "ids_pages": pages,
            "ids_rows": rows,
            "ids_stats_text": await self.scraper._text(popup, C.SEL_STATS),
            "ids_status_text": await self.scraper._text(popup,
                                                       C.SEL_INSTRUCTIONS),
            "ids_error": await self.scraper._text(popup, C.SEL_ERROR),
            "ids_info": await self.scraper._text(popup, C.SEL_INFO),
            "has_start": await self.scraper._has(popup, C.SEL_START_CRAWL),
            "has_stop": await self.scraper._has(popup, C.SEL_STOP_CRAWL),
            "has_wrong_table": await self.scraper._has(popup,
                                                      C.SEL_WRONG_TABLE),
            "has_next_button": await self.scraper._has(popup,
                                                      C.SEL_NEXT_BUTTON),
            "has_xlsx": await self.scraper._has(popup, C.SEL_XLSX),
        }
        info["says_table_not_found"] = (
            "table not found" in (info["ids_error"] or "").lower())
        # The popup's own bound tab, from its query string.
        try:
            query = urllib.parse.parse_qs(
                urllib.parse.urlparse(popup.url).query)
            info["popup_bound_tabid"] = (query.get("tabid") or [None])[0]
            info["popup_bound_url"] = (query.get("url") or [None])[0]
        except Exception:
            pass
        if label:
            info["label"] = label
        return info

    async def tab_id(self, page):
        try:
            return await self.browser.chrome_tab_id(page)
        except Exception as exc:
            return "unavailable (%s)" % exc

    # -------------------------------------------------------------- phases --
    async def phase_open_search(self, page, query):
        """Land on the requested Users search and PROVE it before going on.

        The login check comes FIRST because browser_manager's
        github_login_name() navigates the tab to https://github.com to read
        the meta tag.  Calling it after the search navigation is what made
        this diagnostic inspect the homepage.
        """
        say("")
        say("=" * 66)
        say("PHASE 0 - open the GitHub Users search")
        say("=" * 66)

        # Done first, and deliberately: this navigates the tab to github.com.
        login = await self.browser.github_login_name(page)
        say("Signed in as        : %s" % (login or "NOT SIGNED IN"))

        url = search_url_for(query)
        say("Search query        : %s" % query)
        say("Search URL          : %s" % url)
        await page.goto(url, wait_until="domcontentloaded",
                        timeout=C.NAVIGATION_TIMEOUT)
        try:
            await page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        # The SPA can still be swapping the results in after domcontentloaded.
        for _attempt in range(20):
            probe = await self.probe(page)
            if probe.get("resultsListPresent") or probe.get("userLinkCount"):
                break
            await asyncio.sleep(0.5)

        state = await self.scraper.capture_page_state(page, reason="page 1")
        verdict, why = classify_search_page(state)
        final_url = state.get("url") or page.url
        say("Final browser URL   : %s" % final_url)
        say("Page classification : %s (%s)" % (verdict, why))

        ok, reason = verify_search_landing(url, final_url, (verdict, why))
        self.report["phases"]["0_open_search"] = {
            "login": login, "query": query, "search_url": url,
            "final_url": final_url, "classification": verdict, "why": why,
            "landed": ok, "reason": reason,
        }
        if not ok:
            say("")
            say("STOPPING: %s" % reason)
            say("")
            say("Nothing was diagnosed. The page in the browser is not the")
            say("Users search this run asked for, so any capture from it")
            say("would describe the wrong page.")
            if "HOMEPAGE" in reason or "LOGIN" in reason:
                say("")
                say("If this is a sign-in problem, authenticate this rank")
                say("once and try again:")
                say("  python3 py --login"
                    % self.browser.context_paths.rank)
            return None
        say("")
        say("Confirmed: a valid GitHub Users results page.")
        return state

    async def phase_page_one(self, page, query, state):
        say("")
        say("=" * 66)
        say("PHASE A - page 1 baseline")
        say("=" * 66)
        verdict, why = classify_search_page(state)
        login = self.report["phases"]["0_open_search"]["login"]
        probe = await self.probe(page)
        stamp = await page.evaluate(STAMP_JS)
        tab = await self.tab_id(page)

        popup = await self.scraper.open_instant_data_scraper(page)
        ids = await self.popup_state(popup, "page 1, popup just opened")

        say("page classification : %s (%s)" % (verdict, why))
        say("browser page number : %s" % page_number(probe.get("url", "")))
        say("readyState          : %s" % probe.get("readyState"))
        say("results list        : present=%s children=%s"
            % (probe.get("resultsListPresent"),
               probe.get("resultsListChildren")))
        say("user links (results): %s" % probe.get("userLinkCount"))
        say("reported total users: %s" % probe.get("reportedTotalUsers"))
        say("chrome tab id       : %s" % tab)
        say("DOM stamped         : %s (%s row nodes)"
            % (stamp.get("stamped"), stamp.get("children")))
        say("")
        say("IDS on page 1       : pages=%s rows=%s" % (ids["ids_pages"],
                                                        ids["ids_rows"]))
        say("  status            : %s" % (ids["ids_status_text"] or "-"))
        say("  error             : %s" % (ids["ids_error"] or "(none)"))
        say("  Table not found   : %s" % ids["says_table_not_found"])
        say("  buttons           : start=%s stop=%s wrongTable=%s xlsx=%s"
            % (ids["has_start"], ids["has_stop"], ids["has_wrong_table"],
               ids["has_xlsx"]))
        say("  popup bound tabid : %s" % ids.get("popup_bound_tabid"))

        self.report["phases"]["A_page_one"] = {
            "login": login, "classification": verdict, "why": why,
            "probe": probe, "stamp": stamp, "tab_id": tab, "ids": ids,
        }
        return popup, probe

    async def phase_navigate(self, page, popup, before_probe):
        """Click Next once, then watch page 2 settle, second by second."""
        say("")
        say("=" * 66)
        say("PHASE B - navigate page 1 -> page 2 and watch it settle")
        say("=" * 66)
        target = None
        for selector in ('a[rel="next"]', 'a[aria-label="Next Page"]',
                         'a[aria-label="Next page"]', "a.next_page"):
            node = page.locator(selector)
            try:
                if await node.count() == 1 and await node.first.is_visible():
                    target = selector
                    break
            except Exception:
                continue
        if target is None:
            say("NO usable Next control on page 1 -- cannot continue.")
            self.report["phases"]["B_navigate"] = {"error": "no next control"}
            return None
        say('clicking %s' % target)
        await page.locator(target).first.click(timeout=10000)

        samples = []
        stable_since = None
        last_signature = None
        for tick in range(31):                      # up to ~15s at 0.5s
            await asyncio.sleep(0.5)
            probe = await self.probe(page)
            ids = await self.popup_state(popup)
            elapsed = round((tick + 1) * 0.5, 1)
            signature = (probe.get("url"), probe.get("resultsListChildren"),
                         probe.get("userLinkCount"),
                         probe.get("domSignature"))
            if signature == last_signature:
                if stable_since is None:
                    stable_since = elapsed
            else:
                stable_since = None
                last_signature = signature
            samples.append({"t": elapsed, "probe": probe, "ids": ids})
            if tick in (0, 1, 3, 5, 9, 19, 30):
                say("  t=%4.1fs page=%s ready=%-9s children=%-3s users=%-3s "
                    "IDS pages=%-3s rows=%-3s err=%s"
                    % (elapsed, page_number(probe.get("url", "")),
                       probe.get("readyState"),
                       probe.get("resultsListChildren"),
                       probe.get("userLinkCount"),
                       ids.get("ids_pages"), ids.get("ids_rows"),
                       (ids.get("ids_error") or "-")[:22]))
            if stable_since is not None and elapsed - stable_since >= 2.0:
                break

        final = samples[-1]["probe"] if samples else {}
        say("")
        say("page 2 URL              : %s" % final.get("url"))
        say("page 2 readyState       : %s" % final.get("readyState"))
        say("page 2 result children  : %s" % final.get("resultsListChildren"))
        say("page 2 user links       : %s" % final.get("userLinkCount"))
        say("reported total users    : %s" % final.get("reportedTotalUsers"))
        say("DOM stabilised after    : %s"
            % ("%.1fs" % stable_since if stable_since is not None
               else "did not stabilise within 15s"))
        say("stamped list survived   : %s  (stamped rows surviving: %s)"
            % (final.get("stampedListSurvived"),
               final.get("stampedRowsSurviving")))
        say("  -> GitHub %s the results DOM"
            % ("MUTATED" if final.get("stampedListSurvived")
               else "REPLACED"))
        say("mutations observed      : %s" % final.get("mutationsObserved"))
        skeletons = {k: v for k, v in (final.get("skeletons") or {}).items()
                     if v}
        say("loading/skeleton nodes  : %s" % (skeletons or "none"))
        ids_final = samples[-1]["ids"] if samples else {}
        say("")
        say("IDS after navigation    : pages=%s rows=%s"
            % (ids_final.get("ids_pages"), ids_final.get("ids_rows")))
        say("  status                : %s"
            % (ids_final.get("ids_status_text") or "-"))
        say("  error                 : %s"
            % (ids_final.get("ids_error") or "(none)"))
        say("  Table not found       : %s"
            % ids_final.get("says_table_not_found"))
        say("  wrongTable available  : %s"
            % ids_final.get("has_wrong_table"))
        say("  popup bound url       : %s"
            % ids_final.get("popup_bound_url"))
        say("  popup still open      : %s" % (not popup.is_closed()))
        say("  chrome tab id now     : %s" % await self.tab_id(page))

        self.report["phases"]["B_navigate"] = {
            "next_selector": target,
            "stable_after_seconds": stable_since,
            "samples": samples,
        }
        return samples

    async def phase_fresh_ids(self, page, old_popup, label, reload_first):
        """THE question: a fresh IDS on page 2 -- does it see the rows?"""
        say("")
        say("=" * 66)
        say("PHASE %s - fresh Instant Data Scraper on page 2%s"
            % ("C" if not reload_first else "D",
               " (after a reload)" if reload_first else ""))
        say("=" * 66)
        if old_popup is not None and not old_popup.is_closed():
            await old_popup.close()
            say("closed the page-1 popup")
        if reload_first:
            await page.reload(wait_until="domcontentloaded",
                              timeout=C.NAVIGATION_TIMEOUT)
            try:
                await page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass
            say("reloaded the tab (still %s)" % page.url)
        await asyncio.sleep(2)

        probe = await self.probe(page)
        say("page is on              : page %s, %s user link(s), %s children"
            % (page_number(probe.get("url", "")), probe.get("userLinkCount"),
               probe.get("resultsListChildren")))
        tab = await self.tab_id(page)
        popup = await self.scraper.open_instant_data_scraper(page)
        # Give the content script the same grace the production code gives it.
        detected = await self.scraper.wait_for_table(popup, timeout=20000)
        ids = await self.popup_state(popup, label)

        say("")
        say("FRESH IDS RESULT        : table detected = %s" % detected)
        say("  pages                 : %s" % ids["ids_pages"])
        say("  rows                  : %s" % ids["ids_rows"])
        say("  status                : %s" % (ids["ids_status_text"] or "-"))
        say("  error                 : %s" % (ids["ids_error"] or "(none)"))
        say("  Table not found       : %s" % ids["says_table_not_found"])
        say("  buttons               : start=%s wrongTable=%s xlsx=%s"
            % (ids["has_start"], ids["has_wrong_table"], ids["has_xlsx"]))
        say("  popup bound tabid/url : %s / %s"
            % (ids.get("popup_bound_tabid"), ids.get("popup_bound_url")))
        say("  chrome tab id         : %s" % tab)

        self.report["phases"]["%s_fresh_ids%s" % (
            "C" if not reload_first else "D",
            "_after_reload" if reload_first else "")] = {
            "detected": detected, "probe": probe, "ids": ids, "tab_id": tab}
        return popup, detected, ids, probe


async def mode_diagnose_page2(query, fresh_profile):
    if fresh_profile:
        import tempfile
        profile_dir = os.path.join(tempfile.mkdtemp(prefix="diag_profile_"),
                                   "pw_profile")
        say("Using a THROWAWAY profile at %s (you will be logged out)"
            % profile_dir)
    else:
        profile_dir = C.PROFILE_DIR
        say("Using this folder's profile: %s" % profile_dir)
        say("Chromium will write cache/cookies there, as any session does.")
        say("No scraper state, store or collected data is touched.")
    os.makedirs(profile_dir, exist_ok=True)

    browser = BrowserManager(headless=False, profile_dir=profile_dir)
    diagnostic = None
    try:
        await browser.start(with_extension=True)
        page = await browser.page()
        await asyncio.sleep(2)
        await browser.close_extra_pages(keep=(page,))
        scraper = InstantDataScraper(browser, state=None, say=say)
        diagnostic = Diagnostic(browser, scraper)

        landed = await diagnostic.phase_open_search(page, query)
        if landed is None:
            return 1
        popup, probe = await diagnostic.phase_page_one(page, query, landed)
        samples = await diagnostic.phase_navigate(page, popup, probe)
        if samples is None:
            return 1
        popup2, detected_spa, ids_spa, probe2 = \
            await diagnostic.phase_fresh_ids(page, popup, "fresh after SPA",
                                             reload_first=False)
        popup3, detected_reload, ids_reload, probe3 = \
            await diagnostic.phase_fresh_ids(page, popup2,
                                             "fresh after reload",
                                             reload_first=True)

        # ------------------------------------------------------- verdict --
        say("")
        say("=" * 66)
        say("VERDICT")
        say("=" * 66)
        expected = None
        total = probe.get("reportedTotalUsers")
        if total:
            try:
                expected = int(str(total).replace(",", ""))
            except ValueError:
                expected = None
        page2_users = probe2.get("userLinkCount") or 0
        say("GitHub reports total users    : %s" % (total or "not stated"))
        if expected:
            say("expected pages at 10/page     : %d" % -(-expected // 10))
        say("page 2 shows                  : %d user link(s)" % page2_users)
        say("fresh IDS after SPA nav       : detected=%s rows=%s"
            % (detected_spa, ids_spa["ids_rows"]))
        say("fresh IDS after a reload      : detected=%s rows=%s"
            % (detected_reload, ids_reload["ids_rows"]))
        say("")
        if detected_spa or detected_reload:
            say("=> A FRESH IDS CAN SEE PAGE 2.")
            say("   IDS loses its table across its OWN navigation, not")
            say("   because page 2 is undetectable.")
            say("   The page-by-page architecture is the right fix:")
            say("     Playwright navigates and synchronises,")
            say("     IDS is reinitialised per page and exports that page,")
            say("     page exports are merged into ONE keyword XLSX.")
            if detected_reload and not detected_spa:
                say("   NOTE: only the RELOADED page worked, so the per-page")
                say("         step must reload (or hard-navigate) rather")
                say("         than rely on the SPA transition.")
        else:
            say("=> A FRESH IDS CANNOT SEE PAGE 2 EITHER.")
            say("   Page-by-page merging would fail the same way.")
            say("   The problem is page 2's markup or IDS's table")
            say("   configuration, not navigation. Do NOT build the")
            say("   page-by-page fallback on this evidence.")
        say("")
        report_path = os.path.join(
            C.TEMP_DOWNLOAD_DIR,
            "page2_diagnosis_%s.json"
            % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        diagnostic.report["verdict"] = {
            "reported_total_users": total,
            "page2_user_links": page2_users,
            "fresh_ids_after_spa": bool(detected_spa),
            "fresh_ids_after_reload": bool(detected_reload),
        }
        with open(report_path, "w", encoding="utf-8") as handle:
            json.dump(diagnostic.report, handle, indent=2, default=str)
        say("Full capture written to:")
        say("  %s" % report_path)
        say("(that is a rank-local temp directory; it is not output data)")
        return 0
    finally:
        await browser.close()


# ==========================================================================
# DIAGNOSE IDS ROWS
# ==========================================================================
# Why does fresh IDS see fewer rows than the GitHub results DOM?
#
#     python3 diagnose_ids_rows.py "Engineering Student Visakhapatnam" --page 2
#     python3 diagnose_ids_rows.py "University Student Vijayawada"     --page 3
#     python3 diagnose_ids_rows.py "College Student Vijayawada"        --page 1
#
# For ONE page of a Users search this prints, side by side:
#
#     GitHub reported total
#     current page number
#     GitHub result-DOM user count, and every username/profile URL in the DOM
#     fresh IDS detected row count, and every username/profile URL IDS exported
#     which DOM users are MISSING from the IDS export
#     which IDS rows are UNUSABLE (yield no GitHub identity) and why
#     the DOM structure (tag / testid / classes) of an EXTRACTED user vs a
#         MISSING user, so a markup difference is visible
#
# It runs the SAME production export the scraper now uses: establish the correct
# table, stabilise it, do the final pre-write preview check, then write the page
# XLSX straight from the VERIFIED preview grid (export_page_from_preview -- NO
# #xlsx click), and read it back to compare identity sets.
#
# It is strictly read-only: it opens the page, clicks Next to reach the target
# page, opens a fresh IDS popup, writes to a scratch TEMP directory, and reads it
# back.  It never touches scraper state, never marks a SearchUnit complete, never
# deletes staging, and writes only into a scratch temp directory.
#
# It uses the rank's persistent profile (so the logged-in DOM is faithful);
# --fresh-profile uses a throwaway (logged out) instead.
# --------------------------------------------------------------------------
#!/usr/bin/env python3



# Per-user DOM detail: for every results-area profile link, the anchor's own
# shape and that of its enclosing card, so extracted vs missing users can be
# compared structurally.
DOM_USERS_JS = r"""
() => {
  const reserved = new Set(['search','login','settings','topics','orgs',
      'marketplace','features','pricing','about','security','enterprise',
      'notifications','issues','pulls','codespaces','sponsors','explore',
      'trending','collections','events','new','organizations','users']);
  const CHROME = 'header, nav, footer, [role="banner"], [role="navigation"],'
               + ' [role="contentinfo"], .Header, .header, .footer';
  const list = document.querySelector('[data-testid="results-list"]')
            || document.querySelector('mode_diagnose_ids_rows') || document.body;

  function describe(el) {
    if (!el) return null;
    return {
      tag: el.tagName ? el.tagName.toLowerCase() : null,
      testid: el.getAttribute ? el.getAttribute('data-testid') : null,
      classes: (el.className && el.className.baseVal !== undefined)
               ? el.className.baseVal
               : (typeof el.className === 'string' ? el.className : ''),
    };
  }

  const seen = new Map();
  const anchors = list.querySelectorAll('a[href]');
  anchors.forEach(a => {
    const href = a.getAttribute('href') || '';
    const m = href.match(/^\/([A-Za-z0-9][A-Za-z0-9-]{0,38})$/);
    if (!m || reserved.has(m[1].toLowerCase())) return;
    if (a.closest(CHROME)) return;
    const user = m[1];
    if (seen.has(user.toLowerCase())) return;
    // the card is the nearest list-item-ish ancestor
    let card = a;
    for (let i = 0; i < 6 && card && card.parentElement; i++) {
      card = card.parentElement;
      const t = (card.getAttribute && card.getAttribute('data-testid')) || '';
      if (t.indexOf('results') !== -1 || t.indexOf('user') !== -1
          || (card.tagName && card.tagName.toLowerCase() === 'li')) break;
    }
    seen.set(user.toLowerCase(), {
      username: user,
      href: href,
      hasImg: !!a.querySelector('img'),
      anchorText: (a.textContent || '').trim().slice(0, 40),
      anchor: describe(a),
      card: describe(card),
    });
  });
  return {
    url: location.href,
    resultsListPresent: !!document.querySelector('[data-testid="results-list"]'),
    users: Array.from(seen.values()),
  };
}
"""


def _dir_detail(directory):
    rows = []
    if not os.path.isdir(directory):
        return rows
    for name in sorted(os.listdir(directory)):
        path = os.path.join(directory, name)
        try:
            st = os.stat(path)
        except OSError:
            continue
        rows.append({"name": name, "size": st.st_size,
                     "mtime": st.st_mtime, "inode": st.st_ino})
    return rows


def _sha256(path, chunk=1 << 20):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(chunk), b""):
            digest.update(block)
    return digest.hexdigest()




def _grid_summary(grids, expected_ids):
    """(identities, row_count) for the ONE grid we would export.

    The grid whose OWN rows reproduce the expected identity set exactly, or --
    when none does -- the single largest grid (never a union across grids).
    """
    if not grids:
        return set(), 0
    chosen = None
    for grid in grids:
        if grid["identities"] == expected_ids:
            chosen = grid
            break
    if chosen is None:
        chosen = max(grids, key=lambda g: len(g["identities"]))
    return chosen["identities"], len(chosen["rows"])


def _usernames_from_export(path):
    """(usable, unusable): parsed rows.  usable -> {username: profile_url},
    unusable -> [raw cell tuples that yield no GitHub identity]."""
    from openpyxl import load_workbook
    usable, unusable = {}, []
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in book.worksheets:
            header = None
            for raw in sheet.iter_rows(values_only=True):
                cells = ["" if v is None else str(v).strip() for v in raw]
                if header is None:
                    header = cells
                    continue
                if not any(cells):
                    continue
                user = XlsxReader._username_from(cells)
                if user and valid_username(user):
                    usable[user.lower()] = "https://github.com/%s" % user
                else:
                    unusable.append(cells)
    finally:
        book.close()
    return usable, unusable


async def reach_page(page, scraper, query, target_page):
    """Open the search and click Next until on target_page.  Returns state."""
    url = C.SEARCH_URL.format(query=urllib.parse.quote_plus(query))
    await page.goto(url, wait_until="domcontentloaded",
                    timeout=C.NAVIGATION_TIMEOUT)
    ok, state = await scraper.wait_for_stable_results(page)
    number = page_number(state.get("url") or page.url)
    while number < target_page:
        clicked = False
        for selector in C.STABLE_NEXT_SELECTORS:
            node = page.locator(selector)
            try:
                if await node.count() == 1 and await node.first.is_visible():
                    await node.first.click(timeout=15000)
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            raise SystemExit("could not reach page %d (no Next on page %d)"
                             % (target_page, number))
        ok, state = await scraper.wait_for_stable_results(page)
        landed = page_number(state.get("url") or page.url)
        if landed <= number:
            raise SystemExit("Next did not advance past page %d" % number)
        number = landed
    return state


async def mode_diagnose_ids_rows(query, target_page, fresh_profile):
    if fresh_profile:
        profile_dir = os.path.join(tempfile.mkdtemp(prefix="diag_profile_"),
                                   "pw_profile")
        say("Using a THROWAWAY profile (logged out): %s" % profile_dir)
    else:
        profile_dir = C.PROFILE_DIR
        say("Using this folder's profile: %s" % profile_dir)
    os.makedirs(profile_dir, exist_ok=True)

    scratch = tempfile.mkdtemp(prefix="diag_ids_")
    browser = BrowserManager(headless=False, profile_dir=profile_dir)
    try:
        await browser.start(with_extension=True)
        page = await browser.page()
        await asyncio.sleep(2)
        await browser.close_extra_pages(keep=(page,))
        scraper = InstantDataScraper(browser, state=None, say=say)

        say("")
        say("Opening %r and reaching page %d ..." % (query, target_page))
        state = await reach_page(page, scraper, query, target_page)
        verdict, why = classify_search_page(state)
        expected_total, mode = expected_total_users(state)

        dom = await page.evaluate(DOM_USERS_JS)
        dom_users = {u["username"].lower(): u for u in dom["users"]}

        # Fresh IDS on the CURRENT page.
        popup = await scraper.open_instant_data_scraper(page)
        detected = await scraper.wait_for_table(popup, timeout=20000)
        _p, ids_rows = await scraper.read_stats(popup)

        expected_ids = scraper.dom_result_identities(state)

        # ---- reporting variables, initialised BEFORE the export phase, so a
        # ---- failure anywhere below can never leave one undefined (the
        # ---- previous crash: NameError: name 'exported' is not defined).
        exported = 0
        usable, unusable = {}, []
        export_path = os.path.join(scratch, "page_%03d.xlsx" % target_page)
        export_error = None
        export_ok = False
        recovered_sha = ""
        new_artifacts = []
        struct_rows = 0
        struct_ids = set()

        # ---- EXACTLY the production flow the scraper now uses -------------
        #   establish the correct table   (ONE structured grid matches the DOM)
        #   -> stabilize_results_table     (that grid's rows hold still)
        #   -> final pre-write grid check
        #   -> export_page_from_preview: write THAT grid's rows to XLSX (NO
        #      #xlsx click) -> read the written workbook back and compare.
        # Every identity below is derived from ONE structured grid's rows --
        # never a flat/global union of cells across Handsontable clones.
        initial_ids, _rows = _grid_summary(
            await scraper.read_preview_grids(popup), expected_ids)
        established = await scraper.establish_results_table(popup, page,
                                                            expected_ids)
        after_establish_ids, _rows = _grid_summary(
            await scraper.read_preview_grids(popup), expected_ids)
        stabilized = await scraper.stabilize_results_table(popup, page,
                                                           expected_ids)
        # Structured-grid state after stabilisation: the grid we would export.
        final_grids = await scraper.read_preview_grids(popup)
        struct_ids, struct_rows = _grid_summary(final_grids, expected_ids)
        after_stability_ids = struct_ids
        selected_grid = scraper._select_grid(final_grids, expected_ids)
        final_ids = await scraper.preview_identities(popup, expected_ids)
        final_matches = bool(expected_ids) and final_ids == expected_ids

        dl_dir = C.KEYWORD_DOWNLOAD_DIR
        before_snapshot = _dir_detail(dl_dir)
        if not stabilized:
            export_error = ("table did not remain stable; production would "
                            "refuse to export (nothing written)")
        elif not final_matches:
            export_error = ("final pre-write preview != DOM identities; "
                            "production would refuse to export")
        else:
            try:
                # Write straight from the verified preview -- no #xlsx click.
                exported = await scraper.export_page_from_preview(
                    popup, export_path, expected_ids)
                if exported is None or exported < 0:
                    export_error = ("preview export refused: preview "
                                    "identities did not match the DOM")
                    exported = 0
                else:
                    export_ok = True
            except BrowserLost as exc:
                export_error = "browser lost during export: %s" % exc
                exported = 0
            except Exception as exc:                       # never crash
                export_error = "%s: %s" % (type(exc).__name__, exc)
                exported = 0

        after_snapshot = _dir_detail(dl_dir)
        new_artifacts = [row for row in after_snapshot
                         if row["name"] not in {r["name"]
                                                for r in before_snapshot}]

        if export_ok and os.path.exists(export_path):
            usable, unusable = _usernames_from_export(export_path)
            recovered_sha = _sha256(export_path)
        elif export_ok:
            # the writer reported success but nothing is at the path we asked
            # for -- report a failed export rather than crash.
            export_ok = False
            export_error = ("preview export reported success but no file at %s"
                            % export_path)

        exported_id_set = set(usable)
        exact_match = export_ok and exported_id_set == set(dom_users)
        missing = [u for k, u in dom_users.items() if k not in usable]
        extra = [v for k, v in usable.items() if k not in dom_users]

        # ------------------------------------------------------- report ---
        say("")
        say("=" * 66)
        say("DIAGNOSIS: %s  (page %d)" % (query, target_page))
        say("=" * 66)
        say("classification         : %s (%s)" % (verdict, why))
        say("GitHub reported total  : %s (%s)"
            % (expected_total if expected_total else state.get(
                "reportedTotalUsers"), mode))
        say("current page number    : %d"
            % page_number(state.get("url") or page.url))
        say("GitHub DOM users       : %d" % len(dom_users))
        say("IDS detected rows      : %s" % (ids_rows if detected else "0"))
        say("IDS initial usable     : %d (before Try-another-table)"
            % len(initial_ids))
        say("IDS after establish    : %d usable  -> %s"
            % (len(after_establish_ids),
               "EXACT DOM MATCH" if established else "NOT ESTABLISHED"))
        say("IDS after stability    : %d usable  -> %s"
            % (len(after_stability_ids),
               "STABLE" if stabilized else "NOT STABLE"))
        say("IDS stable             : %d" % len(after_stability_ids))
        say("Structured grid rows      : %d" % struct_rows)
        say("Structured grid usable    : %d" % len(struct_ids))
        say("Structured grid identities: %s"
            % (", ".join(sorted(struct_ids)) if struct_ids else "(none)"))
        say("Final pre-write usable : %d  -> %s"
            % (len(final_ids),
               "matches DOM" if final_matches else "does NOT match DOM"))
        if export_ok:
            say("export method          : verified preview (no #xlsx click)")
            say("IDS export status      : OK")
            say("written XLSX rows      : %d (raw)" % max(exported, 0))
        else:
            say("IDS export status      : FAILED")
            say("Export error           : %s" % (export_error or "unknown"))
            say("written XLSX rows      : N/A")
        say("written XLSX usable    : %d" % len(usable))
        say("IDS unusable rows      : %d" % len(unusable))
        say("MISSING                : %d" % len(missing))
        say("Exact export match     : %s" % ("YES" if exact_match else "NO"))

        say("")
        say("NO-#xlsx PROOF (Chromium never asked to download)")
        say("  download dir            : %s" % dl_dir)
        say("  new download artifact(s): %d  (expected 0 -- #xlsx not clicked)"
            % len(new_artifacts))
        for row in new_artifacts:
            say("      %-40s size=%d mtime=%.0f inode=%s"
                % (row["name"], row["size"], row["mtime"], row["inode"]))
        say("  written file            : %s"
            % (os.path.basename(export_path)
               if (export_ok and os.path.exists(export_path)) else "(none)"))
        if recovered_sha:
            say("      sha256 %s  size %d"
                % (recovered_sha, os.path.getsize(export_path)))

        # ---- ROUND-TRIP PROOF: grid rows -> XLSX -> XlsxReader --------------
        # A: the selected grid as the parser saw it (headers + rows + per-row
        #    identities).  B: the written workbook, straight from openpyxl.
        # C: what XlsxReader recovers per row, and why a row yields nothing.
        import student_scraper as _identity
        import student_scraper as _xr
        say("")
        say("ROUND-TRIP PROOF (grid -> write_preview_xlsx -> XlsxReader)")
        if selected_grid is not None:
            say("  A. BEFORE WRITE -- selected structured grid")
            say("     headers: %r" % (selected_grid["headers"],))
            for i, grow in enumerate(selected_grid["rows"]):
                rid = sorted(scraper._identities_from_cells(grow))
                say("     row %d %r" % (i, grow))
                say("           identities: %s" % (rid or "(none)"))
        else:
            say("  A. no single structured grid matched the DOM")
        if export_ok and os.path.exists(export_path):
            from openpyxl import load_workbook
            book = load_workbook(export_path, read_only=True, data_only=True)
            wb_rows = []
            for sheet in book.worksheets:
                for raw in sheet.iter_rows(values_only=True):
                    wb_rows.append(["" if v is None else str(v) for v in raw])
                break
            book.close()
            say("  B. AFTER WRITE -- openpyxl workbook")
            if wb_rows:
                say("     headers: %r" % (wb_rows[0],))
                for i, wrow in enumerate(wb_rows[1:]):
                    say("     row %d %r" % (i, wrow))
            say("  C. AFTER XlsxReader")
            reader = _xr.XlsxReader()
            got = reader.read(export_path, "")
            say("     recovered %d candidate(s): %s"
                % (len(got), sorted(_identity.normalize_username(c.username)
                                    for c in got)))
            # per-row identity + why a row is unusable
            for i, wrow in enumerate(wb_rows[1:]):
                user = _xr.XlsxReader._username_from(wrow)
                if user and _xr.valid_username(user):
                    say("     row %d -> %s" % (i, user))
                else:
                    say("     row %d -> UNUSABLE (no github url in row)" % i)

        say("")
        say("GitHub DOM usernames   :")
        for u in dom_users.values():
            say("    %-30s %s" % (u["username"], u["href"]))
        say("IDS export usernames   :")
        for k, url in usable.items():
            say("    %-30s %s" % (url.rsplit("/", 1)[-1], url))
        say("")
        say("MISSING from IDS (in DOM, not exported): %d" % len(missing))
        for u in missing:
            say("    %-30s img=%s card.testid=%s card.tag=%s"
                % (u["username"], u["hasImg"],
                   (u["card"] or {}).get("testid"),
                   (u["card"] or {}).get("tag")))
        say("IN IDS but not DOM (unexpected): %d" % len(extra))
        for u in extra:
            say("    %s" % u.rsplit("/", 1)[-1])
        say("")
        say("UNUSABLE IDS rows (no GitHub identity), with the row cells:")
        for cells in unusable:
            say("    %s" % [c[:40] for c in cells if c][:6])

        # Structure comparison: an extracted user vs a missing user.
        say("")
        say("DOM STRUCTURE COMPARISON")
        extracted_sample = next((u for k, u in dom_users.items()
                                 if k in usable), None)
        missing_sample = missing[0] if missing else None
        for label, sample in (("extracted", extracted_sample),
                              ("missing  ", missing_sample)):
            if sample is None:
                say("  %s : (none on this page)" % label)
                continue
            say("  %s : %s" % (label, sample["username"]))
            say("      anchor : %s" % json.dumps(sample["anchor"]))
            say("      card   : %s" % json.dumps(sample["card"]))
            say("      hasImg : %s  anchorText : %r"
                % (sample["hasImg"], sample["anchorText"]))

        report_path = os.path.join(C.TEMP_DOWNLOAD_DIR,
                                   "ids_diag_page%d.json" % target_page)
        with open(report_path, "w", encoding="utf-8") as handle:
            json.dump({"query": query, "page": target_page,
                       "reported_total": expected_total,
                       "established": established,
                       "stabilized": stabilized,
                       "final_matches": final_matches,
                       "export_ok": export_ok,
                       "export_error": export_error,
                       "exported_rows": exported,
                       "exact_match": exact_match,
                       "dom_users": list(dom_users),
                       "ids_usable": list(usable),
                       "ids_unusable_count": len(unusable),
                       "missing": [u["username"] for u in missing],
                       "dom": dom}, handle, indent=2, default=str)
        say("")
        say("Full capture: %s" % report_path)
        say("(scratch export dir %s can be deleted)" % scratch)
        return 0
    finally:
        try:
            await browser.close()
        finally:
            shutil.rmtree(scratch, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
