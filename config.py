"""Central configuration for the GitHub student/person data extractor.

Everything discovered by inspecting this machine is recorded here with the
evidence that produced it, so no value in the workflow code is a guess.

Extension facts (from
  ~/snap/chromium/common/chromium/Default/Extensions/
      ofaokhiedipichpaobibbnahnkdoiiah/1.7.1_0/manifest.json):

  name                  "Instant Data Scraper"
  version               1.7.1
  manifest_version      3
  background            service_worker = src/background.js
  action                default_icon only -- there is NO default_popup
  commands              ABSENT

Because the manifest has no "commands" key, Ctrl+G cannot be an extension
command; it can only be a browser-level _execute_action binding.  Chrome's
toolbar is browser chrome, so page.keyboard.press("Control+G") can never
reach it from Playwright.  background.js instead reveals the real entry
point:

  fs(tab) => windows.create({
      url: runtime.getURL("src/popup.html?" + URLSearchParams({
              tabid: String(tab.id), url: tab.url})),
      type: "popup", width: 720, height: 690})

so the identical UI is obtained by opening that URL directly, which is what
open_instant_data_scraper() does.  Ctrl+G is still attempted first so the
real shortcut is used whenever the browser honours it.

Control ids below were read out of src/popup.js (React bundle, literal
`id:"..."` props) -- they are the extension's own ids, not guesses.
"""

import hashlib
import os
import tempfile

# ---------------------------------------------------------------- paths ----
# DERIVED FROM THIS FILE, never written down.  naukri does exactly this --
# BASE_DIR = os.path.dirname(os.path.abspath(__file__)) -- and the comment
# below has always claimed the same thing, but the path used to be the
# literal "/home/anand/github_profile_extractor".
#
# A literal breaks the "one folder is one rank" rule it was written to serve:
# a second clone, or the same clone on a server at /opt or under another
# user, ran its own code while reading and writing the FIRST folder's files.
# Not a crash -- it would create that tree wherever it was told to and put
# the collected students in it, which is worse than a crash.
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

# --- one folder is one rank ------------------------------------------------
# Every path a running process writes to sits next to this file, exactly as
# the naukri scraper does it.  There is no runtime/rank_NNN/ layering and no
# RankContext any more, because a rank is not a subdirectory: a rank is a
# whole clone of this project, on its own laptop, with its own `rank = N` at
# the top of student_scraper.py.
#
# That is how the job fetching scripts are already deployed -- one git
# repository per rank, cloned into its own folder -- so two ranks can no more
# share a Playwright profile or a state file than two separate checkouts can.
RUNTIME_ROOT = PROJECT_ROOT

# Chromium's user-data-dir.  It has to EXIST -- an unpacked extension can
# ONLY be loaded by launch_persistent_context(), and Instant Data Scraper is
# how every result page gets crawled -- but it does not have to live in this
# folder, and it does not have to survive the run.
#
# naukri keeps its whole login in naukri_auth.json and hands it to
# new_context(storage_state=...), so that project has no profile directory at
# all.  launch_persistent_context() takes no storage_state, so the same result
# is reached from the other side: the profile is made DISPOSABLE.  It is
# created in the system temp directory, the session is injected into it from
# github_auth.json at startup, the refreshed session is written back out at
# shutdown, and the directory is deleted.
#
# This folder therefore holds exactly what naukri_script holds -- one portable
# JSON -- and the ~1 GB of Chromium cache never accumulates here, because
# there is nothing here for it to accumulate in.
#
# The name is derived from this folder's path, so two clones on one laptop get
# two different profiles: Chromium holds a SingletonLock on a user-data-dir
# and the second launch would fail outright.
# /var/tmp, NOT /tmp.  MEASURED on this machine: /tmp is tmpfs, so a profile
# there lives in RAM and competes with Chromium's own ~900 MB, and
# /usr/lib/tmpfiles.d/tmp.conf carries "q /tmp 1777 root root 10d" -- the
# system deletes /tmp entries after ten days, open handles included, which a
# run lasting longer than that could be hit by mid-flight.  /var/tmp is
# disk-backed with a 30-day policy.  TMPDIR still wins when it is set, for a
# deployment that wants the profile somewhere specific.
PROFILE_PARENT = (os.environ.get("TMPDIR")
                  or ("/var/tmp" if os.path.isdir("/var/tmp")
                      else tempfile.gettempdir()))
PROFILE_TAG = hashlib.sha1(PROJECT_ROOT.encode("utf-8")).hexdigest()[:12]
PROFILE_DIR = os.path.join(PROFILE_PARENT, "gpe_profile_%s" % PROFILE_TAG)

# The profile is deleted at shutdown and again before the next first launch.
# Tests that want to keep one set this False.
PROFILE_DISPOSABLE = True

# THE GITHUB SESSION, as naukri stores it.
#
# find_naukri_job.py keeps its login in naukri_auth.json and hands it to
# new_context(storage_state=...).  That exact line is not available here:
# launch_persistent_context() takes no storage_state argument, because the
# profile directory IS its storage.  So the same idea is done in two steps --
# --login writes the cookies out with context.storage_state(path=...), and
# every run reads them back with context.add_cookies().
#
# The point of the file is that it is now the ONLY copy.  The profile is
# disposable and lives outside this folder, so this file is what carries the
# login -- between runs, and to another laptop.  Every shutdown writes the
# refreshed session back here; github.com cookies only, which is ~2 KB.
#
# It used to be a convenience next to a ~1 GB profile that held the real
# session, and it went stale whenever GitHub rotated user_session mid-run.
#
# It IS the session: anyone holding it is signed in as you, no password
# needed.  Written 0600 and gitignored for that reason.
AUTH_FILE = os.path.join(PROJECT_ROOT, "github_auth.json")

# Crash-safe resume point.
STATE_FILE = os.path.join(PROJECT_ROOT, "scraper_state.json")

# Which entry in the search list is being worked on now.  Written AFTER the
# current row is finished, so a crash resumes on the next one.
CURRENT_ROW_FILE = os.path.join(PROJECT_ROOT, "current_row.json")

# Profiles set aside after repeated transient failures.  Never a blacklist.
DEFERRED_FILE = os.path.join(PROJECT_ROOT, "deferred_profiles.json")
TRANSIENT_FILE = os.path.join(PROJECT_ROOT, "transient_failures.json")

# Persisted so restarting the command cannot bypass a block.
THROTTLE_FILE = os.path.join(PROJECT_ROOT, "throttle_state.json")
RATE_LIMIT_FILE = os.path.join(PROJECT_ROOT, "rate_limit_state.json")

# The raw record for the ONE person being processed right now.
TEMP_STUDENT_FILE = os.path.join(PROJECT_ROOT, "temporary.txt")
TEMP_STUDENT_JSON = os.path.join(PROJECT_ROOT, "temporary.json")
TEMP_DOWNLOAD_DIR = os.path.join(PROJECT_ROOT, "temp_downloads")

# One Instant Data Scraper XLSX per search query.
KEYWORD_DOWNLOAD_DIR = os.path.join(PROJECT_ROOT, "keyword_downloads")

# Page-by-page crawl staging: one subdirectory per search holding
# page_001.xlsx, page_002.xlsx ...  Kept out of temp_downloads/ because
# temp_store sweeps that directory clean between students.
CRAWL_STAGING_DIR = os.path.join(PROJECT_ROOT, "crawl_staging")

# Where --retry-current-search puts the XLSX it sets aside.
RETRY_BACKUP_DIR = os.path.join(PROJECT_ROOT, "retry_backup")

LOG_DIR = os.path.join(PROJECT_ROOT, "logs")
LOG_FILE = os.path.join(LOG_DIR, "student_scraper.log")

# HOW LONG A LOG IS KEPT.
#
# The log used to be one file that was never rotated.  The main loop does not
# end -- it wraps at the last search and runs the whole list again -- so on a
# server that file grew for as long as the process lived.  MEASURED on this
# folder's own log: 54 KB/hour over a four-hour span that was mostly idle,
# which is ~38 MB/month and ~0.5 GB/year of pure growth, faster under load.
#
# The log now rolls at midnight and this many dated files are kept behind the
# live one:
#
#     student_scraper.log              today, being written
#     student_scraper.log.2026-10-07   yesterday
#     ...                              and so on, oldest first to be deleted
#
# At each midnight roll the oldest beyond this count is deleted, so the logs
# directory stays bounded on its own with nothing to run by hand.
LOG_RETENTION_DAYS = 10

# This folder's result store: what makes a duplicate a duplicate.
STORE_DB = os.path.join(PROJECT_ROOT, "students.db")

# PERMANENT skips (confirmed 404, deleted account, organisation).
SKIPPED_FILE = os.path.join(PROJECT_ROOT, "skipped_accounts.txt")
FAILED_FILE = os.path.join(PROJECT_ROOT, "failed_profiles.txt")

# Directories this folder owns and creates on startup.
#
# CRAWL_STAGING_DIR is deliberately NOT here: it is created by staging_for()
# only when a search actually needs page-by-page staging.  A single-page or
# zero-result search leaves no staging directory at all, and several checks
# read "the directory exists" as "a crawl was staged".
RUNTIME_DIRECTORIES = (
    PROFILE_DIR,
    TEMP_DOWNLOAD_DIR,
    KEYWORD_DOWNLOAD_DIR,
    LOG_DIR,
)


def ensure_runtime_dirs():
    """Create this folder's working directories.  Safe to call repeatedly."""
    for path in RUNTIME_DIRECTORIES:
        os.makedirs(path, exist_ok=True)


def staging_for(slug):
    """crawl_staging/<search-slug>/ for one search."""
    return os.path.join(CRAWL_STAGING_DIR, slug or "search")


# --- collected output: current_data.xlsx and its cycle files ---------------
# Rows are appended to DETAIL_OUTPUT_FILE.  Once it holds CYCLE_SIZE rows it
# is copied to current_data_cycle1.xlsx, current_data_cycle2.xlsx ... and
# emptied.  A cycle file is sealed and never written to again, which is what
# makes it safe for a collector to copy it away while the scraper keeps
# running.

# --- cumulative duplicate history -----------------------------------------
# Identity keys only -- no personal data -- for everyone collected in any
# PAST campaign, so a new campaign never re-extracts them.
#
# LIMITATION, AND IT MATTERS: with the SQLite backend this file is LOCAL to
# one laptop.  It does NOT give 8 laptops a shared view of history.  Between
# campaigns the rank stores from every laptop must be merged, globally
# de-duplicated, built into one master identities.db, and that same file
# copied back to every laptop.  Only then does every rank start the next
# campaign from the same history.  A central PostgreSQL store would remove
# the manual step; until then the step is real and must actually be done.
HISTORY_DIR = os.path.join(PROJECT_ROOT, "history")
HISTORY_DB = os.path.join(HISTORY_DIR, "identities.db")

# --- imported pre-redesign data --------------------------------------------
# The 278 students already collected into student_data_1/2/3.xlsx are real
# records and belong in the final export, so --migrate-output imports them
# into a store of their own.  It is NOT a rank: no scraper writes it, it has
# no Playwright profile and no state file.  The source workbooks are
# PRESERVED, never deleted.
IMPORTED_DIR = os.path.join(PROJECT_ROOT, "imported")
LEGACY_IMPORT_DB = os.path.join(IMPORTED_DIR, "legacy_students.db")

# --- final export ----------------------------------------------------------
# ONE workbook.  There is no rows-per-file rotation any more: the store is
# the dataset and this is a rendering of it.  Scraper ranks NEVER write it.
FINAL_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "Student Data original")
FINAL_WORKBOOK = "student_data.xlsx"
FINAL_WORKBOOK_PATH = os.path.join(FINAL_OUTPUT_DIR, FINAL_WORKBOOK)

# The collected rows live with the rest of the output, in
# "Student Data original/", not loose at the top of the project.  The cycle
# files are derived from this path, so they land in the same folder:
#
#     Student Data original/current_data.xlsx          <- filling
#     Student Data original/current_data_cycle1.xlsx   <- sealed
#     Student Data original/current_data_cycle2.xlsx
#
# Nothing else in that folder is confused by them: the legacy-import scan
# matches only student_data_<N>.xlsx, and the merged workbook is
# student_data.xlsx.
DETAIL_OUTPUT_FILE = os.path.join(FINAL_OUTPUT_DIR, "current_data.xlsx")

# The cycle COUNTER stays beside the script with the other state files.  It
# is state, not output -- and keeping it out of the output folder means a
# folder you copy away or clear never takes the counter with it, which is
# what would make cycle numbering restart and overwrite a sealed batch.
CYCLE_STATE_FILE = os.path.join(PROJECT_ROOT, "cycle_state.json")
CYCLE_SIZE = 100

# An xlsx worksheet holds 1,048,576 rows; one is the header.  Exceeding it is
# reported, never silently truncated.
XLSX_MAX_DATA_ROWS = 1_048_575
XLSX_WARN_DATA_ROWS = 1_000_000

# --- keywords --------------------------------------------------------------

# The pre-redesign manually expanded "<Base Keyword> <City>" files have been
# removed: the city is appended by Python (the keyword list crossed with
# locations.py), never stored in a file.  load_keywords() still exists for
# ad-hoc use, but there are no default files for it to read.
KEYWORD_FILES = []

# Stamped into scraper_state.json for information only.  v5 is the simple
# flat state: no campaign fingerprints, no plan comparison, nothing that can
# refuse to resume.  A state written by an older version is read as far as it
# makes sense and the missing keys are filled with defaults.
STATE_VERSION = 5

# ------------------------------------------------------------ extension ----
EXTENSION_ID = "ofaokhiedipichpaobibbnahnkdoiiah"

# Where to look for the unpacked extension.  The first directory that holds a
# manifest.json naming "Instant Data Scraper" wins; IDS_EXTENSION_DIR beats
# all of them.  Auto-discovery means a snap/Chrome extension update that
# bumps the version directory does not break the run.
EXTENSION_SEARCH_ROOTS = [
    os.path.expanduser("~/snap/chromium/common/chromium"),
    os.path.expanduser("~/.config/chromium"),
    os.path.expanduser("~/.config/google-chrome"),
    os.path.expanduser("~/.config/opera"),
    os.path.expanduser("~/.config/BraveSoftware/Brave-Browser"),
    os.path.expanduser("~/.config/microsoft-edge"),
]
EXTENSION_DIR_OVERRIDE = os.environ.get("IDS_EXTENSION_DIR", "")

EXTENSION_POPUP_PATH = "src/popup.html"

# --- popup control ids, read from src/popup.js -----------------------------
SEL_START_CRAWL = "#startScraping"     # children: "Start crawling"
SEL_STOP_CRAWL = "#stopScraping"       # children: "Stop crawling"
SEL_NEXT_BUTTON = "#nextButton"        # children: 'Locate "Next" button'
SEL_WRONG_TABLE = "#wrongTable"        # "Try another table"
SEL_XLSX = "#xlsx"                     # "XLSX"
SEL_STATS = "#stats"                   # Pages scraped / Rows collected / ...
SEL_INFO = "#info"                     # alert container
SEL_INSTRUCTIONS = "#instructions"
SEL_ERROR = "#error"
SEL_CRAWL_DELAY = "#crawlDelay"        # "Min delay" seconds
SEL_MAX_WAIT = "#maxWait"              # "Max delay" seconds

# popup.js: Mr() sets this exact text when a crawl ends by itself.
CRAWL_STOPPED_TEXT = "Crawling stopped"

# The extension keeps the learned pagination selector at
#   localStorage["nextSelector:" + hostName]    (popup.js: fr()/dr())
# on the chrome-extension origin.  That mechanism is preserved exactly --
# the extension genuinely requires it -- but we choose WHAT goes into it.
NEXT_SELECTOR_STORAGE_KEY = "nextSelector:%s"

# Preferred pagination selectors, most stable first.  The extension's own
# "Locate Next" flow produces a full CSS path built from GitHub's hashed
# module classes, e.g.
#   div.logged-in>...>nav.prc-Pagination-PaginationContainer-HEhGo>
#   div.prc-Pagination-TablePaginationSteps-cAXs3>a.prc-Pagination-Page-Etgqf
# which breaks on every GitHub CSS release and, worse, is ambiguous: that
# final class matches the Previous link as well as Next.  These semantic
# selectors match exactly one element and survive restyling.
STABLE_NEXT_SELECTORS = [
    'a[rel="next"]',
    'a[aria-label="Next Page"]',
    'a[aria-label="Next page"]',
    "a.next_page",
]

# Crawl pacing written into the extension's own inputs.
# The extension defaults to 1 second.  A 28-page run at that pace took
# about 35 seconds and GitHub answered with HTTP 429 on page 29, losing the
# whole crawl.  Three seconds keeps a long crawl under the throttle while
# still finishing a 30-page keyword in roughly a minute and a half.
CRAWL_MIN_DELAY_SEC = 3
CRAWL_MAX_DELAY_SEC = 20

# --------------------------------------------------------------- limits ----
NAVIGATION_TIMEOUT = 45_000

# How many times one results page may be asked to settle before the search
# is abandoned.
#
# OBSERVED, 2026-10-08: a Delhi crawl exported pages 1-24 at a steady 12-13
# seconds each, then page 25 did not stabilise inside the 45s window and the
# whole search was thrown away -- 24 good pages preserved in staging but no
# keyword XLSX, waiting for a human to re-run it.  There was no rate limit
# and no block in that run, and page 25 served normally when checked
# afterwards.
#
# That clone's OWN log showed no rate limit, which is why this first read as
# "GitHub was briefly slow".  Its SIBLINGS tell a different story: at 13:01,
# the minute before the stall, Script2 logged 16 rate-limit lines, Script5
# 16 and Script3 eight.  GitHub rate-limits per ACCOUNT, and four clones were
# sharing one user_session -- so the limit was real and simply surfaced in
# another folder's log.  The machine was also short of memory at the time
# (~650 MB available, swap full, four Chromiums at ~1.3 GB peak each).
#
# Both make a 12-second page load miss a 45-second window, and retrying helps
# either one.  Neither is FIXED by retrying: the cures are fewer concurrent
# clones and not sharing one session across them.
#
# One page is reloaded and re-waited rather than losing the search to a
# blip.  Three attempts is at most ~135s of waiting against a crawl that had
# already cost five minutes.  A page that is genuinely unreachable still
# fails, just three times slower.
PAGE_STABLE_ATTEMPTS = 3

# How many times the SEARCH page itself may be opened before the search is
# given up.
#
# OBSERVED, 2026-10-08 (Script3, rank 3): Page.goto on
# "in:Network Technician location:Chikkamagaluru" exceeded the 45s timeout
# and the TimeoutError was never caught -- it unwound through collect(),
# do_search_unit(), _run() and main() and killed the process with a
# traceback, losing 48 collected students from the visible workbook and
# leaving no "run the same command to resume" message.
#
# Everything else in collect() already degrades politely: a dead browser is
# restarted, a non-users page returns FAILED, a zero-result page completes.
# Only the first navigation could still crash the run.  It is now retried,
# and a navigation that keeps failing returns FAILED like every other
# collection failure -- which stops the rank cleanly with its state intact.
SEARCH_NAV_ATTEMPTS = 3
EXTENSION_TIMEOUT = 30_000
DOWNLOAD_TIMEOUT = 120_000
PROFILE_TIMEOUT = 30_000
PORTFOLIO_TIMEOUT = 25_000
RESUME_TIMEOUT = 45_000

# How many times Chromium may be relaunched before the run gives up.  A
# restart is cheap (the session is reinjected from github_auth.json) but an
# endless restart loop would hide a real problem.
BROWSER_MAX_RESTARTS = 5

# ...and how many students must finish cleanly before that budget is handed
# back.  THE BUDGET IS PER STREAK, NOT PER RUN.
#
# It used to be per run: self.restarts was zeroed in __init__ and nowhere
# else, so five deaths spread over a month exhausted it exactly as surely as
# five in a row, and a server left running for weeks would stop on a browser
# that was behaving perfectly well -- "the restart budget is exhausted" after
# four crashes three weeks and ten thousand students ago.
#
# A browser that carries this many students without dying is not the problem
# the budget exists to catch, so the count goes back to zero.  A browser that
# really is broken never reaches the threshold and still stops the run.
BROWSER_RESTART_FORGIVE_AFTER = 25

PROFILE_RETRIES = 3        # navigation attempts inside ONE fetch
# How many times a profile may fail transiently across the whole run before
# the keyword stops.  The profile is never blacklisted -- the next run
# retries it from zero-ish, because the count only grows on real failures.
PROFILE_ATTEMPTS_MAX = 3
# After this many SEPARATE transient failures (counted across runs in
# transient_failures.json), a profile is moved to the rank's deferred queue
# (runtime/rank_NNN/deferred_profiles.json) so one repeatedly-unavailable
# profile can no longer block the whole rank.  It is retried at end-of-rank and
# by --retry-deferred; it is never blacklisted, deleted, or counted as saved.
MAX_DEFERRED_FAILURES = 3
# GLOBAL THROTTLE CIRCUIT BREAKER.  A single unavailable profile is deferred;
# but when this many DIFFERENT profiles hit a GitHub rate-limit/challenge
# within the window, that is an account/session-wide block, not three bad
# profiles.  The rank and the whole device then stop -- we do not keep probing
# GitHub.  Evidence is persisted per rank so restarting every few minutes
# cannot bypass it; a successful profile load clears the streak.
GLOBAL_THROTTLE_DISTINCT_PROFILES = 3
GLOBAL_THROTTLE_WINDOW_MINUTES = 10
# GitHub search results can change while paginating, so validated pages can
# yield a unique count that differs from the header total.  The behaviour is
# strict and NOT configurable: after one controlled reconciliation a persistent
# mismatch STOPs and staging is preserved.  Only
# --accept-stable-pagination-mismatch may accept a set proven identical across
# two fully validated crawls.
# CONTACTABLE-ONLY COLLECTION.  When True, a profile is saved only if at
# least ONE way to reach the person was found -- an email OR a phone number
# (from GitHub, the portfolio, the resume, the README, or the source row).
# Someone with neither is fetched, reported as skipped, and NOT written to the
# store or the workbook: a row nobody can be contacted through is not a lead.
#
# Both fields are already cleaned by the merger -- placeholder and
# non-personal addresses (noreply@, no-reply@, ...) are discarded and phone
# numbers must match a real format -- so a non-empty value means a real one.
REQUIRE_CONTACT = True

# The record fields that count as "contactable".  Any ONE of them is enough.
CONTACT_FIELDS = ("Email Address", "Phone Number")

# Kept as an alias so an older config or command that still says
# REQUIRE_EMAIL keeps working; the behaviour is REQUIRE_CONTACT.
REQUIRE_EMAIL = REQUIRE_CONTACT
# ======================================================= device health =====
# A run of tens of thousands of searches is unattended for hours, so two
# machine conditions are checked at every safe point between searches.
#
# BATTERY.  A laptop that dies mid-run loses the browser session and stops
# without writing a resume point.  On battery power:
#     <= PAUSE  -> the run WAITS for a charger (it does not fail)
#     <= CRITICAL -> the run STOPS cleanly while there is still time to save
# Mains power is always fine: a plugged-in laptop at 5% is charging, so
# plugging in resumes a paused run immediately without waiting for a percent.
# RESUME is the charge a run paused on battery must reach to continue.  The
# operator asked for 21: pause at 20, resume as soon as it reads 21.  That
# leaves no hysteresis band -- unplugging again just above the threshold can
# pause the run a second time within minutes -- which is the intended
# trade-off.  Raise it (25 was the previous value) to buy flap protection.
BATTERY_PAUSE_PERCENT = 20
BATTERY_RESUME_PERCENT = 21
BATTERY_CRITICAL_PERCENT = 10

# NETWORK.  With no connectivity every search fails, and those failures would
# be recorded as transient profile failures -- enough of them would park good
# profiles in the deferred queue for no reason.  So the run pauses while the
# network is down and resumes by itself the moment it comes back.
# A plain TCP connect, so nothing is sent to GitHub and no rate-limit budget
# is spent.  NOTE this answers "is github.com reachable", which covers a
# dropped link, DNS failure and an unreachable host; GitHub answering but
# erroring is a different condition, handled by the rate-limit/throttle path.
HEALTH_HOST = "github.com"
HEALTH_PORT = 443
HEALTH_TIMEOUT_SECONDS = 5
# How often a paused run re-checks.
HEALTH_RECHECK_SECONDS = 30
# How long a paused run waits before giving up and stopping safely.
# None = wait indefinitely and resume automatically, which is what the
# operator asked for: an outage overnight should find the run still there in
# the morning, not stopped.  Waiting cannot strand a dying laptop -- the
# battery keeps draining while paused and BATTERY_CRITICAL_PERCENT still
# stops it cleanly.  Set a number of minutes to bound the wait instead.
HEALTH_MAX_WAIT_MINUTES = None

# HOW OFTEN THE ONE VISIBLE WORKBOOK IS REFRESHED.
# Student Data original/student_data.xlsx used to be rebuilt only when a rank
# FINISHED.  A rank now owns ~41,400 searches, which is days of running, so
# the operator would see no spreadsheet for days even though rows were being
# collected.  It is therefore also refreshed every EXPORT_EVERY_N_STUDENTS
# saves, and whenever a run stops for any reason (Ctrl+C, rate-limit
# cooldown, low battery), so the file is never more than this many records
# behind.  SQLite stays the source of truth; the workbook is a rendering of
# it, and a failed export never affects collection.
# 0 disables periodic refresh (completion-only, the old behaviour).
# A GITHUB RATE LIMIT ON A PROFILE: COOL DOWN RATHER THAN STOP.
# A rate limit on the SEARCH page or during pagination already escalates the
# persisted cooldown (see COOLDOWN_MINUTES) and retries the same search.  A rate
# limit on a PROFILE used to behave differently: the profile stayed
# retryable, the search could not complete, and the whole rank stopped
# waiting for a human to rerun it -- which for a multi-day unattended run
# meant a restart every 20-30 minutes.
#
# With this True the two are treated the same: wait out the cooldown, then
# retry the SAME search from its saved resume point.  GitHub is still
# respected -- this waits longer, it does not retry harder -- and the level-4
# exhaustion stop remains the backstop, so a device that is genuinely blocked
# still stops instead of probing forever.
PROFILE_RATE_LIMIT_COOLDOWN = True

EXPORT_EVERY_N_STUDENTS = 25

EXPORT_RETRIES = 2

# Crawl-completion tuning (see instant_data_scraper.wait_for_crawl_complete).
CRAWL_STABLE_SECONDS = 30      # stats unchanged this long == stable
CRAWL_POLL_SECONDS = 2.0
CRAWL_MAX_SECONDS = 3 * 60 * 60
CRAWL_NO_START_SECONDS = 180   # nothing happened at all -> treat as failure
# How many times a lost table may be re-established inside one crawl.
CRAWL_TABLE_RECOVERIES = 2
# A page cap is a testing/safety valve only, passed in with --max-pages.
# Without it the extension reaches the real last page, which is production.

PACING_PROFILE = (1.0, 2.5)

# ------------------------------------------------------ conservative pacing --
# Not here: the pacing delays, the work/break cycle, the night stop and the
# keyword-shuffle switch are all at the TOP of student_scraper.py, below the
# imports, where an operator can see them at a glance.  What stays here is
# what is not a per-laptop knob -- the cooldown ladder below, the crawl
# tuning above.

# keyword_cycle.json holds the lap number the keyword order is derived from.
KEYWORD_CYCLE_FILE = os.path.join(PROJECT_ROOT, "keyword_cycle.json")

# ------------------------------------------------- rate-limit cooldown -------
# When GitHub actually returns a rate-limit/challenge, normal pacing is not
# enough.  These are OUR cooldown minutes per escalation level (1..4).
#
# They were 15/30/60/120.  Measured against this workload that was far too
# long: a rank runs at about FOUR requests per minute -- roughly 300x under
# GitHub's documented 30-searches-per-minute allowance -- and the search
# limit resets in about 60 seconds, so a 15-minute first wait cost ~15x more
# than the limit it was waiting out.  The ladder still escalates steeply, so
# a device that is genuinely blocked backs off hard; it just no longer treats
# a one-minute limit as a quarter-hour outage.
COOLDOWN_MINUTES = [1, 5, 15, 30]
MAX_COOLDOWN_LEVEL = 4         # the last defined rung of the ladder

# WHAT HAPPENS WHEN THE LADDER RUNS OUT.
#
# True  -- stop the device safely after MAX_COOLDOWN_LEVEL.  State is
#          preserved and a human reruns it.
# False -- never stop.  The wait stays at the top of the ladder
#          (COOLDOWN_MINUTES[-1]) and the SAME search is retried for as long
#          as it takes.  This is exactly how a GitHub 5xx outage is already
#          handled -- see RateLimitCooldown.exhausted() -- so nothing new is
#          invented, the same path is simply opened to rate limits too.
#
# False, because the point of this project is an unattended multi-day run: a
# rank that stopped at 02:00 on a 51-minute ladder collects nothing until
# somebody notices in the morning.
#
# THE COST, stated plainly: a genuinely blocked or suspended account is now
# retried forever instead of stopping.  At COOLDOWN_MINUTES[-1] that is one
# attempt every 30 minutes -- 48 a day, which is a trickle, not a hammering
# -- but nothing will tell you the account is dead except the log.  Watch
# logs/student_scraper.log for a cooldown that never clears.
STOP_WHEN_COOLDOWN_EXHAUSTED = False

# WHEN GITHUB TELLS US HOW LONG TO WAIT, BELIEVE IT.
# A Retry-After header is GitHub stating its own reset time, so it is used as
# given -- including when it is SHORTER than the ladder above.  Previously the
# code took max(policy, header), which overrode GitHub's own guidance in the
# wrong direction: a "retry after 60 seconds" still waited 15 minutes.
# Set False to go back to never shortening.
TRUST_RETRY_AFTER = True

# ---------------------------------------------------------- low memory ----
# Measured with tests/resource_probe.py on this machine: Python RSS is flat
# (48 MB across students), file descriptors are flat and temp_downloads
# stays empty -- there is no Python-side leak, so no forced gc is used.
# Chromium's total RSS does climb ~150-250 MB over the first students and
# then plateaus, which is cache rather than a leak.  Low-memory mode targets
# exactly the two things that are real: surplus renderer processes, and one
# heavy document staying resident between students.
LOW_MEMORY_MODE = True

# Release the working tab's document every N students (0 disables).
RELEASE_PAGE_EVERY = 5

SEARCH_URL = "https://github.com/search?q={query}&type=users"

# HOW A SEARCH QUERY IS BUILT.  The operator's required format:
#
#     in:computer science student location:Delhi
#
# The city goes in GitHub's location: qualifier so a hit means the person's
# PROFILE LOCATION is that city, not merely that the city name appears in
# their bio.  A city containing a space is quoted (location:"Greater Noida"),
# because without quotes GitHub would read only the first word as the
# location and the rest as free text.
#
# RECORDED MEASUREMENT, so nobody has to re-derive it:  GitHub's in:
# qualifier officially accepts only login/name/email.  Given a phrase it
# DISCARDS the term -- `in:B.Tech Computer Science location:delhi` and
# `Computer Science location:delhi` both returned 2,336 users.  The in:
# prefix is used here because the operator requires this exact format; be
# aware the prefixed word may therefore not narrow the search.
SEARCH_QUERY_TEMPLATE = 'in:{keyword} location:{city}'


# --------------------------------------------------------- failure marks ---
# Text that means GitHub is throttling/challenging us.  Never bypassed --
# seeing any of these stops the run and preserves every temporary file.
BLOCK_MARKERS = [
    "rate limit", "secondary rate limit", "abuse detection",
    "temporarily blocked", "unusual traffic", "verify your identity",
    "please complete the security check", "access to this site has been",
    "too many requests",
    # GitHub's search SPA does not use any of the wordings above when it
    # throttles a fast crawl.  Observed live at page 29 of a 28-page run:
    #     title: "Error 429"
    #     body : "Unable to load page. Status: 429
    #             Message: Please reload page and try again"
    # Missing these is why a rate limit surfaced only as the extension's
    # generic "Table not found".
    "unable to load page",
    "please reload page and try again",
]
BLOCK_STATUS = {429, 403, 500, 502, 503, 504}
# Of those, the ones that mean GITHUB IS BROKEN rather than "you are being
# throttled".  An outage is not our fault and not a restriction to respect,
# so the run waits it out indefinitely and retries the SAME unit instead of
# stopping: the operator does not want a GitHub hiccup at 3am to end the job.
# 429/403 keep the conservative escalate-then-stop ladder.
OUTAGE_STATUS = {500, 502, 503, 504}

# The SPA renders its own error screen with this title.
BLOCK_TITLE_RE = r"^\s*Error\s+(\d{3})\s*$"
# ...and this body line, which carries the real status code.
BLOCK_STATUS_RE = r"status:\s*(\d{3})"

# ------------------------------------------------------- resume handling ---
RESUME_MAX_BYTES = 25 * 1024 * 1024
RESUME_ALLOWED_EXT = {".pdf", ".doc", ".docx", ".txt", ".rtf", ".odt"}
RESUME_ALLOWED_MIME = {
    "application/pdf", "application/x-pdf",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.oasis.opendocument.text",
    "application/rtf", "text/rtf",
    "text/plain", "text/html", "application/xhtml+xml",
    "application/octet-stream",
}
# Never parsed, never executed, never stored.
RESUME_FORBIDDEN_EXT = {
    ".exe", ".sh", ".bat", ".cmd", ".com", ".msi", ".dll", ".so", ".dmg",
    ".apk", ".jar", ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz",
    ".iso", ".deb", ".rpm", ".bin", ".run", ".ps1", ".py", ".js",
}

# Portfolio crawl scope: homepage + at most these follow-ups (spec 25).
PORTFOLIO_MAX_SUBPAGES = 3

# --------------------------------------------------------------- output ----
# EXACT final column order.  Do not rename or reorder (spec 56).
FINAL_COLUMNS = [
    "Student Name",
    "Email Address",
    "Phone Number",
    "WhatsApp Number",
    "Current City",
    "State",
    "Country",
    "PIN Code",
    "Current Education Level",
    "Current Course",
    "Current Specialization",
    "Current Year / Semester",
    "Expected Graduation Year",
    "College Name",
    "University / Affiliation University",
    "College City",
    "College State",
    "UG Degree",
    "UG Specialization",
    "UG College / Institute Name",
    "UG University",
    "UG Graduation Year",
    "PG Degree",
    "PG Specialization",
    "PG College / Institute Name",
    "PG University",
    "PG Graduation Year",
    "Primary Skills",
    "Secondary Skills",
    "Programming Languages",
    "Technologies / Frameworks",
    "Tools / Software Known",
    "Total Experience",
    "Internship Experience",
    "Current Employment Status",
    "Current Company",
    "Current Job Role",
    "LinkedIn Profile",
    "GitHub Profile",
    "Portfolio Website",
    "Resume URL",
    "Career Interest",
    "Preferred Job Role",
    "Preferred Industry",
    "Data Source",
    "Source URL",
    "Source Platform",
    "Data Collection Date",
    "Email Validation Status",
    "Phone Validation Status",
    "Data Verified Status",
    "Remarks / Notes",
    "source_file",
    "country",
    "pin_code",
    "experience",
    "current_company",
    "current_role",
    "career_interest",
    "preferred_role",
    "preferred_industry",
    "linkedin",
    "github",
    "portfolio",
    "resume_url",
]

# Alias column -> canonical column.  Both are always written (spec 57).
ALIAS_COLUMNS = {
    "country": "Country",
    "pin_code": "PIN Code",
    "experience": "Total Experience",
    "current_company": "Current Company",
    "current_role": "Current Job Role",
    "career_interest": "Career Interest",
    "preferred_role": "Preferred Job Role",
    "preferred_industry": "Preferred Industry",
    "linkedin": "LinkedIn Profile",
    "github": "GitHub Profile",
    "portfolio": "Portfolio Website",
    "resume_url": "Resume URL",
}

# Single-segment github.com paths that are never user profiles.
RESERVED_PATHS = {
    "about", "account", "apps", "blog", "business", "collections", "contact",
    "codespaces", "customer-stories", "dashboard", "discussions", "download",
    "education", "enterprise", "events", "explore", "features", "git-guides",
    "github", "home", "issues", "join", "login", "logout", "marketplace",
    "mobile", "new", "nonprofit", "notifications", "open-source",
    "organizations", "orgs", "premium", "pricing", "pulls", "readme",
    "resources", "search", "security", "session", "sessions", "settings",
    "signup", "site", "solutions", "sponsors", "stars", "team", "topics",
    "trending", "watching", "wiki", "pull", "commits", "tree", "blob",
    "releases", "tags", "compare", "graphs", "network", "wiki", "projects",
    "packages", "sponsors", "copilot", "repos", "users", "assets",
    "favicon.ico", "robots.txt", "sitemap.xml", "auth", "connect",
}

# Chromium is launched as the full build (the headless shell cannot load
# extensions), so its footprint is trimmed explicitly.
CHROMIUM_LEAN_ARGS = [
    "--disable-dev-shm-usage",
    "--disable-gpu",
    "--disable-software-rasterizer",
    "--disable-background-networking",
    "--disable-background-timer-throttling",
    "--disable-renderer-backgrounding",
    "--disable-breakpad",
    "--disable-component-update",
    "--disable-sync",
    "--no-first-run",
    "--no-default-browser-check",
    "--hide-crash-restore-bubble",
    # The profile is thrown away at shutdown, so this only bounds the peak
    # during one run -- but one run is what grew the old .pw_profile to a
    # gigabyte, so it is bounded here too.  50 MB; Chromium evicts oldest
    # first once it is full.
    "--disk-cache-size=52428800",
]

# Added only when LOW_MEMORY_MODE is on.  Both are ordinary Chromium
# scheduling flags, not security trade-offs: Site Isolation stays enabled
# and nothing Instant Data Scraper needs is disabled.  Profiles are visited
# one at a time, so a small renderer pool is enough and each renderer saved
# is roughly 60-90 MB on this machine.
CHROMIUM_LOW_MEMORY_ARGS = [
    "--renderer-process-limit=3",
    "--disable-back-forward-cache",
]


def ensure_dirs():
    """Create every directory this folder writes to.

    HISTORY_DIR is deliberately NOT created here.  Nothing in a run opens the
    cumulative identity index -- open_history() has no call sites and
    remember() is reached only from the test suite -- so creating the
    directory on every startup made an empty folder that no code ever used.
    On a read-only or restricted deployment that makedirs() was a startup
    failure for a feature the run does not have.

    If the index is ever wired up, _connect() already calls makedirs() on the
    database's own parent before opening it for writing, so the directory
    appears exactly when something first needs it.
    """
    ensure_runtime_dirs()
    os.makedirs(FINAL_OUTPUT_DIR, exist_ok=True)


def find_extension_dir():
    """Return the unpacked Instant Data Scraper directory, or "".

    Verified by reading manifest.json, never by assuming a version folder.
    """
    if EXTENSION_DIR_OVERRIDE:
        return EXTENSION_DIR_OVERRIDE if _is_ids(EXTENSION_DIR_OVERRIDE) else ""
    # Roots are searched in order, so the browser listed first wins; within
    # one root the highest version directory wins.
    for root in EXTENSION_SEARCH_ROOTS:
        if not os.path.isdir(root):
            continue
        found = []
        for dirpath, dirnames, _files in os.walk(root):
            if os.path.basename(dirpath) != EXTENSION_ID:
                continue
            for version in sorted(dirnames):
                candidate = os.path.join(dirpath, version)
                if _is_ids(candidate):
                    found.append(candidate)
            dirnames[:] = []
        if found:
            return sorted(
                found, key=lambda p: _version_key(os.path.basename(p)))[-1]
    return ""


def _version_key(name):
    core = name.split("_")[0]
    parts = []
    for chunk in core.split("."):
        parts.append(int(chunk) if chunk.isdigit() else 0)
    return parts


def _is_ids(path):
    manifest = os.path.join(path, "manifest.json")
    if not os.path.isfile(manifest):
        return False
    try:
        import json
        with open(manifest, encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        return False
    return "instant data scraper" in str(data.get("name", "")).lower()
