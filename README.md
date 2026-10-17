# GitHub Student Data Extractor

Collects **public** student/person information starting from a GitHub Users
search, using the **Instant Data Scraper** browser extension to gather every
search-result page, then visiting each person's GitHub profile, public
portfolio and public resume.

```
new_mylib.py  (keyword_by_category  x  city_by_states_original)
  -> this rank's cities     (rank / num_sys at the top of student_scraper.py)
  -> crossed with every keyword
  -> for each search:
       GitHub Users search  (https://github.com/search?q=...&type=users)
       -> Instant Data Scraper -> Start Crawling -> ALL result pages
       -> ONE XLSX per search
       -> for each row: GitHub profile -> raw README -> portfolio -> resume
            -> temporary.txt -> extract / normalise
            -> master_index.xlsx (duplicate index) + current_data.xlsx
            -> clear temporary.txt
       -> every row resolved -> delete that XLSX -> next search
  -> every 100 people: current_data.xlsx -> current_data_cycleN.xlsx
  -> reaching the last search wraps back to the first, forever
```

**One search = one crawl = one XLSX.** Searches are **generated**, never
written out by hand: 3,450 base keywords x 871 cities = **3,004,950
searches**. Editing `keyword_by_category` in `new_mylib.py` regenerates the
list.

**One folder = one rank.** Work division is two variables at the top of
`student_scraper.py` and nothing else:

```python
rank = 1          # which copy of this script this folder is
num_sys = 5       # how many copies are running altogether
```

Clone this folder onto each laptop and set `rank` to 1, 2, 3 ... leaving
`num_sys` the same everywhere. Each copy takes its own contiguous slice of
the city list, so no two copies ever search the same city. There is no
`--rank` flag, no campaign, no manifest and no `runtime/rank_NNN/`: a rank is
not a subdirectory, it is a whole checkout — exactly how the job fetching
scripts are deployed.

Everything a run writes sits next to `student_scraper.py`:
`current_row.json`, `scraper_state.json`, `students.db`, `current_data.xlsx`.
The Chromium profile is the one exception — it is disposable and lives in the
system temp directory, so the login travels in `github_auth.json` alone,
exactly as `naukri_auth.json` carries the Naukri login.

## The safety rule

```
SUCCESS                -> clean up
ANY UNRESOLVED FAILURE -> keep the XLSX, the state and the temp files
```

A crash, Ctrl+C, rate limit, timeout or a failed Excel write never deletes
anything. Cleanup is deliberately **not** in a `finally:` block.

`current_row.json` moves on only after a search has fully completed, so an
interrupted search is re-run rather than skipped.

## The files

**Three** Python files, down from forty.

| file | lines | what it is |
|---|---|---|
| `student_scraper.py` | 16,753 | **the main script** — everything the run needs |
| `new_mylib.py` | 5,614 | **the cities and keywords** |
| `config.py` | 881 | paths, schema, internal tuning |

The diagnostics that used to be `diagnose_page.py`, `diagnose_page2.py` and
`diagnose_ids_rows.py` are now flags on the main script (`--diagnose-page`,
`--diagnose-page2`, `--diagnose-ids-rows`), and the keyword helpers that were
`student_search_keywords.py` read `new_mylib.py` directly.

The same shape as `naukri_script`: one script you run, one file holding the
data it works through.

### new_mylib.py — the data

Two dictionaries in it are the live source of the whole search plan:

| dictionary | gives | size |
|---|---|---|
| `city_by_states_original` | 871 cities | 242 lines |
| `keyword_by_category` | 3,450 base keywords | 2,903 lines |

`student_scraper.py` **parses** this file as text rather than importing it:

```python
CITY_SOURCE_FILE = "new_mylib.py"
CITY_SOURCE_DICT = "city_by_states_original"
KEYWORD_SOURCE_FILE = "new_mylib.py"
KEYWORD_SOURCE_DICT = "keyword_by_category"
```

Parsed, not imported, for a concrete reason: the file also carries naukri-era
helpers that need `Levenshtein` and `gtts`, which this project does not
install, so importing it would fail. Reading the two dictionaries out of the
text sidesteps that entirely and means the file can be edited freely without
anything else having to change.

To change what gets searched, edit the dictionaries in `new_mylib.py` and run
`--dry-run`.

### student_scraper.py — the script

It was forty modules. Thirty of them were merged into this one file in
dependency order, each keeping its original heading as a banner so the
boundaries are still visible:

```
  the search plan      rank/num_sys, get_cities_by_rank, build_search_list
  the run              Workflow: collect -> process -> save -> next search
  THE LIBRARY          browser, extension, extractors, store, writers, pacing
  the modes            --dry-run, --login, --retry-current-search, ...
```

**Nothing was rewritten in the merge.** It deleted the imports the modules
needed from each other, dropped the forwarding wrappers that existed only so
one module could call another by its own name, and renamed the few functions
that clashed. 384 definitions were compared against their originals line for
line and all 384 matched.

| was | is now | why |
|---|---|---|
| `identity.is_placeholder_email` | `identity_is_placeholder_email` | stricter than the extractor's: it decides whether an address may establish IDENTITY |
| `device_health.describe` | `describe_health` | unrelated report builders |
| `transient_failures.describe` | `describe_transient_failure` | |
| `result_store._now` | `_now_datetime` | returns a datetime; the other `_now` returns an ISO string |
| `locations.CITIES` | `SEARCH_CITIES` | where we LOOK |
| `location_extractor.CITIES` | `KNOWN_CITIES` | what a PERSON wrote |

That last pair matters most. One is the search dataset, the other is the
vocabulary for reading a person's own location, and the whole project depends
on never confusing them. `normalizer.parse_location` also survives, as
`parse_location_tuple`: it returned `.as_tuple()` and callers unpack three
values from it, so it is an adapter rather than a wrapper.

## Requirements

Already present on this machine; no `pip install` is needed.

| Need | Used |
|---|---|
| Browser automation | `playwright` (async API) + the full Chromium build |
| Excel | `openpyxl` |
| PDF resumes | `pypdf` / `pdfminer` if installed, else `pdftotext` (poppler) |
| DOCX / ODT resumes | `python-docx` if installed, else stdlib `zipfile` |
| Legacy `.doc` | `antiword` / `catdoc` / `soffice`, if present |

Selenium, PyAutoGUI, physical mouse control and the desktop clipboard are
**not** used anywhere.

## Commands

### Reports — these write nothing

```bash
# check the production keyword file and show the resolved list
python3 student_scraper.py --validate-keywords
```
```bash
# this copy's share of the cities, what every other copy gets, and the
# resume point.  Also warns if num_sys leaves some copies with no work.
python3 student_scraper.py --dry-run
```
```bash
# check a workbook against the 65-column contract
python3 student_scraper.py --validate-output
```

### Running

```bash
# one-time per folder: log this copy's profile into GitHub (visible window)
python3 student_scraper.py --login
```
```bash
# collect, forever
python3 student_scraper.py
```
```bash
# diagnose the extension end to end
python3 student_scraper.py --check-extension
```
```bash
# testing helpers: visible browser, stop after N people, cap the crawl
python3 student_scraper.py --show --limit 3 --max-pages 2
```
```bash
# process an already-downloaded XLSX without re-crawling
python3 student_scraper.py --process-existing
```

### Recovery — reports unless `--apply`

```bash
# re-collect the search this copy is currently on: clears ONLY that search's
# collection/processing state and moves its XLSX aside.  current_row.json,
# students.db and every cycle file are untouched.
python3 student_scraper.py --retry-current-search
python3 student_scraper.py --retry-current-search --apply
```
```bash
# one more conservative attempt at this copy's deferred profiles
python3 student_scraper.py --retry-deferred
```

### Output

Collection output needs no command: rows land in `current_data.xlsx` as they
are found, and every 100 rows it is sealed into the next
`current_data_cycleN.xlsx` and emptied. A cycle file is finished the moment
it is created and is never reopened, so a collector can copy it away while
the scraper keeps running.

`Student Data original/master_index.xlsx` is the duplicate index — it is what
makes "have I already collected this person?" answerable before any fetching.
One row per person, keyed by GitHub username, with `first_seen_at`,
`last_seen_at` and `seen_count`.

> **`--export-final` does not work right now.** It still merges `students.db`
> files, and nothing has written `students.db` since the master workbook
> replaced it. Merging several laptops' people into one workbook is a matter of
> merging their `master_index.xlsx` files; the command has not been rewritten
> to do that yet. Deciding what it should do is the one open item in this
> project.

```bash
python3 student_scraper.py --export-final \
    --store /media/usb/rank2/students.db        # reads a file no longer written
```
```bash
# which stores would be merged, and how many people each holds
python3 student_scraper.py --export-final --list-stores
```

Being explicit matters: scanning for stores would silently produce a workbook
missing whichever laptop had not been copied over yet, and a merge that
quietly leaves a laptop out looks exactly like a merge that worked.

## Running several copies

### The work division

```python
rank = 1          # top of student_scraper.py
num_sys = 5
```

`get_cities_by_rank()` cuts the city list into `num_sys` contiguous blocks and
this copy takes the one at its own position:

```python
def get_cities_by_rank(all_cities, rank, num_sys):
    len_city  = len(all_cities)
    one_unit  = int(len_city / num_sys) + 1
    first_idx = one_unit * (rank - 1)
    last_idx  = one_unit * rank
    return all_cities[first_idx:last_idx]
```

It is a pure function of `(len, rank, num_sys)`, so every copy works out its
own share without talking to any other copy, and the same numbers always
produce the same split.

**Whole cities, never half a city.** Every keyword of a city lands on one
copy, so no two copies crawl the same `location:` query. `tests/test_locations.py`
proves it for `num_sys` of 1, 2, 5 and 8.

**One name, one city — naukri's rule.** The list is
`flatten(city_by_states_original)` then `dict.fromkeys`, exactly as
`find_naukri_job.py` builds it. So "Bengaluru" and "Bangalore" are two
separate cities, not one city with two spellings, and they can land on
different ranks. Two copies can then crawl overlapping GitHub results for what
is really one place. Each folder's `master_index.xlsx` absorbs its own repeats;
between two laptops the overlap goes when their masters are merged.

The other half of that rule: **a name `new_mylib.py` does not list is not
searched.** The old alias tables added thirteen by hand — Bangalore, Mysore,
Mangalore, Vizag, Trivandrum, Calicut, Baroda, Belgaum, Trichy, Hanamkonda,
Bokaro, New Delhi, Chhatrapati Sambhajinagar. They are gone. To search a
spelling, add it to `city_by_states_original`; it then becomes its own city,
the same way `Gurgaon` and `Gurugram` are both already listed and both
searched.

With 871 cities and `num_sys = 5`:

| rank | cities | searches |
|---|---|---|
| 1 | 175 | 603,750 |
| 2 | 175 | 603,750 |
| 3 | 175 | 603,750 |
| 4 | 175 | 603,750 |
| 5 | 171 | 589,950 |

871 of 871 cities covered.

### A caveat about large `num_sys`

The formula rounds the block size up, so it never loses a city, but once
`num_sys` grows past about a tenth of the city count the last copies get an
empty slice and sit idle. With 871 cities that starts at `num_sys = 80`
(1 idle copy) and reaches 435 idle copies at `num_sys = 871`. No work is ever
lost — the earlier copies cover every city — but those copies have nothing to
do. `--dry-run` prints exactly which ranks would be idle, and a copy that
starts with no cities says so and exits rather than pretending to run.

### Deploying

```bash
# on each laptop
git clone <this repo> github_profile_extractor
cd github_profile_extractor
# edit the two lines at the top of student_scraper.py:  rank = N
python3 student_scraper.py --dry-run     # confirm this copy's share
python3 student_scraper.py --login       # once, per copy
python3 student_scraper.py               # collect, forever
```

Nothing is shared between the copies at run time. Each has its own Chromium
profile, its own `current_row.json`, its own `master_index.xlsx` and its own
`current_data.xlsx`. Duplicates *between* laptops are resolved only when their
masters are merged — see **Storage** below.

### Where it resumes from

`current_row.json` holds the index of the next search to run:

```json
{
    "current_row": 41287
}
```

It is written **after** a search finishes, so a crash resumes on the search
*after* the one that completed and never repeats it. Reaching the last search
wraps the index back to 0 and the whole list runs again — a city crawled
months ago has new people in it by the time the loop returns to it.

An index past the end of the list — the list shrank because a city or keyword
was removed — resets to 0 rather than being used, because an index past the
end would otherwise skip the entire run.

`scraper_state.json` is a separate file and tracks only the search currently
in hand: whether its XLSX has been collected, how far down it the processing
has got, and which person is being handled. Keeping the two apart is what
makes resuming simple — the position is one integer in its own file.

`keyword_cycle.json` holds the third number: which **lap** this folder is on,
which is what decides the keyword order (see *Random keyword order* below).
Only the number is stored, and the order is re-derived from it, so a copy
killed half way through lap 7 resumes lap 7 in the order it started it:

```json
{
    "keyword_cycle": 7,
    "shuffled": true,
    "first_keywords": ["Python Developer fresher", "B.Tech IT 3rd year"]
}
```

The `first_keywords` and `shuffle_seed` keys are a preview for reading by eye;
nothing depends on them. All three files are per-laptop and gitignored, and
`reset.sh` clears them together.

### Changing the target audience

Edit `keyword_by_category` in `new_mylib.py` and run
`--validate-keywords`, then `--dry-run`. Nothing refuses to resume: the old
version compared campaign fingerprints and would rather stop than continue
against a changed plan, and all of that is gone. Changing the keywords simply
changes which searches each copy walks, and `current_row.json` is reset to 0
if the list got shorter than the saved index.

## Storage

```
5 copies / 5 laptops
      |
   <each folder>/Student Data original/current_data.xlsx            <- the output
                 -> current_data_cycleN.xlsx                        <- sealed at 100
   <each folder>/Student Data original/master_index.xlsx            <- duplicate index
      |
   merge the masters  (global de-duplication -- see Storage)
      |
   Student Data original/student_data.xlsx                          <- on demand
```

Two files, two jobs:

| file | what it is for |
|---|---|
| `current_data.xlsx` + `current_data_cycleN.xlsx` | the collected output, written as it is found |
| `students.db` | the duplicate index: "have I already collected this person?" |

The spreadsheet is written only **after** the database accepted the row, so a
person can never reach `current_data.xlsx` twice: the UNIQUE index, not the
code, is what decides they are new.

Uniqueness is enforced by the **database** —
`INSERT … ON CONFLICT DO NOTHING` against UNIQUE indexes on the normalised
username, profile URL and real email. A SELECT-then-INSERT would be a race
even in one process.

Claims are **leases**, never locks: a crashed copy's claim expires and the
candidate becomes retryable. Nothing about a failure blacklists anybody.

A failed spreadsheet write never loses a person. `students.db` has already
accepted the row, so the failure is reported loudly and collection continues —
a silent failure here would look like the scraper was finding nobody.

### What a local SQLite store does NOT give you

**There is no real-time global duplicate prevention across laptops.** Copy 1
on laptop 1 and copy 4 on laptop 4 cannot see each other. In practice this
costs far less than it sounds, because the city split already keeps them in
different places: two copies only ever meet on a person who appears in two
cities' result sets. Where they do meet, both rows exist and the **export**
collapses them. Duplicate *prevention* is per copy; duplicate *resolution* is
global and happens at export.

`master_index.xlsx` is **local to one laptop**. Sharing it means:

```
every laptop's master_index.xlsx  ->  merge  ->  global de-duplicate
  ->  one master  ->  copy that same file to each laptop
```

That is a manual step, exactly as it is in `naukri_script`.

> **`history/` is dead.** It held a cumulative SQLite identity index for the
> old multi-campaign design. `IdentityIndex` and `open_history()` are still in
> `student_scraper.py` but nothing calls them, the directory is empty, and
> `history/identities.db` is never created.

## Safety guards

**LibreOffice locks.** If the merged workbook is open, the **export** stops
before writing anything and names it. A lock file is never deleted
automatically, not even when it looks stale — a wrong guess costs somebody's
edits. An orphaned lock (no matching workbook) is reported as a warning only.
Collection never writes that workbook, so somebody having `student_data.xlsx`
open does not block it. `current_data.xlsx` is this copy's own file and
nothing else opens it.

**Where the knobs are.** Every setting an operator tunes is **directly below
the imports in `student_scraper.py`** — 22 lines, one comment each, in the
order `find_naukri_job.py` has them, and above the first line of machinery.
Nothing about waiting is in `config.py` any more; it keeps paths, the output
schema, the rate-limit cooldown ladder and crawl tuning.

```python
rank = 1
num_sys = 15

# -- night stop: no collecting between these hours (local, 24-hour clock) --
NIGHT_PAUSE_ENABLED = False     # True = ON, False = OFF
NIGHT_START_HOUR = 23           # collecting stops at this hour
NIGHT_END_HOUR = 5              # collecting resumes at this hour

# -- keyword order ---------------------------------------------------------
KEYWORD_SHUFFLE_ENABLED = True  # False = new_mylib.py's file order

# -- request pacing: every wait is a fresh random draw MIN..MAX, in seconds
SEARCH_DELAY_MIN = 15           # between two completed searches
SEARCH_DELAY_MAX = 20
ZERO_RESULT_DELAY_MIN = 5       # after a search that matched nobody
ZERO_RESULT_DELAY_MAX = 7
PAGE_DELAY_MIN = 60             # between two results pages of one search
PAGE_DELAY_MAX = 180
PROFILE_DELAY_MIN = 5           # between two people
PROFILE_DELAY_MAX = 25
JITTER_FRACTION = 0.35          # retry backoffs are spread up by 0-35%
LAP_COOLDOWN_MINUTES = 45       # cooldown when the whole list has been walked

# -- work / break: both drawn fresh every time, counting WORKING time only -
WORK_PERIOD_MIN_HOURS = 2       # work this long...
WORK_PERIOD_MAX_HOURS = 6
BREAK_MIN_MINUTES = 30          # ...then stop this long
BREAK_MAX_MINUTES = 60
BROWSER_RESTART_PAUSE_SECONDS = 600   # a pause this long closes Chromium
```

`config.KEYWORD_CYCLE_FILE` stays in `config.py`, because it is a path.

Nothing an operator tunes is in `config.py` any more. The one related entry
still there is `KEYWORD_CYCLE_FILE` — the path of the file holding the lap
number, so it lives with the other paths.

**Request pacing.** Every delay is a fresh uniform draw between a MIN and a
MAX at the top of `student_scraper.py`, and the numbers are
`find_naukri_job.py`'s — that script
has been pacing a live campaign for months, so its ranges are the ones that
have actually been operated:

| between | here | naukri constant |
|---|---|---|
| two people | 5–25 s | `JOB_CARD_DELAY_*_SECONDS` |
| two results pages | 60–180 s | `LISTING_PAGE_DELAY_*_SECONDS` |
| two searches | 15–20 s | `HX_KEYWORD_GAP_SECONDS` (15 s + jitter) |
| after an empty search | 5–7 s | `HX_CITY_GAP_SECONDS` (5 s + jitter) |
| two laps | 45 min | `HX_CYCLE_COOLDOWN_MINUTES` |

Change any of them and the next start picks it up; no other edit is needed
anywhere in the file.

A results page every 1–3 minutes is the slow end on purpose: it is the
request that returns the most data for the least effort, and it is where a
long campaign buys its safety. The retry backoffs are the only fixed-interval
waits left and they are jittered up by 0–`JITTER_FRACTION` (0.35), because a
perfectly uniform retry interval is distinguishable from a person browsing.
An empty search is deliberately the cheap one — most keyword × city pairs
match nobody. Raising throughput means lowering `PAGE_DELAY_*` and
`PROFILE_DELAY_*`; nothing else has to change.

**Random keyword order.** The search list is cities → spellings → keywords,
and the keyword part of it is **shuffled once per lap** instead of always
running `new_mylib.py`'s file order — `find_naukri_job.py`'s per-cycle keyword
shuffle, on this dataset. A fixed order would send the same keyword first for
every city, on every laptop, for the life of the campaign.

The draw is seeded from `(rank, lap number)`, not from the clock, which buys
three things: a restart does not reshuffle (only the lap number is saved — see
*Where it resumes from*), rank 1 and rank 4 run the same lap in **different**
orders instead of hitting GitHub with the same keyword at the same moment, and
the order is reproducible from one integer. It is a permutation and that is
asserted rather than assumed, so a lap can never silently drop a keyword.
`--dry-run` previews the order the next run will actually use.

**The night stop and the work/break cycle.** Both are
`find_naukri_job.py`'s, and both are checked **only between two searches** — a
search in progress is never abandoned part way through:

| guard | default | names |
|---|---|---|
| night stop | off, 23:00–05:00 | `NIGHT_PAUSE_ENABLED`, `NIGHT_START_HOUR`, `NIGHT_END_HOUR` |
| work period | 2–6 h, drawn fresh each time | `WORK_PERIOD_MIN_HOURS`, `WORK_PERIOD_MAX_HOURS` |
| break | 30–60 min, drawn fresh each time | `BREAK_MIN_MINUTES`, `BREAK_MAX_MINUTES` |

All of it is at the **top of `student_scraper.py`**, directly under `rank` and
`num_sys` — see *Where the knobs are* below.

A pause of 10 minutes or more (`BROWSER_RESTART_PAUSE_SECONDS`) **closes
Chromium** for the wait and relaunches it afterwards rather than hold an idle
browser and an ageing GitHub session open; the persistent profile keeps the
login, so the relaunch costs a page load. `current_row.json` is already
written, so the process can be killed in the middle of a pause and the next
start resumes in the same place.

The work clock counts **working** time, not wall-clock time. It is stopped
around the night stop, the break itself, the rate-limit cooldowns and the
battery/network gate, so "4 hours" means four hours of actual collecting — a
copy that spent its morning waiting out cooldowns has not earned a break.
`WORK_PERIOD_MAX_HOURS = 0` switches the work/break cycle off entirely, and
`KEYWORD_SHUFFLE_ENABLED = False` restores `new_mylib.py`'s file order.

**Transient profile failures.** A timeout, 5xx or dropped connection is
recorded in `transient_failures.json` with username, GitHub URL,
search query, source XLSX, source row, reason, attempt count and timestamp.
The profile is retried up to `PROFILE_ATTEMPTS_MAX` (3); after that the search
stops with the XLSX and state preserved. Nothing is ever blacklisted — the entry is
deleted the moment that profile succeeds. Permanent skips (confirmed 404,
deleted account, organisation) go to `skipped_accounts.txt` instead.

**Page classification.** Before the extension is opened, the search page is
classified from GitHub's own DOM into one explicit state:

| State | Behaviour |
|---|---|
`RESULTS_MULTI_PAGE` | pagination crawl → export → process |
`RESULTS_SINGLE_PAGE` | export the one page → process |
`ZERO_RESULTS` | **search COMPLETE**, no XLSX, no candidates, advance |
`LOGIN_REQUIRED` | fail, preserve |
`HTTP_ERROR` / `BLOCKED` | stop the run, preserve |
`GITHUB_ERROR` / `UNKNOWN_INVALID` | fail, preserve |

The order of checks matters and is fixed: login → HTTP error → blocked →
GitHub error → is this a Users search → count results → explicit empty state →
unknown. Every one of the failure states *also* renders zero user results, so
"how many users?" is asked last. Visible results always beat empty-state
wording.

**A valid search with zero users is a success, not a failure.** It completes
its search unit with `rows = 0`, creates no XLSX, saves no student, leaves no
temporary data, and does **not** consume `--limit` — the limit counts students
saved, not searches visited. Treating it as a failure made one empty query
block its entire rank on every rerun.

Result counting excludes GitHub's own chrome (header, nav, footer, the
viewer's avatar and sponsor links), because those are all single-segment
profile links and counting them would make an empty page look full.

`python3 student_scraper.py --diagnose-page "<query>"` classifies one page and
prints every
signal behind the verdict. It opens no extension, downloads nothing and
touches no state.

**Multi-page searches are collected page by page.** GitHub replaces the
results DOM when it paginates, and the extension popup stays bound to the page
it was opened on — a fresh popup on page 2 sees all its users, the page-1 popup
sees none. So the extension is reinitialised per page:

```
page 1   fresh IDS -> export -> crawl_staging/<slug>/page_001.xlsx
Next     Playwright clicks, waits for the results DOM to settle
page 2   fresh IDS -> export -> page_002.xlsx
merge    dedupe by GitHub identity -> ONE keyword_downloads/<slug>.xlsx
```

Instant Data Scraper is still the extractor; Playwright only does the
navigation and synchronisation. A popup holds its rows only while it lives, so
exporting each page *before* replacing the popup is forced by the extension,
not chosen.

**Atomicity.** The keyword XLSX is written once, after every page succeeds. If
any page fails, staging is preserved, no keyword XLSX is created, and
`collection.complete` stays `False`. Staging lives outside `temp_downloads/`
because `temp_store` sweeps that directory between students.

**A page export survives a browser hiccup.** Chromium writes downloads to
`keyword_downloads/`, not into staging, so page-level recovery scans the
*browser's* download directory (passed explicitly) rather than the staging
destination — the bug that turned a survivable `save_as` failure into a fatal
`BrowserLost`. Recovery snapshots that directory **before** the click, so it
can only ever adopt *this* page's export: a previous page's file, or a stale
one, was already in the snapshot; more than one recoverable file refuses as
ambiguous.

**Completeness is judged by opening the file, not by its name.** In this
environment (snap Chromium under Playwright) a *finished* download is often
left as an unrenamed `<guid>.crdownload` in the downloads directory —
`save_as` never runs the rename because the browser dies first. Those files
are complete, valid workbooks. So recovery waits for a candidate's size to go
stable and then tries to **open** it: one that opens is recovered (whatever
its extension); one that will not open is a genuine partial and is refused.
An extension-only check discarded finished exports and was the regression that
made every page look incomplete.

**The crawl survives the browser dying between pages.** A page counts as done
only once its staged file validates. If the browser dies after a page is
validated, the collector restarts Chromium on the same rank profile,
reconfirms login, navigates straight to the next page's URL, and continues —
the whole search does not need one browser process to survive it. If the
browser dies *before* a page is validated, that page is retried; if the
restart budget (`BROWSER_MAX_RESTARTS`) is exhausted, staging is preserved and
the run stops. A validated page is never re-exported.

**Before any export, IDS's selected table is validated against the DOM.**
Instant Data Scraper auto-detects a table, and on a small results page it can
pick the wrong one — a table that is mostly page chrome (one real user plus
rows like "Skip to content"). So before exporting, the identities in IDS's
own preview grid are compared to the live DOM identities, and "Try another
table" is cycled until they are **exactly equal** — correctness is
identity-set equality, never raw row count or "table detected". This happens
pre-export because export can kill Chromium in this environment; cycling by
exporting each candidate is not an option. The search is bounded (stops when a
preview signature repeats, when the control disappears, or after 8 candidates)
and on failure yields `[INCOMPLETE] IDS could not establish the GitHub Users
result table` with nothing exported. Both single-page and every page of a
multi-page crawl use it. This is the **first** of two protections; the second
is:

**Two guards defend the export, because IDS's model can shift after the
preview is validated.** The extension source confirms the preview grid
(`Or() -> sr(mn.data)`) and the XLSX export (`fl() -> cr(mn.data)`) are built
from the *same* `mn.data`, so they cannot differ at one instant — but
`mn.data` is mutated asynchronously by "Try another table" (a content-script
`getTableData` plus a server-processing round-trip), so a table validated a
moment ago can shift before the XLSX click. Observed live on a 4-user page:
the preview was 4/4 but the export held one user plus page chrome. So:

* **before export** — the preview is re-read immediately before the XLSX
  click and must still equal the DOM identities;
* **after export** — the exported file's identities must **exactly equal** the
  page's DOM identities (a set, not a count), which rejects an export whose
  model diverged from the validated preview.

Either guard failing yields `INCOMPLETE`, preserves the evidence, and does not
advance the search.

**Single-page collection enforces the same exact count.** A page GitHub says
holds N users is not complete until N *usable* identities are collected — the
same guard multi-page uses. If GitHub reports "5 result(s), no next page" but
Instant Data Scraper yields only 3 usable rows, the result is `INCOMPLETE`
(the export is preserved as `<slug>.partial.xlsx`, `collection.complete` stays
False, the search does not advance), not `[OK] 3 rows`. Two counts are tracked
and both reported: **raw** rows exported and **usable** identities parsed from
them — a row that yields no GitHub username or profile URL does not count
toward completeness.

**Completeness is verified two ways**, and which one was used is recorded:

| Mode | When | Rule |
|---|---|---|
`exact-total` | GitHub states a small exact figure (`19 users`) | unique rows must **equal** it — a shortfall *and* an excess both refuse |
`page-completeness` | figure abbreviated (`9.1k`), absent, or above GitHub's 1,000-result cap | every reachable page through the last must be exported |

A merge that fails verification is deleted, so no unverified file can be
mistaken for a finished collection.

**Retries restart from page 1** and archive the previous attempt's staging to
`<slug>.attempt_<timestamp>/` first, so yesterday's `page_002` can never be
merged with today's `page_001`.

**Pagination must be proved, not assumed.** If GitHub offered a live Next
before the crawl, the crawl may only be called complete when both hold:

| Guarantee | Catches |
|---|---|
`transitions >= 1` | a crawler that never navigated at all |
`extension pages >= browser page` | a crawler that navigated but did **not scrape** the page it reached |

The second exists because of a real loss: a 2-page, 19-user search was
accepted after 10 rows and one page. The browser *had* reached `&p=2` — so a
transition check alone would have passed it — but the extension reported
`Pages scraped: 1` and `Table not found`, with nine users visible on screen
and uncollected. From the page alone, "the last page, with results on it" is
indistinguishable from "the extension failed to scrape the last page"; only
the crawl's own history separates them.

Failing either check yields `INCOMPLETE`: the rows already collected are kept
as `<search>.partial.xlsx`, the real XLSX is **not** written, the search unit
is **not** marked complete, and nothing is deleted. The extension's own table
recovery is attempted first, since it reports "continue crawling" in exactly
this state.

A page that cannot be inspected at all (dead renderer, detached frame) is
`TABLE_LOST`, never "last page reached" — previously an uninspectable page
read as a completed crawl, a second route to silent truncation.

Per-transition logging is concise — one line per page turn, not per poll:

```
[CRAWL] before start: page 1, rows 0, Next present: yes
[CRAWL] transition 1: page 1 -> 2, rows 10 -> 20
[CRAWL] transition 2: page 2 -> 3, rows 20 -> 26
[CRAWL] final: Next present: no, pages visited 3, transitions 2, extension pages 3, rows 26
```

**Interrupted downloads.** Chromium writes a download in full and only then
renames it, so a renderer that dies mid-save leaves a complete
`.crdownload`. Rather than throw away a finished crawl, the export recovers
that file and validates it before use. The scan is limited to the directory of
the destination it was given — this folder's `keyword_downloads/` — so one
copy can never adopt the export another copy just finished. (Copies are
separate checkouts, so that is already true of the filesystem; passing the
directory explicitly keeps it true of the code.)

**Ctrl+C.** Affects only the copy it was sent to: it saves its own state,
preserves its own search XLSX and unresolved work, releases its own claims and
closes its own Chromium. `current_row.json` still points at the interrupted
search, so the same command resumes it.

**Cleanup stays in the folder.** A copy removes only its own `temporary.txt`,
`temp_downloads/` and completed search XLSX, and only under the success rule
above.

**Retrying one search.** When a search was accepted as collected but the
collection is known to be wrong, `--retry-current-search` re-collects exactly
that one search:

```bash
python3 student_scraper.py --retry-current-search
python3 student_scraper.py --retry-current-search --apply
```

It reports first and changes nothing without `--apply`. With `--apply` it
clears **only** that search's `collection` and `processing` blocks, in **one
atomic StateManager write** — two writes would briefly leave a state on disk
claiming rows from a file that had already been moved — and moves the XLSX
(and any `.partial`) into `retry_backup/` with a timestamp, rather than
deleting it.

Untouched: `current_row.json`, the search identity, `students.db`,
`current_data.xlsx` and every cycle file, the history index, the Playwright
profile, learned pagination selectors, transient failures and skipped
accounts. This is a retry of the current search, not a rewind of the list.

People already saved from that search stay in the store, so re-collection
skips them as duplicates and only genuinely new people cost anything.

Never edit `scraper_state.json` or `current_row.json` by hand: it bypasses the
atomic write.

## Paths

Everything one process writes sits **next to `student_scraper.py`**, in the
folder that copy was cloned into. There is no `runtime/rank_NNN/` layering,
because a rank is not a subdirectory — it is a whole checkout. Two copies can
no more share a path than two separate clones can.

| What | Where |
|---|---|
| Browser profile | `$TMPDIR/gpe_profile_<hash>` — disposable, deleted at shutdown |
| **GitHub session** | `github_auth.json` (0600, gitignored) |
| Resume point | `current_row.json` (atomic writes) |
| Current search's state | `scraper_state.json` (atomic writes) |
| **Collected output** | `Student Data original/current_data.xlsx` |
| **Sealed batches** | `Student Data original/current_data_cycleN.xlsx` |
| Cycle counter | `cycle_state.json` |
| Duplicate index | `students.db` |
| Search XLSX (temporary) | `keyword_downloads/` |
| Page-by-page staging | `crawl_staging/<search-slug>/` |
| Current person's raw text | `temporary.txt` (+ `.json`) |
| Resume downloads (temporary) | `temp_downloads/` |
| Transient failures | `transient_failures.json` |
| Deferred profiles | `deferred_profiles.json` |
| Throttle / cooldown | `throttle_state.json`, `rate_limit_state.json` |
| Permanent skips | `skipped_accounts.txt` |
| Set aside by `--retry-current-search` | `retry_backup/` |
| Log | `logs/student_scraper.log` |
| Duplicate index | `Student Data original/master_index.xlsx` |
| Merged workbook (on demand) | `Student Data original/student_data.xlsx` |

`config.py` holds every one of these as a module-level constant, so there is
one place to look and nothing derives a path at run time.

## Instant Data Scraper

Discovered by inspecting the installed extension, not assumed:

* id `ofaokhiedipichpaobibbnahnkdoiiah`, version 1.7.1, manifest v3
* the manifest has **no `commands` key**, so `Ctrl+G` can only be a
  browser-level `_execute_action` binding, which `page.keyboard.press()`
  cannot reach. `Ctrl+G` is still attempted first; the fallback opens the
  exact window `background.js` would have created:
  `chrome-extension://<id>/src/popup.html?tabid=<id>&url=<page>`
* control ids read out of `src/popup.js`: `#startScraping`, `#stopScraping`,
  `#nextButton`, `#wrongTable`, `#xlsx`, `#stats`, `#instructions`, `#error`

### Pagination

The extension needs a "Next" selector before it can crawl past page 1. It is
learned **once** automatically (`Locate "Next"` + a click on GitHub's own
Next control) and mirrored into `scraper_state.json`, so it survives even a
wiped browser profile. No manual step is ever required.

### Knowing when a crawl finished

Never one signal. All of these must hold together:

* `#stopScraping` gone (rendered only while the extension is crawling)
* `#startScraping` back
* `#instructions` says `Crawling stopped`
* pages **and** rows unchanged for 30 s

A frozen network looks like "no change" but never clears the crawling flag,
so a stall cannot be mistaken for success. Rate limits, CAPTCHAs and 5xx
stop the run safely; none of them are bypassed.

## Memory

Measured on this machine with `tests/resource_probe.py` (8 students, real
profiles, extension loaded):

| | at start | after 8 students |
|---|---|---|
| Python RSS | 47 MB | 48.6 MB |
| Chromium PSS | 394 MB | ~935 MB (plateauing) |
| Browser pages | 1 | 1 |
| Open file descriptors | 10 | 10 |
| `temp_downloads/` | 0 | 0 |

**No leak.** Python RSS, page count, descriptors and temp files are flat;
Chromium's growth is cache and levels off. The cache is capped at **50 MB**
(`--disk-cache-size`) and the whole profile is deleted at shutdown, so it is
disk held during a run, not disk kept. It is also not in this folder: an
earlier build left the profile here and it reached **1 GB**.

The search list adds about **45 MB**: 603,750 `(city, spelling, keyword)`
tuples at `num_sys = 5`, built in 0.3 s at startup. They are tuples rather
than finished `SearchUnit` objects precisely so the strings are shared --
175 city objects and 3,450 keyword strings, referenced 603,750 times between
them. `make_unit()` builds the object for one search when it is that
search's turn.

`LOW_MEMORY_MODE = True` (config.py) does three things, all measured:

* closes surplus browser pages — Instant Data Scraper opens its own options
  page and leaves it open for the whole session, costing a renderer
* empties the working tab every `RELEASE_PAGE_EVERY` students so one heavy
  profile does not stay resident
* adds `--renderer-process-limit=3` and `--disable-back-forward-cache`
  (13 → 11 processes, about 60 MB off the RSS sum)

No forced garbage collection: profiling showed Python is already flat.
Site Isolation stays enabled and nothing the extension needs is disabled.

```bash
python3 tests/resource_probe.py --students 8 --extension
```

## Output schema

**65 columns, exactly as specified — unchanged.** Alias columns (`country`,
`pin_code`, `experience`, `linkedin`, `github`, …) always mirror their
canonical column.

Search metadata never touches them. `rank`, `search_country/state/city/query`
and `source_xlsx` are **columns of the store**, not of the workbook, so
`Data Source` and `Remarks / Notes` keep exactly the meaning they always had.

### Search location is not the student's location

```
"Fresher Hyderabad"  ->  where we LOOKED
Current City         ->  where the person says they are
```

Somebody found by `Fresher Hyderabad` whose profile says Pune is a **Pune**
record. Somebody with no location evidence at all gets a **blank**
`Current City` — never the city that happened to find them. `Current City`,
`State` and `Country` come only from the profile, resume or portfolio, via
`location_extractor`, which does not import the search dataset.

## Output files

### While collecting: current_data.xlsx and its cycle files

```
Student Data original/
    current_data.xlsx            <- appended to, one row per person
        reaches 100 rows?
            copy  ->  current_data_cycle1.xlsx    (sealed, never reopened)
            empty ->  current_data.xlsx           (keeps going)

    current_data_cycle1.xlsx, current_data_cycle2.xlsx, ...
    student_data.xlsx            <- the merged workbook, on demand
```

The collected rows sit with the rest of the output rather than loose at the
top of the project. `cycle_state.json` -- the counter -- deliberately stays
beside the script: it is state, not output, so a folder you copy away or
clear never takes the counter with it. If it did, numbering would restart at
1 and the next rollover would overwrite a batch you had already taken.

This is the same output contract the job fetching scripts already use, so
whatever collects from them can collect from this folder without being told
anything new.

* **A sealed file is finished.** `current_data.xlsx` is being written to; a
  `cycleN.xlsx` never is. That is what lets a collector copy one away at any
  moment. A collector should therefore look only for names ending
  `cycleN.xlsx` and leave `current_data.xlsx` alone.
* **Every write is atomic** — built in full, saved to a sibling temp file,
  then `os.replace`d. A crash leaves either the previous workbook or the new
  one, which matters because the previous one is the only copy of up to 100
  people.
* **The cycle number lives in `cycle_state.json`,** not in a directory
  listing. Deriving it from the directory would restart at 1 as soon as a
  collector removed the files it had taken, and `current_data_cycle1.xlsx`
  would then be written twice — the second silently replacing a batch of 100
  people that had already been handed over. If a computed name is somehow
  still on disk, the number advances rather than overwriting it.
* **An unreadable workbook is moved aside, not deleted.** A corrupt
  `current_data.xlsx` still holds up to 100 real people; it is renamed
  `…​.corrupt_<timestamp>` so the loss is visible rather than silent.
* All 65 columns are written, in `config.FINAL_COLUMNS` order.

It all lives in the CYCLE WRITER section of `student_scraper.py`, and
`tests/test_cycle_writer.py` covers
the rollover, the numbering, the sealing and the damaged-input case.

### On demand: one merged workbook

```
Student Data original/student_data.xlsx      one file, every person
```

* Written **only** by `--export-final`, from the stores you name. See the
  warning in **Several laptops** above: it reads `students.db`, which is no
  longer written.
* Atomic: written to a sibling temp file, fsynced, then `os.replace`d. An
  interrupted export cannot corrupt the existing workbook.
* Refuses while a LibreOffice lock is active. A lock is **never** removed
  automatically — not even when it looks stale — because the cost of being
  wrong is somebody's spreadsheet.
* De-duplicated globally by normalised username, then profile URL, then real
  email. Placeholder addresses (`test@test.com`, `your_email@example.com`, …)
  never establish identity.
* Conflicts resolved deterministically: more populated fields wins, then
  earlier `collected_at`, then lower rank. The same inputs always give the
  same workbook.
* An xlsx worksheet holds 1,048,576 rows, one being the header. Exceeding it
  **stops with a message** and writes nothing; it never truncates silently.

There is **no** rows-per-file rotation: `ROWS_PER_FILE`, `_target_file()`,
the numbering-gap logic and `RotationError` are **removed**, not switched off,
so nothing half-dead remains to be re-enabled by accident.

## Tests

```bash
python3 tests/run_all.py            # every file, one summary
python3 tests/run_all.py -v         # with each file's own output
```

The suite runs in about **75 seconds**. It used to take ten minutes, because
several files drive the real `Workflow` with only the browser stubbed and so
executed the real pacing delays — and those are deliberately slow (a results
page every 1–3 minutes). `sandbox.enter()` now collapses the pacing and the
scheduled stops for the duration of a test, so the delays are exercised as
*decisions* without being waited out: `test_lifecycle.py` went from 355 s to
2 s. A test that cares about the production numbers themselves reads them
through `sandbox.pristine()`.

There is no pytest on this machine and `tests/` is not a package, so each file
is also a standalone script:

```bash
python3 -m py_compile *.py
python3 tests/test_locations.py        # aliases; every spelling of a city
                                       # stays with ONE rank; search city
                                       # must not leak into a person's record
python3 tests/test_cycle_writer.py     # current_data.xlsx, the 100-row
                                       # rollover, cycle numbering, sealing
python3 tests/test_result_store.py     # claims, a real multi-process race
python3 tests/test_final_export.py     # one workbook, 65 columns, atomic, locks
python3 tests/test_lifecycle.py        # crash / resume / cleanup safety rule
python3 tests/test_workflow.py         # duplicates, XLSX reading, state
python3 tests/test_extraction.py       # regex + normalisation, +/- cases
python3 tests/test_merge.py            # sources -> temporary.txt -> final row
python3 tests/test_hardening.py        # live-bug regressions, locks
python3 tests/test_browser_recovery.py # the browser dying mid-run
python3 tests/test_schedule_and_shuffle.py
                                       # the keyword order, the night stop,
                                       # the work/break cycle, the pacing
```

`tests/sandbox.py` is the shared fixture. `sandbox.enter(tmpdir)` points every
per-folder path in `config.py` at a throwaway directory for one test and
`sandbox.restore()` puts them back, so no test can write into the real project
folder. It also builds `SearchUnit`s through the **same** `make_unit()` and
`get_cities_by_rank()` production uses, so a test cannot pass against an
ordering production does not produce.

Real Instant Data Scraper exports live in `tests/fixtures/` so the suite does
not depend on anything in `~/Downloads`.

## Scope

Only publicly visible information relevant to education, skills, career and
professional contact is collected. Government IDs, financial details and
private data are never sought and are ignored if they appear incidentally.
Logins are never bypassed and private resumes are skipped with a remark.
