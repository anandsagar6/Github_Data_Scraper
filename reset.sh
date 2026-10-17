#!/bin/sh
# reset.sh -- put this clone back into its freshly-cloned state.
#
# After this runs, the next `python3 student_scraper.py` starts from SEARCH 1
# with an empty store: no resume point, no collected rows, no duplicate
# index, no logs, no half-downloaded workbooks.  That is the whole point --
# it is a FRESH RUN, not a continuation.
#
# WHAT IS CLEARED  (everything a run creates and can create again)
#
#   state beside the script    scraper_state.json  current_row.json
#                              cycle_state.json  keyword_cycle.json
#                              deferred_profiles.json
#                              transient_failures.json  throttle_state.json
#                              rate_limit_state.json  temporary.txt
#                              temporary.json  skipped_accounts.txt
#                              failed_profiles.txt
#   the result store           students.db (+ -wal, -shm, -journal) -- the
#                              duplicate index, so everyone met from here on
#                              is collected again
#   atomic-write leftovers     *.tmp and the hidden .<name>.tmp siblings
#   collected output           "Student Data original"/current_data.xlsx,
#                              current_data_cycleN.xlsx, student_data.xlsx,
#                              master_index.xlsx and the .corrupt_/.partial/
#                              .verify/.probe spill files
#   directories EMPTIED,       temp_downloads/  keyword_downloads/  logs/
#   the folder itself kept     and "Student Data original"/.  The contents
#                              go; the folder stays, because deleting it
#                              hands it back with WIDER permissions on this
#                              clone -- section 4 has the measurement.
#   directories DELETED        crawl_staging/  retry_backup/  history/
#                              imported/  __pycache__/  tests/__pycache__/
#                              .pytest_cache/  .pw_profile/
#   the browser profile        ($TMPDIR|/var/tmp|/tmp)/gpe_profile_<hash>,
#                              which is where Instant Data Scraper keeps its
#                              learned "nextSelector:<host>" pagination
#                              selector as well as the Chromium cache
#
# WHAT IS KEPT, AND WHY
#
#   code and configuration     student_scraper.py  config.py  new_mylib.py
#                              (the keywords and the cities)  requirements.txt
#                              README.md  Documentation/  tests/  this file
#   github_auth.json           it IS the GitHub login and holds no crawl
#                              position.  Deleting it costs an interactive
#                              --login on a machine with a display.
#   rank.json                  which rank this laptop is.  Configuration,
#                              not crawl state -- deleting it would quietly
#                              turn laptop 3 back into laptop 1 and make two
#                              machines crawl the same cities.
#   github_profile.txt         a captured profile tests/test_merge.py reads.
#   "Student Data original"/student_data_<N>.xlsx
#                              the pre-redesign workbooks.  They hold real
#                              collected students, legacy_workbooks() only
#                              ever READS them, and nothing in this project
#                              can create them again -- so a reset that
#                              deleted them would destroy data for good.
#                              The older reset.sh did exactly that.
#   *.xlsx.legacy, backup_*/   same reasoning: backups of real data that no
#   loose *.xlsx at the top    run writes and no run can rebuild.  Nothing
#                              reads them either, so keeping them does not
#                              make the next run any less fresh.  They are
#                              listed at the end so they are never a
#                              surprise -- delete them by hand if you want.
#   .~lock.<name># files       LibreOffice's own locks.  This project never
#                              deletes them; they are reported instead.
#
# USAGE
#
#     cd /home/anand/Script1 && ./reset.sh        # do it
#     cd /home/anand/Script1 && ./reset.sh -n     # show it, change nothing
#
# ONE CLONE AT A TIME.  Each folder is its own rank with its own state, so
# resetting Script1 does nothing to Script2 or Script3.
set -e

DRY=0
for arg in "$@"; do
    case "$arg" in
        -n|--dry-run) DRY=1 ;;
        -h|--help)
            # The comment block above IS the help text, printed by reading
            # this file.  A fixed `sed -n '2,70p'` went stale the first time
            # a line was added to the header; stopping at the first line
            # that is not a comment cannot.
            awk 'NR > 1 { if ($0 !~ /^#/) exit; sub(/^# ?/, ""); print }' "$0"
            exit 0
            ;;
        *)
            echo "reset: unknown option '$arg' (try -n, --dry-run, --help)" >&2
            exit 2
            ;;
    esac
done

# THE TARGET IS THE DIRECTORY YOU ARE STANDING IN, not where this file
# lives.  It started out sitting inside each clone, so dirname "$0" was the
# clone; installed on PATH as `reset` that would resolve to ~/.local/bin and
# try to empty that instead.  $PWD is what "reset this clone" means.
HERE=$(pwd)
# The same folder with every symlink resolved.  config.py builds PROJECT_ROOT
# from os.path.abspath(__file__), which keeps a symlinked path as given but
# becomes the physical path when the script is started as a bare filename --
# so this clone can own a Chromium profile under either spelling, and both
# are checked below.
HERE_P=$(pwd -P)

# Refuse anywhere that is not a clone, so a mistyped cd cannot empty a
# directory that merely happens to have a folder of the same name.
if [ ! -f "$HERE/student_scraper.py" ] || [ ! -f "$HERE/config.py" ]; then
    echo "REFUSING: $HERE is not a scraper clone"
    echo "(no student_scraper.py / config.py here)."
    echo "cd into the clone first, for example:  cd ~/Script2 && reset"
    exit 1
fi

# -- 1. refuse to run under a live crawl ----------------------------------
# Deleting state from under a running process is pointless: it rewrites the
# files on exit and the reset is silently undone.
# A plain `pgrep -f student_scraper.py` is WRONG here: it matches anything
# whose command line merely contains the name -- an editor with the file
# open, a grep, even the shell that launched this script.  The test for a
# real crawl is a PYTHON process that is running THIS clone's script, which
# is true when its working directory is this folder (how a bare
# `python3 student_scraper.py` is started) OR when this folder's own script
# path is one of its arguments (how a crawl launched from elsewhere, or by a
# systemd unit or a cron line, is started).  Matching the cwd alone -- which
# is all the previous version did -- let `cd ~ && python3 ~/Script1/student_
# scraper.py` run straight through the guard.
#
# No pgrep means no check, and a silent no-check on a destructive script is
# the wrong kind of quiet -- so it says so.
if ! command -v pgrep >/dev/null 2>&1; then
    echo "reset: WARNING -- pgrep not found, cannot check for a live crawl."
    echo "reset: make sure student_scraper.py is stopped here first."
fi
for pid in $(pgrep -f "student_scraper" 2>/dev/null); do
    [ "$pid" = "$$" ] && continue
    [ "$pid" = "$PPID" ] && continue
    exe=$(basename "$(readlink "/proc/$pid/exe" 2>/dev/null)" 2>/dev/null)
    case "$exe" in
        python*) ;;
        *) continue ;;
    esac

    mine=""
    cwd=$(readlink "/proc/$pid/cwd" 2>/dev/null) || cwd=""
    if [ -n "$cwd" ] && { [ "$cwd" = "$HERE" ] || [ "$cwd" = "$HERE_P" ]; }
    then
        # A cwd match still has to be a crawl and not, say, a `python3 -c`
        # that happens to sit here, so the script name must be an argument.
        if tr '\0' '\n' < "/proc/$pid/cmdline" 2>/dev/null \
                | grep -q '/\{0,1\}student_scraper\.py$'; then
            mine=1
        fi
    fi
    if [ -z "$mine" ] \
        && tr '\0' '\n' < "/proc/$pid/cmdline" 2>/dev/null \
            | grep -Fqx -e "$HERE/student_scraper.py" \
                        -e "$HERE_P/student_scraper.py"; then
        mine=1
    fi
    [ -n "$mine" ] || continue

    echo "REFUSING: student_scraper.py is running here (pid $pid)."
    echo "Stop it with Ctrl+C -- not kill -9 -- and run this again."
    exit 1
done

# -- 2. how a path is removed ---------------------------------------------
REMOVED=0

# Reported relative to the clone, so the list reads as a list of names and
# not as forty copies of the same long absolute prefix.  Anything outside
# the clone -- the Chromium profile -- keeps its full path, because that is
# the surprising one and it should look like it.
short() {
    case "$1" in
        "$HERE"/*) printf '%s' "${1#"$HERE"/}" ;;
        *) printf '%s' "$1" ;;
    esac
}

gone() {
    # A dangling symlink is not -e, so -L is tested too.  An unmatched glob
    # arrives here as its own literal text and fails both, which is how the
    # loops below stay quiet when there is nothing to delete.
    if [ ! -e "$1" ] && [ ! -L "$1" ]; then
        return 0
    fi
    # Every caller builds "$HERE/<name>" or the profile path, so none of
    # these can happen -- which is the point.  An `rm -rf` that is one unset
    # variable away from "/" or "$HOME" is worth the four lines.
    case "$1" in
        ""|"/"|"$HOME"|"$HOME"/)
            echo "reset: INTERNAL ERROR, refusing to remove '$1'" >&2
            exit 1
            ;;
    esac
    if [ "$DRY" = 1 ]; then
        printf '  would remove  %s\n' "$(short "$1")"
    else
        rm -rf -- "$1"
        printf '  removed       %s\n' "$(short "$1")"
    fi
    REMOVED=$((REMOVED + 1))
}

# Throw away a directory's contents and keep the directory itself, so its
# mode and its ACL survive untouched.  See section 4 for why that matters.
# -mindepth 1 is what spares the directory; -delete implies -depth, so a
# subtree empties from the leaves up.
empty_out() {
    if [ -z "$(find "$1" -mindepth 1 -print -quit 2>/dev/null)" ]; then
        return 0        # already empty; nothing to report
    fi
    if [ "$DRY" = 1 ]; then
        printf '  would empty   %s/\n' "$(short "$1")"
    else
        find "$1" -mindepth 1 -delete
        printf '  emptied       %s/\n' "$(short "$1")"
    fi
    REMOVED=$((REMOVED + 1))
}

# -- 3. state files beside the script -------------------------------------
# students.db-wal and -shm exist because the store runs in WAL mode
# (PRAGMA journal_mode=WAL); -journal is what an older rollback-mode file
# would be called.  Deleting students.db while leaving a -wal behind would
# hand the next run a store whose newest pages describe a database that is
# no longer there.
for f in \
    scraper_state.json \
    current_row.json \
    cycle_state.json \
    keyword_cycle.json \
    deferred_profiles.json \
    transient_failures.json \
    throttle_state.json \
    rate_limit_state.json \
    temporary.txt \
    temporary.json \
    skipped_accounts.txt \
    failed_profiles.txt \
    students.db \
    students.db-wal \
    students.db-shm \
    students.db-journal
do
    gone "$HERE/$f"
done

# Atomic writes land on "<dir>/.<basename>.tmp" and "<path>.tmp" and are
# renamed over the target, so an interrupted write leaves one behind.  The
# hidden spelling is the common one and a plain *.tmp glob never sees it.
for f in "$HERE"/*.tmp "$HERE"/.*.tmp; do
    gone "$f"
done

# -- 4. working directories ------------------------------------------------
# EMPTIED, NOT DELETED -- and that is a permissions decision, not a style
# one.  MEASURED on this clone: it carries a default POSIX ACL
#
#     $ getfacl .
#     default:user:scripts:rwx
#     default:group:abhinav_shared:rwx
#     default:mask::rwx
#
# and a default ACL OVERRIDES the umask for the group bits.  So a directory
# that this script deletes does not come back the way it left:
#
#     before reset   drwx------+   logs, temp_downloads, keyword_downloads,
#                                  "Student Data original"   (mask ---)
#     recreated      drwxrwx---+   mask rwx, so user:scripts and
#                                  group:abhinav_shared get rwx back
#
# PROVED by creating one under the scraper's own startup settings:
#
#     $ python3 -c 'import os; os.umask(0o077); os.mkdir("probe")'
#     $ ls -ld probe   ->   drwxrwx---+
#
# os.umask(0o077) in main() cannot prevent that, and nothing in the run
# chmods afterwards -- so deleting these four would hand every account in
# abhinav_shared read access to collected names, emails and phone numbers,
# which is the one thing a reset must not quietly do.  Emptying them keeps
# the mode and the ACL bit-for-bit.
#
# The mode is LEFT ALONE rather than forced to 0700: a chmod here would
# collapse the ACL mask and revoke the `scripts` account's access, which
# somebody granted on purpose.  Reset's job is to remove data, not to
# re-decide who may read the folder.
#
# Nothing reads these as a signal, so an empty one is as good as an absent
# one: ensure_runtime_dirs() creates temp_downloads/, keyword_downloads/ and
# logs/ on every single startup, and ensure_dirs() does the same for
# "Student Data original"/ -- a directory that is always created cannot mean
# anything by existing.
#
# ("Student Data original"/ is handled on its own in section 5: it is the
# one folder here that holds files which must SURVIVE, so it is cleared by
# name instead of emptied wholesale.)
for d in \
    temp_downloads \
    keyword_downloads \
    logs
do
    if [ -d "$HERE/$d" ]; then
        empty_out "$HERE/$d"
    fi
done

# DELETED outright.  These are different: a run creates them only when it
# actually needs them -- staging_for() makes crawl_staging/<slug>/ for a
# multi-page search, --retry-current-search makes retry_backup/,
# --migrate-output makes imported/, and _connect() makes history/ -- so a
# fresh clone has none of them and the project prefers them absent.
# config.py says so about history/ in as many words: "creating the directory
# on every startup made an empty folder that no code ever used".
#
# The permissions argument above does not apply here.  These do not exist in
# a fresh clone, so whatever ACL they inherit when something next creates
# them is exactly what a fresh clone would have inherited too -- deleting
# them loses nothing that could be restored, and deleting beats emptying for
# the ones that hold scraped pages.
for d in \
    crawl_staging \
    retry_backup \
    history \
    imported \
    __pycache__ \
    tests/__pycache__ \
    .pytest_cache \
    .pw_profile
do
    gone "$HERE/$d"
done

# -- 5. collected output ---------------------------------------------------
# Name by name, NOT `find -delete` on the whole folder: student_data_<N>.xlsx
# lives here too and must survive (see the header).
OUT="$HERE/Student Data original"
if [ -d "$OUT" ]; then
    for f in \
        "$OUT"/current_data.xlsx \
        "$OUT"/current_data_cycle*.xlsx \
        "$OUT"/student_data.xlsx \
        "$OUT"/master_index.xlsx \
        "$OUT"/*.corrupt_* \
        "$OUT"/*.partial.xlsx \
        "$OUT"/*.verify.xlsx \
        "$OUT"/*.probe.xlsx \
        "$OUT"/*.tmp \
        "$OUT"/.*.tmp
    do
        gone "$f"
    done
    # The folder ITSELF stays, empty or not, for the permissions reason in
    # section 4: it is 0700 today, ensure_dirs() recreates it on every
    # startup, and the recreated one would be 0770 under this clone's
    # default ACL -- on the single folder that holds every collected
    # student's name, email and phone number.
fi

# -- 6. this clone's disposable Chromium profile --------------------------
# config.py puts it OUTSIDE the project, at
# <parent>/gpe_profile_<sha1(PROJECT_ROOT)[:12]>, so it is not reached by
# anything above.  A run discards it at first launch anyway, but a crawl
# killed with -9 leaves ~1 GB of cache and a SingletonLock sitting there.
#
# The name is recomputed here and only an EXACT match is deleted: a sibling
# clone's profile is a different hash in the same directory, and guessing
# with gpe_profile_* would take it out too.
#
# NOT touched, on purpose: ResumeParser's tempfile.mkdtemp(prefix="resume_")
# fallback.  A real run never reaches it -- line 12207 hands the parser
# C.TEMP_DOWNLOAD_DIR -- and a mkdtemp name carries nothing that ties it to
# one clone, so deleting /tmp/resume_* would be guessing at another clone's
# or another program's directory.  The system's own temp policy clears it.
if command -v sha1sum >/dev/null 2>&1; then
    # printf, not echo: hashlib.sha1(PROJECT_ROOT.encode()) hashes the path
    # with no trailing newline, and `echo` would append one and produce a
    # different digest -- a name that matches nothing, silently.
    TAG=$(printf '%s' "$HERE" | sha1sum | cut -c1-12)
    TAG_P=$(printf '%s' "$HERE_P" | sha1sum | cut -c1-12)
    # Usually the same folder under both spellings; hashing it twice would
    # report one profile twice and count it twice.
    if [ "$TAG_P" = "$TAG" ]; then
        TAG_P=""
    fi
    # All three candidates are checked rather than recomputing config.py's
    # "TMPDIR or /var/tmp or gettempdir()" choice, because the TMPDIR of the
    # shell that ran the crawl need not be the TMPDIR of this shell.  SEEN
    # stops TMPDIR=/tmp from reporting the same profile twice.
    SEEN=":"
    for parent in "$TMPDIR" /var/tmp /tmp; do
        if [ -z "$parent" ] || [ ! -d "$parent" ]; then
            continue
        fi
        case "$SEEN" in
            *":$parent:"*) continue ;;
        esac
        SEEN="$SEEN$parent:"
        gone "$parent/gpe_profile_$TAG"
        if [ -n "$TAG_P" ]; then
            gone "$parent/gpe_profile_$TAG_P"
        fi
    done
fi

# -- 7. what was deliberately left behind ---------------------------------
# Printed, not deleted.  Each one is either real data no run can rebuild or
# a file that belongs to another program; none of them makes the next run
# any less fresh, because nothing in a run reads them.
note_kept() {
    if [ -e "$1" ] || [ -L "$1" ]; then
        printf '  kept          %s  (%s)\n' "$(short "$1")" "$2"
    fi
}
for f in "$OUT"/student_data_*.xlsx; do
    note_kept "$f" "pre-redesign data, read-only input"
done
for f in "$HERE"/*.xlsx.legacy "$OUT"/*.xlsx.legacy; do
    note_kept "$f" "backup of real data, nothing recreates it"
done
for f in "$HERE"/backup_final_output_* "$HERE"/backup_legacy_migration_*; do
    note_kept "$f" "backup of real data, nothing recreates it"
done
for f in "$HERE"/student_data_*.xlsx "$HERE"/github*.xlsx; do
    note_kept "$f" "loose workbook at the top level, not written by a run"
done
for f in "$HERE"/.~lock.*# "$OUT"/.~lock.*#; do
    note_kept "$f" "LibreOffice lock -- close the file in LibreOffice"
done
note_kept "$HERE/github_auth.json" "the GitHub login; keeps --login unneeded"
# Per-machine CONFIGURATION, not crawl state: it says which rank this laptop
# is, and a reset must not silently turn laptop 3 back into laptop 1.
note_kept "$HERE/rank.json" "this laptop's rank; config, not crawl state"

# -- 8. prove the clone can still run -------------------------------------
# The one thing a reset must never do is take a file the run needs.  These
# are read at startup, every one of them is tracked in git except the login,
# and none of them is created by a run -- so if the list above ever grows a
# line it should not have, this check is what catches it.
MISSING=""
for f in \
    student_scraper.py \
    config.py \
    new_mylib.py \
    requirements.txt \
    github_profile.txt \
    tests/run_all.py
do
    [ -f "$HERE/$f" ] || MISSING="$MISSING $f"
done

if [ -n "$MISSING" ]; then
    echo "reset: FILES THE RUN NEEDS ARE MISSING:$MISSING" >&2
    echo "reset: restore them before running, e.g." >&2
    echo "reset:   git checkout --$MISSING" >&2
    exit 1
fi

if [ ! -f "$HERE/github_auth.json" ]; then
    echo "  note          no github_auth.json; sign in once with"
    echo "                python3 student_scraper.py --login"
fi

# "cleared", not "removed": the count mixes deleted paths with emptied
# directories, and three of the lines above say "emptied" for a reason.
if [ "$DRY" = 1 ]; then
    echo "reset: $(basename "$HERE") -- dry run, $REMOVED path(s) would be cleared"
    echo "       nothing was changed; drop -n to do it"
else
    echo "reset: $(basename "$HERE") -- fresh; $REMOVED path(s) cleared"
    echo "       next run starts at search 1 with an empty store"
fi
