#!/usr/bin/env python3
"""Resource probe: does anything grow per student?

Runs the real profile-processing loop against a small export and samples
browser page count, this process's RSS, open file descriptors and the
temp_downloads backlog at every checkpoint.

Exact flatness is NOT expected -- Chromium caches.  What matters is that
nothing grows monotonically with the number of students.

    python3 tests/resource_probe.py [--students N] [--extension]
"""

import argparse
import asyncio
import os
import random
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook                      # noqa: E402

import config as C                                 # noqa: E402


def rss_mb(pid=None):
    try:
        with open("/proc/%d/status" % (pid or os.getpid())) as handle:
            for line in handle:
                if line.startswith("VmRSS:"):
                    return int(line.split()[1]) / 1024.0
    except OSError:
        pass
    return -1.0


def open_fds():
    try:
        return len(os.listdir("/proc/self/fd"))
    except OSError:
        return -1


def temp_files():
    # There is no global temp dir any more: each rank downloads into its own
    # runtime/rank_NNN/temp_downloads.  Count across every rank's temp dir.
    root = C.RUNTIME_ROOT
    if not os.path.isdir(root):
        return 0
    total = 0
    for name in os.listdir(root):
        temp = os.path.join(root, name, "temp_downloads")
        if os.path.isdir(temp):
            total += len(os.listdir(temp))
    return total


def _pss_mb(pid):
    """Proportional set size: shared pages divided among their sharers.

    Summing RSS across Chromium's process tree double-counts every shared
    library and mapping, inflating the total several-fold.  PSS is the
    honest answer to "how much memory does this tree really cost".
    """
    try:
        with open("/proc/%d/smaps_rollup" % pid) as handle:
            for line in handle:
                if line.startswith("Pss:"):
                    return int(line.split()[1]) / 1024.0
    except OSError:
        pass
    return max(rss_mb(pid), 0.0)


def browser_memory():
    """(pss_mb, rss_sum_mb, process_count) for the Chromium tree."""
    pss = rss = 0.0
    count = 0
    for pid in os.listdir("/proc"):
        if not pid.isdigit():
            continue
        try:
            with open("/proc/%s/cmdline" % pid, "rb") as handle:
                cmd = handle.read().decode("utf-8", "replace")
        except OSError:
            continue
        if "ms-playwright" in cmd and "chrome" in cmd:
            value = _pss_mb(int(pid))
            if value > 0:
                pss += value
                rss += max(rss_mb(int(pid)), 0.0)
                count += 1
    return pss, rss, count


USERS = ["GouravRusiya30", "letscodedanish", "Sheetal0601", "khush1yaaar",
         "yugmalviya1", "yuktirapartiwar", "Malviyaaadarsh", "octocat",
         "torvalds", "gaearon"]


def make_export(path, count):
    book = Workbook()
    sheet = book.active
    sheet.append(["Users-module__avatar__NxLXK src", "prc-Link-Link-9ZwDx href",
                  "Users-module__titleLoginText__J6lmd",
                  "Users-module__titleName__lXTax"])
    # USERS is CYCLED, not sliced.  It used to be USERS[:count], so the
    # probe could never visit more than len(USERS) profiles however large
    # --students was -- which is why the only soak evidence on record was
    # eight students.  Repeats are fine for a leak soak: what is being
    # measured is navigations, renderers and handles per student, not who
    # the student is.  See the note in the verdict about what repeats do
    # and do not prove.
    for index in range(count):
        name = USERS[index % len(USERS)]
        sheet.append(["https://avatars.githubusercontent.com/u/1?v=4",
                      "https://github.com/%s" % name, name, name])
    book.save(path)


async def main(students, with_extension):
    from student_scraper import BrowserManager
    from student_scraper import GitHubProfile
    from student_scraper import SourceBundle
    from student_scraper import PortfolioParser
    from student_scraper import clean_url, extract_portfolio, extract_resume_urls
    from student_scraper import ResumeParser
    from student_scraper import StudentMerger
    from student_scraper import XlsxReader
    import student_scraper as temp_store

    scratch = os.environ.get("PROBE_DIR", "/tmp/resource_probe")
    os.makedirs(scratch, exist_ok=True)
    export = os.path.join(scratch, "probe.xlsx")
    make_export(export, students)

    samples = []

    def sample(label, context=None):
        pss, rss_sum, procs = browser_memory()
        row = {
            "label": label,
            "browser_rss": rss_sum,
            "procs": procs,
            "pages": len(context.pages) if context else 0,
            "rss": rss_mb(),
            "browser": pss,
            "fds": open_fds(),
            "temp": temp_files(),
        }
        samples.append(row)
        print("  %-28s pages=%d procs=%-3d py=%6.1fMB PSS=%7.1fMB "
              "RSSsum=%7.1fMB fds=%-3d temp=%d"
              % (row["label"], row["pages"], row["procs"], row["rss"],
                 row["browser"], row["browser_rss"], row["fds"],
                 row["temp"]), flush=True)
        if context and os.environ.get("PROBE_VERBOSE"):
            for open_page in context.pages:
                print("        page: %s" % open_page.url[:95])

    sample("start (no browser)")
    browser = BrowserManager(headless=True)
    await browser.start(with_extension=with_extension)
    context = browser.context
    sample("after browser start", context)

    candidates = XlsxReader().read(export, "probe")
    # XlsxReader DEDUPES BY USERNAME -- that is production duplicate
    # protection, so a cycled export collapses straight back to
    # len(USERS) candidates however many rows it has.  The repetition
    # therefore has to happen after the reader, not in the workbook.
    #
    # WHAT THIS DOES AND DOES NOT PROVE: repeated profiles fully exercise
    # the per-student machinery -- navigation, a fresh renderer, the DOM
    # parse, portfolio and resume fetches, page release -- which is where
    # a leak would live.  They do NOT reproduce the cache pressure of N
    # distinct profiles, because the second visit to a profile is partly
    # served from Chromium's cache.  Read a flat result as "the loop does
    # not leak", not as "N distinct profiles cost this much".
    if candidates and len(candidates) < students:
        distinct = len(candidates)
        base = list(candidates)
        while len(candidates) < students:
            candidates.append(base[len(candidates) % distinct])
        candidates = candidates[:students]
        print("  (%d distinct profile(s) repeated to %d students)"
              % (distinct, len(candidates)))
    sample("after reading export", context)

    page = await browser.page()
    await browser.close_extra_pages(keep=(page,))
    profiles = GitHubProfile(browser)
    portfolios = PortfolioParser(say=lambda *a, **k: None)
    resumes = ResumeParser(say=lambda *a, **k: None)
    merger = StudentMerger()

    for number, candidate in enumerate(candidates, 1):
        bundle = SourceBundle(candidate=candidate)
        try:
            fields, text = await profiles.fetch(page, candidate.username)
            bundle.github_ok = True
            bundle.github_fields = fields
            bundle.github_text = text
            bundle.github_links = list(fields.get("links", []))
        except Exception as exc:
            print("    (%s: %s)" % (candidate.username, str(exc)[:60]))

        website = bundle.github_fields.get("website", "")
        portfolio = extract_portfolio(bundle.github_links, website)
        if portfolio:
            bundle.portfolio_url = clean_url(portfolio)
            text, links, visited = await portfolios.fetch(
                context, bundle.portfolio_url)
            bundle.portfolio_text, bundle.portfolio_links = text, links
            bundle.portfolio_pages = visited

        urls = extract_resume_urls(
            "\n".join([bundle.github_text, bundle.portfolio_text]),
            bundle.github_links + bundle.portfolio_links)
        if urls:
            bundle.resume_url = urls[0]
            text, saved, _remark = await resumes.fetch(
                context, bundle.resume_url, candidate.username)
            bundle.resume_text, bundle.resume_file = text, saved

        temp_store.write(bundle)
        merger.build(bundle)
        temp_store.clear()
        resumes.discard(bundle.resume_file)
        del bundle
        # Same reclaim the production loop performs.
        if C.LOW_MEMORY_MODE:
            await browser.close_extra_pages(keep=(page,))
            if C.RELEASE_PAGE_EVERY and number % C.RELEASE_PAGE_EVERY == 0:
                await browser.release_page_memory(page)
        sample("after student %d/%d" % (number, len(candidates)), context)
        # Paced like production.  Without this the probe fetches back to
        # back, which is both unlike a real run and unkind to GitHub.
        if number < len(candidates):
            # The pacing lives at the top of student_scraper.py now, not in
            # config, so the probe reads it from there.
            import student_scraper as pacing
            await asyncio.sleep(random.uniform(pacing.PROFILE_DELAY_MIN,
                                               pacing.PROFILE_DELAY_MAX))

    sample("after profile processing", context)
    await browser.close()
    sample("after browser close")

    # ---- verdict
    print()
    per_student = [row for row in samples if row["label"].startswith("after student")]
    if len(per_student) >= 3:
        first, last = per_student[0], per_student[-1]
        print("  pages   : %d -> %d" % (first["pages"], last["pages"]))
        print("  py RSS  : %.1f -> %.1f MB (%+.1f)"
              % (first["rss"], last["rss"], last["rss"] - first["rss"]))
        print("  chromium: PSS %.1f -> %.1f MB (%+.1f) | RSS-sum %.1f -> %.1f"
              % (first["browser"], last["browser"],
                 last["browser"] - first["browser"],
                 first["browser_rss"], last["browser_rss"]))
        print("  fds     : %d -> %d" % (first["fds"], last["fds"]))
        print("  temp    : %d -> %d" % (first["temp"], last["temp"]))
        problems = []
        if last["pages"] > first["pages"]:
            problems.append("browser page count grew")
        if last["fds"] > first["fds"] + 5:
            problems.append("file descriptors grew")
        if last["temp"] > first["temp"]:
            problems.append("temp_downloads accumulated")
        growth = last["rss"] - first["rss"]
        if growth > 4 * len(per_student):
            problems.append("python RSS grew %.1fMB over %d students"
                            % (growth, len(per_student)))

        # IS CHROMIUM PLATEAUING, OR STILL CLIMBING?
        #
        # The verdict used to print Chromium's growth and then ignore it, on
        # the grounds that "Chromium caches" -- so a run could report "no
        # uncontrolled growth detected" while PSS went 394 MB -> 935 MB.
        # Caching does explain a rise that FLATTENS; it does not explain one
        # that keeps going.  So compare the two halves: a plateau spends most
        # of its growth early, a leak keeps spending it.
        half = len(per_student) // 2
        if half >= 3:
            mid = per_student[half]
            total = last["browser"] - first["browser"]
            late = last["browser"] - mid["browser"]
            print("  chromium trend: %+.1f MB over the first half, "
                  "%+.1f MB over the second"
                  % (total - late, late))
            if total > 50 and late > 0.55 * total:
                problems.append(
                    "chromium PSS is still climbing in the second half "
                    "(%+.1f of %+.1f MB) -- not a plateau" % (late, total))
        print()
        print("  VERDICT: %s" % ("; ".join(problems) if problems
                                 else "no uncontrolled growth detected"))
        return 1 if problems else 0
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--students", type=int, default=6)
    parser.add_argument("--extension", action="store_true",
                        help="load Instant Data Scraper, as production does")
    options = parser.parse_args()
    sys.exit(asyncio.run(main(options.students, options.extension)))
