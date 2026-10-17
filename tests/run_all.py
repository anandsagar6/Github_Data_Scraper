#!/usr/bin/env python3
"""Run every test file and report one summary.

There is no pytest on this machine, and tests/ is not a package, so each file
is a standalone script.  This runs them the same way the README does, in one
command, and exits non-zero if anything failed.

    python3 tests/run_all.py
    python3 tests/run_all.py -v        # show each file's own output
"""

import os
import re
import subprocess
import sys
import time

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
COUNT_RE = re.compile(r"^Ran (\d+) test", re.M)


def main(argv):
    verbose = "-v" in argv or "--verbose" in argv
    files = sorted(name for name in os.listdir(TESTS_DIR)
                   if name.startswith("test_") and name.endswith(".py"))
    if not files:
        print("no test files found in %s" % TESTS_DIR)
        return 1

    total_tests = 0
    failed_files = []
    started = time.time()
    print("=" * 62)
    print("TEST SUITE  (%d file(s))" % len(files))
    print("=" * 62)

    for name in files:
        path = os.path.join(TESTS_DIR, name)
        begin = time.time()
        result = subprocess.run([sys.executable, path],
                                capture_output=True, text=True)
        output = (result.stdout or "") + (result.stderr or "")
        match = COUNT_RE.search(output)
        count = int(match.group(1)) if match else 0
        total_tests += count
        ok = result.returncode == 0
        if not ok:
            failed_files.append(name)
        print("  %-30s %4d test(s)  %5.1fs  %s"
              % (name, count, time.time() - begin, "OK" if ok else "FAIL"))
        if verbose or not ok:
            for line in output.splitlines():
                print("      %s" % line)

    print("-" * 62)
    print("  %-30s %4d test(s)  %5.1fs"
          % ("TOTAL", total_tests, time.time() - started))
    if failed_files:
        print("")
        print("FAILED FILES (%d): %s"
              % (len(failed_files), ", ".join(failed_files)))
    print("RESULT: %s" % ("PASS" if not failed_files else "FAIL"))
    return 0 if not failed_files else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
