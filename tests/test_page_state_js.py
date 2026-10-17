"""PAGE_STATE_JS must EXECUTE, not merely parse.

THE BUG THIS FILE EXISTS FOR
----------------------------
    [FAIL] the page could not be inspected:
    Page.evaluate: ReferenceError:
        Cannot access 'resultsList' before initialization

PAGE_STATE_JS declared `resultsList` with `const` near the end, but a
`const region = resultsList || mainRegion` earlier in the same function read
it first.  Using a `const` before its declaration is a temporal-dead-zone
ReferenceError -- at RUNTIME.  It PARSES cleanly, which is exactly why an
earlier node syntax-only check ("const f = <js>; console.log('parses')")
missed it: the function was never called.

So this test CALLS the probe.  It runs the real PAGE_STATE_JS, verbatim from
the module, inside a JS engine (node) against synthetic DOMs, and fails if it
throws.  A parse check could never have caught a TDZ error; only execution
can.

If node is unavailable the JS-execution tests skip (never silently pass), and
a pure-Python guard still asserts every declaration precedes its use.
"""

import os
import re
import shutil
import subprocess
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import student_scraper                                           # noqa: E402


NODE = shutil.which("node") or shutil.which("nodejs")

# A synthetic-DOM harness that CALLS the probe once per scenario and prints a
# one-line result or the exception.  Kept minimal: it models only what
# PAGE_STATE_JS actually touches.
HARNESS = r"""
function makeDom(opts) {
  opts = opts || {};
  const users = opts.users || [];
  const rows = users.map(u => ({
    getAttribute: (k) => k === 'href' ? '/' + u : null,
    closest: () => null,
  }));
  const nav = (opts.navUsers || []).map(u => ({
    getAttribute: (k) => k === 'href' ? '/' + u : null,
    closest: () => ({}),
  }));
  const resultsList = opts.hasList ? {
    children: rows, innerHTML: 'x'.repeat(opts.sig || 100),
  } : null;
  const main = opts.hasMain !== false ? {
    children: rows, innerHTML: 'x'.repeat(opts.sig || 100),
  } : null;
  const nextEl = opts.hasNext ? {
    getAttribute: () => opts.nextDisabled ? 'true' : null,
    hasAttribute: () => false,
  } : null;
  const byShape = {};
  (opts.nextShapes || []).forEach(s => byShape[s] = 1);
  return {
    readyState: 'complete', visibilityState: 'visible',
    title: opts.title || 'User search results',
    body: { innerText: opts.body || '' },
    querySelector: (sel) => {
      if (sel.indexOf('results-list') !== -1 && sel.indexOf(',') === -1)
        return resultsList;
      if (sel === 'a[rel="next"]') return nextEl;
      if (sel.indexOf('main') !== -1 || sel.indexOf('role="main"') !== -1)
        return resultsList || main;
      if (sel.indexOf('form[action') !== -1) return opts.login ? {} : null;
      return null;
    },
    querySelectorAll: (sel) => {
      if (sel === 'a[href]') return rows.concat(nav);
      if (byShape[sel]) return [{}];
      return [];
    },
  };
}
function run(label, opts) {
  global.document = makeDom(opts);
  global.location = {
    href: opts.url || 'https://github.com/search?q=x&type=users',
    search: (opts.url || '?q=x&type=users').replace(/^[^?]*\??/, '?'),
    pathname: '/search',
  };
  global.URLSearchParams = class {
    constructor(s) { this.s = s || ''; }
    get(k) {
      if (k === 'type') return 'users';
      const m = this.s.match(new RegExp(k + '=(\\w+)'));
      return m ? m[1] : null;
    }
  };
  try {
    const out = PROBE();
    process.stdout.write('OK ' + label + ' ' + JSON.stringify({
      children: out.resultsListChildren,
      users: out.resultProfileLinkCount,
      total: out.reportedTotalUsers,
      noResults: out.noResults,
      login: out.loginForm,
      domSig: out.domSignature,
      skeleton: out.skeletonCount,
    }) + '\n');
  } catch (e) {
    process.stdout.write('THREW ' + label + ' ' + e.name + ': '
      + e.message + '\n');
  }
}
run('ten', {hasList:true, users:['a','b','c','d','e','f','g','h','i','j'],
     navUsers:['settings','me'], hasNext:true, nextShapes:['a[rel="next"]'],
     body:'Users 19 (19) results', sig:500});
run('nine', {hasList:true, users:['k','l','m','n','o','p','q','r','s'],
     hasNext:true, nextDisabled:true, body:'Users 19 (19) results',
     url:'?q=x&type=users&p=2', sig:400});
run('zero', {hasList:false, hasMain:true, users:[],
     body:'We couldn’t find any users matching your search'});
run('nocontainer', {hasList:false, hasMain:false, users:[], body:'nothing'});
run('login', {hasList:false, hasMain:true, login:true,
     body:'Sign in to GitHub'});
run('abbrev', {hasList:true, users:['a','b'], hasNext:true,
     body:'Users 9.1k (9.1k) results'});
"""


def run_probe_in_node():
    script = "const PROBE = %s;\n%s" % (student_scraper.PAGE_STATE_JS, HARNESS)
    result = subprocess.run([NODE, "-e", script], capture_output=True,
                            text=True, timeout=30)
    return result


@unittest.skipUnless(NODE, "node is not installed; JS execution not checked")
class TestPageStateJsExecutes(unittest.TestCase):
    """The probe must RUN against real DOMs, not merely parse."""

    @classmethod
    def setUpClass(cls):
        cls.result = run_probe_in_node()
        cls.lines = {}
        for line in (cls.result.stdout or "").splitlines():
            parts = line.split(" ", 2)
            if len(parts) >= 2:
                cls.lines[parts[1]] = line

    def test_the_probe_runs_without_a_runtime_error(self):
        self.assertEqual(self.result.returncode, 0,
                         self.result.stderr or self.result.stdout)
        self.assertNotIn("THREW", self.result.stdout,
                         "PAGE_STATE_JS threw at runtime:\n%s"
                         % self.result.stdout)

    def test_no_reference_error_before_initialization(self):
        """The exact failure mode, named."""
        self.assertNotIn("ReferenceError", self.result.stdout)
        self.assertNotIn("before initialization", self.result.stdout)

    def test_a_valid_ten_result_page_is_read(self):
        self.assertIn("ten", self.lines, self.result.stdout)
        line = self.lines["ten"]
        self.assertTrue(line.startswith("OK"), line)
        self.assertIn('"children":10', line)
        self.assertIn('"users":10', line)      # nav links excluded
        self.assertIn('"total":"19"', line)

    def test_a_valid_nine_result_page_is_read(self):
        line = self.lines.get("nine", "")
        self.assertTrue(line.startswith("OK"), line)
        self.assertIn('"children":9', line)
        self.assertIn('"users":9', line)

    def test_a_zero_results_page_is_read(self):
        line = self.lines.get("zero", "")
        self.assertTrue(line.startswith("OK"), line)
        self.assertIn('"noResults":true', line)
        self.assertIn('"children":0', line)

    def test_a_page_without_a_result_container_is_read(self):
        line = self.lines.get("nocontainer", "")
        self.assertTrue(line.startswith("OK"), line)
        self.assertIn('"children":0', line)
        self.assertIn('"domSig":0', line)       # no region -> zero, not a throw

    def test_a_login_wall_is_read(self):
        line = self.lines.get("login", "")
        self.assertTrue(line.startswith("OK"), line)
        self.assertIn('"login":true', line)

    def test_an_abbreviated_total_is_read(self):
        line = self.lines.get("abbrev", "")
        self.assertTrue(line.startswith("OK"), line)
        self.assertIn('"total":"9.1k"', line)


class TestNoTemporalDeadZone(unittest.TestCase):
    """A pure-Python guard, so the ordering is checked even without node."""

    @staticmethod
    def _code_only(js):
        """PAGE_STATE_JS with comments and strings blanked out.

        The names appear in explanatory comments and in selector strings; a
        TDZ check must look only at executable code, or a comment that merely
        mentions `resultsList` would read as a use-before-declaration.
        Blanked to spaces rather than removed, so byte offsets are preserved
        and error messages still point at the real position.
        """
        out = []
        i, n = 0, len(js)
        while i < n:
            two = js[i:i + 2]
            if two == "//":
                while i < n and js[i] != "\n":
                    out.append(" "); i += 1
            elif two == "/*":
                out.append("  "); i += 2
                while i < n and js[i:i + 2] != "*/":
                    out.append(" " if js[i] != "\n" else "\n"); i += 1
                out.append("  "); i += 2
            elif js[i] in "'\"":
                quote = js[i]; out.append(" "); i += 1
                while i < n and js[i] != quote:
                    if js[i] == "\\" and i + 1 < n:
                        out.append("  "); i += 2; continue
                    out.append(" "); i += 1
                out.append(" "); i += 1
            else:
                out.append(js[i]); i += 1
        return "".join(out)

    def _declared_before_used(self, name):
        js = self._code_only(student_scraper.PAGE_STATE_JS)
        decl = re.search(r"\b(?:const|let|var)\s+%s\b" % name, js)
        self.assertIsNotNone(decl, "%s is never declared" % name)
        first_use = re.search(r"\b%s\b" % name, js)
        self.assertLessEqual(
            decl.start(), first_use.start(),
            "%s is used at offset %d before its declaration at %d "
            "(temporal-dead-zone ReferenceError)"
            % (name, first_use.start(), decl.start()))

    def test_resultslist_is_declared_before_use(self):
        self._declared_before_used("resultsList")

    def test_mainregion_is_declared_before_use(self):
        self._declared_before_used("mainRegion")

    def test_region_is_declared_before_use(self):
        self._declared_before_used("region")

    def test_resultslist_is_declared_exactly_once(self):
        count = len(re.findall(
            r"\bconst\s+resultsList\s*=\s*document\.querySelector",
            student_scraper.PAGE_STATE_JS))
        self.assertEqual(count, 1,
                         "a duplicate/shadow declaration is a hazard")

    def test_no_field_in_the_return_is_undeclared(self):
        """Every identifier used in the return object must be declared."""
        js = self._code_only(student_scraper.PAGE_STATE_JS)
        declared = set(re.findall(r"\b(?:const|let|var)\s+([A-Za-z_]\w*)", js))
        declared |= {"document", "window", "location", "URLSearchParams",
                     "RegExp", "Set", "Array", "true", "false", "null"}
        ret = js[js.index("return {"):js.index("};", js.index("return {"))]
        for value in re.findall(r":\s*([A-Za-z_]\w*)", ret):
            self.assertIn(value, declared,
                          "%s is returned but never declared" % value)


if __name__ == "__main__":
    unittest.main(verbosity=2)
