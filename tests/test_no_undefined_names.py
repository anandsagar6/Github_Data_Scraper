"""Every name the scripts use must actually be defined.

WHY THIS FILE EXISTS
--------------------
Thirty modules were merged into student_scraper.py.  The merge deleted the
`from X import a, b` lines those modules used, which is correct -- the names
are in the same file now -- but a name that arrived under an ALIAS
(`describe_transient_failure as describe`) loses its binding when the import
line goes, and nothing complains until that exact line runs.

That happened: `_report_blocking` called `describe(entry)` and raised
NameError only when a profile had actually exhausted its retries, which is a
path the happy case never reaches.  Python has no compile-time check for it
and there is no pyflakes on this machine, so this file is the check.

It walks every function, works out what is bound in scope -- arguments, local
assignments, imports, comprehension targets, except-handler names, nested
definitions, and anything bound in an ENCLOSING function, which is what makes
closures legal -- and reports any remaining Load of a name that is not a
module global or a builtin.
"""

import ast
import builtins
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = ["student_scraper.py", "config.py"]

BUILTIN = set(dir(builtins)) | {"__file__", "__name__", "__doc__", "__spec__"}


def _direct_children(node):
    """Statements of `node`, not descending into nested function/class bodies."""
    for child in ast.iter_child_nodes(node):
        yield child


def _bound_in(node, top=True):
    """Names bound by `node` itself, without entering a nested scope."""
    out = set()
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
        if top:
            for a in list(node.args.posonlyargs) + list(node.args.args) + \
                     list(node.args.kwonlyargs):
                out.add(a.arg)
            if node.args.vararg:
                out.add(node.args.vararg.arg)
            if node.args.kwarg:
                out.add(node.args.kwarg.arg)
        else:
            out.add(node.name)
            return out
    elif isinstance(node, ast.ClassDef) and not top:
        out.add(node.name)
        return out
    elif isinstance(node, ast.Name) and isinstance(node.ctx, (ast.Store, ast.Del)):
        out.add(node.id)
    elif isinstance(node, (ast.Import, ast.ImportFrom)):
        for a in node.names:
            out.add(a.asname or a.name.split(".")[0])
    elif isinstance(node, ast.ExceptHandler) and node.name:
        out.add(node.name)
    elif isinstance(node, (ast.Global, ast.Nonlocal)):
        out.update(node.names)

    for child in _direct_children(node):
        out |= _bound_in(child, top=False)
    return out


def _loads_in(node, top=True):
    """(lineno, name) Loads in `node`, not entering nested scopes."""
    out = []
    if not top and isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef,
                                     ast.ClassDef, ast.Lambda)):
        return out
    if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
        out.append((node.lineno, node.id))
    for child in _direct_children(node):
        out += _loads_in(child, top=False)
    return out


def _nested(node):
    """Function and lambda scopes directly inside `node`."""
    found = []

    def walk(n, top=True):
        if not top and isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef,
                                      ast.Lambda, ast.ClassDef)):
            found.append(n)
            return
        for child in _direct_children(n):
            walk(child, top=False)

    walk(node)
    return found


def undefined_names(path):
    tree = ast.parse(open(path, encoding="utf-8").read())
    module_scope = _bound_in(tree) | BUILTIN
    problems = []

    def check(node, enclosing):
        if isinstance(node, ast.Lambda):
            scope = set(enclosing)
            for a in list(node.args.posonlyargs) + list(node.args.args) + \
                     list(node.args.kwonlyargs):
                scope.add(a.arg)
            if node.args.vararg:
                scope.add(node.args.vararg.arg)
            if node.args.kwarg:
                scope.add(node.args.kwarg.arg)
        elif isinstance(node, ast.ClassDef):
            scope = enclosing | _bound_in(node)
        else:
            scope = enclosing | _bound_in(node)
        for lineno, name in _loads_in(node):
            if name not in scope:
                problems.append((lineno, name))
        for inner in _nested(node):
            check(inner, scope)

    for top in _nested(tree):
        check(top, module_scope)
    # comprehension targets bind in their own scope -- treat them as bound
    comp_targets = set()
    for n in ast.walk(tree):
        if isinstance(n, ast.comprehension):
            for t in ast.walk(n.target):
                if isinstance(t, ast.Name):
                    comp_targets.add(t.id)
    return sorted({(ln, nm) for ln, nm in problems if nm not in comp_targets})


class TestNoUndefinedNames(unittest.TestCase):
    def test_every_script_name_resolves(self):
        for name in SCRIPTS:
            path = os.path.join(PROJECT, name)
            if not os.path.exists(path):
                continue
            problems = undefined_names(path)
            self.assertEqual(
                problems, [],
                "%s uses names that are never defined: %s\n"
                "An alias lost when an import line was removed looks exactly "
                "like this, and only raises on the line that uses it."
                % (name, ", ".join("%s (line %d)" % (n, ln)
                                   for ln, n in problems[:10])))


class TestEveryImportResolves(unittest.TestCase):
    """An import inside a function body must name a module that exists.

    WHY THIS EXISTS
    ---------------
    Module-level imports fail loudly the moment the file is imported, so they
    cannot rot.  An import inside a function body does not run until that
    function is called -- so a function that nothing calls can carry an import
    of a module that was deleted, and the whole suite stays green.

    That happened.  `mode_debug_extract` did `import mylib as ED` seven times
    over, and `mylib` was deleted when the thirty modules were merged into
    student_scraper.py.  The function was not wired to any CLI flag, so the
    ModuleNotFoundError was unreachable -- but it was still there, waiting for
    whoever wired up a --debug-extract flag.

    Optional third-party dependencies are exempt ONLY when the import sits in
    a try/except with a working fallback, which is how psutil, pdfminer and
    docx are used.  An unguarded import of something not installed fails here.
    """

    def local_modules(self):
        import glob
        names = set()
        for path in (glob.glob(os.path.join(PROJECT, "*.py"))
                     + glob.glob(os.path.join(PROJECT, "tests", "*.py"))):
            names.add(os.path.splitext(os.path.basename(path))[0])
        return names

    def guarded(self, tree):
        """ids of import nodes that sit inside a try: block."""
        safe = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Try):
                for child in node.body:
                    for inner in ast.walk(child):
                        if isinstance(inner, (ast.Import, ast.ImportFrom)):
                            safe.add(id(inner))
        return safe

    def test_no_import_names_a_module_that_does_not_exist(self):
        import importlib.util
        local = self.local_modules()
        problems = []
        for name in SCRIPTS:
            path = os.path.join(PROJECT, name)
            with open(path, encoding="utf-8") as handle:
                tree = ast.parse(handle.read())
            safe = self.guarded(tree)
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    modules = [alias.name for alias in node.names]
                elif isinstance(node, ast.ImportFrom):
                    if node.level or not node.module:
                        continue
                    modules = [node.module]
                else:
                    continue
                if id(node) in safe:
                    continue
                for module in modules:
                    root = module.split(".")[0]
                    if root in local or root in sys.builtin_module_names:
                        continue
                    try:
                        found = importlib.util.find_spec(root) is not None
                    except (ImportError, ValueError, ModuleNotFoundError):
                        found = False
                    if not found:
                        problems.append("%s:%d imports %s"
                                        % (name, node.lineno, module))
        self.assertEqual(
            problems, [],
            "these imports name a module that is not installed and is not a "
            "file in this project. An import inside a function body does not "
            "run until that function is called, so this is invisible at "
            "import time:\n  " + "\n  ".join(problems))


if __name__ == "__main__":
    unittest.main(verbosity=2)
