#!/usr/bin/env python3
"""Repository checks that tests do not cover.

  1. every Python file compiles
  2. doc2kb.sh passes bash -n
  3. the tool's own scripts import only the standard library (and each other)
  4. VERSION holds a version, and CHANGELOG.md has a section for it

stdout: one line per problem and a result line. Exit code 1 when any check fails.
Usage: python3 scripts/check_repo.py [--log-level LEVEL]
"""
import argparse
import ast
import logging
import os
import re
import subprocess
import sys

log = logging.getLogger("check_repo")
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s.%(funcName)s:%(lineno)d - %(message)s"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOL_SCRIPTS = ["doc2kb_core.py", "doc2kb_dedup.py", "doc2kb_web.py"]
LOCAL_MODULES = {"doc2kb_core"}


def python_files():
    for d, dirs, files in os.walk(ROOT):
        dirs[:] = [x for x in dirs if not x.startswith(".") and x != "__pycache__"]
        for f in files:
            if f.endswith(".py"):
                yield os.path.join(d, f)
    for d in (".claude/hooks",):
        for f in sorted(os.listdir(os.path.join(ROOT, d))):
            if f.endswith(".py"):
                yield os.path.join(ROOT, d, f)


def imports(path):
    tree = ast.parse(open(path, encoding="utf-8").read(), path)
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                yield a.name.split(".")[0]
        elif isinstance(node, ast.ImportFrom) and node.module and not node.level:
            yield node.module.split(".")[0]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--log-level", default=os.environ.get("LOG_LEVEL", "WARNING").upper(),
                    choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    a = ap.parse_args()
    logging.basicConfig(level=a.log_level, format=LOG_FORMAT, stream=sys.stderr)
    problems = []
    for path in python_files():
        try:
            compile(open(path, encoding="utf-8").read(), path, "exec")
        except SyntaxError as e:
            problems.append(f"{os.path.relpath(path, ROOT)}: does not compile: {e}")
    p = subprocess.run(["bash", "-n", os.path.join(ROOT, "doc2kb.sh")], capture_output=True, text=True)
    if p.returncode:
        problems.append("doc2kb.sh: bash -n failed: " + p.stderr.strip())
    stdlib = getattr(sys, "stdlib_module_names", None)
    if stdlib is None:
        log.warning("Python %s has no sys.stdlib_module_names; import check skipped", sys.version.split()[0])
    else:
        for name in TOOL_SCRIPTS:
            for mod in imports(os.path.join(ROOT, name)):
                if mod not in stdlib and mod not in LOCAL_MODULES:
                    problems.append(f"{name}: imports {mod}, which is not in the standard library")
    version = open(os.path.join(ROOT, "VERSION"), encoding="utf-8").read().strip()
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        problems.append(f"VERSION: {version!r} is not X.Y.Z")
    elif not re.search(r"^## %s$" % re.escape(version), open(os.path.join(ROOT, "CHANGELOG.md"), encoding="utf-8").read(), re.M):
        problems.append(f"CHANGELOG.md: no '## {version}' section")
    for line in problems:
        print(line)
    print("result: " + (f"{len(problems)} problems" if problems else "ok"))
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
