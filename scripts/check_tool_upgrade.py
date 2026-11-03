#!/usr/bin/env python3
"""Check that tool_upgrade.py is an unchanged copy of the upgrade module named in TOOL_UPGRADE_VERSION.

The upgrade module is kept in walkwithedi/edi-lab as .github/upgrade/tool_upgrade.py (the master copy);
this repository carries a copy. TOOL_UPGRADE_VERSION holds the master's MODULE_VERSION on its first
line and the copy's SHA-256 on the second ("sha256 <hex>").

  check_tool_upgrade.py            offline: the copy's MODULE_VERSION and SHA-256 match TOOL_UPGRADE_VERSION
                                   (scripts/check_repo.py runs this)
  check_tool_upgrade.py DIR        also: DIR is a clone of walkwithedi/edi-lab with history (CI checks it
                                   out when the EDI_LAB_READ_TOKEN secret is set), and some revision of the
                                   master on any branch has that MODULE_VERSION and is byte for byte the copy

To update the copy: copy the master, write its MODULE_VERSION and the new SHA-256 into
TOOL_UPGRADE_VERSION, run the tests, and note it in CHANGELOG.md.

stdout: one result line. stderr: log. Exit codes: 0 ok; 1 copy differs; 2 bad arguments or files.
"""
import argparse
import hashlib
import logging
import os
import re
import subprocess
import sys

log = logging.getLogger("check_tool_upgrade")
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s.%(funcName)s:%(lineno)d - %(message)s"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER = ".github/upgrade/tool_upgrade.py"
VERSION_RE = re.compile(rb"^MODULE_VERSION = '([0-9.]+)'$", re.M)


def module_version(data):
    m = VERSION_RE.search(data)
    return m.group(1).decode() if m else None


def read_pin(root=ROOT):
    """(version, sha256) from TOOL_UPGRADE_VERSION; raises ValueError when it is malformed."""
    with open(os.path.join(root, "TOOL_UPGRADE_VERSION"), encoding="utf-8") as f:
        lines = [x.strip() for x in f.read().splitlines() if x.strip()]
    if len(lines) != 2 or not re.fullmatch(r"\d+(\.\d+)*", lines[0]) or \
            not re.fullmatch(r"sha256 [0-9a-f]{64}", lines[1]):
        raise ValueError("TOOL_UPGRADE_VERSION must hold the version, then 'sha256 <hex>'")
    return lines[0], lines[1].split()[1]


def offline_problems(root=ROOT):
    try:
        version, digest = read_pin(root)
        with open(os.path.join(root, "tool_upgrade.py"), "rb") as f:
            data = f.read()
    except (OSError, ValueError) as e:
        return [str(e)]
    problems = []
    if module_version(data) != version:
        problems.append(f"tool_upgrade.py has MODULE_VERSION {module_version(data)}, TOOL_UPGRADE_VERSION says {version}")
    if hashlib.sha256(data).hexdigest() != digest:
        problems.append("tool_upgrade.py differs from the copy recorded in TOOL_UPGRADE_VERSION "
                        "(the copy is never edited here; copy the master again and update the pin)")
    return problems


def master_problems(master_dir, root=ROOT):
    """Look for a revision of the master that has the pinned version and equals the copy."""
    version, _ = read_pin(root)
    with open(os.path.join(root, "tool_upgrade.py"), "rb") as f:
        data = f.read()
    env = {k: v for k, v in os.environ.items() if not k.startswith("GIT_")}   # not the hook's repository
    r = subprocess.run(["git", "-C", master_dir, "log", "--all", "--format=%H", "--", MASTER],
                       capture_output=True, env=env)
    if r.returncode != 0:
        return [f"{master_dir} is not a clone with history: {r.stderr.decode(errors='replace').strip()}"]
    seen_version = False
    for rev in r.stdout.decode().split():
        show = subprocess.run(["git", "-C", master_dir, "show", f"{rev}:{MASTER}"], capture_output=True, env=env)
        if show.returncode != 0:
            continue
        if module_version(show.stdout) != version:
            continue
        seen_version = True
        if show.stdout == data:
            log.info("matches master revision %s", rev)
            return []
    if not seen_version:
        return [f"no revision of the master has MODULE_VERSION {version}"]
    return [f"tool_upgrade.py differs from every master revision with MODULE_VERSION {version}"]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("dir", nargs="?", help="a clone of walkwithedi/edi-lab with history")
    ap.add_argument("--log-level", default=os.environ.get("LOG_LEVEL", "WARNING").upper(),
                    choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    a = ap.parse_args()
    logging.basicConfig(level=a.log_level, format=LOG_FORMAT, stream=sys.stderr)
    problems = offline_problems()
    if not problems and a.dir:
        if not os.path.isdir(a.dir):
            print(f"result: {a.dir} is not a folder")
            sys.exit(2)
        problems = master_problems(a.dir)
    for p in problems:
        print(p)
    try:
        version = read_pin()[0]
    except (OSError, ValueError):
        version = "?"
    print("result: " + ("; ".join(problems) if problems else
                        f"tool_upgrade.py is the upgrade module {version}" + (" (compared with the master)" if a.dir else "")))
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
