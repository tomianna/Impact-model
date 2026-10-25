#!/usr/bin/env python3
"""PreToolUse hook (impact-model): version A is immutable once it exists.

  *_vA.xlsx     changed by hand or by a shell command  -> deny (rebuild from the spec instead)
  *_vA.xlsx     rebuilt by build_model.py               -> ask  (new data only, never a revision)
  spec_vA.json  changed in any way                      -> ask  (new data or the team's own re-specification)
Creating either file for the first time, and reading them, is allowed. Covers Edit/Write/MultiEdit/NotebookEdit
and Bash, so a sed, a redirect or a python one-liner cannot get round it.
"""
import json
import os
import re
import sys

VA = re.compile(r"""[^\s'"<>|;&]*?(?:_vA\.xlsx|spec_vA\.json)""")
WRITE = re.compile(r"\bsed\s+(-\w*\s+)*-i|\bperl\s+(-\w*\s+)*-i|>>?\s*['\"]?[^\s'\"]*(_vA\.xlsx|spec_vA\.json)|\btee\b|"
                   r"\brm\b|\bmv\b|\bcp\b[^;&|\n]*\s\S*(_vA\.xlsx|spec_vA\.json)['\"]?\s*($|[;&|\n])|\btruncate\b|\.save\(|open\([^)]*['\"][wa]b?\+?['\"]|json\.dump\(|"
                   r"\.write\(|write_text\(|write_bytes\(|shutil\.(copy|move)|os\.(remove|rename|replace|unlink)")
XLSX_MSG = ("Version A workbook is immutable (impact-model): {f}. Never edit a built workbook. Change spec_vA.json and "
            "rebuild with build_model.py; strategy ideas go to revisions.py propose (copies), and only accepted ones "
            "into version B.")
SPEC_MSG = ("Changing the version A spec {f}. OK only for new data or the team re-specifying its own strategy, NOT for "
            "a revision (those go in revisions/proposals.json). Rebuild and re-check afterwards.")
BUILD_MSG = ("Rebuilding version A ({f}). OK after new data or a change the team asked for; never to apply a "
             "revision. Re-run check_model.py afterwards.")


def decide(kind, reason):
    print(json.dumps({"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": kind,
                                             "permissionDecisionReason": reason}}))
    sys.exit(0)


def exists(path, cwd):
    p = os.path.expanduser(path.strip("'\""))
    return os.path.exists(p if os.path.isabs(p) else os.path.join(cwd or "", p))


def main():
    try:
        ev = json.load(sys.stdin)
    except ValueError:
        sys.exit(0)
    tool, ti, cwd = ev.get("tool_name", ""), ev.get("tool_input") or {}, ev.get("cwd", "")
    if tool == "Bash":
        cmd = re.sub(r"""\s-m\s+("(?:[^"\\]|\\.)*"|'[^']*')""", " -m MSG", ti.get("command", ""))   # commit messages are text
        cd = re.search(r"(?:^|&&|;)\s*cd\s+(\"[^\"]+\"|'[^']+'|\S+)", cmd)   # cd <dir> && ...: paths are relative to it
        if cd:
            d = os.path.expanduser(cd.group(1).strip("'\""))
            cwd = d if os.path.isabs(d) else os.path.join(cwd, d)
        files = VA.findall(cmd)
        if not files:
            sys.exit(0)
        xlsx = [f for f in files if f.endswith(".xlsx")]
        spec = [f for f in files if f.endswith(".json")]
        if re.search(r"revisions\.py\s+apply", cmd) and xlsx:
            decide("deny", XLSX_MSG.format(f=xlsx[0]) + " revisions.py apply writes version B (--out *_vB.xlsx).")
        if re.search(r"build_model\.py", cmd) and xlsx:
            # the workbook argument follows the spec: build_model.py spec.json out.xlsx
            out = re.search(r"build_model\.py\s+\S+\s+(\S+_vA\.xlsx)", cmd)
            if out and exists(out.group(1), cwd):
                decide("ask", BUILD_MSG.format(f=out.group(1)))
            if not WRITE.search(re.sub(r"build_model\.py[^;&|]*", "", cmd)):
                sys.exit(0)
        if WRITE.search(cmd):
            hit_x = [f for f in xlsx if exists(f, cwd) or "/" not in f]
            if hit_x:
                decide("deny", XLSX_MSG.format(f=hit_x[0]))
            hit_s = [f for f in spec if exists(f, cwd) or "/" not in f]
            if hit_s:
                decide("ask", SPEC_MSG.format(f=hit_s[0]))
        sys.exit(0)
    f = ti.get("file_path") or ti.get("notebook_path") or ""
    if not f or not os.path.exists(f):
        sys.exit(0)
    name = os.path.basename(f)
    if name.endswith("_vA.xlsx"):
        decide("deny", XLSX_MSG.format(f=f))
    if name == "spec_vA.json":
        decide("ask", SPEC_MSG.format(f=f))
    sys.exit(0)


if __name__ == "__main__":
    main()
