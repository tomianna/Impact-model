#!/usr/bin/env python3
"""PreToolUse hook (impact-model): the team decides, not Claude.

Asks the user before Claude records a revision decision ("status": "accepted" / "rejected" in
revisions/proposals.json) or applies proposals (revisions.py apply). Proposing stays free.
"""
import json
import re
import sys

DECISION = re.compile(r"""\\*["']status\\*["']\s*:\s*\\*["'](accepted|rejected)\\*["']""")
ASSIGN = re.compile(r"""\[\\*["']status\\*["']\]\s*=\s*\\*["'](accepted|rejected)""")   # p["status"] = "accepted"
MSG = ("Recording a revision decision ({what}). The impact-model rule: only the team accepts or rejects proposals. "
       "Approve only if the team has actually decided this.")


def decide(reason):
    print(json.dumps({"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": "ask",
                                             "permissionDecisionReason": reason}}))
    sys.exit(0)


def main():
    try:
        ev = json.load(sys.stdin)
    except ValueError:
        sys.exit(0)
    tool, ti = ev.get("tool_name", ""), ev.get("tool_input") or {}
    if tool == "Bash":
        cmd = re.sub(r"""\s-m\s+("(?:[^"\\]|\\.)*"|'[^']*')""", " -m MSG", ti.get("command", ""))   # commit messages are text
        if re.search(r"revisions\.py\s+apply\b", cmd):
            decide(MSG.format(what="revisions.py apply builds version B from the accepted proposals"))
        hits = DECISION.findall(cmd) + ASSIGN.findall(cmd)
        if "proposals" in cmd and hits:
            decide(MSG.format(what=", ".join(sorted(set(hits))) + " via a shell command"))
        sys.exit(0)
    f = ti.get("file_path") or ""
    if not re.search(r"proposals[^/]*\.json$", f):
        sys.exit(0)
    new = ti.get("content") or ti.get("new_string") or ""
    new += "".join(e.get("new_string", "") for e in ti.get("edits") or [])
    hits = DECISION.findall(new)
    if hits:
        decide(MSG.format(what=", ".join(sorted(set(hits))) + f" in {f}"))
    sys.exit(0)


if __name__ == "__main__":
    main()
