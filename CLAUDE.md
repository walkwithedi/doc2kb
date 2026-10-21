# Working on doc2kb

This repository is public: code, history, commit e-mails, issues, pull requests and Actions logs can be read by anyone. CONTRIBUTING.md has the general rules; these are the ones an assistant must follow on top of them.

## Keeping private data out

- Never add personal data, details of anyone's workplace (organisation, host, system, project or people names, internal paths and addresses, business figures) or credentials. Examples use obviously fictional names (Riverside Clinic, Project Atlas), example.com addresses and the documentation IP ranges.
- Write rules and behaviour, not stories: README files, the spec, code comments, the changelog, commit messages and pull request descriptions say what something does and why in general terms. No accounts of where a rule came from, what happened to whom, or the environment a tool was built for, and no figures taken from real documents.
- scripts/sensitive_scan.py runs before every commit and push (.claude/hooks/pre_commit.py, .githooks) and in CI. When it blocks a commit or push, stop and report the findings to the user. Do not reword, delete or move the text to get past the check, and do not add the marker sensitive-scan: ignore unless the user confirms the line is harmless and agrees to the marker. Private-term findings cannot be allowed with the marker.
- When you notice something that looks private but the scan did not catch (a name, a system, a figure), report it to the user before doing anything else with it.
- The private-terms list (SENSITIVE_TERMS secret, ~/.config/doc2kb/sensitive_terms.txt) never enters the repository, and its terms are never quoted in commits, issues or pull requests.

## Layout and rules

- doc2kb_core.py is the converter and the master copy: other projects copy it from here, so a change to it is made here first. doc2kb.sh is the entry script, doc2kb_dedup.py the clean-up tool, SPEC.md the behaviour contract; a change in behaviour updates SPEC.md and both READMEs.
- Python 3 standard library only, bash for the entry script; no third-party packages and nothing fetched at run time. scripts/check_repo.py enforces the imports.
- Scripts log to stderr in the format the others use and take --log-level.
- VERSION holds the version; CHANGELOG.md has a section for it. Merging a version change to main publishes v<version> with doc2kb-v<version>.zip (.github/workflows/release.yml).
- Run python3 scripts/run_checks.py before committing.

## Design before code

- Any task that changes files or has side effects (commits, pushes, pull requests, remote settings) starts with a design: the goal, the files to change, the approach and trade-offs considered, and how the result will be verified. A small change needs only a few lines.
- The user approves the design before any code is written: through the plan-mode approval when the session is in plan mode, otherwise by asking with AskUserQuestion.
- If the work needs to depart from the approved design (wider scope, a different approach), stop and ask again instead of widening the change.
- Questions and read-only investigation need no design.
- .claude/settings.json sets permissions.defaultMode to plan, so sessions start in plan mode and cannot edit files until a design is approved.
- Sessions started in another mode (auto, for example) are held to the same rule: .claude/hooks/plan_gate.py refuses file edits, Bash commands with side effects (git writes, removing or changing files, redirecting output into files, package installs, gh and curl writes; anything unclear counts as a change) and GitHub write tools until the session's design is approved. Reads and writes to the system temporary directory pass. Approval is an accepted plan-mode exit, or an AskUserQuestion answer that starts with "Approve plan" (or "批准方案"), so the approving option of a design question must start that way. The approval mark lives per session under ~/.claude/plan-approved/, outside the repository; never create or change it yourself. When the gate blocks something harmless, report it to the user instead of working around it.
