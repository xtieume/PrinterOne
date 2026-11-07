---
name: open-pr
description: Use when PrinterOne work is ready to leave the machine — "push", "open a PR", "tạo PR", "đẩy lên" — or when implement-feature / fix-bug reach their PR step.
---

# Open a PR

IRON LAW: push only a branch whose gates are green on the commit being pushed; never `master`, never a merge.

## Steps

1. Branch: never `master`. `git switch -c feat/<topic>` (or `fix/<topic>`) if needed.
2. Identity: commits use the repo's configured `git config user.name/user.email`; never AI attribution (rule
   `03-git-release`).
3. Stage explicit paths (never `git add -A` blindly; no `.testcases/`) and commit with a short English message.
4. Write the PR body (English) to a file first. It must carry the lines CI reads:
   - `Feature: F-00N` for a feature revision, or `Feature: F-00N bug <name>` for a bug run;
   - `Spec-Exempt: <reason>` only when no behaviour changes (docs, tooling) — the owner sees the file list in CI;
   - `Closes #N` for every issue in the feature's `issues` that this PR finishes;
   - what changed for the user, the spec path, checks run and results, what was not verified;
   - each `MANUAL` row still `WAIT` and what it needs (XP/Win7 VM, real printer, browser).
5. Gates on that commit, then push the branch and open the PR:
   ```bash
   scripts/harness lint-rules
   scripts/harness spec-gate --base origin/master --pr-body-file <body.md>   # compares commits: after committing
   scripts/harness check harness-selftest   # only when .harness/, scripts/harness, .claude/ or verify.yml changed
   git push -u origin <branch>
   gh pr create --base master --body-file <body.md>
   ```
   For a feature or bug run, the last goalrun `--verify` must be on this commit with every automated row `PASS`
   and every break `VERIFIED`; only `MANUAL` rows may be `WAIT`.
6. Report the PR URL. The owner reviews and merges; for harness runs the merge is the acceptance (then
   `harness accept`).

## Anti-patterns

- `Spec-Exempt` on a change that alters behaviour — that needs a spec and a feature run.
- Pushing while a gate is red "to see what CI says".
- Merging, approving, commenting on issues, or force-pushing over someone else's commits.
