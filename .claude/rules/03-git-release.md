---
description: Git, PR and release workflow — branch per change, never push or merge master, no AI attribution, English everywhere, public GitHub actions only with the owner's approval, harness Verify before accept.
paths:
  - ".github/**"
  - ".harness/**"
  - ".claude/**"
  - "scripts/**"
  - ".githooks/**"
  - "CLAUDE.md"
  - "AGENTS.md"
---

# 03 — Git, PR and release

- Never push or merge `master` (or `main`); every change goes through a branch and a PR, and the owner reviews and
  merges. The hook denies `git push … master` and `gh pr merge`. For every client, `.githooks/pre-push` refuses a
  push to `master`/`main` once a clone has run `git config core.hooksPath .githooks`.
- No AI attribution anywhere (`Co-Authored-By`, "Generated with …") in commits, PRs, comments, branches.
- Commits, PRs, comments and docs are in English; commit messages are short and say what changed.
- Only the main session commits and pushes; role subagents edit files and hand back.
- Open a PR with skill `open-pr`. The PR body links the issues the feature answers with `Closes #N`.
- Public actions — `gh release`, a tag push, `gh issue comment|close|edit`, `gh pr comment|close|review`, a
  `gh api` write — are drafted first and run only after the owner approves that exact text (the hook asks).
- CI `Verify` (`.github/workflows/verify.yml`): lint-rules, spec-gate, the affected checks and the harness
  self-test (goalrun fetched at the commit pinned in `.harness/profile.json`). The owner checks `Verify` before
  merging; `scripts/harness accept` refuses when `Verify` for the head is not green.
- After editing `.harness/`, `scripts/harness`, `.claude/` or `verify.yml`: run
  `scripts/harness check harness-selftest` and `scripts/harness lint-rules`.
- Images for a PR or issue: give the owner the file path to drag in; never a branch or commit just to host an image.
