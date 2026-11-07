---
name: review-change
description: Use when a PrinterOne diff, branch or PR needs a review before it is accepted — "review", "review PR", "check this diff", "kiểm tra code" — or when implement-feature / fix-bug reach their reviewer step.
---

# Review a change

IRON LAW: every finding is checked against the code before it is reported — a file:line, the input that fails,
and the wrong result. An unverified suspicion goes under "could not verify", not under findings.

The review follows the code-review-expert workflow (SKILL.md + its four references: `solid-checklist.md`,
`security-checklist.md`, `code-quality-checklist.md`, `removal-plan.md`). Use skill `code-review-expert` when
installed; otherwise clone it at the pinned commit:

```sh
review_tmp=$(mktemp -d)
git clone --depth 1 --filter=blob:none --sparse https://github.com/sanyuan0704/sanyuan-skills.git "$review_tmp/sanyuan-skills"
git -C "$review_tmp/sanyuan-skills" sparse-checkout set skills/code-review-expert
git -C "$review_tmp/sanyuan-skills" fetch -q --depth 1 origin 08b6572ef108f22d4e8a3ecf9182a4bbef097744
git -C "$review_tmp/sanyuan-skills" checkout -q FETCH_HEAD
```

## Steps

1. Scope: `git diff origin/master...HEAD` (or the PR), staged and new files included.
2. Rules: `git diff --name-only origin/master...HEAD | scripts/harness rules-for --allow-missing` and read each
   rule. A rule broken is a finding with the rule cited (a Go 1.12+ API inside `internal/core/` is one).
3. Spec: the feature's spec (`.harness/features.json` → `spec`) and the run's `reqs.txt` (`goalrun inspect <run>`
   → `paths.goalrun_dir`). Behaviour the spec does not state, or a REQ with no test, is a finding.
4. Network input: the receiver and the web UI face the LAN. Check size limits, timeouts, allowlist bypass,
   CSRF on form POSTs, HTML escaping in templates, binding to `127.0.0.1` by default.
5. Harness evidence: `goalrun inspect <run>` (`whole_ledger_verified`, `proof_stale`); a ledger row whose check is
   not `scripts/harness check …`, or a test that cannot fail, is a finding.
6. Report P0–P3: `file:line`, impact (which input gives which wrong result), fix direction. Only what breaks
   correctness, security, a rule or a stated requirement; other notes go in an "optional" list. Say what you ran
   and what you could not verify.

Read-only: no edits, no commits, no GitHub comments.
