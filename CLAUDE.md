# PrinterOne — map for agents

RAW TCP print server for Windows, being rewritten in Go (v2). Two editions from one repo: main (Go stable,
Windows 10/11) and legacy (Go 1.10.8, Windows XP–8), sharing `internal/core/` and one web UI. Structure and
ownership: `.claude/rules/00-project-overview.md`. Plan: `docs/specs/2026-11-07-go-rewrite-plan.md`.

## From spec to code

- New feature, changed behaviour, or a GitHub issue asking for one: skill `implement-feature`.
- Bug: skill `fix-bug` — the spec already says what is right, fix the code; the spec is silent, it is a spec change.
- Review: skill `review-change`. Open a PR: skill `open-pr`.
- Work state lives in goalrun (plugin `testcase`, `.testcases/runs/<run>`) and receipts (`.harness/acceptance/`),
  not in chat. New session: `goalrun list` / `goalrun inspect <run>` → continue at `next_action`.
- `scripts/harness` (spec `docs/specs/2026-11-07-harness.md`) has 5 commands: `check` (registry checks, exit
  0 PASS · 1 FAIL · 2 ERROR), `rules-for`, `lint-rules`, `spec-gate`, `accept`. ERROR is never PASS.

## Checks

| Check | What |
|---|---|
| `go-test` | `go test -json ./...` on the host (Windows calls behind stubs) |
| `go-vet` / `win-build` | `go vet` for the host and for `GOOS=windows` |
| `core-go110` | build + test `internal/core/` and the legacy command with Go 1.10.8 in Docker |
| `harness-selftest` | the harness's own tests |

## Path-scoped rules

Rules in `.claude/rules/` load when a matching file is read. Codex: run `scripts/harness rules-for <path>` before
reading or editing a file and read every rule it prints.

| Rule | Area |
|---|---|
| `00-project-overview` | always |
| `01-go-code` | `cmd/`, `internal/`, `go.mod` |
| `02-testing` | `*_test.go`, `testcases.md` |
| `03-git-release` | `.github/`, `.harness/`, `.claude/`, `scripts/` |

Roles (`.claude/agents/`): `implementer`, `test-writer`, `reviewer`. Roles never commit; the main session does.

## Never

- Push or merge `master`; every change goes through a branch + PR and the owner merges (hook
  `.harness/hooks/guard.py`, git hook `.githooks/pre-push`).
- AI attribution in commits, PRs, comments or branches.
- Release, push a tag, or comment/close on GitHub issues and PRs without the owner's approval of that exact text
  (the hook asks first).
- Write anything in the repo in a language other than English.
