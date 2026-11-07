---
description: Tests and test cases — Go tests named by requirement, red before code, testcases.md as the why-record, ledger rows only through the harness check registry, MANUAL rows for XP/Win7 and real printers.
paths:
  - "**/*_test.go"
  - "testcases.md"
---

# 02 — Testing

- Anything that can be checked automatically has a real Go test, named after its requirement:
  `TestREQ_RECV_003_RejectsIPOutsideAllowlist`. Tests in `internal/core/` are Go 1.10 too (no `t.Cleanup`,
  `t.TempDir`, `t.Setenv`: use `defer` and `ioutil.TempDir`).
- `testcases.md` (committed) records WHY each case exists and holds the cases that need a person. Test cases are
  written with skill `testcase` when skill `implement-feature` / `fix-bug` reaches that step; an independent
  review pass after pass 1 is mandatory. `.testcases/` is never committed.
- Tests that touch Windows APIs use a fake behind an interface (spooler, listener) so they run anywhere; the real
  API calls are covered by a Windows check (from F-001) or a MANUAL row.
- An automated ledger row is only `scripts/harness check <id> [--tests <TestName>]` with `go-test` or
  `core-go110`. `windows-e2e` runs only on the CI Windows runner, so its rows are `MANUAL:admin`, signed with the
  URL of the green CI run. `go-vet` and `win-build` are `exit` checks and never go in a ledger. Never a `grep` on source as
  a check.
- `MANUAL:admin` only for what needs a person: running on an XP/Win7 VM, printing on a real printer, looking at
  the UI in IE6/IE8 and in a modern browser.
- Results: 0 PASS · 1 FAIL · 2 ERROR (missing tool, 0 tests ran, build error, panic) · 3 STALE. ERROR is never PASS.
- Bug kind 1: the reproducing test is RED on the old code before the fix and names the REQ it restores.
