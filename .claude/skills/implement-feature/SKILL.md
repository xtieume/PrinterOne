---
name: implement-feature
description: Use when the owner asks to implement, add, build or change a feature or behaviour in PrinterOne — "implement X", "add X", "support X", "handle issue #N", "change the spec", "làm tính năng", "thêm chức năng" — including urgent requests.
---

# Implement a feature — spec to code

IRON LAW: no product behaviour (declarations that only let a test compile are not behaviour) before (1) the spec
is committed with the owner's answers in it, and (2) a test for each requirement has been seen red. "Done" is
goalrun `--verify` exit 0 on the current commit plus a receipt.

**Urgency changes the size of the spec, never the order.** A ten-line spec is fine; skipping it is not.

```
Feature Progress:
- [ ] 0. Branch feat/<topic> from master
- [ ] 1. Spec with atomic REQ list + open questions → owner answered ⛔ (no defaults while waiting)
- [ ] 2. features.json entry (new id, or revision+1, with `issues`) + spec committed ⛔
- [ ] 3. goalrun init F-00N-r<rev> + resume; rules for allowed_paths read
- [ ] 4. docs-review → reqs.txt; testcase → tests RED; ledger ⛔
- [ ] 5. Build; goalrun --only per phase
- [ ] 6. role reviewer; breaks for rows green before the code; goalrun --verify on the commit ⚠️
- [ ] 7. open-pr (Feature: F-00N, Closes #N) → owner signs MANUAL → --verify exit 0 → owner merges
      → scripts/harness accept → receipt PR
```

## Steps

1. **Spec** — `docs/specs/<date>-<topic>.md`, in English: what the user sees, config keys and their defaults,
   behaviour on bad input and on Windows API failure, both editions (main and legacy), the web UI without JS and
   with JS, migration of existing `config.json`, out of scope. Ask open questions with a recommended answer and
   **wait**; write the answers into the spec with the date (a §1 table). End the spec with `## Requirements`: one
   `REQ-<AREA>-NNN` per yes/no fact, written before the spec is committed. Keep it to roughly 10–25 REQs; split the
   feature if it needs more.
2. **Feature** — `.harness/features.json`: `id`, `revision`, `behavior`, `spec`, `acceptance`, `allowed_paths`
   (code + tests + `testcases.md`), `required_checks`, `risk`, `issues`, `dependencies`. A changed accepted spec =
   same id, revision+1. Commit spec + entry together.
3. **Run** — skill `goalrun` with run id `F-00N-r<rev>` (`init … --spec <spec>`, `resume`). Then
   `scripts/harness rules-for --allow-missing <paths you will touch>` and read every rule printed.
4. **Requirements and tests** — inside the run: skill `docs-review` checks the spec's REQ list and writes
   `reqs.txt`. Skill `testcase` writes TCs, `testcases.md` rows and the real Go tests; run each and record it RED.
   Follow rule `02-testing`. Ledger `check` is only `scripts/harness check go-test --tests <TestName>`,
   `scripts/harness check core-go110 --tests <TestName>` or `MANUAL:admin`. Rows whose test was seen red first
   take a waiver line in the run's `ledger.tsv` (`# verify-ok: <id> — test-first, seen red on <date>`); do not
   reword `what` after that.
5. **Build** — role `implementer` per phase with the paths; role `test-writer` for missing coverage. Code under
   `internal/core/` must pass `scripts/harness check core-go110`.
6. **Prove** — role `reviewer` on the diff. Rows that were green before the code get a `break` now, written by a
   subagent that sees `what`, the spec and the product source, never the check. Commit; goalrun `--verify` on
   that commit. Every automated row ends `PASS` with its break `VERIFIED`; `MANUAL` rows may still be `WAIT`.
7. **Deliver** — skill `open-pr` with `Feature: F-00N` and `Closes #N` for each issue in `issues`. The controller
   records owner answers with `--sign`, reruns `--verify` until exit 0. The owner merging the PR is the
   acceptance: `git fetch origin master pull/<N>/head`, check out that head, `scripts/harness accept F-00N --pr <N>`,
   then commit the receipt on a new branch from `master` through a new PR. Comments on the issues are a separate,
   owner-approved step.

## Rationalizations

| Thought | Reality |
|---|---|
| "Ask, but start with defaults so we don't lose time" | Code built on a guessed answer is rework. Wait; meanwhile read code and rules. |
| "Write tests while coding" | A test never seen red proves nothing. Red first, per requirement. |
| "It only matters on Windows, skip the test" | Put the Windows call behind an interface; test the logic with a fake. |
| "Legacy can come later" | Core code is shared; `core-go110` runs on every core change. |
| "Spec can be committed with the code" | Then nothing proves the code followed the spec. Spec commit comes first. |
| "MANUAL rows are still WAIT, so no PR yet" | Rows needing a person are signed after the PR. The PR needs automated rows proven. |

Stop when blocked (3 failed verifies, owner unreachable): goalrun `checkpoint` + `release` with the next action.
