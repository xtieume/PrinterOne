---
name: fix-bug
description: Use when something in PrinterOne misbehaves — a GitHub issue reporting a bug, a log, a print job that never arrives, "bug", "lỗi", "doesn't work", "sửa lỗi" — including when the owner says it is a one-line fix.
---

# Fix a bug — against the spec, never around it

IRON LAW: classify first. If the accepted spec already states the right behaviour, fix the code with a test seen
red. If the spec is silent, it is NOT a bug fix: the owner decides the behaviour and it becomes a feature revision.

```
Bug Progress:
- [ ] 1. Reproduced; evidence recorded (a Go test, a log, or the reporter's steps on a VM)
- [ ] 2. Classified against the accepted receipt's REQs ⛔
- [ ] 3. Kind 1: goalrun init + resume, run id F-00N-r<rev>-bug-<name>; reqs.txt = violated REQs
- [ ] 4. Reproducing test seen RED ⛔; testcases.md row with the bug as reason
- [ ] 5. Fix inside allowed_paths; role reviewer
- [ ] 6. goalrun --verify on the commit → open-pr (Feature: F-00N bug <name>, Closes #N if an issue)
      → owner signs MANUAL → --verify exit 0 → owner merges → harness accept
```

## Classify

Find the feature (`.harness/features.json`) and its receipt (`.harness/acceptance/F-00N/*.json` → `reqs`, `spec`).

| The accepted spec… | Kind | Route |
|---|---|---|
| states the right behaviour (a REQ in the receipt) | 1 | steps 3–6 here |
| is silent or wrong about this case | 2 | ask the owner what should happen, then skill `implement-feature` (revision+1) |
| feature not accepted yet | — | fix inside that feature's own run |

## Kind 1 steps

- Skill `goalrun`: `init` + `resume` with run id `F-00N-r<rev>-bug-<name>`; `reqs.txt` holds only REQs from the
  receipt.
- The reproducing test goes in first and must fail on the current code; record command and output.
- Ledger row: `scripts/harness check go-test --tests <TestName>` (or `core-go110`), a `break` that plants the bug
  back. Follow rule `02-testing`.
- `scripts/harness accept F-00N --bug <name> --pr <N>` refuses a REQ the receipt does not have.

## Rationalizations

| Thought | Reality |
|---|---|
| "Spec says nothing, I'll pick a sensible default and note it in the PR" | The owner decides behaviour. Silent spec = kind 2. |
| "I'll add one sentence to the spec inside the fix" | A spec edit in a bug fix is a hidden revision; `spec-gate` fails it. |
| "Owner said one line, skip the process" | One line still needs a red test and a REQ it restores. |
| "It only happens on XP, can't test it" | Reproduce the logic with a fake; the XP run is a MANUAL row. |
