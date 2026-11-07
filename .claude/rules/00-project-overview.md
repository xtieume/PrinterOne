---
description: PrinterOne map — RAW TCP print server for Windows rewritten in Go, two editions (main and legacy XP), shared core, web UI; who owns what and the spec-first rule every change follows.
---

# 00 — Project overview

PrinterOne receives RAW print data over TCP (port 9100 by default) and writes it straight to a Windows printer
through winspool. The Python v1 code is gone from `master`; check out tag `v1.0.1` to see it. Overall plan:
`docs/specs/2026-11-07-go-rewrite-plan.md`.

## Two editions

| Edition | Toolchain | Runs on | Entry |
|---|---|---|---|
| main | Go stable | Windows 10/11 x64, x86 | `cmd/printerone` |
| legacy | Go 1.10.8 (Docker, `scripts/go110`) | XP SP3, Vista, 7, 8 (x86) | `cmd/printerone-legacy` |

## Who owns what

| Path | Contents | Constraint |
|---|---|---|
| `internal/core/` | Everything both editions share: config, TCP receiver, spooler, allowlist, web UI, API | Go 1.10 syntax and stdlib only (rule `01-go-code`) |
| `internal/` (outside `core/`) | Main edition only: tray, richer logging | Go stable |
| `cmd/` | The two `main` packages: flags and wiring only | |
| `static/` (from F-005) | Web UI: server-rendered HTML, `base.css` (IE6+), `modern.css` inside `@supports (display:grid)`, `app.js` as `type="module"` | Every action works without JS |
| `scripts/harness`, `.harness/`, `.claude/` | The harness | rule `03-git-release` |

## Changes go from spec to code

1. Every behaviour change starts from a spec in `docs/specs/` and a feature in `.harness/features.json` (its
   `issues` field lists the GitHub issues it answers). Skill `implement-feature` drives the whole path.
2. A bug the accepted spec already covers: skill `fix-bug`. A case the spec is silent on: the owner decides,
   and the feature gets a new revision.
3. CI `Verify` runs `scripts/harness spec-gate`: files under `cmd/`, `internal/`, `static/` must sit inside the
   `allowed_paths` of the feature the PR body declares (`Feature: F-00N`), or the PR carries
   `Spec-Exempt: <reason>` when no behaviour changes.

## Public repository

Issues, outside PRs, comments and releases are public actions: draft them, the owner approves, then post (the
hook asks first). Everything in the repo is written in English. Never put the owner's email or company identity
in the repo.
