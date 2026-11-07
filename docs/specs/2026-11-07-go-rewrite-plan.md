# PrinterOne v2 — Go rewrite plan

Status: approved by the owner on 2026-11-07. This is the overall plan; each feature below gets its own spec with
atomic requirements before any code (skill `implement-feature`).

## 1. Decisions (2026-11-07)

| Topic | Decision |
|---|---|
| Language | Go. One static `.exe`, no runtime, cross-compiled from macOS, stdlib covers TCP, HTTP, templates. C# (.NET 4.0) and C++ were considered: C++ gives the widest compatibility but costs 2–3× the effort and is memory-unsafe on a network-facing service; C# needs .NET installed on XP. |
| Editions | Two builds. **main**: Go stable, Windows 10/11, x64 and x86. **legacy**: Go 1.10.8 (last Go for XP), Windows XP SP3 to 8, x86. Go 1.10 binaries also run on 7/8, so two builds cover everything. |
| Shared code | `internal/core/` holds everything both editions need and is written in the Go 1.10 subset. CI builds it with Go 1.10.8 (`core-go110`), so a newer API there fails fast. Main-only code (tray, richer logging) lives outside `core/`. |
| UI | One web UI for both editions, no native app UI. The binary serves it on `127.0.0.1`; starting the exe opens the browser. |
| UI levels | Same functions, two looks. Server-rendered HTML where every action is a plain `<form>` POST (works in IE6 with no JS). `base.css` is minimal; `modern.css` is wrapped in `@supports (display:grid)` and `app.js` loads as `<script type="module">`, so old browsers (IE6–11, Chrome 49 / Firefox 52 on XP) ignore both. JS only enhances (live log via SSE, inline toggles) and calls the same endpoints the forms use. |
| Remote admin | Off by default. Turning it on in config requires a password. |
| Raw port auth (#2) | IP/CIDR allowlist per printer. No token on the RAW port: RAW 9100 has no header, so a token would break every standard printer driver. Tokens apply to the HTTP API only. |
| Python v1 | Deleted from `master` in phase 0. Tag `v1.0.1` and its GitHub release keep it. v2 is new software: no v1 config migration, no v1 auto-start cleanup, no rollback path (owner, 2026-11-07). |
| Dropped from v1 | RAW→PDF conversion (reportlab, test-client only), killing the process that holds a port, the HKCU Run-key auto-start. |
| Spec details | The owner delegates details below their answers to Claude (2026-11-07); each spec records them as such. |
| Harness | Adapted from the LaunchPadTV harness (spec-first, goalrun-driven, receipts). See `docs/specs/2026-11-07-harness.md`. |
| Language of the repo | English for every file, commit, PR and comment. |

## 2. Architecture

```
printerone.exe            # double-click: start server + open web UI
printerone.exe serve      # console mode (debug)
printerone.exe service …  # install | uninstall | run as a Windows service (no login needed)
printerone.exe tray       # main edition only: tray icon that talks to the HTTP API
```

- **Receiver**: one TCP listener per configured printer; checks the client IP against that printer's allowlist,
  reads the job with a size limit and timeout, writes it to the spooler.
- **Spooler**: winspool `OpenPrinter` / `StartDocPrinter` / `WritePrinter` through `syscall.NewLazyDLL` (no
  `x/sys`, so Go 1.10 builds it). A fake implements the same interface for tests on macOS/Linux.
- **Web UI + API**: `net/http` + `html/template`. REST for printers, config, jobs and status; SSE for the live
  log; form endpoints are the same handlers the JS calls.
- **Assets**: Go 1.10 has no `embed`, so `static/` is turned into a generated Go file in `internal/core/webui`
  and a check keeps it in sync with `static/`.
- **Tray**: separate process, because a Windows service runs in session 0 and cannot show an icon. Legacy has
  no tray: the exe opens the UI and a Start-menu shortcut is created.
- **Logs**: one central folder (`%ProgramData%\PrinterOne\logs`), daily rotation, files opened with share-read so
  they can be tailed while the server runs, path configurable.
- **Config**: JSON with `version: 2` and a `printers` list; looked up at `--config`, next to the exe, then ProgramData.

## 3. Project structure

```
PrinterOne/
├── cmd/
│   ├── printerone/            # main edition
│   └── printerone-legacy/     # XP–8 edition, Go 1.10
├── internal/
│   ├── core/                  # Go 1.10 subset, both editions
│   │   ├── buildinfo/  config/  receiver/  spool/  auth/  webui/  api/  service/
│   └── tray/                  # main only
├── static/                    # web UI sources (templates, base.css, modern.css, app.js)
├── docs/specs/                # specs, one per feature revision
├── docs/client-setup.md       # F-007
├── scripts/harness, scripts/go110
├── .harness/  .claude/  .githooks/  .github/workflows/
└── testcases.md
```

## 4. GitHub issues and PRs

| Item | State | Ask | Plan |
|---|---|---|---|
| #1 | open | Multiple printer profiles | F-003; PR says `Closes #1` |
| #6 | open | Share several printers from one server | F-003; PR says `Closes #6` |
| #2 | open | Client authentication | F-004 allowlist; the issue reply explains why not a token on the RAW port |
| #3 | open | Log location, locked files, too many files | F-002; PR says `Closes #3` |
| #7 | open | Windows XP support | F-006 legacy edition with the minimal web UI; README credits the ashtray01 Go fork |
| #4 | closed | macOS cannot print through PrinterOne | No code change; F-007 client setup guide (the client needs the printer's own driver; PrinterOne forwards RAW) |
| PR #5 | open since 2026-01 | Multi-printer, one port each, in Python (+366/−252, commits `dist/server.exe`) | Not merged: Python is gone and it commits a binary. Same idea as F-003. Closed with thanks and credit when v2.0 ships |

Features carry an `issues` list in `.harness/features.json`; skill `open-pr` writes `Closes #N`. Comments on
issues and closing PR #5 are drafted and posted only after the owner approves the text.

## 5. Features

| Feature | Scope | Issues | Depends on |
|---|---|---|---|
| F-001 | Core: config, receive RAW on TCP, print via winspool, `serve` / `printers` / `version`, Windows e2e check in CI | — | — |
| F-008 | Windows service (install / uninstall / run, auto-start) with its log file, firewall rule added and removed with it | — | F-001 |
| F-002 | Central log folder, daily rotation, share-read files, configurable path | #3 | F-001 |
| F-003 | Several printers, one TCP port each | #1, #6 | F-001 |
| F-004 | Per-printer IP/CIDR allowlist | #2 | F-003 |
| F-005 | Shared web UI with two visual levels, live log, API with token, tray (main) | — | F-003 |
| F-006 | Legacy build (XP–8): packaging, Start-menu shortcut, MANUAL runs on XP and Win7 VMs | #7 | F-004, F-005 |
| F-007 | `docs/client-setup.md`: Windows (Standard TCP/IP port), macOS (HP JetDirect – Socket), Linux (CUPS `socket://`), driver note | #4 | F-001 |

Each spec stays around 10–25 atomic REQs; a bigger one is split.

## 6. Phases

| Phase | Work | Done when |
|---|---|---|
| 0 | Go skeleton (`cmd/`, `internal/core/buildinfo`), harness (script, hook, rules, skills, roles, self-test), CI `verify.yml`, Python removed, README notice | `lint-rules`, `go-test`, `go-vet`, `win-build`, `core-go110`, `harness-selftest` PASS; PR opened with `Spec-Exempt` |
| 1 | F-001 through the full harness loop; first real run, then tune the skills | receipt for F-001 |
| 2 | F-008 → F-002 → F-003 → F-004, one PR each | receipts |
| 3 | F-005 | receipt; MANUAL rows signed in IE6/IE8 and a modern browser |
| 4 | F-006 (+ F-007 alongside) | MANUAL rows signed on XP and Win7 VMs |
| 5 | `release.yml` builds both exes on a `v*` tag; v2.0.0 release; issue replies and PR #5 closed (owner-approved) | release published |

## 7. Checks (registry in `.harness/profile.json`)

| Check | Runs | Where |
|---|---|---|
| `go-test` | `go test -json ./...` | any OS (Windows calls behind stubs) |
| `go-vet` | `go vet ./...` | any OS |
| `win-build` | `GOOS=windows go vet ./...` | any OS |
| `core-go110` | vet + `GOOS=windows GOARCH=386` build + test of `internal/core/` and the legacy command with Go 1.10.8 | Docker |
| `harness-selftest` | the harness's own tests | any OS |
| `windows-e2e` (F-001) | `windows-latest`: a Generic / Text Only printer printing to a file, send RAW to 9100, compare the file | CI |
| `webui-assets` (F-005) | generated asset file matches `static/` | any OS |

XP and Windows 7 runs, real printers and browser looks are `MANUAL:admin` rows signed by the owner.

## 8. Out of scope

- LPD, IPP or AirPrint servers; PrinterOne stays a RAW socket server.
- Converting print data (PDF, PCL) on the server.
- A native desktop UI or a mobile app (the HTTP API leaves the door open).
- Code signing of the exe (revisit if antivirus false positives are reported).
