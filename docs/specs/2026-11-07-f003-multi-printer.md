# F-003 — Several printers in one server

Answers issues #1 and #6 (and the idea of PR #5). Plan: `docs/specs/2026-11-07-go-rewrite-plan.md`. Builds on F-001
(config with a `printers` list of exactly one) and F-008 (firewall rule over a port list).

## 0. Scope

One PrinterOne process shares several printers, each on its own TCP port. Both editions, `internal/core/`.

Not in F-003: picking the printer from the job content, templates or data formats per profile (issue #1 mentions
them; out of scope), per-printer allowlist (F-004).

## 1. Decisions

| # | Question | Answer |
|---|---|---|
| Q1 | Routing | One port per printer (issue #6 question 3, PR #5) — decided by Claude on the owner's delegation (2026-11-07). |
| Q2 | Limits with several printers | Decided by Claude (2026-11-07): the 8 / 2 receive slots and the 256 MB limit are shared by the whole process, so memory stays bounded however many printers there are. |

## 2. Behaviour

- Config: `printers` holds 1 to 32 entries. Each entry is validated as in F-001 §2.1 with its index in the field
  name (`printers[2].port`). Additionally: two entries with the same `port` → exit 2 naming the second
  `printers[i].port`; two entries with the same `name` (after the `default` rule; with several printers an empty
  name becomes `printer<i+1>`) → exit 2 naming the second `printers[i].name`. The same Windows printer may appear on
  two ports.
- Start-up: every port is bound before any is served. If one port is busy, nothing listens: exit 1 naming that port
  (F-001 §2.2).
- Each port prints to its own entry's printer. Jobs for the same printer are written one after another; jobs for
  different printers may print at the same time.
- Receive slots (8 main, 2 legacy) and the 256 MB limit are shared by all ports.
- Log lines name the printer of the port the job came in on (F-001 §2.5 already carries `"<printer>"`); start-up
  lines list every `name → printer on 0.0.0.0:<port>`.
- Ctrl+C / service stop apply to every port at once.
- F-008: `service install` opens every configured port in the firewall rule (one rule, port list on Vista+; one
  port opening per port on XP); uninstall closes them all.

## 3. Out of scope

Choosing the printer from the job data or the client, load balancing, printer groups, per-printer limits.

## Requirements

- **REQ-MULTI-001** A config with 2–32 valid printers loads; 0 or 33 printers exits 2 naming `printers`.
- **REQ-MULTI-002** An invalid field of any entry exits 2 naming it with its index (`printers[2].port`).
- **REQ-MULTI-003** Two entries with the same port exit 2 naming the second entry's `printers[i].port`.
- **REQ-MULTI-004** Two entries with the same name exit 2 naming the second `printers[i].name`; with several
  printers an empty name becomes `printer<i+1>`.
- **REQ-MULTI-005** A job sent to an entry's port is printed on that entry's printer only.
- **REQ-MULTI-006** Jobs for two different printers can be printed at the same time; jobs for one printer are
  written one after another.
- **REQ-MULTI-007** If one configured port is busy, nothing listens and the program exits 1 naming that port.
- **REQ-MULTI-008** The receive slots and the 256 MB limit are shared by all ports.
- **REQ-MULTI-009** Start-up lines list every printer with its name and port; connection lines name the printer of
  the port used.
- **REQ-MULTI-010** Ctrl+C and service stop stop every port, with F-001 shutdown rules on each.
- **REQ-MULTI-011** `service install` opens every configured port in the firewall and uninstall closes them all.
- **REQ-MULTI-012** The multi-printer code builds and its tests pass with Go 1.10.8.
