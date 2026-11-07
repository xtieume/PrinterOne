---
description: Go code — the Go 1.10 subset for internal/core and the legacy command, stdlib first, no new dependencies without the owner, Windows code behind build tags with a stub so tests run on macOS/Linux.
paths:
  - "cmd/**"
  - "internal/**"
  - "go.mod"
  - "scripts/go110"
---

# 01 — Go code

## `internal/core/` and `cmd/printerone-legacy/`: Go 1.10 only

The legacy edition is built with Go 1.10.8, the last Go that runs on Windows XP. These two areas must **not** use:

- generics, `any`, `min`/`max`/`clear`, `for i := range n` (type aliases from 1.9 and `context` from 1.7 are fine);
- `embed`, `io/fs`, `log/slog`, `errors.Is/As/Join`, `fmt.Errorf("%w")`, `io.ReadAll`, `os.ReadFile`,
  `os.WriteFile` (use `ioutil`), `strings.Cut`, `net/netip`;
- `golang.org/x/sys` or any other module: winspool, services and the registry go through `syscall.NewLazyDLL`.

`scripts/harness check core-go110` builds and tests exactly these areas with Go 1.10.8; one newer feature turns
it red. Never split a core package into "new" and "old" copies.

## Everything else

- Stdlib first. A new dependency in `go.mod` needs the owner's approval (already approved: a systray library for
  the tray).
- Windows-only code lives in `*_windows.go`, always with an `*_other.go` stub or fake so `go test ./...` runs on
  macOS and Linux.
- Never kill the process holding a port; report a busy port clearly.
- Network input is untrusted: bound sizes, set read timeouts, never panic on input.
- `gofmt` clean; `go vet` passes for `GOOS=windows` too (check `win-build`).
