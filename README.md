# PrinterOne 🖨️

[![GitHub stars](https://img.shields.io/github/stars/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/network)
[![GitHub issues](https://img.shields.io/github/issues/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/issues)
[![GitHub release](https://img.shields.io/github/v/release/xtieume/PrinterOne?style=flat-square)](https://github.com/xtieume/PrinterOne/releases)

Share printers installed on a Windows PC over the network: PrinterOne listens on a TCP port (9100 by default) and
hands every job it receives, unchanged, to the Windows spooler as a RAW print job — like the "JetDirect / port 9100"
interface of a network printer.

Version 2 is a rewrite in Go: one small `.exe`, nothing to install, managed from a web page.

## Features

- **Several printers in one server**, each on its own port (issues #1, #6).
- **Runs as a Windows service** — starts with the PC, no login needed; the firewall rule is added and removed for you.
- **Web UI** at `http://127.0.0.1:9180/` to add printers, send a test page and watch the live log. Every page also
  works in Internet Explorer 6 without JavaScript; newer browsers get a nicer look.
- **Per-printer IP allowlist** (addresses and CIDR blocks, IPv4/IPv6) (issue #2).
- **One central log** in `%ProgramData%\PrinterOne\logs`, readable while PrinterOne runs, rotated daily (issue #3).
- **Windows XP to Windows 11**: a legacy build for XP SP3, Vista, 7 and 8 (issue #7) and a main build for 10/11.
- Safe limits: 256 MB per job, 60 s idle timeout, a bounded number of jobs in memory.

## Download

From [Releases](https://github.com/xtieume/PrinterOne/releases):

| File | For |
|---|---|
| `PrinterOne-<version>-windows-x64.exe` | Windows 10 / 11, 64-bit |
| `PrinterOne-<version>-windows-x86.exe` | Windows 10 / 11, 32-bit |
| `PrinterOne-<version>-legacy-x86.exe` | Windows XP SP3, Vista, 7, 8 |

`SHA256SUMS` lists the checksums. Version 1 (Python) is still available as
[v1.0.1](https://github.com/xtieume/PrinterOne/releases/tag/v1.0.1); v2 does not read the v1 `config.json`.

## Quick start

1. Put the `.exe` in a folder of its own (for example `C:\PrinterOne\printerone.exe`).
2. Double-click it. Your browser opens the PrinterOne page; add a printer (choose the Windows printer and a port).
3. To keep it running after reboots, open a Command Prompt **as administrator** and run:

   ```bat
   C:\PrinterOne\printerone.exe service install
   ```

4. Set up the client computers: see **[docs/client-setup.md](docs/client-setup.md)** (Windows, macOS, Linux, scripts).
   The client needs the printer's own driver — PrinterOne does not convert print data.

## Commands

| Command | Does |
|---|---|
| `printerone` | serve and open the web UI in the browser |
| `printerone serve [--config file]` | run in this console (Ctrl+C stops after printing jobs already received) |
| `printerone printers` | list the printers installed on this PC |
| `printerone logs` | print the path of the log file |
| `printerone service install [--config file]` | install and start the Windows service, open the firewall (admin) |
| `printerone service uninstall` | stop and remove the service and the firewall rule (admin); config and logs stay |
| `printerone version` | print the version |

## Configuration

`config.json` is read from `--config`, else next to the exe, else `%ProgramData%\PrinterOne\config.json`
(`%ALLUSERSPROFILE%\Application Data\PrinterOne` on XP). The web UI writes it for you; by hand it looks like:

```json
{
  "version": 2,
  "printers": [
    {"name": "receipts", "printer": "EPSON TM-T82", "port": 9100},
    {"name": "labels", "printer": "Zebra ZD420", "port": 9101, "allow": ["192.168.1.0/24", "10.0.0.5"]}
  ],
  "log": {"dir": "D:\\PrinterOneLogs", "keep_days": 30},
  "ui": {"port": 9180, "remote": false, "password": ""}
}
```

| Key | Meaning |
|---|---|
| `printers[].printer` | Windows printer name, as `printerone printers` shows it |
| `printers[].port` | TCP port for this printer (1–65535, one per printer); default 9100 |
| `printers[].name` | short name shown in logs and the UI |
| `printers[].allow` | optional list of client addresses / CIDR blocks; missing = everyone |
| `log.dir`, `log.keep_days` | log folder (default `<data dir>\logs`) and days of rotated logs to keep (default 30) |
| `ui.port` | web UI port (default 9180) |
| `ui.remote`, `ui.password` | allow the web UI from other PCs (HTTP Basic, user `admin`, password of 8+ characters) |

An invalid config stops PrinterOne with a message naming the file and the field.

## Logs

One line per connection in `printerone.log`:

```text
2026-11-07T16:00:00+07:00 192.168.1.5:51234 1024B "EPSON TM-T82" printed
```

Results: `printed`, `empty`, `rejected: not allowed`, `dropped: too large | idle | reset | shutdown`,
`failed: <windows error code> <message>`. `printerone logs` shows where the file is.

## Troubleshooting

See the table at the end of [docs/client-setup.md](docs/client-setup.md). The usual causes: the firewall (the
service opens it; `serve` in a console does not), a wrong driver on the client, or the client's IP missing from an
allowlist.

## Building from source

Go 1.22 or newer for the main build, Docker for the legacy build (Go 1.10.8):

```sh
go test ./...
scripts/build-release v2.0.0-dev   # writes dist/*.exe and SHA256SUMS
```

The code shared by both editions lives in `internal/core` and must stay compatible with Go 1.10 (the last Go that runs
on Windows XP); CI builds and tests it with Go 1.10.8. Contributor notes: `CLAUDE.md`, specs in `docs/specs/`.

## Credits

- [ashtray01/printerone](https://github.com/ashtray01/printerone) showed that a Go build runs well on Windows XP (#7).
- @DiegoERuiz's PR #5 proposed one port per printer, which v2 adopted.

## Contributing

Issues and pull requests are welcome. Please describe the behaviour you want in the issue first; changes go through
a spec in `docs/specs/` and tests.

## Support

- Bugs and feature requests: [GitHub Issues](https://github.com/xtieume/PrinterOne/issues)
- Email: xtieume@gmail.com
