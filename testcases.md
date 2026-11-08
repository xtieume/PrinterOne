# Test cases

Why each test case exists, and the cases that need a person (rule `02-testing`). One row per TC; the test itself
carries the REQ id in its name. Rows are added by skill `testcase` inside a feature or bug run.

| TC | REQ | Test / how | Why it exists |
|---|---|---|---|
| TC-CFG-001 | REQ-CFG-001 | `TestREQ_CFG_001_FlagWinsOverExeAndDataDir` | `--config` wins over exe-folder and data-dir files; catches: exe folder or data dir checked before the flag |
| TC-CFG-002 | REQ-CFG-001 | `TestREQ_CFG_001_ExeDirWinsOverDataDir` | Exe folder wins over data dir; catches: data dir checked before the exe folder |
| TC-CFG-003 | REQ-CFG-001 | `TestREQ_CFG_001_DataDirUsedWhenNothingElse` | Data dir used when nothing else exists; catches: data-dir branch never consulted |
| TC-CFG-004 | REQ-CFG-001 | `TestREQ_CFG_001_MissingFlagFileDoesNotFallThrough` | Missing `--config` file does not fall through; catches: falling through to the exe-folder file |
| TC-CFG-005 | REQ-CFG-001 | `TestREQ_CFG_001_ServeFindsConfigInProgramData` | `serve` without `--config` finds `%ProgramData%\PrinterOne\config.json`; catches: CLI not wiring `DataDir(env.Getenv)` / `ExeDir` into the lookup |
| TC-CFG-006 | REQ-CFG-001 | `TestREQ_CFG_001_ServeMissingConfigFlagExits2` | `serve --config <missing>` exits 2; catches: error printed but server started anyway |
| TC-CFG-007 | REQ-CFG-002 | `TestREQ_CFG_002_DataDirFromProgramData` | Data dir from `%ProgramData%`; catches: `ALLUSERSPROFILE` preferred when both set |
| TC-CFG-008 | REQ-CFG-002 | `TestREQ_CFG_002_DataDirXPFallback` | XP fallback when `%ProgramData%` is unset; catches: no XP branch (returns `PrinterOne` relative path) or `Application Data` dropped |
| TC-CFG-009 | REQ-CFG-003 | `TestREQ_CFG_003_ValidFileLoadsWithDefaults/TC-CFG-009` | Full valid file loads every field; catches: fields swapped (`name` read as printer) |
| TC-CFG-010 | REQ-CFG-003 | `TestREQ_CFG_003_ValidFileLoadsWithDefaults/TC-CFG-010` | Missing `name` becomes `default`; catches: `name` left `""` when the key is absent |
| TC-CFG-011 | REQ-CFG-003 | `TestREQ_CFG_003_ValidFileLoadsWithDefaults/TC-CFG-011` | Empty `name` becomes `default`; catches: only the missing key defaulted, `""` kept |
| TC-CFG-012 | REQ-CFG-003 | `TestREQ_CFG_003_ValidFileLoadsWithDefaults/TC-CFG-012` | Missing `port` becomes 9100; catches: zero value `0` kept (then rejected or bound to a random port) |
| TC-CFG-013 | REQ-CFG-003 | `TestREQ_CFG_003_ValidFileLoadsWithDefaults/TC-CFG-013` | Unknown keys ignored at top level and in the entry; catches: strict decoding (`DisallowUnknownFields`) |
| TC-CFG-014 | REQ-CFG-005 | `TestREQ_CFG_005_PortBoundsAccepted` | Ports 1 and 65535 load; catches: exclusive bounds (`> 1`, `< 65535`) |
| TC-CFG-015 | REQ-CFG-004 | `TestREQ_CFG_004_UnreadableOrBadJSONNamesPathAndPosition/TC-CFG-015` | Unreadable file names the path; catches: OS error dropped, or read error returned as non-`*Error` or swallowed |
| TC-CFG-016 | REQ-CFG-004 | `TestREQ_CFG_004_UnreadableOrBadJSONNamesPathAndPosition/TC-CFG-016` | Syntax error reports path and line:column; catches: byte offset (`34`) reported instead of line:column (`3:16`) |
| TC-CFG-017 | REQ-CFG-004 | `TestREQ_CFG_004_UnreadableOrBadJSONNamesPathAndPosition/TC-CFG-017` | Truncated multi-line JSON reports path and line:column; catches: unexpected-EOF error mapped to a default `1:0`/`1:1` |
| TC-CFG-018 | REQ-CFG-004 | `TestREQ_CFG_004_ServeBadJSONExits2` | `serve` with bad JSON exits 2 with one stderr line; catches: multi-line error dump, or exit 1 |
| TC-CFG-019 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-019` | `version` missing (a v1 file); catches: missing key read as zero and accepted (v1 file silently loaded) |
| TC-CFG-020 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-020` | `version` 1; catches: version not checked at all |
| TC-CFG-021 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-021` | `version` the string `"2"`; catches: lenient string-to-number decoding |
| TC-CFG-022 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-022` | `printers` missing; catches: nil slice indexed → panic |
| TC-CFG-023 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-023` | `printers` not a list; catches: decode error reported without the field name |
| TC-CFG-024 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-024` | `printers` empty; catches: `len > 1` checked but not `len == 0` |
| TC-CFG-025 | REQ-CFG-005 | — (superseded) | Two printers was an error in F-001; F-003 makes several printers valid, so the case moved to TC-MULTI-* (2026-11-07) |
| TC-CFG-026 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-026` | Entry not an object; catches: decode error without field |
| TC-CFG-027 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-027` | `printer` missing; catches: empty printer accepted, every job fails later |
| TC-CFG-028 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-028` | `printer` empty string; catches: only the missing key checked |
| TC-CFG-029 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-029` | `printer` not a string; catches: type error reported without field |
| TC-CFG-030 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-030` | `port` a string; catches: lenient string port accepted |
| TC-CFG-031 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-031` | `port` fractional; catches: float truncated to 9100 |
| TC-CFG-032 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-032` | `port` 0; catches: 0 treated as "missing" → 9100, or bound to a random port |
| TC-CFG-033 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-033` | `port` 65536; catches: upper bound not checked |
| TC-CFG-034 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-034` | `port` negative; catches: only upper bound checked |
| TC-CFG-035 | REQ-CFG-005 | `TestREQ_CFG_005_ServeInvalidFieldExits2` | `serve` with an invalid field exits 2, nothing listens; catches: validation in Load but CLI ignores the error |
| TC-CFG-036 | REQ-CFG-006 | `TestREQ_CFG_006_NoConfigNamesDataDirAndHint` | No config anywhere names the data dir and hint; catches: generic "file not found" naming the exe folder |
| TC-CFG-037 | REQ-CFG-006 | `TestREQ_CFG_006_ServeNoConfigExits2` | `serve` with no config anywhere exits 2; catches: server started on default 9100 without config |
| TC-RECV-001 | REQ-RECV-001 | `TestREQ_RECV_001_BinaryJobPrintedExactly` | Binary job printed with exactly its bytes; catches: text-mode conversion (CRLF, NUL/0x1A trimming) |
| TC-RECV-002 | REQ-RECV-001 | `TestREQ_RECV_001_SeveralWritesAreOneJob` | Job sent in several writes is one job; catches: one job per read/segment |
| TC-RECV-003 | REQ-RECV-002 | `TestREQ_RECV_002_ConcurrentClientsSeparateSerialJobs` | Two concurrent clients → two jobs, never interleaved; catches: jobs merged by printer, or Print called concurrently |
| TC-RECV-004 | REQ-RECV-003 | `TestREQ_RECV_003_LimitBoundary/TC-RECV-004` | Job of exactly MaxBytes is printed; catches: off-by-one `>=` limit |
| TC-RECV-005 | REQ-RECV-003 | `TestREQ_RECV_003_LimitBoundary/TC-RECV-005` | Job of MaxBytes+1 is dropped and its connection closed; catches: limit checked only at end of job / partial job printed / truncated to 1024 and printed |
| TC-RECV-006 | REQ-RECV-003 | `TestREQ_RECV_003_ProductionLimitIs268435456` | Production limit is 268,435,456 bytes; catches: 256,000,000 or 2^28-1 |
| TC-RECV-007 | REQ-RECV-004 | `TestREQ_RECV_004_IdleConnectionDropped/TC-RECV-007` | Silent before the first byte → closed, nothing printed; catches: idle timer only armed after the first byte |
| TC-RECV-008 | REQ-RECV-004 | `TestREQ_RECV_004_IdleConnectionDropped/TC-RECV-008` | Silent after the first byte → closed, partial job not printed; catches: partial job printed on timeout |
| TC-RECV-009 | REQ-RECV-004 | `TestREQ_RECV_004_IdleConnectionDropped/TC-RECV-009` | Gaps shorter than Idle never drop the job; catches: absolute 60 s deadline from accept instead of per-gap |
| TC-RECV-010 | REQ-RECV-004 | `TestREQ_RECV_004_ProductionIdleIs60s` | Production idle timeout is 60 s; catches: 60 ms / 30 s |
| TC-RECV-011 | REQ-RECV-005 | `TestREQ_RECV_005_ResetPrintsNothing` | Reset mid-transfer prints nothing; catches: reset treated like EOF → partial job printed |
| TC-RECV-012 | REQ-RECV-005 | `TestREQ_RECV_005_ResetPrintsNothing` | Server keeps serving after a reset; catches: accept loop exits on a connection error |
| TC-RECV-013 | REQ-RECV-006 | `TestREQ_RECV_006_EmptyConnectionPrintsNothing` | 0-byte connection prints nothing; catches: empty job sent to the printer (blank page) |
| TC-RECV-014 | REQ-RECV-007 | `TestREQ_RECV_007_ClientOverSlotLimitWaits` | Over the slot limit a client waits, then is served promptly; catches: unlimited slots / Slots-1 (off by one) / third client refused or only served after an idle drop |
| TC-RECV-015 | REQ-RECV-007 | `TestREQ_RECV_007_LegacyEditionReceivesTwoAtOnce` | Legacy edition receives exactly 2 at once; catches: CLI uses 8 slots (or 1) for the legacy edition |
| TC-RECV-016 | REQ-RECV-007 | `TestREQ_RECV_007_MainEditionReceivesEightAtOnce` | Main edition receives exactly 8 at once; catches: main limited to 2 or 7, or unlimited |
| TC-RECV-017 | REQ-RECV-007 | `TestREQ_RECV_007_ProductionSlots` | Production slot counts; catches: swapped |
| TC-RECV-018 | REQ-RECV-008 | `TestREQ_RECV_008_BusyPortReturnsPortBusyError` | `Listen` on a busy port returns `*PortBusyError`; catches: raw `net.OpError` returned, port not named |
| TC-RECV-019 | REQ-RECV-008 | `TestREQ_RECV_008_ServeBusyPortExits1` | `serve` on a busy port exits 1, occupant untouched; catches: exit 2, or occupant closed / process killed |
| TC-RECV-020 | REQ-RECV-009 | `TestREQ_RECV_009_HalfCloseEndsJob` | Half-close ends the job and it is printed; catches: job waits for full close (hangs until idle drop) |
| TC-SPOOL-001 | REQ-SPOOL-001 | `TestREQ_SPOOL_001_CallSequence` | Call sequence of one job; catches: EndDoc before EndPage, missing StartPage, datatype `TEXT` |
| TC-SPOOL-002 | REQ-SPOOL-001 | `TestREQ_SPOOL_001_BytesUnchanged` | Bytes reach WritePrinter unchanged; catches: CRLF translation of 0x0A or truncation at 0x00/0x1A |
| TC-SPOOL-003 | REQ-SPOOL-001 | `TestREQ_SPOOL_001_ShortWritesContinued` | Short writes are continued; catches: single Write ignoring `n` |
| TC-SPOOL-004 | REQ-SPOOL-001 | `TestREQ_SPOOL_001_DocNameHasClientIPWithoutPort` | Document name carries the client IP without port; catches: `ip:port` in the name |
| TC-SPOOL-005 | REQ-SPOOL-001 | `TestREQ_SPOOL_001_ServePrintsThroughSpooler` | `serve` prints a job through the spooler; catches: config `name` used as printer |
| TC-SPOOL-006 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailingCallAbortAndCloseRules/TC-SPOOL-006` | OpenPrinter fails; catches: Close on a handle never opened |
| TC-SPOOL-007 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailingCallAbortAndCloseRules/TC-SPOOL-007` | StartDocPrinter fails; catches: Close skipped when StartDoc fails |
| TC-SPOOL-008 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailingCallAbortAndCloseRules/TC-SPOOL-008` | StartPagePrinter fails; catches: doc left open (no Abort) |
| TC-SPOOL-009 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailingCallAbortAndCloseRules/TC-SPOOL-009` | WritePrinter fails; catches: EndDoc on failure (prints a partial job) |
| TC-SPOOL-010 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailingCallAbortAndCloseRules/TC-SPOOL-010` | EndPagePrinter fails; catches: error ignored, returns nil |
| TC-SPOOL-011 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailingCallAbortAndCloseRules/TC-SPOOL-011` | EndDocPrinter fails; catches: error ignored, returns nil |
| TC-SPOOL-012 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_CloseFailureAfterEndDocStillPrinted` | ClosePrinter fails after EndDocPrinter succeeded: the spooler already holds the whole job, so it counts as printed (no Abort); decided 2026-11-07 |
| TC-SPOOL-013 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_FailedJobDoesNotStopNext` | A failed job does not stop the next one; catches: server stops / fails all later jobs |
| TC-SPOOL-014 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_PrinterNotInstalledEachJobFails` | Printer not installed: serve starts, each job fails; catches: printer checked at start-up (exit, or a probe that only warns) |
| TC-SPOOL-015 | REQ-SPOOL-003 | `TestREQ_SPOOL_003_PrintersCommand/TC-SPOOL-015` | `printers` lists one name per line; catches: names joined on one line / a name dropped |
| TC-SPOOL-016 | REQ-SPOOL-003 | `TestREQ_SPOOL_003_PrintersCommand/TC-SPOOL-016` | `printers` with none; catches: "no printers" treated as an error (exit 1) |
| TC-SPOOL-017 | REQ-SPOOL-003 | `TestREQ_SPOOL_003_PrintersCommand/TC-SPOOL-017` | `printers` when the spooler fails; catches: exit 0 with empty list |
| TC-SPOOL-018 | REQ-SPOOL-003 | `TestREQ_SPOOL_003_ListReturnsNames` | `spool.List` returns the API's names; catches: List swallowing the EnumPrinters error |
| TC-CLI-001 | REQ-CLI-001 | `TestREQ_CLI_001_ServeStartsAndStopsOnCtrlC` | `serve` starts, logs start-up lines, stops on Ctrl+C; catches: bound to 127.0.0.1 only / no start-up line |
| TC-CLI-002 | REQ-CLI-001 | `TestREQ_CLI_001_NoArgsRunsServe` | No args behaves as `serve`; catches: no args prints usage (exit 2) |
| TC-CLI-003 | REQ-CLI-001 | `TestREQ_CLI_001_CtrlCFinishesPrintingJobAndStopsAccepting` | Ctrl+C waits for the job being printed, then exits 0 and stops accepting; catches: exit at once losing the job / keeps accepting |
| TC-CLI-004 | REQ-CLI-001 | `TestREQ_CLI_001_SecondCtrlCExits1AtOnce` | Second Ctrl+C exits 1 at once; catches: second signal ignored (hang) |
| TC-CLI-005 | REQ-CLI-002 | `TestREQ_CLI_002_VersionNamesEdition` | `version` main; catches: edition or parentheses missing from the string |
| TC-CLI-006 | REQ-CLI-002 | `TestREQ_CLI_002_VersionNamesEdition` | `version` legacy; catches: edition hard-coded `main` |
| TC-CLI-007 | REQ-CLI-003 | `TestREQ_CLI_003_UsageErrorsExit2/TC-CLI-007` | Unknown command; catches: unknown command treated as serve |
| TC-CLI-008 | REQ-CLI-003 | `TestREQ_CLI_003_UsageErrorsExit2/TC-CLI-008` | Unknown flag; catches: unknown flag ignored and serve starts |
| TC-CLI-009 | REQ-CLI-003 | `TestREQ_CLI_003_UsageErrorsExit2/TC-CLI-009` | Flag without its value; catches: missing value read as `""` → normal lookup |
| TC-CLI-010 | REQ-CLI-004 | `TestREQ_CLI_004_VersionAndPrintersNeedNoConfig/TC-CLI-010` | `version` with no config anywhere; catches: config loaded before the command switch |
| TC-CLI-011 | REQ-CLI-004 | `TestREQ_CLI_004_VersionAndPrintersNeedNoConfig/TC-CLI-011` | `printers` with no config anywhere; catches: config loaded before `printers` (the hint's own command fails) |
| TC-CLI-012 | REQ-CLI-004 | `TestREQ_CLI_004_VersionAndPrintersNeedNoConfig/TC-CLI-012` | `version` with a broken config next to the exe; catches: config read and only "missing" tolerated |
| TC-CLI-013 | REQ-CLI-005 | `TestREQ_CLI_005_WaitingJobPrintedOnCtrlC` | A received job waiting for the printer is printed at Ctrl+C; catches: waiting job discarded on shutdown, or printed by the shutdown goroutine beside the running one |
| TC-LOG-001 | REQ-LOG-001 | `TestREQ_LOG_001_LineFormatPerResult/TC-LOG-001` | `printed` line format; catches: UTC (`2026-11-07T03:00:00Z`) instead of local / port missing / bytes missing `B` |
| TC-LOG-003 | REQ-LOG-001 | `TestREQ_LOG_001_LineFormatPerResult/TC-LOG-003` | `dropped: too large` line; catches: no line for drops |
| TC-LOG-004 | REQ-LOG-001 | `TestREQ_LOG_001_LineFormatPerResult/TC-LOG-004` | `dropped: idle` line; catches: bytes reported as 0 for a partial job |
| TC-LOG-006 | REQ-LOG-001 | `TestREQ_LOG_001_LineFormatPerResult/TC-LOG-006` | `failed:` line; catches: result `failed` without code and message, or the error struct printed with `%+v` |
| TC-LOG-007 | REQ-LOG-001 | `TestREQ_LOG_001_LineFormatPerResult/TC-LOG-007` | One line per connection, no more; catches: extra lines per read / none for empty |
| TC-LOG-008 | REQ-LOG-001 | `TestREQ_LOG_001_ServeLogsPrinterValueNotName` | `serve` log uses the config `printer` value; catches: `name` printed instead of `printer`; production `Now` in UTC; a `log` prefix added before the line |
| TC-LOG-009 | REQ-LOG-002 | `TestREQ_LOG_002_ReceivingConnectionDroppedOnShutdown` | Receiving connection at Ctrl+C → `dropped: shutdown`; catches: receiving connection drained and printed on shutdown |
| TC-E2E-001 | REQ-E2E-001 | `TestREQ_E2E_001_BytesArriveInPrinterFileUnchanged` — `scripts/harness check windows-e2e` (windows-latest only) | Real spooler: bytes arrive in the output file unchanged; catches: fake-only coverage hiding a winspool binding bug |
| TC-LEG-001 | REQ-LEG-001 (moved to F-006) | MANUAL:admin — run `printerone-legacy version`, `printerone-legacy printers`, `printerone-legacy serve`; send one text job to port 9100 from another machine (firewall opened) or from `127.0.0.1` | Legacy exe on Windows XP SP3; catches: an API missing on XP (a call newer than XP) or a Go 1.10 runtime problem |
| TC-LEG-002 | REQ-LEG-001 (moved to F-006) | MANUAL:admin — same three commands and one job as TC-LEG-001 | Legacy exe on Windows 7; catches: behaviour that works on XP but not on 7 (UAC, ProgramData path) |
| TC-LEG-003 | REQ-LEG-002 | `scripts/harness check core-go110` | `internal/core` and legacy command build and test with Go 1.10.8; catches: a newer stdlib call (`os.ReadFile`, `t.Cleanup`) in core |
| TC-CFG-038 | REQ-CFG-001 | `TestREQ_CFG_001_InvalidExeDirFileDoesNotFallThrough` | Invalid exe-folder file does not fall through to the data dir; catches: "try each place until one loads" lookup |
| TC-RECV-021 | REQ-RECV-003 | `TestREQ_RECV_003_ServeDropsJobOverProductionLimit` | `serve` wires the production limit: 268,435,457 bytes are dropped; catches: CLI leaves `MaxBytes` unset (no limit) or a larger limit |
| TC-RECV-022 | REQ-RECV-004 | `TestREQ_RECV_004_ServeDropsConnectionSilentFor60s` | `serve` wires the production idle timeout: 60 s; catches: CLI leaves `Idle` unset (never drops) or uses 30 s / 55 s / 65 s |
| TC-RECV-023 | REQ-RECV-004 | `TestREQ_RECV_004_JobWaitingForPrinterNotIdleDropped` | A received job waiting for the printer longer than Idle is still printed; catches: idle timer still running while a fully received job waits for the printer |
| TC-RECV-024 | REQ-RECV-007 | `TestREQ_RECV_007_BacklogWaitLongerThanIdleNotDropped` | A client waiting for a slot longer than Idle is not dropped; catches: connection accepted and idle timer started before a slot is free |
| TC-RECV-025 | REQ-RECV-006 | `TestREQ_RECV_006_HalfCloseWithNoBytesIsEmpty` | Half-close with 0 bytes is empty; catches: half-close treated as a job and an empty document printed |
| TC-CLI-014 | REQ-CLI-003 | `TestREQ_CLI_003_UsageErrorsExit2/TC-CLI-014` | Unknown flag on `version`; catches: flags ignored on commands without flags |
| TC-CLI-015 | REQ-CLI-003 | `TestREQ_CLI_003_UsageErrorsExit2/TC-CLI-015` | Unknown flag on `printers`; catches: flags ignored on commands without flags |
| TC-LOG-010 | REQ-LOG-001 | `TestREQ_LOG_001_ConcurrentConnectionsWriteWholeLines` | Concurrent connections write whole lines; catches: a line written in several Write calls (interleaves under concurrency) |
| TC-CFG-039 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-039` | `version` 3; catches: `version < 2` check instead of `!= 2` |
| TC-CFG-040 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-040` | Entry `null`; catches: `null` decoded into a zero struct, error names the sub-field |
| TC-RECV-026 | REQ-RECV-007 | `TestREQ_RECV_007_DroppedConnectionsFreeTheirSlot` | Dropped and failed connections free their slot; catches: slot released only after a successful print (leak on drop/fail paths) |
| TC-SPOOL-019 | REQ-SPOOL-002 | `TestREQ_SPOOL_002_CloseCalledEvenWhenAbortFails` | ClosePrinter still called when AbortPrinter fails; catches: early return on the Abort error leaks the handle |
| TC-LOG-011 | REQ-LOG-002 | `TestREQ_LOG_002_SilentConnectionDroppedOnShutdown` | Accepted connection with 0 bytes at Ctrl+C → `dropped: shutdown`; catches: socket closed to unblock Read, then classified as `empty` |
| TC-CFG-041 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-041` | `port` 74636 (wraps to 9100 in uint16); catches: port cast to uint16 before the range check (74636 → 9100 accepted) |
| TC-LOG-012 | REQ-LOG-002 | `TestREQ_LOG_002_ReceivingDroppedWhilePrintingOnShutdown` | Ctrl+C with one job printing and another connection still receiving; catches: shutdown waits for the print queue before dropping receivers, so B completes and prints |
| TC-CLI-016 | REQ-CLI-005 | `TestREQ_CLI_005_FailedDrainedJobStillExits0` | A drained job that fails at Ctrl+C still exits 0; catches: exit 1 when any drained job failed |
| TC-E2E-002 | REQ-SPOOL-003 | `TestREQ_SPOOL_003_WindowsPrintersListsRealPrinter` — `scripts/harness check windows-e2e` (windows-latest only) | `printers` on a real spooler lists the e2e printer; catches: EnumPrinters size probe (ERROR_INSUFFICIENT_BUFFER 122) treated as a failure; fake-only coverage |
| TC-E2E-003 | REQ-RECV-008 | `TestREQ_RECV_008_WindowsBusyPortExits1` — `scripts/harness check windows-e2e` (windows-latest only) | Busy port detected with real winsock errors; catches: busy port detected by comparing to `syscall.EADDRINUSE` (never equals WSAEADDRINUSE 10048 on Windows) |
| TC-E2E-004 | REQ-RECV-005 | `TestREQ_RECV_005_WindowsResetLoggedAsReset` — `scripts/harness check windows-e2e` (windows-latest only) | Reset logged with real winsock errors; catches: reset detected by comparing to `syscall.ECONNRESET` (never equals WSAECONNRESET 10054 on Windows) |
| TC-LEG-004 | REQ-LEG-001 (moved to F-006) | MANUAL:admin — from two clients at once send 268,435,456 bytes each to port 9100 | Two maximum-size jobs at once on the 32-bit legacy exe; catches: buffer growth by doubling (≈512 MB per job) exhausting 32-bit address space |
| TC-RECV-027 | REQ-RECV-003 | `TestREQ_RECV_003_ServePrintsJobOfExactlyProductionLimit` | `serve` prints a job of exactly 268,435,456 bytes; catches: CLI wires a smaller limit (256,000,000 or MaxJobBytes-1) |
| TC-CLI-017 | REQ-CLI-001 | `TestREQ_CLI_001_CtrlCDoesNotServeClientWaitingForSlot` | Ctrl+C does not serve a client still waiting for a slot; catches: accept-then-wait-for-slot: the waiting client takes a freed slot after Ctrl+C and is printed |
| TC-E2E-005 | REQ-CLI-001 | `TestREQ_CLI_001_WindowsCtrlBreakExits0` — `scripts/harness check windows-e2e` (windows-latest only) | Ctrl+Break on Windows reaches the shutdown path; catches: no `signal.Notify(os.Interrupt)` in `cmd/` (killed, non-zero exit) |
| TC-E2E-006 | REQ-CFG-001 | `TestREQ_CFG_001_WindowsConfigNextToExeFound` — `scripts/harness check windows-e2e` (windows-latest only) | `config.json` next to the exe is found from another working directory; catches: exe folder taken from the working directory or `os.Args[0]` instead of `os.Executable()` |
| TC-E2E-007 | REQ-CLI-002 | `TestREQ_CLI_002_WindowsExesNameTheirEdition` — `scripts/harness check windows-e2e` (windows-latest only) | Both real exes name their edition; catches: one `main` package passes the other edition |
| TC-E2E-008 | REQ-LOG-001 | `TestREQ_LOG_001_WindowsFailedLineFromRealSpooler` — `scripts/harness check windows-e2e` (windows-latest only) | `failed:` line from a real spooler error; catches: `syscall.Errno` returned instead of `*spool.Error` (code lost), or FormatMessage `\r\n` kept (two lines) |
| TC-CFG-042 | REQ-CFG-001 | `TestREQ_CFG_001_UnreadableOrEmptyExeDirFileDoesNotFallThrough/TC-CFG-042` | Unreadable exe-folder file does not fall through; catches: lookup skipping any file it cannot read |
| TC-CFG-043 | REQ-CFG-003 | `TestREQ_CFG_003_ValidFileLoadsWithDefaults/TC-CFG-043` | Keys are case-sensitive: `Port`/`Printer` are unknown keys; catches: `encoding/json` case-insensitive field matching (`Port` read as port → error, `Printer` overrides `printer`) |
| TC-CFG-044 | REQ-CFG-001 | `TestREQ_CFG_001_UnreadableOrEmptyExeDirFileDoesNotFallThrough` | 0-byte exe-folder file does not fall through; catches: empty file treated as absent |
| TC-LOG-013 | REQ-LOG-001 | `TestREQ_LOG_001_PrinterNameLoggedVerbatim` | Printer value logged verbatim between the quotes; catches: `%q` formatting (`"\\\\srv\\Canon \"A4\""`) |
| TC-CLI-018 | REQ-CLI-003 | `TestREQ_CLI_003_UsageErrorsExit2/TC-CLI-018` | `-h` is an unknown flag; catches: `flag.ErrHelp` mapped to exit 0 |
| TC-CFG-040 | REQ-CFG-004 | `TestREQ_CFG_004_ServeUnreadableExits2` | Ledger audit: only bad JSON was proven through `serve`; an unreadable config (a directory) must also exit 2 with one stderr line and start nothing |
| TC-CLI-020 | REQ-CLI-001 | `TestREQ_CLI_001_CtrlCDropsReceivingConnection` | Ledger audit: the CLI-001 row did not run any test of "Ctrl+C drops receiving connections" |
| TC-LOG-020 | REQ-LOG-001 | `TestREQ_LOG_001_FullLineForEmptyResetShutdown` | Ledger audit: the LOG-001 row never matched the full line for `empty`, `dropped: reset`, `dropped: shutdown` |
| TC-RECV-028 | REQ-RECV-010 | `TestREQ_RECV_010_AcceptErrorRetriesAndPrintsNextJob` | review finding P1-1, 2026-11-07 |
| TC-RECV-029 | REQ-RECV-007 | `TestREQ_RECV_007_ClientWaitsInBacklogUntilSlotFrees` | review finding P1-2, 2026-11-07 |
| TC-RECV-030 | REQ-RECV-007 | `TestREQ_RECV_007_PoolServerWaitsInBacklogUntilSlotFrees` | review finding, 2026-11-07 |
| TC-LOG-021 | REQ-LOG-001 | `TestREQ_LOG_001_ConnectionWaitingForSlotLoggedOnShutdown` | review finding, 2026-11-07 |
| TC-CFG-045 | REQ-CFG-003 | `TestREQ_CFG_003_LeadingUTF8BOMIgnored` | review finding P2-2, 2026-11-07 |
| TC-CFG-046 | REQ-CFG-005 | `TestREQ_CFG_005_TopLevelMustBeObject` | review finding P3-4, 2026-11-07 |
| TC-SVC-001 | REQ-SVC-001 | `TestREQ_SVC_001_InstallWithoutAdminChangesNothing` | Install without admin changes nothing; catches: admin check missing or done after mutating the SCM |
| TC-SVC-002 | REQ-SVC-001 | `TestREQ_SVC_001_AdminCheckedBeforeConfig` | Admin checked before the config; catches: config validated before the admin check |
| TC-SVC-003 | REQ-SVC-001 | `TestREQ_SVC_001_UninstallWithoutAdminChangesNothing` | Uninstall without admin changes nothing; catches: uninstall path skips the admin check |
| TC-SVC-004 | REQ-SVC-002 | `TestREQ_SVC_002_InvalidConfigExits2ChangesNothing` | Bad JSON config exits 2, nothing changed; catches: rule or service written before validation; exit 1 |
| TC-SVC-005 | REQ-SVC-002 | `TestREQ_SVC_002_InvalidConfigExits2ChangesNothing` | Missing `--config` file exits 2; catches: missing flag file falls through to exe-folder config |
| TC-SVC-006 | REQ-SVC-002 | `TestREQ_SVC_002_ConfigLookedUpLikeServe` | Config found next to the exe like `serve`; catches: install requiring `--config`; lookup not shared with serve |
| TC-SVC-007 | REQ-SVC-003 | `TestREQ_SVC_003_CreatesServiceWithSettings` | Fresh install creates the service with §2.1 settings; catches: unquoted paths (exe with a space); `serve` instead of `service run`; manual start |
| TC-SVC-008 | REQ-SVC-003 | `TestREQ_SVC_003_RelativeConfigBecomesAbsolute` | Relative `--config` becomes absolute in the command line; catches: relative path stored; service (cwd System32) cannot find it |
| TC-SVC-009 | REQ-SVC-004 | `TestREQ_SVC_004_ExistingServiceUpdated` | Existing service updated, not recreated; catches: Create on existing (ERROR_SERVICE_EXISTS → exit 1); delete-and-recreate |
| TC-SVC-010 | REQ-SVC-005 | `TestREQ_SVC_005_RuleReplacedDeleteThenAdd` | Rule replaced: delete every PrinterOne rule, then one add; catches: add without delete (duplicate rules); add before delete (deletes the new rule) |
| TC-SVC-011 | REQ-SVC-006 | `TestREQ_SVC_006_StartsAndWaitsForRunning` | Stopped service started, waited for, success line; catches: success printed right after Start without waiting |
| TC-SVC-012 | REQ-SVC-006 | `TestREQ_SVC_006_RunningServiceRestarted` | Running service is restarted; catches: Start on a running service (ERROR_SERVICE_ALREADY_RUNNING) or no restart (old settings keep running) |
| TC-SVC-013 | REQ-SVC-006 | `TestREQ_SVC_006_RunningWithin30sSucceeds` | Running reached at 29 s still succeeds; catches: timeout shorter than 30 s |
| TC-SVC-014 | REQ-SVC-007 | `TestREQ_SVC_007_RuleAddFailureRollsBack` | Rule add failure rolls back a created service; catches: service left installed; error swallowed |
| TC-SVC-015 | REQ-SVC-007 | `TestREQ_SVC_007_StartFailureRollsBack` | Start failure rolls back service and rule; catches: rule left open; service left |
| TC-SVC-016 | REQ-SVC-007 | `TestREQ_SVC_007_ExistingServiceNotDeletedOnFailure` | Start failure keeps a pre-existing service; catches: rollback deletes a service it did not create |
| TC-SVC-017 | REQ-SVC-007 | `TestREQ_SVC_007_StartTimeoutRollsBack` | Never running within 30 s is a start failure; catches: waits forever; gives up early |
| TC-SVC-018 | REQ-SVC-008 | `TestREQ_SVC_008_UninstallStopsWaitsDeletesRemovesRule` | Uninstall stops, waits, deletes, removes rule; catches: Delete while still running (marked for deletion); rule left |
| TC-SVC-019 | REQ-SVC-008 | `TestREQ_SVC_008_NothingInstalledExits0` | Nothing installed still exits 0; catches: ERROR_SERVICE_DOES_NOT_EXIST reported as failure |
| TC-SVC-020 | REQ-SVC-008 | `TestREQ_SVC_008_DeleteFailureExits1` | Delete failure is reported; catches: failure swallowed, success line printed |
| TC-SVC-021 | REQ-SVC-008 | `TestREQ_SVC_017_WindowsServiceInstallPrintUninstall` (windows-e2e, MANUAL:admin with CI run URL) | Uninstall keeps config and logs (assertion inside the TC-SVC-039 run); catches: uninstall cleans the data dir |
| TC-SVC-022 | REQ-SVC-009 | `TestREQ_SVC_009_ConsoleRunExits2` | `service run` from a console exits 2; catches: falls back to serving in the console; usage text |
| TC-SVC-023 | REQ-SVC-010 | `TestREQ_SVC_010_ServesJobUnderSCM` | Under SCM a job is printed; catches: receiver not started; status never Running |
| TC-SVC-024 | REQ-SVC-010 | `TestREQ_SVC_010_LegacyLimitUnderSCM` | Legacy edition limit applies under SCM; catches: main limit (8) or no limit used under the service |
| TC-SVC-025 | REQ-SVC-011 | `TestREQ_SVC_011_LinesAppendedToLogFile` | Start-up and connection lines go to the log file; catches: lines to stdout (lost under SCM); folder not created → open fails |
| TC-SVC-026 | REQ-SVC-011 | `TestREQ_SVC_011_AppendsToExistingLog` | Log file appended, not truncated; catches: `os.Create` truncating each start |
| TC-SVC-027 | REQ-SVC-011 | `TestREQ_SVC_011_FailingLogWriteKeepsServing` | Failing log writes do not stop the service; catches: log error treated as fatal |
| TC-SVC-028 | REQ-SVC-012 | `TestREQ_SVC_012_ConfigErrorStopsWithCode2` | Config error under SCM: logged, exit code 2; catches: exit code 1 / 0; error only on stderr |
| TC-SVC-029 | REQ-SVC-012 | `TestREQ_SVC_012_BusyPortStopsWithCode1` | Busy port under SCM: logged, exit code 1; catches: code 2; occupant killed |
| TC-SVC-030 | REQ-SVC-013 | `TestREQ_SVC_013_StopDropsReceivingPrintsReceived` | Stop drops receiving, prints received, reports in order; catches: no StopPending; partial printed; received job lost |
| TC-SVC-031 | REQ-SVC-013 | `TestREQ_SVC_013_ShutdownLikeStop` | Shutdown request behaves like stop; catches: only ControlStop handled |
| TC-SVC-032 | REQ-SVC-014 | `TestREQ_SVC_014_VistaUsesAdvfirewall` | Vista+ exact advfirewall arguments; catches: major 6 treated as XP (`<= 6`); `dir=out`; missing protocol |
| TC-SVC-033 | REQ-SVC-015 | `TestREQ_SVC_015_XPInstallOpensPort` | XP install opens the port with portopening; catches: advfirewall on XP (does not exist there) |
| TC-SVC-034 | REQ-SVC-015 | `TestREQ_SVC_015_XPUninstallClosesConfiguredPort` | XP uninstall closes the configured port; catches: port taken from nowhere / default 9100 |
| TC-SVC-035 | REQ-SVC-015 | `TestREQ_SVC_015_XPUnreadableConfigWarnsExits0` | XP uninstall with unreadable config warns, exit 0; catches: uninstall aborts with exit 1/2 |
| TC-SVC-036 | REQ-SVC-015 | `TestREQ_SVC_015_XPInstallClosesOldPort` | XP install over an existing service closes the old port; catches: old port left open |
| TC-SVC-037 | REQ-SVC-016 | `TestREQ_SVC_016_NetshFailureReportsOutput` | Uninstall netsh failure reported with its output; catches: any netsh non-zero ignored |
| TC-SVC-038 | REQ-SVC-016 | `TestREQ_SVC_016_NoRulesMatchIsNotFailure` | "No rules match" on removal is not a failure; catches: every non-zero delete treated as failure (fresh install always fails) |
| TC-SVC-039 | REQ-SVC-017 | `TestREQ_SVC_017_WindowsServiceInstallPrintUninstall` (windows-e2e, MANUAL:admin with CI run URL) | Windows e2e install → job → uninstall; catches: fake-only SCM code that never worked against advapi32 |
| TC-SVC-040 | REQ-SVC-018 | `scripts/harness check core-go110` | Go 1.10.8 builds and tests the service code; catches: Go 1.11+ API in service code |
| TC-SVC-041 | REQ-SVC-008 | `TestREQ_SVC_008_StopTimeoutGivesUpAfter30s` | Stop that never completes: uninstall gives up after 30 s; catches: waits forever; deletes a running service |
| TC-SVC-042 | REQ-SVC-010 | `TestREQ_SVC_010_MainLimitUnderSCM` | Main edition limit (8) under SCM; catches: legacy limit 2 always used under the service |
| TC-LOGF-001 | REQ-LOGF-001 | `TestREQ_LOGF_001_ServeWritesToDataDirLogs` | `serve` without `log` writes to `<data dir>/logs`, creating it; catches: folder not created; file written to cwd or exe folder |
| TC-LOGF-002 | REQ-LOGF-001 | `TestREQ_LOGF_001_DefaultDirAndPath` | `DefaultDir` / `Path` names; catches: different folder or file name |
| TC-LOGF-003 | REQ-LOGF-001 | `TestREQ_LOGF_001_StartCreatesNestedFolder` | `Start` creates a folder two levels deep; catches: `os.Mkdir` instead of `MkdirAll` |
| TC-LOGF-004 | REQ-LOGF-001 | `TestREQ_LOGF_001_KeepDaysOnlyKeepsDefaultFolder` | A `log` section with only `keep_days` keeps the default folder; catches: presence of `log` replaces the folder with something non-empty |
| TC-LOGF-005 | REQ-LOGF-002 | `TestREQ_LOGF_002_ServeWritesToLogDir` | `serve` with `log.dir` writes there, not to the data dir; catches: `log.dir` parsed but ignored; folder not created for a configured dir |
| TC-LOGF-006 | REQ-LOGF-002 | `TestREQ_LOGF_002_LoadReturnsConfiguredDir` | `Load` returns the configured folder; catches: default not applied when `log` present |
| TC-LOGF-007 | REQ-LOGF-002 | `TestREQ_LOGF_002_RelativeDirResolvesAgainstConfigFolder` | Relative `log.dir` resolves against the config folder (assumption); catches: resolved against the cwd (System32 for a service) |
| TC-LOGF-008 | REQ-LOGF-003 | `TestREQ_LOGF_003_MissingLogSectionIsValid` | Missing `log` is valid with defaults; catches: `log` required; KeepDays left 0 |
| TC-LOGF-009 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-009` | Empty `log.dir`; catches: empty string accepted as cwd |
| TC-LOGF-010 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-010` | `log.dir` of wrong type; catches: number silently ignored |
| TC-LOGF-011 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-011` | `keep_days` 0; catches: 0 treated as "absent → 30" |
| TC-LOGF-012 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-012` | `keep_days` 3651; catches: upper bound off by one (`<= 3651`) or missing |
| TC-LOGF-013 | REQ-LOGF-003 | `TestREQ_LOGF_003_KeepDaysBoundsValid` | `keep_days` 1 and 3650 valid; catches: bounds exclusive |
| TC-LOGF-014 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-014` | `keep_days` 1.5; catches: float truncated to 1 |
| TC-LOGF-015 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-015` | `keep_days` as string `"30"`; catches: string parsed with Atoi after unquoting |
| TC-LOGF-016 | REQ-LOGF-003 | `TestREQ_LOGF_003_InvalidLogFieldNamesField/TC-LOGF-016` | `log` not an object; catches: non-object silently treated as absent |
| TC-LOGF-017 | REQ-LOGF-003 | `TestREQ_LOGF_003_ServeInvalidKeepDaysExits2` | `serve` with an invalid `keep_days` exits 2; catches: log section validated only by `logs`, not by `serve` |
| TC-LOGF-018 | REQ-LOGF-004 | `TestREQ_LOGF_004_SameLinesInFileAndConsole` | Same lines in file and console; catches: console replaced by the file (or file only start-up lines) |
| TC-LOGF-019 | REQ-LOGF-004 | `TestREQ_LOGF_004_ServeAppendsToExistingFile` | `serve` appends to an existing file; catches: file truncated on open |
| TC-LOGF-020 | REQ-LOGF-004 | `TestREQ_LOGF_004_WritesAppendInOrder` | Writer appends writes in order; catches: overwrite at offset 0; short count returned |
| TC-LOGF-021 | REQ-LOGF-005 | `TestREQ_LOGF_005_ServiceRunWritesToLogDir` | `service run` writes to `log.dir`; catches: service still uses `LogPath(DataDir)` |
| TC-LOGF-022 | REQ-LOGF-005 | `TestREQ_LOGF_005_InvalidKeepDaysLoggedToDefaultFolder` | `service run` with an invalid log section logs the error to the default folder; catches: error written to the folder of an invalid log section (lost when that folder is bad) |
| TC-LOGF-023 | REQ-LOGF-006 | `TestREQ_LOGF_006_TwoWritersNeverSplitLines` | Two writers, concurrent, whole lines; catches: one shared offset without O_APPEND; buffered writes split across calls |
| TC-LOGF-024 | REQ-LOGF-006 | `TestREQ_LOGF_006_FileReadableWhileServing` | File readable while `serve` runs; catches: writer fails while a reader holds the file |
| TC-LOGF-025 | REQ-LOGF-006 | MANUAL:admin — `TestREQ_LOGF_006_WindowsTypeReadsLogWhileServing` (windows-e2e CI run URL) | Windows: `type` reads the log while `serve` runs (e2e, tags e2e,windows); catches: file opened without FILE_SHARE_READ |
| TC-LOGF-026 | REQ-LOGF-007 | `TestREQ_LOGF_007_LocalMidnightRotates` | Local midnight rotates; catches: no rotation; rotated file named with the new day |
| TC-LOGF-027 | REQ-LOGF-007 | `TestREQ_LOGF_007_SameLocalDayAcrossUTCMidnightDoesNotRotate` | Same local day across UTC midnight does not rotate; catches: day compared in UTC |
| TC-LOGF-028 | REQ-LOGF-007 | `TestREQ_LOGF_007_GapOfDaysNamesOldLinesDate` | Gap of several days names the old lines' date; catches: name computed as "yesterday" (now − 1 day) |
| TC-LOGF-029 | REQ-LOGF-007 | `TestREQ_LOGF_007_RestartOnLaterDayRotatesOldFile` | Restart on a later day rotates the old file; catches: day tracked only in memory (old file grows forever) |
| TC-LOGF-030 | REQ-LOGF-007 | `TestREQ_LOGF_007_RestartSameDayDoesNotRotate` | Restart on the same day does not rotate; catches: rotation at every start-up |
| TC-LOGF-031 | REQ-LOGF-007 | `TestREQ_LOGF_007_TwoWritersCrossMidnightRotateOnce` | Two writers cross midnight: one rotation, no line lost; catches: second writer renames the fresh file again (loses `a2` or makes `printerone-2026-11-08.log`) |
| TC-LOGF-032 | REQ-LOGF-007 | `TestREQ_LOGF_007_FailedRenameKeepsLogging` | Failed rename keeps logging; catches: line dropped when the rename fails |
| TC-LOGF-033 | REQ-LOGF-008 | `TestREQ_LOGF_008_StartCleansWithDefault30Days` | Start-up clean-up with default 30 days; catches: `<=` instead of `<`; age from mtime (both fresh) |
| TC-LOGF-034 | REQ-LOGF-008 | `TestREQ_LOGF_008_OtherFilesNeverTouched` | Other files never touched; catches: glob `printerone-*` deletes non-dates; mtime-based age |
| TC-LOGF-035 | REQ-LOGF-008 | `TestREQ_LOGF_008_CleanUpAlsoAtRotation` | Clean-up also runs at rotation; catches: clean-up only at start-up |
| TC-LOGF-036 | REQ-LOGF-008 | `TestREQ_LOGF_008_ServeUsesConfiguredKeepDays` | `serve` passes `keep_days` to the clean-up; catches: config keep_days ignored (default 30 used) |
| TC-LOGF-037 | REQ-LOGF-009 | `TestREQ_LOGF_009_FailingWritesWarnOnce` | Failing writes warn once with path and OS error; catches: warning per line; error returned to the caller |
| TC-LOGF-038 | REQ-LOGF-009 | `TestREQ_LOGF_009_ServeKeepsServingWhenFolderCannotBeCreated` | `serve` keeps serving when the folder cannot be created; catches: io.MultiWriter stops console on file error; exit on log failure |
| TC-LOGF-039 | REQ-LOGF-009 | `TestREQ_LOGF_009_NilWarnDoesNotPanic` | Nil Warn (service) does not panic; catches: nil Warn dereferenced |
| TC-LOGF-040 | REQ-LOGF-010 | `TestREQ_LOGF_010_LoadLogWithoutConfigReturnsDefaults`, `TestREQ_LOGF_010_LogsWithoutConfigPrintsDefaultPath` | `logs` with no config prints the default path; catches: non-zero when config missing |
| TC-LOGF-041 | REQ-LOGF-010 | `TestREQ_LOGF_010_LogsWithConfigLogDir` | `logs --config` with `log.dir`; catches: log.dir ignored by `logs` |
| TC-LOGF-042 | REQ-LOGF-010 | `TestREQ_LOGF_010_LoadLogIgnoresPrinters`, `TestREQ_LOGF_010_LogsIgnoresInvalidPrinters` | `logs` ignores an invalid printer list; catches: full `config.Load` used |
| TC-LOGF-043 | REQ-LOGF-010 | `TestREQ_LOGF_010_LogsInvalidLogSectionExits2` | `logs` with an invalid log section exits 2; catches: log section not validated by `logs` |
| TC-LOGF-044 | REQ-LOGF-010 | `TestREQ_LOGF_010_LogsFollowsLookupOrder` | `logs` follows the lookup order (exe folder); catches: lookup skipped, default always printed |
| TC-LOGF-045 | REQ-LOGF-010 | `TestREQ_LOGF_010_LogsBadArgumentsPrintUsage` | `logs` bad arguments; catches: extra args accepted |
| TC-LOGF-046 | REQ-LOGF-010 | `TestREQ_LOGF_010_LogsMissingConfigFileExits2` | `logs --config <missing>` exits 2 (assumption: as `serve`, F-001 REQ-CFG-001); catches: missing named file falls back to the default path |
| TC-LOGF-047 | REQ-LOGF-011 | `TestREQ_LOGF_011_ServeFirstLineNamesLogFile` | `serve` first line names the log; catches: path printed after `Listening on`, or relative |
| TC-LOGF-048 | REQ-LOGF-011 | `TestREQ_LOGF_011_ServiceRunFirstLineNamesLogFile` | `service run` first line names the log; catches: only `serve` updated |
| TC-LOGF-049 | REQ-LOGF-012 | `scripts/harness check core-go110` | Log code builds and tests with Go 1.10.8; catches: Go ≥1.11 API in logfile/config/cli |
| TC-LOGF-050 | REQ-LOGF-005 | `TestREQ_LOGF_005_ServiceRunWritesToLogDir` | `service run` connection lines reach `log.dir`; catches: only start-up lines redirected, receiver still on the default log |
| TC-LOGF-051 | REQ-LOGF-007 | `TestREQ_LOGF_007_TwoWritersRotateAtSameInstant` | Two writers rotate at the same instant: one rotation, no line lost; catches: check-then-rename race renames the fresh file again |
| TC-LOGF-052 | REQ-LOGF-006 | `TestREQ_LOGF_006_TwoProcessesNeverSplitLines` | Two processes, whole lines; catches: in-process mutex + seek-to-end instead of O_APPEND |
| TC-LOGF-053 | REQ-LOGF-007 | `TestREQ_LOGF_007_RotationWithOpenReaderLosesNothing` | Rotation while a reader holds the file loses nothing; catches: rename failure (Windows, open reader) drops the line or truncates |
| TC-LOGF-054 | REQ-LOGF-010 | `TestREQ_LOGF_010_LogsPrintsAbsolutePathForRelativeDir` | `logs` prints an absolute path for a relative `log.dir`; catches: relative path printed as written |
| TC-LOGF-055 | REQ-LOGF-005 | `TestREQ_LOGF_005_PrinterErrorLoggedInValidLogDir` | Printer config error with a valid `log.dir` is logged in `log.dir` (assumption); catches: error sent to the default folder while `printerone logs` points at `<tmp>/S` |
| TC-SVC-043 | REQ-SVC-014 | `TestREQ_SVC_014_InstallRemovalChecksRuleExistence` | review finding, 2026-11-07 |
| TC-SVC-044 | REQ-SVC-006 | `TestREQ_SVC_006_StoppedAfterStartFailsImmediatelyWithExitCode` | review finding, 2026-11-07 |
| TC-SVC-045 | REQ-SVC-006 | `TestREQ_SVC_006_StartPendingServiceWaitsThenRestarts` | review finding, 2026-11-07 |
| TC-SVC-046 | REQ-SVC-008 | `TestREQ_SVC_008_StopPendingUninstallWaitsWithoutResendingStop` | review finding, 2026-11-07 |
| TC-SVC-047 | REQ-SVC-015 | `TestREQ_SVC_015_XPRemovalMatchesPortNumber` | review finding, 2026-11-07 |
| TC-SVC-048 | REQ-SVC-014 | `TestREQ_SVC_014_RemovalSkipsDeleteWhenShowFails` | review finding, 2026-11-07 |
| TC-SVC-049 | REQ-SVC-007 | `TestREQ_SVC_007_RollbackStopsCreatedServiceBeforeDelete` | review finding, 2026-11-07 |
| TC-LOGF-056 | REQ-LOGF-007 | `TestREQ_LOGF_007_BackwardClockNeverMovesOrOverwritesDays` | review finding, 2026-11-07 |
| TC-LOGF-057 | REQ-LOGF-007 | `TestREQ_LOGF_007_ExistingRotationTargetIsAppended` | review finding, 2026-11-07 |
| TC-LOGF-058 | REQ-LOGF-007 | `TestREQ_LOGF_007_FailedRotationRetriedOnLaterWrite` | review finding, 2026-11-07 |
| TC-LOGF-059 | REQ-LOGF-007 | `TestREQ_LOGF_007_TwoIndependentWritersCoordinateRotation` | review finding, 2026-11-07 |
| TC-LOGF-060 | REQ-LOGF-004 | `TestREQ_LOGF_004_FileGetsLineWhenStdoutFails` | review finding, 2026-11-07 |
| TC-LOGF-061 | REQ-LOGF-007 | `TestREQ_LOGF_007_RotationNeverRacesAnOpenAppend` | Go 1.10 CI race, 2026-11-07 |
| TC-MULTI-001 | REQ-MULTI-001 | `TestREQ_MULTI_001_PrinterCount/TC-MULTI-001` | Two printers load in order |
| TC-MULTI-002 | REQ-MULTI-001 | `TestREQ_MULTI_001_PrinterCount/TC-MULTI-002` | 32 printers load |
| TC-MULTI-003 | REQ-MULTI-001 | `TestREQ_MULTI_001_PrinterCount/TC-MULTI-003` | 33 printers exit 2 naming `printers` |
| TC-MULTI-004 | REQ-MULTI-001 | `TestREQ_MULTI_001_PrinterCount/TC-MULTI-004` | 0 printers exit 2 naming `printers` |
| TC-MULTI-005 | REQ-MULTI-001 | `TestREQ_MULTI_001_PrinterCount/TC-MULTI-005` | One printer still loads with name `default` |
| TC-MULTI-006 | REQ-MULTI-002 | `TestREQ_MULTI_002_InvalidEntryNamesIndex/TC-MULTI-006` | Invalid port in entry 2 names its index |
| TC-MULTI-007 | REQ-MULTI-002 | `TestREQ_MULTI_002_InvalidEntryNamesIndex/TC-MULTI-007` | Entry 1 not an object |
| TC-MULTI-008 | REQ-MULTI-002 | `TestREQ_MULTI_002_InvalidEntryNamesIndex/TC-MULTI-008` | Entry 1 `printer` missing / empty / not a string |
| TC-MULTI-009 | REQ-MULTI-002 | `TestREQ_MULTI_002_InvalidEntryNamesIndex/TC-MULTI-009` | Entry 1 `port` wrong type / fraction / 0 |
| TC-MULTI-010 | REQ-MULTI-002 | `TestREQ_MULTI_002_InvalidEntryNamesIndex/TC-MULTI-010` | Two invalid entries: the first by index is named |
| TC-MULTI-011 | REQ-MULTI-002 | `TestREQ_MULTI_002_ServeInvalidEntryExits2` | `serve` with an invalid entry exits 2, one stderr line, nothing listens |
| TC-MULTI-012 | REQ-MULTI-003 | `TestREQ_MULTI_003_DuplicatePortNamesSecond/TC-MULTI-012` | Duplicate port names the second entry |
| TC-MULTI-013 | REQ-MULTI-003 | `TestREQ_MULTI_003_DuplicatePortNamesSecond/TC-MULTI-013` | Default port collides with explicit 9100 |
| TC-MULTI-014 | REQ-MULTI-003 | `TestREQ_MULTI_003_DuplicatePortNamesSecond/TC-MULTI-014` | Field error reported before a duplicate |
| TC-MULTI-015 | REQ-MULTI-004 | `TestREQ_MULTI_004_Names/TC-MULTI-015` | Duplicate name names the second entry |
| TC-MULTI-016 | REQ-MULTI-004 | `TestREQ_MULTI_004_Names/TC-MULTI-016` | Empty names become `printer<i+1>` |
| TC-MULTI-017 | REQ-MULTI-004 | `TestREQ_MULTI_004_Names/TC-MULTI-017` | Missing names become `printer<i+1>` |
| TC-MULTI-018 | REQ-MULTI-004 | `TestREQ_MULTI_004_Names/TC-MULTI-018` | Generated name collides with an explicit one |
| TC-MULTI-019 | REQ-MULTI-004 | `TestREQ_MULTI_004_Names/TC-MULTI-019` | Explicit `default` beside a generated name is not a duplicate |
| TC-MULTI-020 | REQ-MULTI-004 | `TestREQ_MULTI_004_Names/TC-MULTI-020` | Same Windows printer on two ports loads |
| TC-MULTI-021 | REQ-MULTI-005 | `TestREQ_MULTI_005_EachPortPrintsOnItsPrinter` | Each port prints on its own printer |
| TC-MULTI-022 | REQ-MULTI-006 | `TestREQ_MULTI_006_DifferentPrintersPrintAtOnce` | Different printers print at the same time |
| TC-MULTI-023 | REQ-MULTI-006 | `TestREQ_MULTI_006_SameWindowsPrinterOneAfterAnother` | One Windows printer on two ports is written one job after another |
| TC-MULTI-024 | REQ-MULTI-007 | `TestREQ_MULTI_007_SecondPortBusyNothingListens` | Second port busy: exit 1 naming it, first port released |
| TC-MULTI-025 | REQ-MULTI-007 | `TestREQ_MULTI_007_FirstPortBusyNamesIt` | First port busy: message names the busy port, not the other |
| TC-MULTI-026 | REQ-MULTI-007 | `TestREQ_MULTI_007_ServiceBusySecondPortStopsWithCode1` | Under the SCM a busy second port stops with code 1 |
| TC-MULTI-027 | REQ-MULTI-008 | `TestREQ_MULTI_008_SlotPoolSharedByServers` | Receive slots shared by two Servers |
| TC-MULTI-028 | REQ-MULTI-008 | `TestREQ_MULTI_008_LegacySlotsSharedAcrossPorts` | Legacy edition: 2 slots across three ports |
| TC-MULTI-029 | REQ-MULTI-008 | `TestREQ_MULTI_008_JobLimitOnSecondPort` | 256 MB limit applies on the second port |
| TC-MULTI-030 | REQ-MULTI-009 | `TestREQ_MULTI_009_StartUpLinesListEveryPrinter` | Start-up lines list every printer |
| TC-MULTI-031 | REQ-MULTI-009 | `TestREQ_MULTI_009_ConnectionLineNamesPortPrinter` | Connection line names the printer of the port used |
| TC-MULTI-032 | REQ-MULTI-009 | `TestREQ_MULTI_009_ServiceLogListsEveryPrinter` | `service run` log lists every printer |
| TC-MULTI-033 | REQ-MULTI-010 | `TestREQ_MULTI_010_CtrlCDropsReceivingOnEveryPort` | Ctrl+C drops receiving connections on every port |
| TC-MULTI-034 | REQ-MULTI-010 | `TestREQ_MULTI_010_CtrlCPrintsReceivedJobOnSecondPort` | A received job on the second port is printed at Ctrl+C |
| TC-MULTI-035 | REQ-MULTI-010 | `TestREQ_MULTI_010_ServiceStopClosesEveryPort` | Service stop closes every port |
| TC-MULTI-036 | REQ-MULTI-011 | `TestREQ_MULTI_011_VistaInstallOneRuleEveryPort` | Vista install: one rule with every port |
| TC-MULTI-037 | REQ-MULTI-011 | `TestREQ_MULTI_011_XPInstallOpensEveryPort` | XP install opens one port opening per port |
| TC-MULTI-038 | REQ-MULTI-011 | `TestREQ_MULTI_011_XPUninstallClosesEveryPort` | XP uninstall closes every configured port |
| TC-MULTI-039 | REQ-MULTI-011 | `TestREQ_MULTI_011_XPInstallClosesEveryOldPortFirst` | XP install closes every old port before opening new ones |
| TC-MULTI-040 | REQ-MULTI-011 | `TestREQ_MULTI_011_XPRollbackClosesEveryPort` | XP install rollback closes every port it opened |
| TC-MULTI-041 | REQ-MULTI-011 | `TestREQ_MULTI_011_VistaUninstallRemovesRule` | Vista uninstall removes the rule for every port |
| TC-MULTI-042 | REQ-MULTI-012 | `scripts/harness check core-go110` | Go 1.10.8 build and tests |
| TC-MULTI-043 | REQ-MULTI-011 | `TestREQ_MULTI_011_XPAddFailureClosesOpenedPorts` | XP install: `add portopening` fails on the second port, the first is closed again |
| TC-MULTI-044 | REQ-MULTI-008 | `TestREQ_MULTI_008_MainSlotsSharedAcrossPorts` | Main edition: 8 slots across two ports |
| TC-MULTI-045 | REQ-MULTI-002 | `TestREQ_MULTI_002_ServiceInstallInvalidEntryChangesNothing` | `service install` with an invalid entry changes nothing |
| TC-MULTI-046 | REQ-MULTI-008 | `TestREQ_MULTI_008_JobLimitIsPerJobNotPerProcess` | 256 MB is per job, not a process budget (Claude decision) |
| TC-MULTI-047 | REQ-MULTI-009 | `TestREQ_MULTI_009_SinglePrinterStartUpLine` | One printer: one start-up line in the new shape |
| TC-MULTI-048 | REQ-MULTI-008 | `TestREQ_MULTI_008_IdlePortsDoNotStarveBusyPort` | review finding, 2026-11-07 |
| TC-MULTI-049 | REQ-MULTI-010 | `TestREQ_MULTI_010_CtrlCStopsEveryPortWhileOnePrints` | review finding, 2026-11-07 |
| TC-MULTI-050 | REQ-MULTI-006 | `TestREQ_MULTI_006_SamePrinterDifferentCaseOneAfterAnother` | review finding, 2026-11-07 |
| TC-ALLOW-001 | REQ-ALLOW-001 | `TestREQ_ALLOW_001_NoAllowKeyLoadsNilList` | Config without `allow` loads a nil list; catches: missing key parsed as empty list (blocks everyone) |
| TC-ALLOW-002 | REQ-ALLOW-001 | `TestREQ_ALLOW_001_NilListAllowsEveryone` | nil `allow.List` allows every address family; catches: nil treated as deny-all |
| TC-ALLOW-003 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ListedAddressOrBlockMatches/TC-ALLOW-003` | Exact address matches only itself; catches: string prefix match (`10.0.0.50`) |
| TC-ALLOW-004 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ListedAddressOrBlockMatches/TC-ALLOW-004` | /24 includes first and last address, excludes neighbours; catches: off-by-one mask; /16 used |
| TC-ALLOW-005 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ListedAddressOrBlockMatches/TC-ALLOW-005` | Spec example list (block + address); catches: only the first item checked |
| TC-ALLOW-006 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ListedAddressOrBlockMatches/TC-ALLOW-006` | /32 and /0; catches: /0 parsed as "no list"/rejected; /32 widened |
| TC-ALLOW-007 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ListedAddressOrBlockMatches/TC-ALLOW-007` | CIDR with host bits set (Claude decision); catches: host-bits CIDR treated as single address or rejected |
| TC-ALLOW-008 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ListedAddressOrBlockMatches/TC-ALLOW-008` | Duplicate / overlapping items accepted; catches: duplicates rejected as invalid |
| TC-ALLOW-009 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ConfigListIsParsed` | Config list is parsed into `Allow[0]`; catches: key ignored by the loader |
| TC-ALLOW-010 | REQ-ALLOW-001 | `TestREQ_ALLOW_001_NilAllowPrints` | Server with nil Allow prints as in F-001; catches: nil Allow dereferenced / treated as deny |
| TC-ALLOW-011 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_AllowedClientPrintedAndAskedWithItsIP` | Server asks Allow once with the client IP; catches: Allow never consulted |
| TC-ALLOW-012 | REQ-ALLOW-003 | `TestREQ_ALLOW_003_RejectedWithoutReadOrPrint` | Rejected: no Read, no print, one `0B` line, closed at once; catches: data read then discarded; closed only at idle timeout; logged with byte count |
| TC-ALLOW-013 | REQ-ALLOW-003 | `TestREQ_ALLOW_003_RejectedClientHoldsNoSlot` | Rejected client holds no slot; catches: check done after taking the slot and slot kept until the client closes |
| TC-ALLOW-014 | REQ-ALLOW-003 | `TestREQ_ALLOW_003_ManyRejectedThenAllowed` | 5 rejected clients that stay open, then an allowed one; catches: rejected sockets leak slots or goroutines blocking accept |
| TC-ALLOW-015 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_IPv6ClientMatchedByIPv6List` | IPv6 client on `[::1]` matched by IPv6 entries; catches: IPv6 client matched by IPv4 /0; v6 not parsed |
| TC-ALLOW-016 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_MappedIPv4OnDualStackSocket` | IPv4 client on a dual-stack `[::]` socket; catches: mapped address compared as 16 bytes against a 4-byte entry |
| TC-ALLOW-017 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ServeListedClientPrinted` | serve prints a listed client; file change needs restart (also §2 restart rule, REQ-ALLOW-008 line); catches: cli drops the list; hot reload |
| TC-ALLOW-018 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_IPv6AndMappedIPv4/TC-ALLOW-018` | IPv6 block and address; catches: IPv6 mask truncated to 32 bits |
| TC-ALLOW-019 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_IPv6AndMappedIPv4/TC-ALLOW-019` | `::ffff:a.b.c.d` (16-byte) matches the IPv4 entry; catches: byte-length compare without `To4()` |
| TC-ALLOW-020 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_IPv6AndMappedIPv4/TC-ALLOW-020` | No cross-family match; catches: IPv4-compatible `::a00:5` treated as mapped; ::/0 matching IPv4 |
| TC-ALLOW-021 | REQ-ALLOW-005 | `TestREQ_ALLOW_005_EmptyOrWrongTypeAllowNamesField/TC-ALLOW-021` | `allow: []`; catches: empty list accepted (blocks all) or treated as missing |
| TC-ALLOW-022 | REQ-ALLOW-005 | `TestREQ_ALLOW_005_EmptyOrWrongTypeAllowNamesField/TC-ALLOW-022` | `allow` of the wrong type (Claude decision); catches: null treated as missing; string accepted as one item |
| TC-ALLOW-023 | REQ-ALLOW-005 | `TestREQ_ALLOW_005_EmptyOrWrongTypeAllowNamesField/TC-ALLOW-023` | Empty list on entry 1 names its index; catches: index hard-coded 0 |
| TC-ALLOW-024 | REQ-ALLOW-003 | `TestREQ_ALLOW_003_ServeRejectsUnlistedClient` | serve rejects an unlisted client; catches: cli never wires Allow |
| TC-ALLOW-025 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_ParseNamesFirstInvalidItem` | `allow.Parse` rejects malformed items, first by index; catches: hostname resolved; spaces trimmed; zone accepted; last bad index |
| TC-ALLOW-026 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_InvalidItemNamesIndex/TC-ALLOW-026` | Config names the invalid item; catches: whole list named; index off by one |
| TC-ALLOW-027 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_InvalidItemNamesIndex/TC-ALLOW-027` | Non-string item; catches: number silently skipped |
| TC-ALLOW-028 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_InvalidItemNamesIndex/TC-ALLOW-028` | Several invalid items → first by index; catches: last error reported |
| TC-ALLOW-029 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_InvalidItemNamesIndex/TC-ALLOW-029` | Invalid item on entry 1; catches: printer index hard-coded 0 |
| TC-ALLOW-030 | REQ-ALLOW-007 | `TestREQ_ALLOW_007_EachEntryKeepsItsOwnList` | Each entry keeps its own list; catches: one list shared / last wins |
| TC-ALLOW-031 | REQ-ALLOW-008 | `TestREQ_ALLOW_008_StringListsEveryItem` | `List.String` names every item; catches: only the first item shown |
| TC-ALLOW-032 | REQ-ALLOW-005 | `TestREQ_ALLOW_005_ServeEmptyAllowExits2` | serve with `allow: []` exits 2; catches: server starts and blocks everyone |
| TC-ALLOW-033 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_ServeInvalidItemExits2` | serve with an invalid item exits 2; catches: invalid item skipped and server starts |
| TC-ALLOW-034 | REQ-ALLOW-007 | `TestREQ_ALLOW_007_ListAppliesOnlyToItsPort` | List applies only to its own port; catches: first list applied to all ports; last (nil) list applied to all |
| TC-ALLOW-035 | REQ-ALLOW-008 | `TestREQ_ALLOW_008_StartUpLinesShowEachList` | Start-up lines with two printers; catches: one global allow line; lists swapped |
| TC-ALLOW-036 | REQ-ALLOW-008 | `TestREQ_ALLOW_008_SinglePrinterAllowAll` | One printer without `allow` shows `allow: all`; catches: allow shown only for multi-printer configs |
| TC-ALLOW-037 | REQ-ALLOW-009 | `scripts/harness check core-go110` | Allowlist code with Go 1.10.8; catches: `net/netip` or `strings.Cut` used |
| TC-ALLOW-038 | REQ-ALLOW-002 | `TestREQ_ALLOW_002_ServeListedClientPrinted` (restart part) | After a restart the rewritten list applies; catches: list cached across runs / read from a stale copy |
| TC-ALLOW-039 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_LeadingZeroRefusedInBothEditions` | Leading-zero IPv4 refused in both editions (Claude decision); catches: Go 1.10 `ParseIP` reading `010` as 10 on legacy only |
| TC-ALLOW-040 | REQ-ALLOW-003 | `TestREQ_ALLOW_003_ChecksRemoteNotLocalAddress` | The remote address is checked, not the local one; catches: `conn.LocalAddr()` passed to Allow |
| TC-ALLOW-041 | REQ-ALLOW-003 | `TestREQ_ALLOW_003_RejectedAtOnceWhileSlotsBusy` | Refused client closed at once while every slot is busy; catches: waits for a slot before the check |
| TC-ALLOW-042 | REQ-ALLOW-007 | `TestREQ_ALLOW_007_ListOnLaterEntryOnly` | List on a later entry only; catches: `Allow` indexed by entries-with-list (shift) |
| TC-ALLOW-043 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_NilIPNotAllowedByAList` | A list never allows a nil IP; catches: fail-open when the client IP is unknown |
| TC-ALLOW-044 | REQ-ALLOW-004 | `TestREQ_ALLOW_004_MappedItemMatchesIPv4Client` | A mapped item matches the plain IPv4 client; catches: mapped item stored as /128 compared byte-wise |
| TC-ALLOW-045 | REQ-ALLOW-006 | `TestREQ_ALLOW_006_LeadingZeroInEmbeddedIPv4Refused` | review finding, 2026-11-07 |
| TC-UI-001 | REQ-UI-001 | `TestREQ_UI_001_DefaultPort9180RemoteOff` | No `ui` section → port 9180, remote off; catches: default 0 / 80 / 9100 |
| TC-UI-002 | REQ-UI-001 | `TestREQ_UI_001_ConfiguredPortUsed` | `ui.port` 9181 is used; catches: section ignored, always default |
| TC-UI-003 | REQ-UI-001 | `TestREQ_UI_001_BindsLoopbackAndServes` | UI bound to loopback only when remote is off; catches: bound to `0.0.0.0` / `:port` |
| TC-UI-004 | REQ-UI-001 | pending: cli / reload wiring test (build step) | `serve --config <cfg>` serves the UI on ui.port; catches: UI started only by no-args |
| TC-UI-005 | REQ-UI-001 | pending: cli / reload wiring test (build step) | `service run --config <cfg>` serves the UI, opens no browser; catches: browser opened from the service (session 0) |
| TC-UI-006 | REQ-UI-001 | pending: cli / reload wiring test (build step) | UI port busy → exit 1 naming the port, nothing listens; catches: printers served without UI |
| TC-UI-007 | REQ-UI-002 | `TestREQ_UI_002_PortEqualToPrinterPort` | `ui.port` equal to the printer port; catches: only range checked |
| TC-UI-008 | REQ-UI-002 | `TestREQ_UI_002_PortRangeBoundaries` | `ui.port` 0 and 65536 refused, 1 and 65535 accepted; catches: `<=` / `<` off by one |
| TC-UI-009 | REQ-UI-002 | `TestREQ_UI_002_PortWrongType` | `ui.port` of wrong type (`"9180"`, `9180.5`, `null`); catches: string coerced with Atoi |
| TC-UI-010 | REQ-UI-002 | `TestREQ_UI_002_SectionAndRemoteWrongType` | `ui` not an object, `ui.remote` not a boolean; catches: silently default |
| TC-UI-011 | REQ-UI-002 | `TestREQ_UI_002_PortEqualToSecondPrinterPort` | `ui.port` equal to the 2nd printer's port; catches: only `Printers[0]` compared |
| TC-UI-012 | REQ-UI-003 | `TestREQ_UI_003_BrowserCommandIsRundll32` | Browser command is rundll32 FileProtocolHandler; catches: `cmd /c start` (breaks on `&`, XP) |
| TC-UI-013 | REQ-UI-003 | `TestREQ_UI_003_OpenerGetsURLAfterBind` | Opener called once with the UI URL after the bind; catches: opened before listen / with 0.0.0.0 |
| TC-UI-014 | REQ-UI-003 | `TestREQ_UI_003_NoConfigEmptyList` | No config: empty list, pages work; catches: nil Config dereference (500/panic) |
| TC-UI-015 | REQ-UI-003 | `TestREQ_UI_003_FirstSaveWritesConfigAndStartsReceivers` | First printer saved from the empty start writes the file and starts receivers; catches: save refused because no file exists |
| TC-UI-016 | REQ-UI-003 | MANUAL:admin | Real browser opens on double-click |
| TC-UI-017 | REQ-UI-004 | `TestREQ_UI_004_RemotePasswordLength` | remote with 7-char password / no password refused, 8 accepted; catches: `>8` instead of `>=8` |
| TC-UI-018 | REQ-UI-004 | `TestREQ_UI_004_ShortPasswordAcceptedWhenRemoteOff` | Remote off: short password accepted; catches: password checked regardless of remote |
| TC-UI-019 | REQ-UI-004 | `TestREQ_UI_004_RemoteBindsAllInterfaces` | Remote binds 0.0.0.0; catches: still loopback |
| TC-UI-020 | REQ-UI-004 | `TestREQ_UI_004_RemoteNeedsBasicAuth` | Remote: missing / wrong Basic auth → 401 on pages, API and assets; catches: only pages protected; user name not checked |
| TC-UI-021 | REQ-UI-004 | `TestREQ_UI_004_RemoteNeedsBasicAuth` | Remote: `admin` / password passes |
| TC-UI-022 | REQ-UI-004 | `TestREQ_UI_004_LocalAsksNoPassword` | Remote off: no auth asked; catches: auth required whenever a password exists |
| TC-UI-023 | REQ-UI-005 | `TestREQ_UI_005_StatusShowsEveryPrinter` | Status shows each printer's fields, UI and log paths; catches: only first printer rendered; no-allow shown empty |
| TC-UI-024 | REQ-UI-006 | `TestREQ_UI_006_AddEditRemoveWithPlainForms` | Printers page forms are plain POST forms with the token; catches: JS-only buttons |
| TC-UI-025 | REQ-UI-006 | `TestREQ_UI_006_AddEditRemoveWithPlainForms` | Add a printer with a plain form post; catches: CRLF / blank line kept as items |
| TC-UI-026 | REQ-UI-006 | `TestREQ_UI_006_AddEditRemoveWithPlainForms` | Edit a printer; catches: typed name ignored for the select |
| TC-UI-027 | REQ-UI-006 | `TestREQ_UI_006_AddEditRemoveWithPlainForms` | Remove a printer; catches: wrong index removed |
| TC-UI-028 | REQ-UI-007 | `TestREQ_UI_007_BadPortRefusedNextToField` | Port 70000 refused next to the field, nothing written; catches: write-then-validate; message far from field |
| TC-UI-029 | REQ-UI-007 | `TestREQ_UI_007_DuplicatePortRefused` | Duplicate port refused (F-003 rule); catches: form-only validation that misses loader rules |
| TC-UI-030 | REQ-UI-007 | `TestREQ_UI_007_BadAllowItemRefused` | Bad allow item refused (F-004 rule); catches: allow textarea not validated |
| TC-UI-031 | REQ-UI-007 | `TestREQ_UI_007_BadPortRefusedNextToField` | Message names the real config path, not the temp file; catches: temp path leaked |
| TC-UI-032 | REQ-UI-007 | `TestREQ_UI_007_RemovingLastPrinterRefused` | Removing the last printer refused; catches: empty list written (F-003 0 printers) |
| TC-UI-033 | REQ-UI-008 | `TestREQ_UI_008_KeepsUnknownKeysAtomically` | Unknown keys kept at every level, also on the edited entry; catches: struct re-marshal drops unknown keys |
| TC-UI-034 | REQ-UI-008 | `TestREQ_UI_008_KeepsUnknownKeysAtomically` | Atomic write leaves no temp file; catches: temp file name left behind |
| TC-UI-035 | REQ-UI-009 | `TestREQ_UI_009_RestartsOnlyChangedListeners` | First Apply starts every printer |
| TC-UI-036 | REQ-UI-009 | `TestREQ_UI_009_RestartsOnlyChangedListeners` | Port change restarts only that printer; catches: restart-all on any save |
| TC-UI-037 | REQ-UI-009 | `TestREQ_UI_009_RestartsOnlyChangedListeners` | Windows printer change and allow change each restart only theirs; name-only change restarts none; catches: allow / printer not in the change key; name in it |
| TC-UI-038 | REQ-UI-009 | `TestREQ_UI_009_RestartsOnlyChangedListeners` | Added printer started, removed printer stopped, others untouched; catches: index-keyed diff shifts after removal |
| TC-UI-039 | REQ-UI-009 | pending: cli / reload wiring test (build step) | Job in flight while its listener restarts is still printed; catches: listener closed hard |
| TC-UI-040 | REQ-UI-010 | `TestREQ_UI_010_TestPageSendsThroughPort` | Test page sends a job with the printer name and time through the port; catches: written to spooler directly, bypassing the port |
| TC-UI-041 | REQ-UI-010 | `TestREQ_UI_010_UnknownPrinterSendsNothing` | Unknown printer name sends nothing; catches: first printer used as fallback |
| TC-UI-042 | REQ-UI-010 | MANUAL:admin | Test page prints on a real printer |
| TC-UI-043 | REQ-UI-011 | `TestREQ_UI_011_LastTwoHundredLinesNewestLast` | Last 200 of 250 lines, newest last; catches: 201 lines / reversed order |
| TC-UI-044 | REQ-UI-011 | `TestREQ_UI_011_MissingLogFile` | Missing log file → page 200; catches: 500 on ENOENT |
| TC-UI-045 | REQ-UI-012 | `TestREQ_UI_012_SettingsSaveLogAndUI` | Settings save log and UI fields; catches: ui section not written |
| TC-UI-046 | REQ-UI-012 | `TestREQ_UI_012_PasswordNeverSentBack` | Password never sent back (page, API); catches: `value="{{.Password}}"` |
| TC-UI-047 | REQ-UI-012 | `TestREQ_UI_012_SettingsSaveLogAndUI` | UI change says restart is needed; catches: hint always shown / silent |
| TC-UI-048 | REQ-UI-012 | `TestREQ_UI_012_BadKeepDaysRefused` | keep_days 0 refused, nothing written; catches: settings bypass loader |
| TC-UI-049 | REQ-UI-012 | `TestREQ_UI_012_EmptyPasswordKeepsStored` | Empty password field keeps the stored password; catches: password wiped |
| TC-UI-050 | REQ-UI-013 | `TestREQ_UI_013_PostWithoutTokenIs403` | POST without / with wrong token → 403, nothing written; catches: token only on /printers |
| TC-UI-051 | REQ-UI-013 | `TestREQ_UI_013_PostWithoutTokenIs403` | API write without the header → 403; catches: API exempt |
| TC-UI-052 | REQ-UI-013 | `TestREQ_UI_013_TokenRandomPerProcess` | Token random per process; catches: constant token |
| TC-UI-053 | REQ-UI-014 | `TestREQ_UI_014_HostileValuesEscaped` | Hostile printer name / Windows printer escaped; catches: text/template used |
| TC-UI-054 | REQ-UI-014 | `TestREQ_UI_014_HostileValuesEscaped` | Hostile log line escaped; catches: log printed raw |
| TC-UI-055 | REQ-UI-014 | `TestREQ_UI_014_EchoedInputEscaped` | Hostile request value echoed back escaped; catches: input echoed raw in `value=` |
| TC-UI-056 | REQ-UI-015 | `TestREQ_UI_015_BaseCSSOnlyIE6` | base.css has no CSS IE6 lacks; catches: modern CSS creeping into base |
| TC-UI-057 | REQ-UI-015 | `TestREQ_UI_015_ModernCSSInsideSupportsGrid` | modern.css is one `@supports (display: grid)` block; catches: a rule outside the block |
| TC-UI-058 | REQ-UI-015 | `TestREQ_UI_015_ScriptsAreModulesAndAssetsLinked` | Every script tag is `type="module"`; templates link base.css, modern.css, app.js; catches: classic `<script src>` run by IE |
| TC-UI-059 | REQ-UI-016 | `TestREQ_UI_016_StreamSendsNewLines` | SSE stream sends one event per new log line; catches: whole file resent / polling only |
| TC-UI-060 | REQ-UI-016 | MANUAL:admin | With JS: live log, status counters refresh, forms save without reload via the form URLs; catches: JS-only endpoint |
| TC-UI-061 | REQ-UI-017 | `TestREQ_UI_017_GeneratedAssetsMatchStatic` | Generated asset file matches `static/`; catches: stale generated file |
| TC-UI-062 | REQ-UI-018 | `TestREQ_UI_018_ReadEndpoints` | GET API status / printers / installed; catches: HTML returned to API |
| TC-UI-063 | REQ-UI-018 | `TestREQ_UI_018_PostPrintersValidatesAndSaves` | POST /api/printers invalid → 400 error + field, nothing written; catches: 200 with error text |
| TC-UI-064 | REQ-UI-018 | `TestREQ_UI_018_PostPrintersValidatesAndSaves` | POST /api/printers valid saves |
| TC-UI-065 | REQ-UI-018 | `TestREQ_UI_018_LogLinesParam` | `/api/log?lines=N`; catches: `<=1000` off by one; default all |
| TC-UI-066 | REQ-UI-018 | `TestREQ_UI_018_PostTestSendsJob` | POST /api/test sends the job |
| TC-UI-067 | REQ-UI-019 | `scripts/harness check core-go110` | UI code builds and tests pass on Go 1.10.8; catches: embed / io/fs used |
| TC-UI-068 | REQ-UI-020 | MANUAL:admin | IE6 on XP: every page and form works, base look; catches: modern-only markup |
| TC-UI-069 | REQ-UI-020 | MANUAL:admin | IE8 on XP: same as IE6 |
| TC-UI-070 | REQ-UI-020 | MANUAL:admin | Current browser: modern look and live updates |
| TC-UI-071 | REQ-UI-003 | `TestREQ_UI_003_OpenerFailureKeepsUI` | Opener failure does not stop the UI; catches: browser failure aborts start-up |
| TC-UI-072 | REQ-UI-006 | `TestREQ_UI_006_BadIndexWritesNothing` | Edit / remove with a bad index writes nothing; catches: index unchecked (panic / last entry removed) |
| TC-UI-073 | REQ-UI-008 | `TestREQ_UI_008_WriteFailureKeepsOriginal` | Write failure keeps the original file; catches: original truncated before the new one exists |
| TC-UI-074 | REQ-UI-009 | `TestREQ_UI_009_FailedStartKeepsOthersAndRetries` | A listener that fails to start does not stop the others and is retried; catches: failed listener recorded as running |
| TC-UI-075 | REQ-UI-016 | `TestREQ_UI_016_StreamNeedsAuthWhenRemote` | SSE stream needs Basic auth with remote; catches: stream endpoint left open |
| TC-UI-076 | REQ-UI-013 | `TestREQ_UI_013_GetNeverChangesState` | State never changes through GET, even with the token; catches: token checked only on POST but FormValue read on GET (img-tag CSRF) |
| TC-UI-077 | REQ-UI-012 | `TestREQ_UI_012_SettingsValidateUI` | Settings validate `ui.*` like loading; catches: settings bypass the ui rules → service exits 2 on next start |
| TC-UI-078 | REQ-UI-007 | `TestREQ_UI_007_PrinterPortEqualToUIPortRefused` | Printer port equal to ui.port refused; catches: printer form ignores the ui port |
| TC-UI-079 | REQ-UI-012 | `TestREQ_UI_012_UntickedRemoteWritesFalse` | Unticking remote writes `ui.remote` false; catches: absent checkbox never written |
| TC-UI-080 | REQ-UI-008 | `TestREQ_UI_008_ConcurrentSavesLeaveValidFile` | Concurrent saves leave a valid file; catches: shared temp name / unguarded state |
| TC-UI-081 | REQ-UI-012 | `TestREQ_UI_012_SettingsShowCurrentValues` | Settings page shows current values; catches: inputs rendered empty |
| TC-UI-082 | REQ-UI-009 | `TestREQ_UI_009_ApplyFailureShownAfterSave` | Save written but listener start fails → error on the page; catches: error swallowed |
| TC-UI-083 | REQ-UI-001 | pending: cli / reload wiring test (build step) | CLI exit codes and no-args start; catches: config error not wired to exit 2 |
| TC-UI-084 | REQ-UI-014 | `TestREQ_UI_014_AppJSInsertsTextOnly` | app.js inserts server data as text only; catches: live log XSS in the browser |
| TC-UI-085 | REQ-UI-006 | `TestREQ_UI_006_InstalledFailureKeepsTypedName` | Installed-printer lookup failure keeps the form usable; catches: page 500 when the spooler is down |
| TC-UI-086 | REQ-UI-010 | `TestREQ_UI_010_NothingListeningShowsError` | Test job to a port with nothing listening; catches: hang / 500 |
| TC-UI-087 | REQ-UI-004 | `TestREQ_UI_004_PasswordLengthInCharacters` | Password length counted in characters; catches: `len()` counts bytes |
| TC-UI-088 | REQ-UI-015 | `TestREQ_UI_015_NoInlineHandlers` | No inline handlers or `javascript:` URLs in templates; catches: function reachable only through JS |
| TC-UI-089 | REQ-UI-011 | `TestREQ_UI_011_RefreshWithoutJS` | Log page has a refresh that works without JS; catches: refresh only through JS |
| TC-UI-090 | REQ-UI-021 | `TestREQ_UI_021_RejectsForeignHostWhenRemoteOff` | spec Q5, 2026-11-07 |
| TC-UI-091 | REQ-UI-022 | `TestREQ_UI_022_AllResponsesDenyFraming` | spec Q5, 2026-11-07 |
| TC-UI-092 | REQ-UI-008 | `TestREQ_UI_008_AddKeepsOtherPrintersAllowAndUnknownKeys` | review finding, 2026-11-07 |
| TC-UI-093 | REQ-UI-008 | `TestREQ_UI_008_RemoveKeepsSettingsWithTheirPrinter` | review finding, 2026-11-07 |
| TC-UI-094 | REQ-UI-016 | `TestREQ_UI_016_StreamWithPollUnsetDeliversNewLine` | review finding, 2026-11-07 |
| TC-UI-095 | REQ-UI-016 | `TestREQ_UI_016_StreamResumesFromLastEventID` | review finding, 2026-11-07 |
| TC-UI-096 | REQ-UI-008 | `TestREQ_UI_008_SaveTempFileNameHasNoPattern` | review finding, 2026-11-07 |
| TC-UI-097 | REQ-UI-016 | `TestREQ_UI_016_InvalidSaveAnswers400WithFieldMessage` | review finding, 2026-11-07 |
| TC-UI-098 | REQ-UI-004 | `TestREQ_UI_004_WrongLengthPasswordRefused` | review finding, 2026-11-07 |
| TC-UI-099 | REQ-UI-018 | `TestREQ_UI_018_APITestErrorIsJSONMessage` | review finding, 2026-11-07 |
| TC-UI-100 | REQ-UI-007 | `TestREQ_UI_007_OversizedPostRefused` | review finding, 2026-11-07 |
| TC-UI-101 | REQ-UI-012 | `TestREQ_UI_012_FirstSaveNeedsNoRestart` | review finding, 2026-11-07 |
| TC-UI-102 | REQ-UI-009 | `TestREQ_UI_009_PortSwapInOneSave` | review finding, 2026-11-07 |
| TC-UI-103 | REQ-UI-009 | `TestREQ_UI_009_ServiceKeepsRunningWhenUIReplacesListeners` | review finding, 2026-11-07 |
| TC-SVC-050 | REQ-SVC-016 | `TestREQ_SVC_016_XPShowFailureReported` | Break review: a failing XP `netsh firewall show portopening` was silently treated as no ports open; spec §2.4 makes it a reported failure (2026-11-07) |
