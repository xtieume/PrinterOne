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
| TC-CFG-025 | REQ-CFG-005 | `TestREQ_CFG_005_InvalidFieldNamesField/TC-CFG-025` | Two printers (F-001 accepts exactly one); catches: extra entries silently ignored |
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
