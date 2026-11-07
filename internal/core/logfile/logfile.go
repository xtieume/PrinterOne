// Package logfile writes the log lines of `serve` and `service run` to one shared, daily rotated file
// (spec docs/specs/2026-11-07-f002-logs.md §2).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package logfile

import (
	"io"
	"sync"
	"time"
)

// FileName is the current log file in the log folder; rotated files are printerone-YYYY-MM-DD.log.
const FileName = "printerone.log"

// DefaultKeepDays is log.keep_days when the config has none.
const DefaultKeepDays = 30

// DefaultDir is <data dir>\logs, the log folder when the config has no log.dir.
func DefaultDir(dataDir string) string { return "" }

// Path is the absolute path of the current log file in dir.
func Path(dir string) string { return "" }

// Writer appends whole lines to Path(Dir). Each Write is one append to the file, which is never held open between
// writes, so other processes can read it and two writers (or two processes) never split each other's lines.
// After each write the file's modification time is set to Now(), so the local date of its modification time is
// the date of its last line — across restarts and across writers. A Write whose Now() falls on a later local date
// first renames the file to printerone-YYYY-MM-DD.log (that date) and deletes rotated files older than KeepDays:
// printerone-X.log is deleted when X < today − KeepDays; no other file is touched. Write never fails: the first
// error is reported once on Warn.
type Writer struct {
	Dir      string
	KeepDays int              // DefaultKeepDays when 0
	Now      func() time.Time // time.Now when nil
	Warn     io.Writer        // gets one line `<path>: <OS error>` on the first failed write; nil for none

	mu     sync.Mutex
	warned bool
}

// Start creates Dir when missing and deletes rotated files older than KeepDays (start-up clean-up).
func (w *Writer) Start() {}

// Write appends p to the current log file, rotating first when the local day changed. It always returns len(p), nil.
func (w *Writer) Write(p []byte) (int, error) { return 0, nil }
