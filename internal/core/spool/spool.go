// Package spool writes RAW jobs to a Windows printer (spec docs/specs/2026-11-07-f001-core.md §2.3).
// The call sequence lives here, OS-independent; the Windows build supplies the real API through syscall.
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package spool

import "strconv"

// Handle is a printer handle returned by OpenPrinter.
type Handle uintptr

// API is the winspool surface PrinterOne uses. System() returns the real one; tests use a fake.
type API interface {
	OpenPrinter(name string) (Handle, error)
	StartDocPrinter(h Handle, docName, datatype string) error
	StartPagePrinter(h Handle) error
	WritePrinter(h Handle, p []byte) (int, error)
	EndPagePrinter(h Handle) error
	EndDocPrinter(h Handle) error
	AbortPrinter(h Handle) error
	ClosePrinter(h Handle) error
	EnumPrinters() ([]string, error)
}

// Error is a failed Windows call: its error code and system message.
type Error struct {
	Code uint32
	Msg  string
}

func (e *Error) Error() string { return strconv.FormatUint(uint64(e.Code), 10) + " " + e.Msg }

// Print writes data to printer as one RAW document named docName (§2.3 call sequence).
func Print(api API, printer, docName string, data []byte) error {
	return nil // F-001 scaffold: declaration only
}

// List returns the installed printer names.
func List(api API) ([]string, error) {
	return nil, nil // F-001 scaffold: declaration only
}
