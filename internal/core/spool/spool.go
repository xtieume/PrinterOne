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

// Print writes data to printer as one RAW document named docName, following the §2.3 call sequence:
// OpenPrinter, StartDocPrinter(RAW), StartPagePrinter, WritePrinter until every byte is accepted,
// EndPagePrinter, EndDocPrinter, ClosePrinter. A failure before EndDocPrinter succeeds aborts the
// document (AbortPrinter once StartDocPrinter succeeded), ClosePrinter is always called once the
// printer was opened, and a ClosePrinter failure after a successful EndDocPrinter does not fail the
// job because the spooler already holds it.
func Print(api API, printer, docName string, data []byte) error {
	h, err := api.OpenPrinter(printer)
	if err != nil {
		return err
	}
	docStarted := false // AbortPrinter is due for any failure while a document is open
	printed := false
	err = func() error {
		if err := api.StartDocPrinter(h, docName, "RAW"); err != nil {
			return err
		}
		docStarted = true
		if err := api.StartPagePrinter(h); err != nil {
			return err
		}
		for len(data) > 0 {
			n, err := api.WritePrinter(h, data)
			if err != nil {
				return err
			}
			if n <= 0 || n > len(data) {
				return &Error{Code: 0, Msg: "WritePrinter accepted " + strconv.Itoa(n) + " of " + strconv.Itoa(len(data)) + " bytes"}
			}
			data = data[n:]
		}
		if err := api.EndPagePrinter(h); err != nil {
			return err
		}
		if err := api.EndDocPrinter(h); err != nil {
			return err
		}
		printed = true
		return nil
	}()
	if docStarted && !printed {
		api.AbortPrinter(h) // its own failure is reported through err below, ClosePrinter still runs
	}
	api.ClosePrinter(h) // ignored on the failure path (err wins) and after EndDocPrinter succeeded
	if printed {
		return nil
	}
	return err
}

// List returns the installed printer names.
func List(api API) ([]string, error) {
	return api.EnumPrinters()
}
