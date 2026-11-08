//go:build !windows
// +build !windows

package spool

// errNotWindows is what every call returns on a machine without a spooler, so
// `go test` and `go vet` run on macOS and Linux (rule 01-go-code).
// 120 is ERROR_CALL_NOT_IMPLEMENTED.
func errNotWindows() *Error {
	return &Error{Code: 120, Msg: "The spooler is only available on Windows."}
}

// System on a non-Windows OS: no spooler exists, every call fails.
func System() API { return stubAPI{} }

type stubAPI struct{}

func (stubAPI) OpenPrinter(name string) (Handle, error)            { return 0, errNotWindows() }
func (stubAPI) StartDocPrinter(h Handle, docName, dt string) error { return errNotWindows() }
func (stubAPI) StartPagePrinter(h Handle) error                    { return errNotWindows() }
func (stubAPI) WritePrinter(h Handle, p []byte) (int, error)       { return 0, errNotWindows() }
func (stubAPI) EndPagePrinter(h Handle) error                      { return errNotWindows() }
func (stubAPI) EndDocPrinter(h Handle) error                       { return errNotWindows() }
func (stubAPI) AbortPrinter(h Handle) error                        { return errNotWindows() }
func (stubAPI) ClosePrinter(h Handle) error                        { return errNotWindows() }
func (stubAPI) EnumPrinters() ([]string, error)                    { return nil, errNotWindows() }
