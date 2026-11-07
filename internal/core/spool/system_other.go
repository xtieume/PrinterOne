//go:build !windows
// +build !windows

package spool

// System on a non-Windows OS: no spooler. Every call fails, so `go test` and `go vet` run on macOS/Linux.
func System() API {
	return nil // F-001 scaffold: declaration only
}
