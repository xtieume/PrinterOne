// Package receiver accepts RAW print jobs over TCP (spec docs/specs/2026-11-07-f001-core.md §2.2, §2.5).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package receiver

import (
	"io"
	"time"
)

// Limits from the spec.
const (
	MaxJobBytes = 268435456 // 256 MB: a job of exactly this size prints, one byte more is dropped
	IdleTimeout = 60 * time.Second
	SlotsMain   = 8
	SlotsLegacy = 2
)

// PrintFunc prints one fully received job. A *spool.Error makes the log line `failed: <code> <message>`.
type PrintFunc func(docName string, data []byte) error

// Server receives jobs on Addr and hands each complete one to Print, one at a time.
type Server struct {
	Addr     string        // "0.0.0.0:9100"; ":0" in tests
	Printer  string        // config `printer` value, used in log lines
	Print    PrintFunc     // writes one job
	MaxBytes int64         // MaxJobBytes in production
	Idle     time.Duration // IdleTimeout in production
	Slots    int           // connections received at once
	Log      io.Writer     // one line per connection (§2.5)
	Now      func() time.Time
}

// PortBusyError means the port is already in use; the program exits 1.
type PortBusyError struct {
	Port int
}

func (e *PortBusyError) Error() string { return "" } // F-001 scaffold: declaration only

// Listen binds the port. A port in use returns *PortBusyError.
func (s *Server) Listen() error {
	return nil // F-001 scaffold: declaration only
}

// ListenAddr is the bound address after Listen (useful with ":0").
func (s *Server) ListenAddr() string {
	return "" // F-001 scaffold: declaration only
}

// Serve accepts connections until Shutdown, then returns nil.
func (s *Server) Serve() error {
	return nil // F-001 scaffold: declaration only
}

// Shutdown stops accepting, drops connections still receiving (`dropped: shutdown`) and returns after every
// fully received job has been printed.
func (s *Server) Shutdown() {
	// F-001 scaffold: declaration only
}
