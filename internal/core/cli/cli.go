// Package cli is the command line shared by both editions (spec docs/specs/2026-11-07-f001-core.md §2.4).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package cli

import (
	"io"
	"os"

	"github.com/xtieume/PrinterOne/internal/core/spool"
)

// Env is everything Run takes from the process, so tests can drive it.
type Env struct {
	Edition   string              // "main" or "legacy"
	ExeDir    string              // folder of the running exe
	Getenv    func(string) string // os.Getenv in production
	Spool     spool.API           // spool.System() in production
	Stdout    io.Writer
	Stderr    io.Writer
	Interrupt <-chan os.Signal  // Ctrl+C
	Ready     func(addr string) // called once serve is listening; nil in production
}

// Run executes one command line (without the program name) and returns the process exit code.
func Run(args []string, env Env) int {
	return 0 // F-001 scaffold: declaration only
}
