// Package config finds, reads and validates config.json (spec docs/specs/2026-11-07-f001-core.md §2.1).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package config

// DefaultPort is used when a printer entry has no port.
const DefaultPort = 9100

// Printer is one shared printer: the Windows printer name and the TCP port it is served on.
type Printer struct {
	Name    string
	Printer string
	Port    int
}

// Config is a loaded, validated config file.
type Config struct {
	Path     string // the file it was read from
	Printers []Printer
}

// Lookup holds the places Load searches, in order: Flag (the --config value, "" when absent), ExeDir, DataDir.
type Lookup struct {
	Flag    string
	ExeDir  string
	DataDir string
}

// Error is a config problem; Msg is the single line printed to stderr and the program exits with code 2.
type Error struct {
	Msg string
}

func (e *Error) Error() string { return e.Msg }

// DataDir returns %ProgramData%\PrinterOne, or %ALLUSERSPROFILE%\Application Data\PrinterOne when
// %ProgramData% is not set (Windows XP).
func DataDir(getenv func(string) string) string {
	return "" // F-001 scaffold: declaration only
}

// Load finds the config by the lookup order and validates it. Every failure is an *Error.
func Load(l Lookup) (*Config, error) {
	return nil, nil // F-001 scaffold: declaration only
}
