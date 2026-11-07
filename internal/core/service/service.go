// Package service installs, removes and runs PrinterOne as a Windows service with its firewall rule
// (spec docs/specs/2026-11-07-f008-service.md §2). The orchestration lives here, OS-independent, behind the
// small SCM / netsh / dispatcher seams; the Windows build supplies the real ones through syscall.
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package service

import (
	"errors"
	"io"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/spool"
)

// Service settings of spec §2.1 step 3.
const (
	Name        = "PrinterOne"
	DisplayName = "PrinterOne print server"
	Description = "Receives RAW print jobs over TCP and prints them"
	Account     = "LocalSystem"
	StartAuto   = 2 // SERVICE_AUTO_START
	RuleName    = "PrinterOne"
	WaitTimeout = 30 * time.Second // install waits for running, uninstall and stop requests for stopped
)

// Error is a service command failure; Msg is printed to stderr and the program exits with Code (1 or 2),
// like config.Error.
type Error struct {
	Code int
	Msg  string
}

func (e *Error) Error() string { return e.Msg }

// State is a SERVICE_STATUS dwCurrentState value.
type State uint32

const (
	Stopped      State = 1
	StartPending State = 2
	StopPending  State = 3
	Running      State = 4
)

// Settings is the configuration of an installed service.
type Settings struct {
	DisplayName string
	Description string
	BinaryPath  string // the full command line
	StartType   uint32
	Account     string
}

// SCM is the Service Control Manager surface install and uninstall use. System() returns the real one;
// tests use a fake.
type SCM interface {
	IsAdmin() bool
	Exists(name string) (bool, error)
	Create(name string, s Settings) error
	Update(name string, s Settings) error
	Config(name string) (Settings, error)
	Start(name string) error
	Stop(name string) error // sends SERVICE_CONTROL_STOP; does not wait
	Delete(name string) error
	State(name string) (State, error)
}

// Runner runs a program (netsh) and returns its combined output; err is non-nil when it exits non-zero.
type Runner func(name string, args ...string) ([]byte, error)

// Control is a request the SCM sends to a running service.
type Control uint32

const (
	ControlStop     Control = 1 // SERVICE_CONTROL_STOP
	ControlShutdown Control = 5 // SERVICE_CONTROL_SHUTDOWN
)

// Session is the running service's link to the SCM.
type Session interface {
	// Report sets the service status; exitCode is the service-specific exit code (0 = none).
	Report(state State, exitCode uint32)
	// Controls delivers stop and shutdown requests.
	Controls() <-chan Control
}

// Dispatcher connects the process to the SCM and calls main once the SCM starts service name; it returns
// when main returns. It returns ErrNotService when the process was not started by the SCM.
type Dispatcher func(name string, main func(Session)) error

// ErrNotService is what a Dispatcher returns when the process was started from a console.
var ErrNotService = errors.New("not started by the service control manager")

// Deps is everything the service commands take from the machine. System() returns the real ones.
type Deps struct {
	SCM      SCM
	Netsh    Runner
	OSMajor  int // Windows major version: 5 is XP (netsh firewall), 6 and later use netsh advfirewall
	Dispatch Dispatcher
	Sleep    func(time.Duration) // time.Sleep when nil
	Now      func() time.Time    // time.Now when nil
}

// InstallOptions is a validated `service install`.
type InstallOptions struct {
	ExePath    string // absolute path of the running exe
	ConfigPath string // absolute path of the resolved config
	Ports      []int  // configured ports
}

// RunOptions is a `service run --config p`.
type RunOptions struct {
	ConfigPath string
	Edition    string // "main" or "legacy": receiver slot limit
	DataDir    string // the log file is LogPath(DataDir)
	Spool      spool.API
	Ready      func(addr string)                    // called once listening; nil in production
	OpenLog    func(path string) (io.Writer, error) // appends to the file when nil
}

// RequireAdmin returns an exit-1 *Error printing `run as administrator` when the process is not elevated.
func RequireAdmin(d Deps) error { return nil }

// Install creates or updates the service, replaces the firewall rule and starts the service (§2.1 steps 3–5),
// undoing this run's changes when the rule or the start fails.
func Install(o InstallOptions, d Deps) error { return nil }

// Uninstall stops and deletes the service and removes the firewall rule (§2.2). Warnings (XP: ports that could
// not be closed) are returned for stderr; they do not fail the command.
func Uninstall(d Deps) (warnings []string, err error) { return nil, nil }

// Run hands the process to the SCM and serves the configured printer until a stop or shutdown request (§2.3).
// It returns an exit-2 *Error when the process was started from a console.
func Run(o RunOptions, d Deps) error { return nil }

// LogPath is <data dir>\logs\printerone.log.
func LogPath(dataDir string) string { return "" }

// Netsh runs a program and returns its combined output; the production Runner.
func Netsh(name string, args ...string) ([]byte, error) { return nil, nil }
