// Package service installs, removes and runs PrinterOne as a Windows service with its firewall rule
// (spec docs/specs/2026-11-07-f008-service.md §2). The orchestration lives here, OS-independent, behind the
// small SCM / netsh / dispatcher seams; the Windows build supplies the real ones through syscall.
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package service

import (
	"errors"
	"fmt"
	"io"
	"net"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/allow"
	"github.com/xtieume/PrinterOne/internal/core/config"
	"github.com/xtieume/PrinterOne/internal/core/logfile"
	"github.com/xtieume/PrinterOne/internal/core/receiver"
	"github.com/xtieume/PrinterOne/internal/core/spool"
	"github.com/xtieume/PrinterOne/internal/core/webui"
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
	State(name string) (State, uint32, uint32, error)
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
	SCM       SCM
	Netsh     Runner
	OSMajor   int // Windows major version: 5 is XP (netsh firewall), 6 and later use netsh advfirewall
	Dispatch  Dispatcher
	Sleep     func(time.Duration) // time.Sleep when nil
	Now       func() time.Time    // time.Now when nil
	NewTicker func(time.Duration) (<-chan time.Time, func())
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
func RequireAdmin(d Deps) error {
	if d.SCM == nil || !d.SCM.IsAdmin() {
		return &Error{Code: 1, Msg: "run as administrator"}
	}
	return nil
}

// Install creates or updates the service, replaces the firewall rule and starts the service (§2.1 steps 3–5),
// undoing this run's changes when the rule or the start fails.
func Install(o InstallOptions, d Deps) error {
	if err := RequireAdmin(d); err != nil {
		return err
	}
	var err error
	exe := o.ExePath
	if !isWindowsAbsolute(exe) {
		exe, err = filepath.Abs(exe)
		if err != nil {
			return serviceError(err)
		}
	}
	cfg, err := filepath.Abs(o.ConfigPath)
	if err != nil {
		return serviceError(err)
	}
	settings := Settings{DisplayName: DisplayName, Description: Description,
		BinaryPath: `"` + exe + `" service run --config "` + cfg + `"`, StartType: StartAuto, Account: Account}
	exists, err := d.SCM.Exists(Name)
	if err != nil {
		return serviceError(err)
	}
	created := false
	var old Settings
	if exists {
		old, err = d.SCM.Config(Name)
		if err != nil {
			return serviceError(err)
		}
		if err = d.SCM.Update(Name, settings); err != nil {
			return serviceError(err)
		}
	} else {
		if err = d.SCM.Create(Name, settings); err != nil {
			return serviceError(err)
		}
		created = true
	}

	if d.OSMajor == 5 && exists {
		oldPath := serviceConfigPath(old.BinaryPath)
		if oldPath != "" {
			if oldCfg, loadErr := config.Load(config.Lookup{Flag: oldPath}); loadErr == nil {
				oldPorts := make([]int, 0, len(oldCfg.Printers))
				for _, p := range oldCfg.Printers {
					oldPorts = append(oldPorts, p.Port)
				}
				if e := removeXPRules(d, oldPorts); e != nil {
					return rollbackInstall(d, created, nil, serviceError(e))
				}
			}
		}
	}
	if d.OSMajor == 5 {
		for _, port := range o.Ports {
			if _, err = runNetsh(d, "firewall", "add", "portopening", "protocol=TCP", "port="+strconv.Itoa(port), "name="+RuleName); err != nil {
				return rollbackInstall(d, created, func() { removeFirewall(d, o.Ports) }, serviceError(err))
			}
		}
	} else {
		if ruleExists(d) {
			_, err = runNetsh(d, "advfirewall", "firewall", "delete", "rule", "name="+RuleName)
		}
		if err != nil {
			return rollbackInstall(d, created, nil, serviceError(err))
		}
		ports := make([]string, 0, len(o.Ports))
		for _, p := range o.Ports {
			ports = append(ports, strconv.Itoa(p))
		}
		if _, err = runNetsh(d, "advfirewall", "firewall", "add", "rule", "name="+RuleName, "dir=in", "action=allow", "protocol=TCP", "localport="+strings.Join(ports, ",")); err != nil {
			return rollbackInstall(d, created, func() { removeFirewall(d, o.Ports) }, serviceError(err))
		}
	}
	if err = startAndWait(d); err != nil {
		return rollbackInstall(d, created, func() { removeFirewall(d, o.Ports) }, serviceError(err))
	}
	return nil
}

// Uninstall stops and deletes the service and removes the firewall rule (§2.2). Warnings (XP: ports that could
// not be closed) are returned for stderr; they do not fail the command.
func Uninstall(d Deps) (warnings []string, err error) {
	if err = RequireAdmin(d); err != nil {
		return nil, err
	}
	exists, err := d.SCM.Exists(Name)
	if err != nil {
		return nil, serviceError(err)
	}
	var old Settings
	var ports []int
	if exists {
		old, err = d.SCM.Config(Name)
		if err != nil {
			return nil, serviceError(err)
		}
		if d.OSMajor == 5 {
			path := serviceConfigPath(old.BinaryPath)
			cfg, loadErr := config.Load(config.Lookup{Flag: path})
			if loadErr != nil {
				warnings = append(warnings, fmt.Sprintf("could not read service config %s; XP firewall ports cannot be known", path))
			} else {
				for _, p := range cfg.Printers {
					ports = append(ports, p.Port)
				}
			}
		}
		state, _, _, e := d.SCM.State(Name)
		if e != nil {
			return warnings, serviceError(e)
		}
		if state != Stopped {
			if state != StopPending {
				if e = d.SCM.Stop(Name); e != nil {
					return warnings, serviceError(e)
				}
			}
			if e = waitForState(d, Stopped, WaitTimeout); e != nil {
				return warnings, serviceError(e)
			}
		}
		if err = d.SCM.Delete(Name); err != nil {
			return warnings, serviceError(err)
		}
	}
	if d.OSMajor == 5 {
		if e := removeXPRules(d, ports); e != nil {
			return warnings, serviceError(e)
		}
	} else if ruleExists(d) {
		if _, e := runNetsh(d, "advfirewall", "firewall", "delete", "rule", "name="+RuleName); e != nil {
			return warnings, serviceError(e)
		}
	}
	return warnings, nil
}

func removeXPRules(d Deps, ports []int) error {
	out, err := runNetsh(d, "firewall", "show", "portopening")
	if err != nil {
		return err // spec §2.4: a failing XP `show` is a failure reported with netsh's output
	}
	for _, p := range ports {
		if portListed(string(out), p) {
			if _, err = runNetsh(d, "firewall", "delete", "portopening", "protocol=TCP", "port="+strconv.Itoa(p)); err != nil {
				return err
			}
		}
	}
	return nil
}

func portListed(output string, port int) bool {
	want := strconv.Itoa(port)
	for _, field := range strings.Fields(output) {
		field = strings.Trim(field, "|,:;()[]")
		if field == want {
			return true
		}
	}
	return false
}

// Run hands the process to the SCM and serves the configured printer until a stop or shutdown request (§2.3).
// It returns an exit-2 *Error when the process was started from a console.
func Run(o RunOptions, d Deps) error {
	if d.Dispatch == nil {
		return &Error{Code: 2, Msg: "service run is started by Windows; use serve to run in a console"}
	}
	err := d.Dispatch(Name, func(session Session) { runSession(o, session, d) })
	if err == ErrNotService {
		return &Error{Code: 2, Msg: "service run is started by Windows; use serve to run in a console"}
	}
	return err
}

// LogPath is <data dir>\logs\printerone.log.
func LogPath(dataDir string) string { return filepath.Join(dataDir, "logs", "printerone.log") }

// Netsh runs a program and returns its combined output; the production Runner.
func Netsh(name string, args ...string) ([]byte, error) {
	return exec.Command(name, args...).CombinedOutput()
}

func serviceError(err error) *Error {
	if e, ok := err.(*Error); ok {
		return e
	}
	return &Error{Code: 1, Msg: err.Error()}
}

func runNetsh(d Deps, args ...string) ([]byte, error) {
	runner := d.Netsh
	if runner == nil {
		runner = Netsh
	}
	out, err := runner("netsh", args...)
	if err != nil {
		if strings.TrimSpace(string(out)) != "" {
			return out, errors.New(strings.TrimSpace(string(out)))
		}
	}
	return out, err
}

func ruleExists(d Deps) bool {
	_, err := runNetsh(d, "advfirewall", "firewall", "show", "rule", "name="+RuleName)
	return err == nil
}

func removeFirewall(d Deps, ports []int) {
	if d.OSMajor == 5 {
		_ = removeXPRules(d, ports)
	} else {
		if ruleExists(d) {
			runNetsh(d, "advfirewall", "firewall", "delete", "rule", "name="+RuleName)
		}
	}
}

func rollbackInstall(d Deps, created bool, cleanup func(), cause error) error {
	if cleanup != nil {
		cleanup()
	}
	if created {
		if state, _, _, e := d.SCM.State(Name); e == nil && state != Stopped {
			if state != StopPending {
				d.SCM.Stop(Name)
			}
			waitForState(d, Stopped, WaitTimeout)
		}
		d.SCM.Delete(Name)
	}
	return cause
}

func sleep(d Deps, duration time.Duration) {
	if d.Sleep != nil {
		d.Sleep(duration)
	} else {
		time.Sleep(duration)
	}
}

func startAndWait(d Deps) error {
	state, _, _, err := d.SCM.State(Name)
	if err != nil {
		return err
	}
	if state == Running {
		if err = d.SCM.Stop(Name); err != nil {
			return err
		}
		if err = waitForState(d, Stopped, WaitTimeout); err != nil {
			return err
		}
	} else if state == StartPending {
		if err = waitForState(d, Running, WaitTimeout); err != nil {
			return err
		}
		if err = d.SCM.Stop(Name); err != nil {
			return err
		}
		if err = waitForState(d, Stopped, WaitTimeout); err != nil {
			return err
		}
	}
	if err = d.SCM.Start(Name); err != nil {
		return err
	}
	return waitForState(d, Running, WaitTimeout)
}

func waitForState(d Deps, want State, timeout time.Duration) error {
	start := time.Now()
	if d.Now != nil {
		start = d.Now()
	}
	for {
		state, win32, specific, err := d.SCM.State(Name)
		if err != nil {
			return err
		}
		if state == want {
			return nil
		}
		if want == Running && state == Stopped {
			return &Error{Code: 1, Msg: fmt.Sprintf("service stopped while waiting for running (Win32 exit code %d, service-specific exit code %d)", win32, specific)}
		}
		now := time.Now()
		if d.Now != nil {
			now = d.Now()
		}
		if now.Sub(start) >= timeout {
			return &Error{Code: 1, Msg: fmt.Sprintf("timed out waiting for service %s", stateLabel(want))}
		}
		sleep(d, time.Second)
	}
}

func stateLabel(s State) string {
	if s == Running {
		return "running"
	}
	return "stopped"
}

func serviceConfigPath(line string) string {
	marker := "service run --config \""
	i := strings.Index(line, marker)
	if i < 0 {
		return ""
	}
	value := line[i+len(marker):]
	if end := strings.Index(value, "\""); end >= 0 {
		return value[:end]
	}
	return ""
}

func isWindowsAbsolute(path string) bool {
	return len(path) >= 3 && ((path[0] >= 'A' && path[0] <= 'Z') || (path[0] >= 'a' && path[0] <= 'z')) && path[1] == ':' && (path[2] == '\\' || path[2] == '/')
}

type logWriter struct{ w io.Writer }

func (l *logWriter) Write(p []byte) (int, error) {
	if l == nil || l.w == nil {
		return len(p), nil
	}
	n, err := l.w.Write(p)
	if err != nil {
		return len(p), nil
	}
	return n, nil
}

func openLog(o RunOptions, settings config.Log) (io.Writer, func()) {
	dir := settings.Dir
	if dir == "" {
		dir = logfile.DefaultDir(o.DataDir)
	}
	if o.OpenLog != nil {
		p := logfile.Path(dir)
		w, err := o.OpenLog(p)
		if err != nil {
			return &logWriter{}, func() {}
		}
		return &logWriter{w}, func() {
			if c, ok := w.(io.Closer); ok {
				c.Close()
			}
		}
	}
	w := &logfile.Writer{Dir: dir, KeepDays: settings.KeepDays}
	w.Start()
	return &logWriter{w}, func() {}
}

func runSession(o RunOptions, session Session, d Deps) {
	session.Report(StartPending, 0)
	lookup := config.Lookup{Flag: o.ConfigPath}
	settings, logErr := config.LoadLog(lookup)
	if logErr != nil {
		settings = config.Log{}
	}
	log, closeLog := openLog(o, settings)
	defer closeLog()
	logDir := settings.Dir
	if logDir == "" {
		logDir = logfile.DefaultDir(o.DataDir)
	}
	fmt.Fprintln(log, "Log file: "+absoluteLogPath(logfile.Path(logDir)))
	if logErr != nil {
		fmt.Fprintln(log, logErr)
		session.Report(Stopped, 2)
		return
	}
	cfg, err := config.Load(lookup)
	if err != nil {
		fmt.Fprintln(log, err)
		session.Report(Stopped, 2)
		return
	}
	ui, err := config.UIOf(cfg)
	if err != nil {
		fmt.Fprintln(log, err)
		session.Report(Stopped, 2)
		return
	}
	if o.Ready != nil && ui.Port == config.DefaultUIPort {
		ui.Port = 0
	}
	slots := receiver.SlotsMain
	if o.Edition == "legacy" {
		slots = receiver.SlotsLegacy
	}
	pool := receiver.NewSlotPool(slots)
	locks := make(map[string]*sync.Mutex)
	var serversMu sync.Mutex
	var servers []*receiver.Server
	manager := &webui.Listeners{Listen: func(p config.Printer, a allow.List) (func(), error) {
		key := strings.ToLower(p.Printer) // Windows printer names are case-insensitive
		lock := locks[key]
		if lock == nil {
			lock = &sync.Mutex{}
			locks[key] = lock
		}
		var allowFn func(net.IP) bool
		if a != nil {
			allowFn = a.Allows
		}
		srv := &receiver.Server{Addr: "0.0.0.0:" + strconv.Itoa(p.Port), Printer: p.Printer, Slots: slots, Pool: pool, MaxBytes: receiver.MaxJobBytes, Idle: receiver.IdleTimeout, Log: log, Allow: allowFn, PrintLock: lock, Print: func(name string, data []byte) error { return spool.Print(o.Spool, p.Printer, name, data) }}
		if e := srv.Listen(); e != nil {
			return nil, e
		}
		serversMu.Lock()
		servers = append(servers, srv) // ponytail: stopped servers stay listed (one per reload); a second Shutdown is a no-op
		serversMu.Unlock()
		fmt.Fprintf(log, "Listening on %s\nPrinter: %s\n", srv.ListenAddr(), p.Printer)
		allowValue := a.String()
		if allowValue == "" {
			allowValue = "all"
		}
		fmt.Fprintf(log, "%s -> %s on %s allow: %s\n", p.Name, p.Printer, srv.ListenAddr(), allowValue)
		go srv.Serve()
		return func() { srv.Shutdown() }, nil
	}}
	if err = manager.Apply(cfg); err != nil {
		manager.Apply(&config.Config{})
		fmt.Fprintln(log, err)
		session.Report(Stopped, 1)
		return
	}
	logPath := logfile.Path(logDir)
	uiDeps := webui.Deps{Config: cfg, ConfigPath: o.ConfigPath, Installed: func() ([]string, error) { return spool.List(o.Spool) }, Stats: func(name string) (int, string) { return webui.LogStats(logPath, name, time.Now()) }, LogPath: logPath, Apply: manager.Apply, UI: ui, Poll: 500 * time.Millisecond}
	uiAddr, stopUI, err := webui.Start(ui, webui.New(uiDeps), nil)
	if err != nil {
		manager.Apply(&config.Config{})
		fmt.Fprintln(log, err)
		session.Report(Stopped, 1)
		return
	}
	addr := uiAddr
	if len(cfg.Printers) > 0 {
		addr = "0.0.0.0:" + strconv.Itoa(cfg.Printers[0].Port)
	}
	if o.Ready != nil {
		o.Ready(addr)
	}
	session.Report(Running, 0)
	// Only the SCM stops the service: a UI save may replace every listener (REQ-UI-009).
	<-session.Controls()
	session.Report(StopPending, 0)
	stopUI()
	stopped := make(chan struct{})
	go func() {
		serversMu.Lock()
		all := append([]*receiver.Server(nil), servers...)
		serversMu.Unlock()
		receiver.ShutdownAll(all) // every port at once (REQ-MULTI-010)
		manager.Apply(&config.Config{})
		close(stopped)
	}()
	newTicker := d.NewTicker
	if newTicker == nil {
		newTicker = func(d time.Duration) (<-chan time.Time, func()) { t := time.NewTicker(d); return t.C, t.Stop }
	}
	ticks, stopTicker := newTicker(2 * time.Second)
	defer stopTicker()
	checkpoint := uint32(1)
	for {
		select {
		case <-stopped:
			session.Report(Stopped, 0)
			return
		case <-ticks:
			checkpoint++
			if p, ok := session.(interface{ ReportPending(uint32) }); ok {
				p.ReportPending(checkpoint)
			}
		}
	}
}

func allowText(cfg *config.Config, i int) string {
	if i < len(cfg.Allow) {
		return cfg.Allow[i].String()
	}
	return "all"
}

func absoluteLogPath(path string) string {
	p, err := filepath.Abs(path)
	if err != nil {
		return path
	}
	return p
}
