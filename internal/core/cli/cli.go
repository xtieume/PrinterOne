// Package cli is the command line shared by both editions (spec docs/specs/2026-11-07-f001-core.md §2.4).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package cli

import (
	"fmt"
	"io"
	"net"
	"os"
	"path/filepath"
	"strconv"
	"sync"

	"github.com/xtieume/PrinterOne/internal/core/buildinfo"
	"github.com/xtieume/PrinterOne/internal/core/config"
	"github.com/xtieume/PrinterOne/internal/core/logfile"
	"github.com/xtieume/PrinterOne/internal/core/receiver"
	"github.com/xtieume/PrinterOne/internal/core/service"
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
	Ready     func(addr string) // called once serve or service run is listening; nil in production
	ExePath   string            // absolute path of the running exe (service command line)
	Service   service.Deps      // service.System() in production
}

// Run executes one command line (without the program name) and returns the process exit code.
func Run(args []string, env Env) int {
	if env.Stdout == nil {
		env.Stdout = os.Stdout
	}
	if env.Stderr == nil {
		env.Stderr = os.Stderr
	}
	if len(args) == 0 {
		return serve(nil, env)
	}
	switch args[0] {
	case "serve":
		var path string
		if len(args) == 1 {
			return serve(nil, env)
		}
		if len(args) != 3 || args[1] != "--config" || args[2] == "" {
			return usage(env)
		}
		path = args[2]
		return serve(&path, env)
	case "version":
		if len(args) != 1 {
			return usage(env)
		}
		fmt.Fprintln(env.Stdout, buildinfo.String(env.Edition))
		return 0
	case "printers":
		if len(args) != 1 {
			return usage(env)
		}
		printers, err := spool.List(env.Spool)
		if err != nil {
			fmt.Fprintln(env.Stderr, err)
			return 1
		}
		for _, name := range printers {
			fmt.Fprintln(env.Stdout, name)
		}
		return 0
	case "service":
		return serviceCommand(args[1:], env)
	case "logs":
		return logsCommand(args[1:], env)
	default:
		return usage(env)
	}
}

func serviceCommand(args []string, env Env) int {
	if len(args) == 0 {
		return usage(env)
	}
	switch args[0] {
	case "install":
		var path string
		if len(args) == 1 {
			path = ""
		} else if len(args) == 3 && args[1] == "--config" && args[2] != "" {
			path = args[2]
		} else {
			return usage(env)
		}
		if err := service.RequireAdmin(env.Service); err != nil {
			return reportServiceError(err, env)
		}
		dataDir := config.DataDir(env.Getenv)
		cfg, err := config.Load(config.Lookup{Flag: path, ExeDir: env.ExeDir, DataDir: dataDir})
		if err != nil {
			fmt.Fprintln(env.Stderr, err)
			return 2
		}
		exe := env.ExePath
		if exe == "" {
			exe, err = os.Executable()
			if err != nil {
				fmt.Fprintln(env.Stderr, err)
				return 1
			}
		}
		if !isWindowsAbsolute(exe) {
			exe, err = filepath.Abs(exe)
			if err != nil {
				fmt.Fprintln(env.Stderr, err)
				return 1
			}
		}
		cfgPath, err := filepath.Abs(cfg.Path)
		if err != nil {
			fmt.Fprintln(env.Stderr, err)
			return 1
		}
		ports := make([]int, 0, len(cfg.Printers))
		for _, p := range cfg.Printers {
			ports = append(ports, p.Port)
		}
		if err = service.Install(service.InstallOptions{ExePath: exe, ConfigPath: cfgPath, Ports: ports}, env.Service); err != nil {
			return reportServiceError(err, env)
		}
		fmt.Fprintln(env.Stdout, "service PrinterOne installed and running")
		return 0
	case "uninstall":
		if len(args) != 1 {
			return usage(env)
		}
		if err := service.RequireAdmin(env.Service); err != nil {
			return reportServiceError(err, env)
		}
		warnings, err := service.Uninstall(env.Service)
		for _, w := range warnings {
			fmt.Fprintln(env.Stderr, w)
		}
		if err != nil {
			return reportServiceError(err, env)
		}
		fmt.Fprintln(env.Stdout, "service PrinterOne removed")
		return 0
	case "run":
		if len(args) != 3 || args[1] != "--config" || args[2] == "" {
			return usage(env)
		}
		path, err := filepath.Abs(args[2])
		if err != nil {
			fmt.Fprintln(env.Stderr, err)
			return 1
		}
		err = service.Run(service.RunOptions{ConfigPath: path, Edition: env.Edition, DataDir: config.DataDir(env.Getenv), Spool: env.Spool, Ready: env.Ready}, env.Service)
		if err != nil {
			return reportServiceError(err, env)
		}
		return 0
	default:
		return usage(env)
	}
}

// logsCommand is `printerone logs [--config p]`: prints the absolute path of the current log file, exit 0
// (spec docs/specs/2026-11-07-f002-logs.md §2).
func logsCommand(args []string, env Env) int {
	var path string
	if len(args) == 2 && args[0] == "--config" && args[1] != "" {
		path = args[1]
	} else if len(args) != 0 {
		return usage(env)
	}
	l, err := config.LoadLog(config.Lookup{Flag: path, ExeDir: env.ExeDir, DataDir: config.DataDir(env.Getenv)})
	if err != nil {
		fmt.Fprintln(env.Stderr, err)
		return 2
	}
	dir := l.Dir
	if dir == "" {
		dir = logfile.DefaultDir(config.DataDir(env.Getenv))
	}
	p := logfile.Path(dir)
	if !isWindowsAbsolute(p) {
		p, _ = filepath.Abs(p)
	}
	fmt.Fprintln(env.Stdout, p)
	return 0
}

func isWindowsAbsolute(path string) bool {
	return len(path) >= 3 && ((path[0] >= 'A' && path[0] <= 'Z') || (path[0] >= 'a' && path[0] <= 'z')) && path[1] == ':' && (path[2] == '\\' || path[2] == '/')
}

func reportServiceError(err error, env Env) int {
	if e, ok := err.(*service.Error); ok {
		fmt.Fprintln(env.Stderr, e.Msg)
		return e.Code
	}
	fmt.Fprintln(env.Stderr, err)
	return 1
}

func usage(env Env) int {
	fmt.Fprintln(env.Stderr, "usage: printerone [serve [--config path] | printers | logs [--config path] | version | service install [--config path] | service uninstall | service run --config path]")
	return 2
}

func serve(path *string, env Env) int {
	dataDir := config.DataDir(env.Getenv)
	flag := ""
	if path != nil {
		flag = *path
	}
	cfg, err := config.Load(config.Lookup{Flag: flag, ExeDir: env.ExeDir, DataDir: dataDir})
	if err != nil {
		fmt.Fprintln(env.Stderr, err)
		return 2
	}
	logDir := cfg.Log.Dir
	if logDir == "" {
		logDir = logfile.DefaultDir(dataDir)
	}
	logFile := &logfile.Writer{Dir: logDir, KeepDays: cfg.Log.KeepDays, Warn: env.Stderr}
	logFile.Start()
	output := io.MultiWriter(logFile, stdoutIgnoringErrors{env.Stdout})
	absLogPath, _ := filepath.Abs(logfile.Path(logDir))
	fmt.Fprintln(output, "Log file: "+absLogPath)
	slots := receiver.SlotsMain
	if env.Edition == "legacy" {
		slots = receiver.SlotsLegacy
	}
	pool := receiver.NewSlotPool(slots)
	locks := make(map[string]*sync.Mutex)
	servers := make([]*receiver.Server, 0, len(cfg.Printers))
	for i, printer := range cfg.Printers {
		printer := printer
		lock := locks[printer.Printer]
		if lock == nil {
			lock = &sync.Mutex{}
			locks[printer.Printer] = lock
		}
		var allowFn func(net.IP) bool
		if i < len(cfg.Allow) && cfg.Allow[i] != nil {
			allowFn = cfg.Allow[i].Allows
		}
		s := &receiver.Server{Addr: "0.0.0.0:" + strconv.Itoa(printer.Port), Printer: printer.Printer, Slots: slots, Pool: pool,
			MaxBytes: receiver.MaxJobBytes, Idle: receiver.IdleTimeout, Log: output, Allow: allowFn, PrintLock: lock,
			Print: func(name string, data []byte) error { return spool.Print(env.Spool, printer.Printer, name, data) }}
		if err := s.Listen(); err != nil {
			for _, previous := range servers {
				previous.CloseListener()
			}
			var busy *receiver.PortBusyError
			if e, ok := err.(*receiver.PortBusyError); ok {
				busy = e
			}
			if busy != nil {
				fmt.Fprintln(env.Stderr, busy)
				return 1
			}
			fmt.Fprintln(env.Stderr, err)
			return 1
		}
		servers = append(servers, s)
	}
	for i, s := range servers {
		p := cfg.Printers[i]
		fmt.Fprintf(output, "Listening on %s\nPrinter: %s\n", s.ListenAddr(), p.Printer)
		fmt.Fprintf(output, "%s -> %s on %s allow: %s\n", p.Name, p.Printer, s.ListenAddr(), allowText(cfg, i))
	}
	if env.Ready != nil && len(servers) > 0 {
		env.Ready(servers[0].ListenAddr())
	}
	dones := make([]chan struct{}, len(servers))
	for i, s := range servers {
		dones[i] = make(chan struct{})
		go func(s *receiver.Server, done chan struct{}) { s.Serve(); close(done) }(s, dones[i])
	}
	waitAll := func() {
		for i, s := range servers {
			s.Shutdown()
			<-dones[i]
		}
	}
	if env.Interrupt == nil {
		for _, done := range dones {
			<-done
		}
		return 0
	}
	select {
	case <-env.Interrupt:
	case <-dones[0]:
		return 0
	}
	stopped := make(chan struct{})
	go func() { waitAll(); close(stopped) }()
	select {
	case <-stopped:
		return 0
	case <-env.Interrupt:
		return 1
	}
}

func allowText(cfg *config.Config, i int) string {
	if i < len(cfg.Allow) {
		return cfg.Allow[i].String()
	}
	return "all"
}

type stdoutIgnoringErrors struct{ w io.Writer }

func (s stdoutIgnoringErrors) Write(p []byte) (int, error) {
	if s.w != nil {
		_, _ = s.w.Write(p)
	}
	return len(p), nil
}
