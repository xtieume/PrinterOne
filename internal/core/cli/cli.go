// Package cli is the command line shared by both editions (spec docs/specs/2026-11-07-f001-core.md §2.4).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package cli

import (
	"fmt"
	"io"
	"os"
	"strconv"

	"github.com/xtieume/PrinterOne/internal/core/buildinfo"
	"github.com/xtieume/PrinterOne/internal/core/config"
	"github.com/xtieume/PrinterOne/internal/core/receiver"
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
	default:
		return usage(env)
	}
}

func usage(env Env) int {
	fmt.Fprintln(env.Stderr, "usage: printerone [serve [--config path] | printers | version]")
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
	printer := cfg.Printers[0]
	slots := receiver.SlotsMain
	if env.Edition == "legacy" {
		slots = receiver.SlotsLegacy
	}
	s := &receiver.Server{Addr: "0.0.0.0:" + strconv.Itoa(printer.Port), Printer: printer.Printer,
		Slots: slots, MaxBytes: receiver.MaxJobBytes, Idle: receiver.IdleTimeout, Log: env.Stdout,
		Print: func(name string, data []byte) error { return spool.Print(env.Spool, printer.Printer, name, data) }}
	if err := s.Listen(); err != nil {
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
	addr := s.ListenAddr()
	fmt.Fprintf(env.Stdout, "Listening on %s\nPrinter: %s\n", addr, printer.Printer)
	if env.Ready != nil {
		env.Ready(addr)
	}
	done := make(chan struct{})
	var serveErr error
	go func() { serveErr = s.Serve(); close(done) }()
	if env.Interrupt == nil {
		<-done
		if serveErr != nil {
			fmt.Fprintln(env.Stderr, serveErr)
			return 1
		}
		return 0
	}
	select {
	case <-env.Interrupt:
	case <-done:
		if serveErr != nil {
			fmt.Fprintln(env.Stderr, serveErr)
			return 1
		}
		return 0
	}
	stopped := make(chan struct{})
	go func() { s.Shutdown(); <-done; close(stopped) }()
	select {
	case <-stopped:
		return 0
	case <-env.Interrupt:
		return 1
	}
}
