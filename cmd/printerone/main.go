// Command printerone is the main edition (Windows 10/11): core + web UI + tray.
// Phase 0 skeleton: only `version`; the server arrives with F-001 (docs/specs/2026-11-07-go-rewrite-plan.md).
package main

import (
	"os"
	"os/signal"
	"path/filepath"

	"github.com/xtieume/PrinterOne/internal/core/cli"
	"github.com/xtieume/PrinterOne/internal/core/service"
	"github.com/xtieume/PrinterOne/internal/core/spool"
)

func main() {
	exe, _ := os.Executable()
	sig := make(chan os.Signal, 2)
	signal.Notify(sig, os.Interrupt)
	env := cli.Env{Edition: "main", ExeDir: filepath.Dir(exe), Getenv: os.Getenv,
		Spool: spool.System(), ExePath: exe, Service: service.System(), Stdout: os.Stdout, Stderr: os.Stderr, Interrupt: sig}
	os.Exit(cli.Run(os.Args[1:], env))
}
