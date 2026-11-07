// Command printerone-legacy is the Windows XP–8 edition, built with Go 1.10.8 from internal/core only.
// Go 1.10 syntax only (rule 01-go-code); `scripts/harness check core-go110` builds it.
package main

import (
	"os"
	"os/signal"
	"path/filepath"

	"github.com/xtieume/PrinterOne/internal/core/cli"
	"github.com/xtieume/PrinterOne/internal/core/spool"
)

func main() {
	exe, _ := os.Executable()
	sig := make(chan os.Signal, 2)
	signal.Notify(sig, os.Interrupt)
	env := cli.Env{Edition: "legacy", ExeDir: filepath.Dir(exe), Getenv: os.Getenv,
		Spool: spool.System(), Stdout: os.Stdout, Stderr: os.Stderr, Interrupt: sig}
	os.Exit(cli.Run(os.Args[1:], env))
}
