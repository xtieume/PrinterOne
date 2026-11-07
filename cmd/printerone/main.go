// Command printerone is the main edition (Windows 10/11): core + web UI + tray.
// Phase 0 skeleton: only `version`; the server arrives with F-001 (docs/specs/2026-11-07-go-rewrite-plan.md).
package main

import (
	"fmt"
	"os"

	"github.com/xtieume/PrinterOne/internal/core/buildinfo"
)

func main() {
	if len(os.Args) > 1 && os.Args[1] == "version" {
		fmt.Println(buildinfo.String("main"))
		return
	}
	fmt.Fprintln(os.Stderr, "usage: printerone version")
	os.Exit(2)
}
