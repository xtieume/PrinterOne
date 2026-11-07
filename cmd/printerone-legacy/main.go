// Command printerone-legacy is the Windows XP–8 edition, built with Go 1.10.8 from internal/core only.
// Go 1.10 syntax only (rule 01-go-code); `scripts/harness check core-go110` builds it.
package main

import (
	"fmt"
	"os"

	"github.com/xtieume/PrinterOne/internal/core/buildinfo"
)

func main() {
	if len(os.Args) > 1 && os.Args[1] == "version" {
		fmt.Println(buildinfo.String("legacy"))
		return
	}
	fmt.Fprintln(os.Stderr, "usage: printerone-legacy version")
	os.Exit(2)
}
