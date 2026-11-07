// Package buildinfo names the running build. Core package: Go 1.10 syntax only (rule 01-go-code).
package buildinfo

// Version is set at link time: -ldflags "-X github.com/xtieume/PrinterOne/internal/core/buildinfo.Version=v2.0.0".
var Version = "dev"

// String is the one-line identity printed by `printerone version`, e.g. "PrinterOne dev (legacy)".
func String(edition string) string {
	return "PrinterOne " + Version + " (" + edition + ")"
}
