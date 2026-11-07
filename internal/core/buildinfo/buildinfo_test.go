package buildinfo

import "testing"

func TestStringNamesVersionAndEdition(t *testing.T) {
	old := Version
	defer func() { Version = old }()
	Version = "v2.0.0"
	if got, want := String("legacy"), "PrinterOne v2.0.0 (legacy)"; got != want {
		t.Fatalf("String() = %q, want %q", got, want)
	}
}
