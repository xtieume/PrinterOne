//go:build !windows
// +build !windows

package service

// System on a non-Windows OS: there is no SCM, every service call fails, so `go test` and `go vet` run on
// macOS and Linux (rule 01-go-code).
// ponytail: declarations only until the implementation phase of F-008.
func System() Deps { return Deps{} }
