//go:build windows
// +build windows

package service

// System returns the real SCM (advapi32 via syscall.NewLazyDLL), netsh, the OS major version and the
// StartServiceCtrlDispatcher-based dispatcher (syscall.NewCallback for the service main and handler).
// ponytail: declarations only until the implementation phase of F-008.
func System() Deps { return Deps{} }
