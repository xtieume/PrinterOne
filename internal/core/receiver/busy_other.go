//go:build !windows
// +build !windows

package receiver

import (
	"net"
	"os"
	"syscall"
)

func isAddressInUse(err error) bool {
	op, ok := err.(*net.OpError)
	if !ok {
		return false
	}
	se, ok := op.Err.(*os.SyscallError)
	if !ok {
		return false
	}
	e, ok := se.Err.(syscall.Errno)
	return ok && e == syscall.EADDRINUSE
}
