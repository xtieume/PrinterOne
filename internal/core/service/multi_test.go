package service

import (
	"io/ioutil"
	"net"
	"os"
	"strconv"
	"strings"
	"testing"
	"time"
)

// F-003 tests of `service run` with several printers (spec docs/specs/2026-11-07-f003-multi-printer.md §2). TC ids:
// .testcases/runs/F-003-r1 and testcases.md. The fake spooler knows one printer, so both entries use testPrinter.

// multiRunRig is a run rig whose config has entries front (r.port) and back (a second free port).
func multiRunRig(t *testing.T) (*runRig, int) {
	r := newRunRig(t, "main")
	p2 := freePort(t)
	body := `{"version":2,"printers":[{"name":"front","printer":"` + testPrinter + `","port":` + strconv.Itoa(r.port) +
		`},{"name":"back","printer":"` + testPrinter + `","port":` + strconv.Itoa(p2) + `}]}`
	if err := ioutil.WriteFile(r.opts.ConfigPath, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	return r, p2
}

func portRefused(port int) bool {
	c, err := net.DialTimeout("tcp", "127.0.0.1:"+strconv.Itoa(port), 500*time.Millisecond)
	if err != nil {
		return true
	}
	c.Close()
	return false
}

// TC-MULTI-026
func TestREQ_MULTI_007_ServiceBusySecondPortStopsWithCode1(t *testing.T) {
	r, p2 := multiRunRig(t)
	defer os.RemoveAll(r.dir) // Run has returned; cleanup would wait 5 s for it again
	busy, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(p2))
	if err != nil {
		t.Fatal(err)
	}
	defer busy.Close()
	r.start()
	select {
	case <-r.done:
	case <-time.After(5 * time.Second):
		t.Fatalf("service still running with port %d busy; reports %v", p2, r.session.list())
	}
	if r.last() != (report{Stopped, 1}) || !strings.Contains(r.log(), strconv.Itoa(p2)) {
		t.Fatalf("reports %v, log %q; want Stopped code 1 and the log naming port %d", r.session.list(), r.log(), p2)
	}
	l, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(r.port))
	if err != nil {
		t.Fatalf("port %d still bound after the service stopped: %v", r.port, err)
	}
	l.Close()
}

// TC-MULTI-032
func TestREQ_MULTI_009_ServiceLogListsEveryPrinter(t *testing.T) {
	r, p2 := multiRunRig(t)
	defer r.cleanup()
	r.start()
	r.waitReady()
	for _, want := range []string{
		"front -> " + testPrinter + " on 0.0.0.0:" + strconv.Itoa(r.port),
		"back -> " + testPrinter + " on 0.0.0.0:" + strconv.Itoa(p2),
	} {
		if !strings.Contains(r.log(), want) {
			t.Fatalf("log %q does not contain %q", r.log(), want)
		}
	}
}

// TC-MULTI-035
func TestREQ_MULTI_010_ServiceStopClosesEveryPort(t *testing.T) {
	r, p2 := multiRunRig(t)
	defer os.RemoveAll(r.dir)
	r.start()
	r.waitReady()
	if !r.stop(ControlStop, 5*time.Second) {
		t.Fatalf("Run did not return after ControlStop; reports %v", r.session.list())
	}
	if r.last() != (report{Stopped, 0}) || !portRefused(r.port) || !portRefused(p2) {
		t.Fatalf("reports %v; p1 refused=%v p2 refused=%v; want Stopped(0) and both ports closed",
			r.session.list(), portRefused(r.port), portRefused(p2))
	}
}
