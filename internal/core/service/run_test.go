package service

import (
	"bytes"
	"errors"
	"io"
	"io/ioutil"
	"math/rand"
	"net"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/spool"
)

// Tests of `service run` under the SCM (spec docs/specs/2026-11-07-f008-service.md §2.3), with a fake dispatcher,
// session and spooler. TC ids: .testcases/runs/F-008-r1 and testcases.md.

const testPrinter = "Canon MF220 Series"

// ---- fake spooler: records finished RAW documents ----

type fakeSpool struct {
	mu   sync.Mutex
	cur  map[spool.Handle]*bytes.Buffer
	next spool.Handle
	docs []string
}

func newFakeSpool() *fakeSpool { return &fakeSpool{cur: map[spool.Handle]*bytes.Buffer{}} }

func (f *fakeSpool) OpenPrinter(name string) (spool.Handle, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	if name != testPrinter {
		return 0, &spool.Error{Code: 1801, Msg: "The printer name is invalid."}
	}
	f.next++
	f.cur[f.next] = &bytes.Buffer{}
	return f.next, nil
}
func (f *fakeSpool) StartDocPrinter(h spool.Handle, n, dt string) error { return nil }
func (f *fakeSpool) StartPagePrinter(h spool.Handle) error              { return nil }
func (f *fakeSpool) WritePrinter(h spool.Handle, p []byte) (int, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.cur[h].Write(p)
	return len(p), nil
}
func (f *fakeSpool) EndPagePrinter(h spool.Handle) error { return nil }
func (f *fakeSpool) EndDocPrinter(h spool.Handle) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.docs = append(f.docs, f.cur[h].String())
	return nil
}
func (f *fakeSpool) AbortPrinter(h spool.Handle) error { return nil }
func (f *fakeSpool) ClosePrinter(h spool.Handle) error { return nil }
func (f *fakeSpool) EnumPrinters() ([]string, error)   { return []string{testPrinter}, nil }

func (f *fakeSpool) printed(doc string) bool {
	f.mu.Lock()
	defer f.mu.Unlock()
	for _, d := range f.docs {
		if d == doc {
			return true
		}
	}
	return false
}

func (f *fakeSpool) waitPrinted(doc string, d time.Duration) bool {
	for deadline := time.Now().Add(d); time.Now().Before(deadline); time.Sleep(20 * time.Millisecond) {
		if f.printed(doc) {
			return true
		}
	}
	return f.printed(doc)
}

// ---- fake session and dispatcher ----

type report struct {
	state State
	code  uint32
}

type fakeSession struct {
	mu      sync.Mutex
	reports []report
	ctl     chan Control
}

func (s *fakeSession) Report(state State, code uint32) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.reports = append(s.reports, report{state, code})
}
func (s *fakeSession) Controls() <-chan Control { return s.ctl }

func (s *fakeSession) list() []report {
	s.mu.Lock()
	defer s.mu.Unlock()
	return append([]report(nil), s.reports...)
}

func (s *fakeSession) saw(state State) bool {
	for _, r := range s.list() {
		if r.state == state {
			return true
		}
	}
	return false
}

// ---- rig ----

type runRig struct {
	t       *testing.T
	dir     string
	port    int
	spool   *fakeSpool
	session *fakeSession
	opts    RunOptions
	ready   chan string
	done    chan error
}

// Go 1.10 does not seed math/rand, so each test process seeds its own source.
var portRand = rand.New(rand.NewSource(time.Now().UnixNano() + int64(os.Getpid())))

func freePort(t *testing.T) int {
	t.Helper()
	// Pick from 20000-29999, outside the ephemeral ranges of macOS, Linux and Windows: a port taken from ":0"
	// and released can be handed straight to another test process's ":0" listener (seen as a flaky
	// "address already in use").
	for i := 0; i < 200; i++ {
		port := 20000 + portRand.Intn(10000)
		l, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(port))
		if err == nil {
			l.Close()
			return port
		}
	}
	t.Fatal("no free port in 20000-29999")
	return 0
}

func configJSON(port int) string {
	return `{"version":2,"printers":[{"name":"default","printer":"` + testPrinter + `","port":` + strconv.Itoa(port) + `}]}`
}

func newRunRig(t *testing.T, edition string) *runRig {
	dir, err := ioutil.TempDir("", "printerone-run")
	if err != nil {
		t.Fatal(err)
	}
	r := &runRig{t: t, dir: dir, port: freePort(t), spool: newFakeSpool(),
		session: &fakeSession{ctl: make(chan Control, 1)}, ready: make(chan string, 1), done: make(chan error, 1)}
	cfg := filepath.Join(dir, "config.json")
	if err := ioutil.WriteFile(cfg, []byte(configJSON(r.port)), 0644); err != nil {
		t.Fatal(err)
	}
	r.opts = RunOptions{ConfigPath: cfg, Edition: edition, DataDir: filepath.Join(dir, "ProgramData", "PrinterOne"),
		Spool: r.spool, Ready: func(a string) { r.ready <- a }}
	return r
}

func (r *runRig) logPath() string { return filepath.Join(r.opts.DataDir, "logs", "printerone.log") }

func (r *runRig) log() string {
	b, _ := ioutil.ReadFile(r.logPath())
	return string(b)
}

func (r *runRig) start() {
	d := Deps{Dispatch: func(name string, main func(Session)) error {
		if name != Name {
			return errors.New("dispatched under name " + name)
		}
		main(r.session)
		return nil
	}}
	go func() { r.done <- Run(r.opts, d) }()
}

// waitReady fails the test unless the service listens and reports Running within 5 s.
func (r *runRig) waitReady() string {
	r.t.Helper()
	select {
	case a := <-r.ready:
		for deadline := time.Now().Add(5 * time.Second); time.Now().Before(deadline); time.Sleep(10 * time.Millisecond) {
			if r.session.saw(Running) {
				return a
			}
		}
		r.t.Fatalf("listening on %s but Running never reported; reports %v", a, r.session.list())
	case err := <-r.done:
		r.t.Fatalf("Run returned %v before listening; reports %v; log %q", err, r.session.list(), r.log())
	case <-time.After(5 * time.Second):
		r.t.Fatalf("not listening within 5 s; reports %v", r.session.list())
	}
	return ""
}

// stop sends ControlStop and waits up to d for Run to return.
func (r *runRig) stop(c Control, d time.Duration) bool {
	select {
	case r.session.ctl <- c:
	default:
	}
	select {
	case <-r.done:
		return true
	case <-time.After(d):
		return false
	}
}

func (r *runRig) cleanup() {
	r.stop(ControlStop, 5*time.Second)
	os.RemoveAll(r.dir)
}

func (r *runRig) addr() string { return "127.0.0.1:" + strconv.Itoa(r.port) }

func (r *runRig) send(data string) {
	r.t.Helper()
	c, err := net.DialTimeout("tcp", r.addr(), 2*time.Second)
	if err != nil {
		r.t.Fatal(err)
	}
	c.Write([]byte(data))
	c.Close()
}

// openPartial connects and sends data without ending the job.
func (r *runRig) openPartial(data string) net.Conn {
	r.t.Helper()
	c, err := net.DialTimeout("tcp", r.addr(), 2*time.Second)
	if err != nil {
		r.t.Fatal(err)
	}
	c.Write([]byte(data))
	return c
}

func (r *runRig) last() report {
	l := r.session.list()
	if len(l) == 0 {
		return report{}
	}
	return l[len(l)-1]
}

// ---- REQ-SVC-010 ----

// TC-SVC-023
func TestREQ_SVC_010_ServesJobUnderSCM(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	r.start()
	r.waitReady()
	r.send("hello")
	if !r.spool.waitPrinted("hello", 5*time.Second) {
		t.Fatalf("job `hello` not printed under the SCM; log %q", r.log())
	}
}

// TC-SVC-024
func TestREQ_SVC_010_LegacyLimitUnderSCM(t *testing.T) {
	r := newRunRig(t, "legacy")
	defer r.cleanup()
	r.start()
	r.waitReady()
	a, b := r.openPartial("a"), r.openPartial("b")
	time.Sleep(200 * time.Millisecond)
	r.send("job3")
	if r.spool.waitPrinted("job3", time.Second) {
		a.Close()
		b.Close()
		t.Fatal("legacy: `job3` printed while 2 connections were receiving; want the limit 2")
	}
	a.Close()
	b.Close()
	if !r.spool.waitPrinted("job3", 5*time.Second) {
		t.Fatalf("legacy: `job3` not printed after the 2 connections ended; log %q", r.log())
	}
}

// TC-SVC-042
func TestREQ_SVC_010_MainLimitUnderSCM(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	r.start()
	r.waitReady()
	a, b := r.openPartial("a"), r.openPartial("b")
	defer a.Close()
	defer b.Close()
	time.Sleep(200 * time.Millisecond)
	r.send("job3")
	if !r.spool.waitPrinted("job3", 2*time.Second) {
		t.Fatalf("main: `job3` not printed with only 2 of 8 connections busy; log %q", r.log())
	}
}

// ---- REQ-SVC-011 ----

// TC-SVC-025
func TestREQ_SVC_011_LinesAppendedToLogFile(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	r.start()
	r.waitReady()
	r.send("hello")
	r.spool.waitPrinted("hello", 5*time.Second)
	var log string
	for deadline := time.Now().Add(5 * time.Second); time.Now().Before(deadline); time.Sleep(20 * time.Millisecond) {
		if log = r.log(); strings.Contains(log, "printed") {
			break
		}
	}
	if !strings.Contains(log, "Listening on") || !strings.Contains(log, "Printer: "+testPrinter) ||
		!strings.Contains(log, `5B "`+testPrinter+`" printed`) {
		t.Fatalf("%s:\n%q\nwant the start-up lines and the connection line `5B \"%s\" printed`", r.logPath(), log, testPrinter)
	}
}

// TC-SVC-026
func TestREQ_SVC_011_AppendsToExistingLog(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	os.MkdirAll(filepath.Dir(r.logPath()), 0755)
	if err := ioutil.WriteFile(r.logPath(), []byte("earlier line\n"), 0644); err != nil {
		t.Fatal(err)
	}
	r.start()
	r.waitReady()
	if log := r.log(); !strings.HasPrefix(log, "earlier line\n") || !strings.Contains(log, "Listening on") {
		t.Fatalf("log %q; want the earlier line kept and the start-up lines appended", log)
	}
}

type failingWriter struct{}

func (failingWriter) Write(p []byte) (int, error) { return 0, errors.New("disk full") }

// TC-SVC-027
func TestREQ_SVC_011_FailingLogWriteKeepsServing(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	opened := false
	r.opts.OpenLog = func(path string) (io.Writer, error) { opened = true; return failingWriter{}, nil }
	r.start()
	r.waitReady()
	r.send("hello")
	if !r.spool.waitPrinted("hello", 5*time.Second) || r.session.saw(Stopped) || !opened {
		t.Fatalf("opened %v, reports %v; want the job printed and no Stopped while log writes fail", opened, r.session.list())
	}
}

// ---- REQ-SVC-012 ----

// TC-SVC-028
func TestREQ_SVC_012_ConfigErrorStopsWithCode2(t *testing.T) {
	r := newRunRig(t, "main")
	defer os.RemoveAll(r.dir)
	ioutil.WriteFile(r.opts.ConfigPath, []byte(`{"version":1}`), 0644)
	r.start()
	select {
	case <-r.done:
	case <-time.After(5 * time.Second):
		t.Fatal("Run still running with an invalid config")
	}
	log := r.log()
	if r.session.saw(Running) || r.last() != (report{Stopped, 2}) || !strings.Contains(log, r.opts.ConfigPath) || !strings.Contains(log, "version") {
		t.Fatalf("reports %v, log %q; want never Running, last Stopped with code 2, the config error in the log", r.session.list(), log)
	}
}

// TC-SVC-029
func TestREQ_SVC_012_BusyPortStopsWithCode1(t *testing.T) {
	r := newRunRig(t, "main")
	defer os.RemoveAll(r.dir)
	occ, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(r.port))
	if err != nil {
		t.Fatal(err)
	}
	defer occ.Close()
	r.start()
	select {
	case <-r.done:
	case <-time.After(5 * time.Second):
		t.Fatal("Run still running on a busy port")
	}
	want := "port " + strconv.Itoa(r.port) + " is already in use"
	if r.session.saw(Running) || r.last() != (report{Stopped, 1}) || !strings.Contains(r.log(), want) {
		t.Fatalf("reports %v, log %q; want last Stopped with code 1 and %q in the log", r.session.list(), r.log(), want)
	}
}

// ---- REQ-SVC-013 ----

func stopPendingThenStopped(l []report) bool {
	sp := -1
	for i, r := range l {
		if r.state == StopPending && sp < 0 {
			sp = i
		}
	}
	return sp >= 0 && len(l) > sp+1 && l[len(l)-1] == (report{Stopped, 0})
}

// TC-SVC-030
func TestREQ_SVC_013_StopDropsReceivingPrintsReceived(t *testing.T) {
	r := newRunRig(t, "main")
	defer os.RemoveAll(r.dir)
	r.start()
	r.waitReady()
	partial := r.openPartial("partial")
	defer partial.Close()
	r.send("done")
	if !r.spool.waitPrinted("done", 5*time.Second) {
		t.Fatalf("`done` not printed before the stop; log %q", r.log())
	}
	if !r.stop(ControlStop, 30*time.Second) {
		t.Fatalf("Run did not return within 30 s of a stop request; reports %v", r.session.list())
	}
	if l := r.session.list(); !stopPendingThenStopped(l) || r.spool.printed("partial") || !strings.Contains(r.log(), "dropped: shutdown") {
		t.Fatalf("reports %v, log %q; want StopPending then Stopped(0), `partial` dropped (`dropped: shutdown`), not printed", l, r.log())
	}
	if c, err := net.DialTimeout("tcp", r.addr(), time.Second); err == nil {
		c.Close()
		t.Fatal("port still accepts after the service stopped")
	}
}

// TC-SVC-031
func TestREQ_SVC_013_ShutdownLikeStop(t *testing.T) {
	r := newRunRig(t, "main")
	defer os.RemoveAll(r.dir)
	r.start()
	r.waitReady()
	r.send("done")
	time.Sleep(200 * time.Millisecond) // received; maybe not printed yet
	if !r.stop(ControlShutdown, 30*time.Second) {
		t.Fatalf("Run did not return within 30 s of a shutdown request; reports %v", r.session.list())
	}
	if l := r.session.list(); !stopPendingThenStopped(l) || !r.spool.printed("done") {
		t.Fatalf("reports %v, printed `done` %v; want StopPending then Stopped(0) and the received job printed", l, r.spool.printed("done"))
	}
}
