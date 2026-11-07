package cli

import (
	"bytes"
	"io/ioutil"
	"math/rand"
	"net"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/buildinfo"
	"github.com/xtieume/PrinterOne/internal/core/spool"
)

// TestMain fixes the local zone once, before any server goroutine runs: a test that swaps time.Local while
// another test's server is still logging is a data race (seen under -race on 2026-11-07).
func TestMain(m *testing.M) {
	time.Local = time.FixedZone("UTC+7", 7*3600)
	os.Exit(m.Run())
}

const printerName = "Canon MF220 Series"

// ---- fake spooler (fixture S-fake for the CLI) ----

type doc struct {
	printer, name, datatype string
	data                    []byte
	ended, closed, aborted  bool
}

type fakeSpool struct {
	mu       sync.Mutex
	docs     []*doc
	cur      map[spool.Handle]*doc
	next     spool.Handle
	opens    []string
	openErr  error
	openFail int // when > 0, the openFail-th OpenPrinter call and later ones fail with openErr
	enum     []string
	enumErr  error
	gate     chan struct{} // when non-nil, every WritePrinter waits for one receive
	countAll bool          // keep only the byte count of written data (big jobs)
	count    int64
	writing  chan struct{} // signalled (non-blocking) when a WritePrinter starts
	startDoc int
	open     int // documents started and not yet closed
	maxOpen  int
	enumCall int
}

func newFakeSpool() *fakeSpool {
	return &fakeSpool{cur: map[spool.Handle]*doc{}, writing: make(chan struct{}, 16)}
}

func (f *fakeSpool) OpenPrinter(name string) (spool.Handle, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.opens = append(f.opens, name)
	if f.openErr != nil && (f.openFail == 0 || len(f.opens) >= f.openFail) {
		return 0, f.openErr
	}
	f.next++
	f.cur[f.next] = &doc{printer: name}
	return f.next, nil
}
func (f *fakeSpool) StartDocPrinter(h spool.Handle, name, datatype string) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.startDoc++
	f.open++
	if f.open > f.maxOpen {
		f.maxOpen = f.open
	}
	d := f.cur[h]
	d.name, d.datatype = name, datatype
	f.docs = append(f.docs, d)
	return nil
}
func (f *fakeSpool) StartPagePrinter(h spool.Handle) error { return nil }
func (f *fakeSpool) WritePrinter(h spool.Handle, p []byte) (int, error) {
	select {
	case f.writing <- struct{}{}:
	default:
	}
	if f.gate != nil {
		<-f.gate
	}
	f.mu.Lock()
	defer f.mu.Unlock()
	if f.countAll {
		f.count += int64(len(p))
		return len(p), nil
	}
	f.cur[h].data = append(f.cur[h].data, p...)
	return len(p), nil
}
func (f *fakeSpool) EndPagePrinter(h spool.Handle) error { return nil }
func (f *fakeSpool) EndDocPrinter(h spool.Handle) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.cur[h].ended = true
	return nil
}
func (f *fakeSpool) AbortPrinter(h spool.Handle) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.cur[h].aborted = true
	return nil
}
func (f *fakeSpool) ClosePrinter(h spool.Handle) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	if d := f.cur[h]; d != nil {
		if d.name != "" && !d.closed {
			f.open--
		}
		d.closed = true
	}
	return nil
}
func (f *fakeSpool) EnumPrinters() ([]string, error) {
	f.mu.Lock()
	f.enumCall++
	f.mu.Unlock()
	return f.enum, f.enumErr
}

// done returns the documents that were ended and closed.
func (f *fakeSpool) done() []doc {
	f.mu.Lock()
	defer f.mu.Unlock()
	var out []doc
	for _, d := range f.docs {
		if d.ended && d.closed {
			out = append(out, *d)
		}
	}
	return out
}

func (f *fakeSpool) waitDone(n int, d time.Duration) []doc {
	deadline := time.Now().Add(d)
	for {
		docs := f.done()
		if len(docs) >= n || time.Now().After(deadline) {
			return docs
		}
		time.Sleep(10 * time.Millisecond)
	}
}

// ---- Run harness (fixture C-std) ----

type syncBuf struct {
	mu sync.Mutex
	b  bytes.Buffer
}

func (s *syncBuf) Write(p []byte) (int, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.b.Write(p)
}
func (s *syncBuf) String() string {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.b.String()
}

type rig struct {
	t              *testing.T
	dir, exe, pd   string
	cfg            string // C-std config file (for --config)
	port           int
	sp             *fakeSpool
	stdout, stderr syncBuf
	sig            chan os.Signal
	ready          chan string
	exit           chan int
	env            Env
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
	return `{"version":2,"printers":[{"name":"front","printer":"` + printerName + `","port":` + strconv.Itoa(port) + `}]}`
}

func newRig(t *testing.T, edition string) *rig {
	t.Helper()
	dir, err := ioutil.TempDir("", "clitest")
	if err != nil {
		t.Fatal(err)
	}
	r := &rig{t: t, dir: dir, exe: filepath.Join(dir, "exe"), pd: filepath.Join(dir, "pd"), port: freePort(t),
		sp: newFakeSpool(), sig: make(chan os.Signal, 4), ready: make(chan string, 4), exit: make(chan int, 1)}
	os.MkdirAll(r.exe, 0755)
	os.MkdirAll(r.pd, 0755)
	r.cfg = r.write(filepath.Join(dir, "cfg.json"), configJSON(r.port))
	env := map[string]string{"ProgramData": r.pd}
	r.env = Env{Edition: edition, ExeDir: r.exe, Getenv: func(k string) string { return env[k] }, Spool: r.sp,
		Stdout: &r.stdout, Stderr: &r.stderr, Interrupt: r.sig, Ready: func(a string) { r.ready <- a }}
	return r
}

func (r *rig) cleanup() { os.RemoveAll(r.dir) }

func (r *rig) write(path, body string) string {
	os.MkdirAll(filepath.Dir(path), 0755)
	if err := ioutil.WriteFile(path, []byte(body), 0644); err != nil {
		r.t.Fatal(err)
	}
	return path
}

func (r *rig) start(args ...string) {
	go func() { r.exit <- Run(args, r.env) }()
}

// run runs a command that must return by itself within 5 s.
func (r *rig) run(args ...string) int {
	r.t.Helper()
	r.start(args...)
	code, ok := r.waitExit(5 * time.Second)
	if !ok {
		r.sig <- os.Interrupt
		r.t.Fatalf("Run(%q) did not return within 5 s (server started?)", args)
	}
	return code
}

func (r *rig) waitExit(d time.Duration) (int, bool) {
	select {
	case c := <-r.exit:
		return c, true
	case <-time.After(d):
		return 0, false
	}
}

func (r *rig) waitReady() string {
	r.t.Helper()
	select {
	case a := <-r.ready:
		return a
	case c := <-r.exit:
		r.t.Fatalf("Run returned %d before Ready; stderr %q", c, r.stderr.String())
	case <-time.After(5 * time.Second):
		r.t.Fatal("Ready not called within 5 s")
	}
	return ""
}

func (r *rig) noReady() {
	r.t.Helper()
	select {
	case a := <-r.ready:
		r.t.Fatalf("Ready(%q) called: the server started", a)
	default:
	}
}

func (r *rig) addr() string { return "127.0.0.1:" + strconv.Itoa(r.port) }

func (r *rig) dial() *net.TCPConn {
	r.t.Helper()
	c, err := net.DialTimeout("tcp", r.addr(), 2*time.Second)
	if err != nil {
		r.t.Fatalf("dial %s: %v", r.addr(), err)
	}
	return c.(*net.TCPConn)
}

func (r *rig) send(data string) *net.TCPConn {
	r.t.Helper()
	c := r.dial()
	if _, err := c.Write([]byte(data)); err != nil {
		r.t.Fatal(err)
	}
	return c
}

// stop sends one Ctrl+C and expects exit 0 within 5 s.
func (r *rig) stop() {
	r.t.Helper()
	r.sig <- os.Interrupt
	code, ok := r.waitExit(5 * time.Second)
	if !ok || code != 0 {
		r.t.Fatalf("after Ctrl+C: returned=%v code=%d, want 0", ok, code)
	}
}

func (r *rig) connLines() []string {
	var out []string
	for _, l := range strings.Split(r.stdout.String(), "\n") {
		if strings.Contains(l, " 127.0.0.1:") && strings.Contains(l, "B \"") {
			out = append(out, l)
		}
	}
	return out
}

func (r *rig) waitConnLines(n int) []string { return r.waitConnLinesFor(n, 3*time.Second) }

func (r *rig) waitConnLinesFor(n int, d time.Duration) []string {
	deadline := time.Now().Add(d)
	for {
		l := r.connLines()
		if len(l) >= n || time.Now().After(deadline) {
			return l
		}
		time.Sleep(10 * time.Millisecond)
	}
}

func oneStderrLine(t *testing.T, s string, wants ...string) {
	t.Helper()
	if strings.Count(s, "\n") != 1 || !strings.HasSuffix(s, "\n") {
		t.Fatalf("stderr %q, want exactly one line", s)
	}
	for _, w := range wants {
		if !strings.Contains(s, w) {
			t.Fatalf("stderr %q does not contain %q", s, w)
		}
	}
}

func nonLoopbackIPv4(t *testing.T) string {
	t.Helper()
	addrs, err := net.InterfaceAddrs()
	if err != nil {
		t.Fatal(err)
	}
	for _, a := range addrs {
		if n, ok := a.(*net.IPNet); ok && !n.IP.IsLoopback() && n.IP.To4() != nil && !n.IP.IsLinkLocalUnicast() {
			return n.IP.String()
		}
	}
	t.Fatal("this host has no non-loopback IPv4 address; TC-CLI-001 needs one")
	return ""
}

// ---- REQ-CFG ----

// TC-CFG-005
func TestREQ_CFG_001_ServeFindsConfigInProgramData(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(filepath.Join(r.pd, "PrinterOne", "config.json"), configJSON(r.port))
	r.start("serve")
	if a := r.waitReady(); a != "0.0.0.0:"+strconv.Itoa(r.port) {
		t.Fatalf("Ready(%q), want 0.0.0.0:%d", a, r.port)
	}
	r.dial().Close()
	r.stop()
}

// TC-CFG-006
func TestREQ_CFG_001_ServeMissingConfigFlagExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(filepath.Join(r.exe, "config.json"), configJSON(r.port))
	missing := filepath.Join(r.dir, "nope.json")
	if code := r.run("serve", "--config", missing); code != 2 {
		t.Fatalf("exit %d, want 2", code)
	}
	oneStderrLine(t, r.stderr.String(), missing)
	r.noReady()
}

// TC-CFG-018
func TestREQ_CFG_004_ServeBadJSONExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, "{bad")
	if code := r.run("serve", "--config", r.cfg); code != 2 {
		t.Fatalf("exit %d, want 2", code)
	}
	oneStderrLine(t, r.stderr.String(), r.cfg)
	if !regexp.MustCompile(`\b1:[0-9]+\b`).MatchString(r.stderr.String()) {
		t.Fatalf("stderr %q has no line:column 1:<col>", r.stderr.String())
	}
	r.noReady()
}

// TC-CFG-035
func TestREQ_CFG_005_ServeInvalidFieldExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, `{"version":2,"printers":[{"printer":"P","port":65536}]}`)
	if code := r.run("serve", "--config", r.cfg); code != 2 {
		t.Fatalf("exit %d, want 2", code)
	}
	oneStderrLine(t, r.stderr.String(), r.cfg, "printers[0].port")
	r.noReady()
}

// TC-CFG-037
func TestREQ_CFG_006_ServeNoConfigExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	if code := r.run("serve"); code != 2 {
		t.Fatalf("exit %d, want 2", code)
	}
	oneStderrLine(t, r.stderr.String(), filepath.Join(r.pd, "PrinterOne"), "printerone printers")
	r.noReady()
}

// ---- REQ-RECV ----

// TC-RECV-021
func TestREQ_RECV_003_ServeDropsJobOverProductionLimit(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.dial()
	chunk := make([]byte, 1<<20)
	left := 268435457
	for left > 0 {
		n := len(chunk)
		if n > left {
			n = left
		}
		if _, err := c.Write(chunk[:n]); err != nil {
			break // the server may close the connection as soon as it has seen one byte too many
		}
		left -= n
	}
	c.CloseWrite()
	lines := r.waitConnLines(1)
	c.Close()
	if len(lines) != 1 || !strings.HasSuffix(lines[0], "dropped: too large") {
		t.Fatalf("connection lines %q, want one `dropped: too large`", lines)
	}
	r.sp.mu.Lock()
	n := r.sp.startDoc
	r.sp.mu.Unlock()
	if n != 0 {
		t.Fatalf("StartDocPrinter called %d times for a dropped job", n)
	}
	r.stop()
}

// holdSlots proves the edition receives exactly `slots` connections at once: with slots-1 held a probe job
// prints at once; with all slots held an extra job waits; once they close it prints.
// TC-RECV-022
func TestREQ_RECV_004_ServeDropsConnectionSilentFor60s(t *testing.T) {
	t.Parallel() // ~62 s of waiting; runs beside the other tests
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("hi")
	defer c.Close()
	time.Sleep(57 * time.Second)
	if l := r.connLines(); len(l) != 0 {
		t.Fatalf("connection ended within 57 s of silence (%q), want the 60 s idle timeout", l)
	}
	lines := r.waitConnLinesFor(1, 6*time.Second)
	if len(lines) != 1 || !strings.HasSuffix(lines[0], "dropped: idle") {
		t.Fatalf("connection lines %q after 63 s, want one `dropped: idle`", lines)
	}
	r.stop()
}

// TC-RECV-027
func TestREQ_RECV_003_ServePrintsJobOfExactlyProductionLimit(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.countAll = true
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.dial()
	chunk := make([]byte, 1<<20)
	for i := 0; i < 256; i++ { // 256 MiB = 268,435,456 bytes
		if _, err := c.Write(chunk); err != nil {
			t.Fatalf("write %d MiB: %v (job dropped below the limit?)", i, err)
		}
	}
	c.Close()
	lines := r.waitConnLinesFor(1, 30*time.Second)
	if len(lines) != 1 || !strings.HasSuffix(lines[0], " 268435456B \""+printerName+"\" printed") {
		t.Fatalf("connection lines %q, want one `268435456B ... printed`", lines)
	}
	r.sp.mu.Lock()
	n := r.sp.count
	r.sp.mu.Unlock()
	if n != 268435456 {
		t.Fatalf("spooler got %d bytes, want 268435456", n)
	}
	r.stop()
}

func holdSlots(t *testing.T, r *rig, edition string, slots int) {
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	var held []*net.TCPConn
	for i := 0; i < slots-1; i++ {
		held = append(held, r.send("h"))
	}
	time.Sleep(100 * time.Millisecond)
	r.send("probe").Close()
	if docs := r.sp.waitDone(1, 2*time.Second); len(docs) != 1 || string(docs[0].data) != "probe" {
		t.Fatalf("%s: with %d of %d slots busy the probe job was not printed (%d docs)", edition, slots-1, slots, len(docs))
	}
	held = append(held, r.send("h"))
	time.Sleep(100 * time.Millisecond)
	r.send("extra").Close()
	time.Sleep(400 * time.Millisecond)
	for _, d := range r.sp.done() {
		if string(d.data) == "extra" {
			t.Fatalf("%s: extra client's job printed while %d slots were busy", edition, slots)
		}
	}
	for _, h := range held {
		h.Close()
	}
	docs := r.sp.waitDone(slots+2, 3*time.Second)
	found := false
	for _, d := range docs {
		found = found || string(d.data) == "extra"
	}
	if !found {
		t.Fatalf("%s: extra client's job not printed after slots freed (%d docs)", edition, len(docs))
	}
	r.stop()
}

// TC-RECV-015
func TestREQ_RECV_007_LegacyEditionReceivesTwoAtOnce(t *testing.T) {
	r := newRig(t, "legacy")
	defer r.cleanup()
	holdSlots(t, r, "legacy", 2)
}

// TC-RECV-016
func TestREQ_RECV_007_MainEditionReceivesEightAtOnce(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	holdSlots(t, r, "main", 8)
}

// TC-RECV-019
func TestREQ_RECV_008_ServeBusyPortExits1(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	occ, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(r.port))
	if err != nil {
		t.Fatal(err)
	}
	defer occ.Close()
	if code := r.run("serve", "--config", r.cfg); code != 1 {
		t.Fatalf("exit %d, want 1", code)
	}
	if !strings.Contains(r.stderr.String(), strconv.Itoa(r.port)) {
		t.Fatalf("stderr %q does not name port %d", r.stderr.String(), r.port)
	}
	r.noReady()
	c, err := net.DialTimeout("tcp", r.addr(), time.Second)
	if err != nil {
		t.Fatalf("occupant no longer accepts: %v", err)
	}
	c.Close()
}

// ---- REQ-SPOOL ----

// TC-SPOOL-005
func TestREQ_SPOOL_001_ServePrintsThroughSpooler(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("hello\n").Close()
	docs := r.sp.waitDone(1, 3*time.Second)
	if len(docs) != 1 {
		t.Fatalf("%d documents, want 1", len(docs))
	}
	d := docs[0]
	if d.printer != printerName || d.name != "PrinterOne job from 127.0.0.1" || d.datatype != "RAW" || string(d.data) != "hello\n" {
		t.Fatalf("document %+v, want printer %q, name `PrinterOne job from 127.0.0.1`, RAW, data hello\\n", d, printerName)
	}
	r.stop()
}

// TC-SPOOL-014
func TestREQ_SPOOL_002_PrinterNotInstalledEachJobFails(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.openErr = &spool.Error{Code: 1801, Msg: "The printer name is invalid."}
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	time.Sleep(100 * time.Millisecond)
	r.sp.mu.Lock()
	early := len(r.sp.opens)
	r.sp.mu.Unlock()
	if early != 0 {
		t.Fatalf("OpenPrinter called %d times before any job: the printer must not be checked at start-up", early)
	}
	r.send("a").Close()
	r.waitConnLines(1)
	r.send("b").Close()
	lines := r.waitConnLines(2)
	if len(lines) != 2 {
		t.Fatalf("connection lines %q, want 2", lines)
	}
	for _, l := range lines {
		if !strings.HasSuffix(l, "failed: 1801 The printer name is invalid.") {
			t.Fatalf("line %q, want `failed: 1801 The printer name is invalid.`", l)
		}
	}
	r.stop()
}

// TC-SPOOL-015..017
func TestREQ_SPOOL_003_PrintersCommand(t *testing.T) {
	t.Run("TC-SPOOL-015", func(t *testing.T) {
		r := newRig(t, "main")
		defer r.cleanup()
		r.sp.enum = []string{"Microsoft Print to PDF", "Canon MF220 Series"}
		code := r.run("printers")
		lines := strings.Split(strings.TrimSuffix(r.stdout.String(), "\n"), "\n")
		sort.Strings(lines) // any order: the spec sets none
		if code != 0 || !strings.HasSuffix(r.stdout.String(), "\n") || strings.Join(lines, "|") != "Canon MF220 Series|Microsoft Print to PDF" {
			t.Fatalf("exit %d stdout %q", code, r.stdout.String())
		}
	})
	t.Run("TC-SPOOL-016", func(t *testing.T) {
		r := newRig(t, "main")
		defer r.cleanup()
		r.sp.enum = []string{}
		if code := r.run("printers"); code != 0 || r.stdout.String() != "" || r.sp.enumCall == 0 {
			t.Fatalf("exit %d stdout %q EnumPrinters calls %d, want 0, empty, the spooler queried", code, r.stdout.String(), r.sp.enumCall)
		}
	})
	t.Run("TC-SPOOL-017", func(t *testing.T) {
		r := newRig(t, "main")
		defer r.cleanup()
		r.sp.enumErr = &spool.Error{Code: 1722, Msg: "The RPC server is unavailable."}
		code := r.run("printers")
		if code != 1 || r.stdout.String() != "" || !strings.Contains(r.stderr.String(), "1722") ||
			!strings.Contains(r.stderr.String(), "The RPC server is unavailable.") {
			t.Fatalf("exit %d stdout %q stderr %q, want 1, empty, the Windows error", code, r.stdout.String(), r.stderr.String())
		}
	})
}

// ---- REQ-CLI ----

// TC-CLI-001
func TestREQ_CLI_001_ServeStartsAndStopsOnCtrlC(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	want := "0.0.0.0:" + strconv.Itoa(r.port)
	if a := r.waitReady(); a != want {
		t.Fatalf("Ready(%q), want %q", a, want)
	}
	out := r.stdout.String()
	if !strings.Contains(out, want) || !strings.Contains(out, printerName) {
		t.Fatalf("start-up stdout %q, want the address %q and printer %q", out, want, printerName)
	}
	// Bound to every IPv4 interface, not only loopback: dial a non-loopback address of this host.
	ip := nonLoopbackIPv4(t)
	c, err := net.DialTimeout("tcp", ip+":"+strconv.Itoa(r.port), 2*time.Second)
	if err != nil {
		t.Fatalf("dial %s:%d (non-loopback): %v — listening on loopback only?", ip, r.port, err)
	}
	c.Close()
	r.stop()
}

// TC-CLI-002
func TestREQ_CLI_001_NoArgsRunsServe(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(filepath.Join(r.exe, "config.json"), configJSON(r.port))
	r.start()
	r.waitReady()
	r.send("x").Close()
	if docs := r.sp.waitDone(1, 3*time.Second); len(docs) != 1 {
		t.Fatalf("%d documents, want 1", len(docs))
	}
	r.stop()
}

// TC-CLI-003
func TestREQ_CLI_001_CtrlCFinishesPrintingJobAndStopsAccepting(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.gate = make(chan struct{})
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("slow").Close()
	select {
	case <-r.sp.writing:
	case <-time.After(3 * time.Second):
		t.Fatal("WritePrinter never started")
	}
	r.sig <- os.Interrupt
	if code, ok := r.waitExit(300 * time.Millisecond); ok {
		t.Fatalf("Run returned %d while a job was still being written", code)
	}
	refused := false
	for i := 0; i < 100 && !refused; i++ {
		c, err := net.DialTimeout("tcp", r.addr(), 200*time.Millisecond)
		if err != nil {
			refused = true
		} else {
			c.Close()
			time.Sleep(10 * time.Millisecond)
		}
	}
	close(r.sp.gate)
	if !refused {
		t.Fatal("still accepting connections after Ctrl+C")
	}
	code, ok := r.waitExit(5 * time.Second)
	if !ok || code != 0 {
		t.Fatalf("returned=%v code=%d, want 0", ok, code)
	}
	if docs := r.sp.done(); len(docs) != 1 || string(docs[0].data) != "slow" {
		t.Fatalf("documents %d, want the job `slow` ended and closed", len(docs))
	}
}

// TC-CLI-004
func TestREQ_CLI_001_SecondCtrlCExits1AtOnce(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.gate = make(chan struct{}) // never released
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("slow").Close()
	select {
	case <-r.sp.writing:
	case <-time.After(3 * time.Second):
		t.Fatal("WritePrinter never started")
	}
	r.sig <- os.Interrupt
	time.Sleep(100 * time.Millisecond)
	r.sig <- os.Interrupt
	code, ok := r.waitExit(time.Second)
	if !ok || code != 1 {
		t.Fatalf("returned=%v code=%d within 1 s of the second Ctrl+C, want 1", ok, code)
	}
}

// TC-CLI-005, TC-CLI-006
func TestREQ_CLI_002_VersionNamesEdition(t *testing.T) {
	old := buildinfo.Version
	defer func() { buildinfo.Version = old }()
	buildinfo.Version = "v2.0.0"
	for _, ed := range []string{"main", "legacy"} {
		r := newRig(t, ed)
		code := r.run("version")
		r.cleanup()
		if want := "PrinterOne v2.0.0 (" + ed + ")\n"; code != 0 || r.stdout.String() != want {
			t.Fatalf("%s: exit %d stdout %q, want 0 %q", ed, code, r.stdout.String(), want)
		}
	}
}

// TC-CLI-007..009, TC-CLI-014, TC-CLI-015, TC-CLI-018
func TestREQ_CLI_003_UsageErrorsExit2(t *testing.T) {
	cases := []struct {
		tc   string
		args []string
	}{
		{"TC-CLI-007", []string{"frobnicate"}},
		{"TC-CLI-008", []string{"serve", "--bogus"}},
		{"TC-CLI-009", []string{"serve", "--config"}},
		{"TC-CLI-014", []string{"version", "--bogus"}},
		{"TC-CLI-015", []string{"printers", "--bogus"}},
		{"TC-CLI-018", []string{"serve", "-h"}}, // no help flag in the spec: unknown flag
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) {
			r := newRig(t, "main")
			defer r.cleanup()
			r.write(filepath.Join(r.exe, "config.json"), configJSON(r.port)) // a serve fall-back would start
			code := r.run(c.args...)
			if code != 2 || !strings.Contains(strings.ToLower(r.stderr.String()), "usage") || r.stdout.String() != "" {
				t.Fatalf("%q: exit %d stdout %q stderr %q, want 2, empty stdout, usage on stderr",
					c.args, code, r.stdout.String(), r.stderr.String())
			}
			r.noReady()
		})
	}
}

// TC-CLI-010..012
func TestREQ_CLI_004_VersionAndPrintersNeedNoConfig(t *testing.T) {
	t.Run("TC-CLI-010", func(t *testing.T) {
		r := newRig(t, "main")
		defer r.cleanup()
		want := buildinfo.String("main") + "\n"
		if code := r.run("version"); code != 0 || r.stderr.String() != "" || r.stdout.String() != want {
			t.Fatalf("exit %d stdout %q stderr %q, want 0, %q, nothing", code, r.stdout.String(), r.stderr.String(), want)
		}
	})
	t.Run("TC-CLI-011", func(t *testing.T) {
		r := newRig(t, "main")
		defer r.cleanup()
		r.sp.enum = []string{"A"}
		if code := r.run("printers"); code != 0 || r.stdout.String() != "A\n" {
			t.Fatalf("exit %d stdout %q stderr %q, want 0 and A", code, r.stdout.String(), r.stderr.String())
		}
	})
	t.Run("TC-CLI-012", func(t *testing.T) {
		r := newRig(t, "main")
		defer r.cleanup()
		r.write(filepath.Join(r.exe, "config.json"), "{bad")
		want := buildinfo.String("main") + "\n"
		if code := r.run("version"); code != 0 || r.stdout.String() != want {
			t.Fatalf("exit %d stdout %q with a broken config next to the exe, want 0 and %q", code, r.stdout.String(), want)
		}
	})
}

// TC-CLI-013
func TestREQ_CLI_005_WaitingJobPrintedOnCtrlC(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.gate = make(chan struct{})
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("aaa").Close()
	select {
	case <-r.sp.writing:
	case <-time.After(3 * time.Second):
		t.Fatal("WritePrinter never started")
	}
	r.send("bbb").Close()
	time.Sleep(300 * time.Millisecond) // bbb fully received, waiting for the printer
	r.sig <- os.Interrupt
	close(r.sp.gate)
	code, ok := r.waitExit(5 * time.Second)
	if !ok || code != 0 {
		t.Fatalf("returned=%v code=%d, want 0", ok, code)
	}
	docs := r.sp.done()
	if len(docs) != 2 || string(docs[0].data)+string(docs[1].data) != "aaabbb" {
		t.Fatalf("%d documents printed, want aaa then bbb", len(docs))
	}
	lines := r.connLines()
	if len(lines) != 2 || !strings.HasSuffix(lines[0], "printed") || !strings.HasSuffix(lines[1], "printed") {
		t.Fatalf("connection lines %q, want two `printed`", lines)
	}
	if r.sp.maxOpen != 1 {
		t.Fatalf("%d documents open at once during shutdown, want 1 (jobs one after another)", r.sp.maxOpen)
	}
}

// ---- REQ-LOG ----

// TC-LOG-008
func TestREQ_LOG_001_ServeLogsPrinterValueNotName(t *testing.T) {
	// TestMain sets time.Local to UTC+7: the production Now must log local time, not UTC
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("x")
	local := c.LocalAddr().String()
	c.Close()
	lines := r.waitConnLines(1)
	re := regexp.MustCompile(`^\S+ ` + regexp.QuoteMeta(local) + ` 1B "` + regexp.QuoteMeta(printerName) + `" printed$`)
	if len(lines) != 1 || !re.MatchString(lines[0]) {
		t.Fatalf("connection lines %q, want one matching %s", lines, re)
	}
	if _, err := time.Parse(time.RFC3339, strings.Fields(lines[0])[0]); err != nil || !strings.HasSuffix(strings.Fields(lines[0])[0], "+07:00") {
		t.Fatalf("time field %q: want RFC 3339 in local time (+07:00 in this test), err %v", strings.Fields(lines[0])[0], err)
	}
	r.stop()
}

// TC-LOG-009
func TestREQ_LOG_002_ReceivingConnectionDroppedOnShutdown(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("partial")
	defer c.Close()
	time.Sleep(300 * time.Millisecond) // accepted and `partial` read
	r.stop()
	lines := r.connLines()
	if len(lines) != 1 || !strings.HasSuffix(lines[0], " 7B \""+printerName+"\" dropped: shutdown") {
		t.Fatalf("connection lines %q, want one `7B ... dropped: shutdown`", lines)
	}
	r.sp.mu.Lock()
	n := r.sp.startDoc
	r.sp.mu.Unlock()
	if n != 0 {
		t.Fatalf("StartDocPrinter called %d times for a dropped connection", n)
	}
	c.SetReadDeadline(time.Now().Add(2 * time.Second))
	if _, err := c.Read(make([]byte, 1)); err == nil {
		t.Fatal("client connection still open after shutdown")
	} else if ne, ok := err.(net.Error); ok && ne.Timeout() {
		t.Fatal("client connection still open after shutdown")
	}
}

// TC-LOG-011
func TestREQ_LOG_002_SilentConnectionDroppedOnShutdown(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.dial() // accepted, 0 bytes so far
	defer c.Close()
	time.Sleep(300 * time.Millisecond)
	r.stop()
	lines := r.connLines()
	if len(lines) != 1 || !strings.HasSuffix(lines[0], " 0B \""+printerName+"\" dropped: shutdown") {
		t.Fatalf("connection lines %q, want one `0B ... dropped: shutdown` (not `empty`)", lines)
	}
}

// TC-LOG-012
func TestREQ_LOG_002_ReceivingDroppedWhilePrintingOnShutdown(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.gate = make(chan struct{})
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("aaa").Close()
	select {
	case <-r.sp.writing:
	case <-time.After(3 * time.Second):
		t.Fatal("WritePrinter never started")
	}
	b := r.send("partial")
	defer b.Close()
	time.Sleep(300 * time.Millisecond)
	r.sig <- os.Interrupt
	time.Sleep(200 * time.Millisecond)
	b.Write([]byte("rest")) // finishes sending only after Ctrl+C; may fail if already dropped
	b.CloseWrite()
	close(r.sp.gate)
	code, ok := r.waitExit(5 * time.Second)
	if !ok || code != 0 {
		t.Fatalf("returned=%v code=%d, want 0", ok, code)
	}
	if docs := r.sp.done(); len(docs) != 1 || string(docs[0].data) != "aaa" {
		t.Fatalf("%d documents, want only aaa (the receiving connection must not print)", len(docs))
	}
	found := false
	for _, l := range r.connLines() {
		found = found || strings.HasSuffix(l, "dropped: shutdown")
	}
	if !found {
		t.Fatalf("connection lines %q, want one `dropped: shutdown`", r.connLines())
	}
}

// TC-CLI-016
func TestREQ_CLI_005_FailedDrainedJobStillExits0(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.sp.gate = make(chan struct{})
	r.sp.openErr, r.sp.openFail = &spool.Error{Code: 5, Msg: "Access is denied."}, 2 // job B's OpenPrinter fails
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("aaa").Close()
	select {
	case <-r.sp.writing:
	case <-time.After(3 * time.Second):
		t.Fatal("WritePrinter never started")
	}
	r.send("bbb").Close()
	time.Sleep(300 * time.Millisecond)
	r.sig <- os.Interrupt
	close(r.sp.gate)
	code, ok := r.waitExit(5 * time.Second)
	if !ok || code != 0 {
		t.Fatalf("returned=%v code=%d, want 0 although a drained job failed", ok, code)
	}
	lines := r.connLines()
	if len(lines) != 2 || !strings.HasSuffix(lines[0], "printed") || !strings.HasSuffix(lines[1], "failed: 5 Access is denied.") {
		t.Fatalf("connection lines %q, want printed then failed", lines)
	}
}

// TC-CLI-017
func TestREQ_CLI_001_CtrlCDoesNotServeClientWaitingForSlot(t *testing.T) {
	r := newRig(t, "legacy") // 2 slots
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	h1, h2 := r.send("h"), r.send("h")
	defer h1.Close()
	defer h2.Close()
	time.Sleep(100 * time.Millisecond)
	w := r.send("waiting") // no slot: waits
	w.CloseWrite()
	defer w.Close()
	time.Sleep(300 * time.Millisecond)
	r.sig <- os.Interrupt
	code, ok := r.waitExit(5 * time.Second)
	if !ok || code != 0 {
		t.Fatalf("returned=%v code=%d within 5 s of Ctrl+C, want 0 (a waiting client must not be served after Ctrl+C)", ok, code)
	}
	if docs := r.sp.done(); len(docs) != 0 {
		t.Fatalf("%d documents printed after Ctrl+C, want none", len(docs))
	}
}

// TC-CFG-040 (ledger audit 2026-11-07): an unreadable config fails `serve` the same way bad JSON does.
func TestREQ_CFG_004_ServeUnreadableExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "config-is-a-dir")
	if err := os.Mkdir(dir, 0755); err != nil {
		t.Fatal(err)
	}
	if code := r.run("serve", "--config", dir); code != 2 {
		t.Fatalf("exit %d, want 2", code)
	}
	oneStderrLine(t, r.stderr.String(), dir)
	r.noReady()
}

// TC-CLI-020 (ledger audit 2026-11-07): Ctrl+C closes a connection that is still receiving and prints nothing of it.
func TestREQ_CLI_001_CtrlCDropsReceivingConnection(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("half a job")
	defer c.Close()
	time.Sleep(300 * time.Millisecond)
	r.stop()
	if docs := r.sp.done(); len(docs) != 0 {
		t.Fatalf("printed %d documents, want none from a connection still receiving at Ctrl+C", len(docs))
	}
	c.SetReadDeadline(time.Now().Add(2 * time.Second))
	if _, err := c.Read(make([]byte, 1)); err == nil {
		t.Fatal("client connection still open after Ctrl+C")
	} else if ne, ok := err.(net.Error); ok && ne.Timeout() {
		t.Fatal("client connection still open after Ctrl+C")
	}
}

// TC-LOG-020 (ledger audit 2026-11-07): the full §2.5 line for `empty`, `dropped: reset` and `dropped: shutdown`.
func TestREQ_LOG_001_FullLineForEmptyResetShutdown(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	line := func(local string, n int, result string) *regexp.Regexp {
		return regexp.MustCompile(`^\S+ ` + regexp.QuoteMeta(local) + ` ` + strconv.Itoa(n) + `B "` +
			regexp.QuoteMeta(printerName) + `" ` + regexp.QuoteMeta(result) + `$`)
	}
	empty := r.dial()
	emptyAddr := empty.LocalAddr().String()
	empty.Close()
	reset := r.send("ab")
	resetAddr := reset.LocalAddr().String()
	time.Sleep(200 * time.Millisecond)
	reset.SetLinger(0)
	reset.Close()
	r.waitConnLines(2)
	open := r.send("xyz")
	openAddr := open.LocalAddr().String()
	defer open.Close()
	time.Sleep(300 * time.Millisecond)
	r.stop()
	want := []*regexp.Regexp{line(emptyAddr, 0, "empty"), line(resetAddr, 2, "dropped: reset"),
		line(openAddr, 3, "dropped: shutdown")}
	lines := r.connLines()
	for _, re := range want {
		found := false
		for _, l := range lines {
			found = found || re.MatchString(l)
		}
		if !found {
			t.Errorf("no connection line matching %s in %q", re, lines)
		}
	}
}
