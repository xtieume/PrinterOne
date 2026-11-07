package receiver

import (
	"bytes"
	"io"
	"net"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/spool"
)

const printerName = "Canon MF220 Series"

// fixedNow is in time.Local as it is when called (TestREQ_LOG_001 sets a non-UTC zone).
func fixedNow() time.Time { return time.Date(2026, 11, 7, 10, 0, 0, 123456789, time.Local) } // ns: RFC3339 has no fraction

type job struct {
	doc  string
	data []byte
}

// rig is fixture R-std: a Server on 127.0.0.1:0 with a recording PrintFunc and a log buffer.
type rig struct {
	t        *testing.T
	srv      *Server
	mu       sync.Mutex
	jobs     []job
	log      bytes.Buffer
	writes   []string // each Log.Write call
	inFlight int
	maxIn    int
	delay    time.Duration
	failNext []error // errors returned by the next Print calls, in order
}

func (r *rig) Write(p []byte) (int, error) {
	r.mu.Lock()
	defer r.mu.Unlock()
	r.writes = append(r.writes, string(p))
	return r.log.Write(p)
}

func (r *rig) print(doc string, data []byte) error {
	r.mu.Lock()
	r.inFlight++
	if r.inFlight > r.maxIn {
		r.maxIn = r.inFlight
	}
	var err error
	if len(r.failNext) > 0 {
		err, r.failNext = r.failNext[0], r.failNext[1:]
	}
	r.mu.Unlock()
	time.Sleep(r.delay)
	r.mu.Lock()
	r.inFlight--
	if err == nil {
		r.jobs = append(r.jobs, job{doc, append([]byte(nil), data...)})
	}
	r.mu.Unlock()
	return err
}

func newRig(t *testing.T, mod func(*rig)) *rig {
	t.Helper()
	r := &rig{t: t}
	r.srv = &Server{Addr: "127.0.0.1:0", Printer: printerName, MaxBytes: 1024, Idle: 300 * time.Millisecond,
		Slots: 8, Log: r, Now: fixedNow}
	r.srv.Print = r.print
	if mod != nil {
		mod(r)
	}
	if err := r.srv.Listen(); err != nil {
		t.Fatalf("Listen: %v", err)
	}
	go r.srv.Serve()
	return r
}

func (r *rig) close() { r.srv.Shutdown() }

func (r *rig) dial() *net.TCPConn {
	r.t.Helper()
	c, err := net.DialTimeout("tcp", r.srv.ListenAddr(), 2*time.Second)
	if err != nil {
		r.t.Fatalf("dial %q: %v", r.srv.ListenAddr(), err)
	}
	return c.(*net.TCPConn)
}

func (r *rig) send(data []byte) *net.TCPConn {
	r.t.Helper()
	c := r.dial()
	if len(data) > 0 {
		if _, err := c.Write(data); err != nil {
			r.t.Fatalf("write: %v", err)
		}
	}
	return c
}

func (r *rig) snapshot() ([]job, []string) {
	r.mu.Lock()
	defer r.mu.Unlock()
	jobs := append([]job(nil), r.jobs...)
	s := strings.TrimRight(r.log.String(), "\n")
	var lines []string
	if s != "" {
		lines = strings.Split(s, "\n")
	}
	return jobs, lines
}

// waitLines waits until the log has n lines (or 3 s) and returns jobs and lines.
func (r *rig) waitLines(n int) ([]job, []string) {
	deadline := time.Now().Add(3 * time.Second)
	for {
		jobs, lines := r.snapshot()
		if len(lines) >= n || time.Now().After(deadline) {
			return jobs, lines
		}
		time.Sleep(10 * time.Millisecond)
	}
}

func (r *rig) waitJobs(n int) []job {
	deadline := time.Now().Add(3 * time.Second)
	for {
		jobs, _ := r.snapshot()
		if len(jobs) >= n || time.Now().After(deadline) {
			return jobs
		}
		time.Sleep(10 * time.Millisecond)
	}
}

// serverClosed reports whether the server closed c (EOF or reset) within d, as opposed to a read timeout.
func serverClosed(c net.Conn, d time.Duration) bool {
	c.SetReadDeadline(time.Now().Add(d))
	buf := make([]byte, 64)
	for {
		_, err := c.Read(buf)
		if err == nil {
			continue
		}
		if ne, ok := err.(net.Error); ok && ne.Timeout() {
			return false
		}
		return true
	}
}

func lastLineEnds(t *testing.T, lines []string, want string) {
	t.Helper()
	if len(lines) != 1 || !strings.HasSuffix(lines[0], want) {
		t.Fatalf("log lines %q, want exactly one ending %q", lines, want)
	}
}

// TC-RECV-001
func TestREQ_RECV_001_BinaryJobPrintedExactly(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	data := []byte{0x00, 0x0D, 0x0A, 0x1B, 0x40, 0xFF, 0x1A, 0x0A, 0x41}
	r.send(data).Close()
	_, _ = r.waitLines(1)
	jobs, _ := r.snapshot()
	if len(jobs) != 1 || !bytes.Equal(jobs[0].data, data) {
		t.Fatalf("jobs %v, want one job % x", jobs, data)
	}
}

// TC-RECV-002
func TestREQ_RECV_001_SeveralWritesAreOneJob(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	c := r.dial()
	for _, p := range []string{"AAA", "BBB", "CCC"} {
		c.Write([]byte(p))
		time.Sleep(50 * time.Millisecond)
	}
	c.Close()
	r.waitLines(1)
	jobs, _ := r.snapshot()
	if len(jobs) != 1 || string(jobs[0].data) != "AAABBBCCC" {
		t.Fatalf("jobs %v, want one job AAABBBCCC", jobs)
	}
}

// TC-RECV-003
func TestREQ_RECV_002_ConcurrentClientsSeparateSerialJobs(t *testing.T) {
	r := newRig(t, func(r *rig) { r.delay = 100 * time.Millisecond })
	defer r.close()
	a, b := r.dial(), r.dial()
	for i := 0; i < 4; i++ {
		a.Write(bytes.Repeat([]byte("A"), 50))
		b.Write(bytes.Repeat([]byte("B"), 50))
		time.Sleep(10 * time.Millisecond)
	}
	a.Close()
	b.Close()
	r.waitLines(2)
	jobs, _ := r.snapshot()
	if len(jobs) != 2 {
		t.Fatalf("got %d jobs, want 2", len(jobs))
	}
	got := map[string]bool{string(jobs[0].data): true, string(jobs[1].data): true}
	if !got[strings.Repeat("A", 200)] || !got[strings.Repeat("B", 200)] {
		t.Fatalf("jobs %q, want A×200 and B×200", []string{string(jobs[0].data), string(jobs[1].data)})
	}
	r.mu.Lock()
	defer r.mu.Unlock()
	if r.maxIn != 1 {
		t.Fatalf("Print ran %d at once, want 1", r.maxIn)
	}
}

// TC-RECV-004, TC-RECV-005
func TestREQ_RECV_003_LimitBoundary(t *testing.T) {
	t.Run("TC-RECV-004", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		r.send(bytes.Repeat([]byte{'x'}, 1024)).Close()
		jobs := r.waitJobs(1)
		if len(jobs) != 1 || len(jobs[0].data) != 1024 {
			t.Fatalf("jobs %d, want one job of 1024 bytes", len(jobs))
		}
	})
	t.Run("TC-RECV-005", func(t *testing.T) {
		r := newRig(t, func(r *rig) { r.srv.Idle = 5 * time.Second }) // the close must come from the limit, not idle
		defer r.close()
		c := r.send(bytes.Repeat([]byte{'x'}, 1025))
		defer c.Close()
		if !serverClosed(c, 2*time.Second) {
			t.Fatal("connection of 1025 bytes not closed by the server")
		}
		jobs, lines := r.waitLines(1)
		if len(jobs) != 0 {
			t.Fatalf("%d jobs printed, want none", len(jobs))
		}
		lastLineEnds(t, lines, "dropped: too large")
	})
}

// TC-RECV-006
func TestREQ_RECV_003_ProductionLimitIs268435456(t *testing.T) {
	if MaxJobBytes != 268435456 {
		t.Fatalf("MaxJobBytes = %d, want 268435456", MaxJobBytes)
	}
}

// TC-RECV-007..009
func TestREQ_RECV_004_IdleConnectionDropped(t *testing.T) {
	t.Run("TC-RECV-007", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		c := r.dial()
		defer c.Close()
		if !serverClosed(c, 3*time.Second) {
			t.Fatal("silent connection not closed within 3 s (Idle 300 ms)")
		}
		jobs, lines := r.waitLines(1)
		if len(jobs) != 0 {
			t.Fatalf("%d jobs printed, want none", len(jobs))
		}
		lastLineEnds(t, lines, "dropped: idle")
	})
	t.Run("TC-RECV-008", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		c := r.send([]byte("hello"))
		defer c.Close()
		if !serverClosed(c, 3*time.Second) {
			t.Fatal("connection silent after its first bytes not closed within 3 s")
		}
		jobs, lines := r.waitLines(1)
		if len(jobs) != 0 {
			t.Fatalf("partial job printed: %q", jobs[0].data)
		}
		lastLineEnds(t, lines, "dropped: idle")
	})
	t.Run("TC-RECV-009", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		c := r.dial()
		for i := 0; i < 8; i++ {
			c.Write([]byte{'0' + byte(i)})
			time.Sleep(100 * time.Millisecond)
		}
		c.Close()
		jobs := r.waitJobs(1)
		if len(jobs) != 1 || string(jobs[0].data) != "01234567" {
			t.Fatalf("jobs %v, want one job 01234567 (gaps < Idle)", jobs)
		}
	})
}

// TC-RECV-023
func TestREQ_RECV_004_JobWaitingForPrinterNotIdleDropped(t *testing.T) {
	r := newRig(t, func(r *rig) { r.delay = 800 * time.Millisecond }) // printer busy far longer than Idle 300 ms
	defer r.close()
	r.send([]byte("first")).Close()
	time.Sleep(50 * time.Millisecond)
	r.send([]byte("second")).Close()
	jobs, lines := r.waitLines(2)
	if len(jobs) != 2 {
		t.Fatalf("jobs %d, log %q: a fully received job waiting for the printer must still print", len(jobs), lines)
	}
}

// TC-RECV-010
func TestREQ_RECV_004_ProductionIdleIs60s(t *testing.T) {
	if IdleTimeout != 60*time.Second {
		t.Fatalf("IdleTimeout = %v, want 60s", IdleTimeout)
	}
}

func reset(t *testing.T, c *net.TCPConn) {
	t.Helper()
	if err := c.SetLinger(0); err != nil {
		t.Fatal(err)
	}
	c.Close()
}

// TC-RECV-011, TC-RECV-012
func TestREQ_RECV_005_ResetPrintsNothing(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	c := r.send([]byte("partial"))
	time.Sleep(100 * time.Millisecond)
	reset(t, c)
	jobs, lines := r.waitLines(1)
	if len(jobs) != 0 {
		t.Fatalf("reset job printed: %q", jobs[0].data)
	}
	lastLineEnds(t, lines, " 7B \""+printerName+"\" dropped: reset")

	r.send([]byte("next")).Close() // TC-RECV-012
	jobs = r.waitJobs(1)
	if len(jobs) != 1 || string(jobs[0].data) != "next" {
		t.Fatalf("jobs after reset %v, want one job next", jobs)
	}
}

// TC-RECV-013
func TestREQ_RECV_006_EmptyConnectionPrintsNothing(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	r.dial().Close()
	jobs, lines := r.waitLines(1)
	if len(jobs) != 0 {
		t.Fatalf("%d jobs printed for an empty connection", len(jobs))
	}
	lastLineEnds(t, lines, " 0B \""+printerName+"\" empty")
}

// TC-RECV-014
func TestREQ_RECV_007_ClientOverSlotLimitWaits(t *testing.T) {
	r := newRig(t, func(r *rig) { r.srv.Slots = 2; r.srv.Idle = 5 * time.Second })
	defer r.close()
	h1 := r.send([]byte("h"))
	time.Sleep(100 * time.Millisecond)
	r.send([]byte("second")).Close() // one slot still free: received and printed at once
	if jobs := r.waitJobs(1); len(jobs) != 1 || string(jobs[0].data) != "second" {
		t.Fatalf("jobs %v with one slot free, want second printed", jobs)
	}
	h2 := r.send([]byte("h"))
	defer h2.Close()
	time.Sleep(100 * time.Millisecond)
	r.send([]byte("third")).Close()
	time.Sleep(400 * time.Millisecond)
	if jobs, _ := r.snapshot(); len(jobs) != 1 {
		t.Fatalf("job %q printed while both slots were busy", jobs[len(jobs)-1].data)
	}
	h1.Close()
	time.Sleep(time.Second) // well under Idle (5 s): third must not depend on an idle drop
	jobs, _ := r.snapshot()
	got := map[string]bool{}
	for _, j := range jobs {
		got[string(j.data)] = true
	}
	if !got["third"] || !got["h"] {
		t.Fatalf("jobs %v within 1 s of a slot freeing, want h and third", jobs)
	}
}

// TC-RECV-024
func TestREQ_RECV_007_BacklogWaitLongerThanIdleNotDropped(t *testing.T) {
	r := newRig(t, func(r *rig) { r.srv.Slots = 1 })
	defer r.close()
	h := r.dial()
	time.Sleep(50 * time.Millisecond)
	r.send([]byte("w")).Close() // waits for the only slot
	for i := 0; i < 8; i++ {    // holder stays active 800 ms, gaps 100 ms < Idle 300 ms
		h.Write([]byte{'h'})
		time.Sleep(100 * time.Millisecond)
	}
	h.Close()
	jobs := r.waitJobs(2)
	got := map[string]bool{}
	for _, j := range jobs {
		got[string(j.data)] = true
	}
	if !got["w"] {
		_, lines := r.snapshot()
		t.Fatalf("waiting client's job not printed after waiting 800 ms > Idle; log %q", lines)
	}
}

// TC-RECV-026
func TestREQ_RECV_007_DroppedConnectionsFreeTheirSlot(t *testing.T) {
	r := newRig(t, func(r *rig) {
		r.srv.Slots = 1
		r.failNext = []error{&spool.Error{Code: 5, Msg: "Access is denied."}}
	})
	defer r.close()
	r.send([]byte("fails")).Close() // failed print
	r.waitLines(1)
	c := r.send([]byte("partial")) // reset
	time.Sleep(100 * time.Millisecond)
	reset(t, c)
	r.waitLines(2)
	r.send([]byte("silent")) // idle drop after 300 ms
	r.waitLines(3)
	r.send(bytes.Repeat([]byte{'x'}, 1025)) // too large
	r.waitLines(4)
	r.dial().Close() // empty
	r.waitLines(5)
	r.send([]byte("ok")).Close()
	jobs, lines := r.waitLines(6)
	if len(jobs) != 1 || string(jobs[0].data) != "ok" {
		t.Fatalf("jobs %v, log %q: with 1 slot, `ok` must print after five dropped/failed connections", jobs, lines)
	}
}

// TC-RECV-017
func TestREQ_RECV_007_ProductionSlots(t *testing.T) {
	if SlotsMain != 8 || SlotsLegacy != 2 {
		t.Fatalf("SlotsMain=%d SlotsLegacy=%d, want 8 and 2", SlotsMain, SlotsLegacy)
	}
}

// TC-RECV-018
func TestREQ_RECV_008_BusyPortReturnsPortBusyError(t *testing.T) {
	occ, err := net.Listen("tcp", "0.0.0.0:0")
	if err != nil {
		t.Fatal(err)
	}
	defer occ.Close()
	port := occ.Addr().(*net.TCPAddr).Port
	s := &Server{Addr: "0.0.0.0:" + strconv.Itoa(port), Printer: printerName, Print: func(string, []byte) error { return nil },
		MaxBytes: 1024, Idle: time.Second, Slots: 1, Log: io.MultiWriter()}
	err = s.Listen()
	pb, ok := err.(*PortBusyError)
	if !ok || pb.Port != port {
		t.Fatalf("Listen error %T %v, want *PortBusyError{Port: %d}", err, err, port)
	}
	if !strings.Contains(pb.Error(), strconv.Itoa(port)) {
		t.Fatalf("Error() %q does not name port %d", pb.Error(), port)
	}
	c, err := net.DialTimeout("tcp", "127.0.0.1:"+strconv.Itoa(port), time.Second)
	if err != nil {
		t.Fatalf("occupant no longer accepts: %v", err)
	}
	c.Close()
}

// TC-RECV-020
func TestREQ_RECV_009_HalfCloseEndsJob(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	c := r.send([]byte("half"))
	defer c.Close()
	if err := c.CloseWrite(); err != nil {
		t.Fatal(err)
	}
	deadline := time.Now().Add(2 * time.Second)
	var jobs []job
	for time.Now().Before(deadline) { // shorter than any idle drop would matter: job must print, not drop
		if jobs, _ = r.snapshot(); len(jobs) > 0 {
			break
		}
		time.Sleep(10 * time.Millisecond)
	}
	if len(jobs) != 1 || string(jobs[0].data) != "half" {
		t.Fatalf("jobs %v after half-close, want one job half", jobs)
	}
}

// TC-RECV-025
func TestREQ_RECV_006_HalfCloseWithNoBytesIsEmpty(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	c := r.dial()
	defer c.Close()
	c.CloseWrite()
	jobs, lines := r.waitLines(1)
	if len(jobs) != 0 {
		t.Fatalf("%d jobs printed for a 0-byte half-close", len(jobs))
	}
	lastLineEnds(t, lines, " 0B \""+printerName+"\" empty")
}

// TC-SPOOL-004
func TestREQ_SPOOL_001_DocNameHasClientIPWithoutPort(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	r.send([]byte("x")).Close()
	jobs := r.waitJobs(1)
	if len(jobs) != 1 || jobs[0].doc != "PrinterOne job from 127.0.0.1" {
		t.Fatalf("jobs %v, want doc name %q", jobs, "PrinterOne job from 127.0.0.1")
	}
}

// TC-SPOOL-013
func TestREQ_SPOOL_002_FailedJobDoesNotStopNext(t *testing.T) {
	r := newRig(t, func(r *rig) { r.failNext = []error{&spool.Error{Code: 5, Msg: "Access is denied."}} })
	defer r.close()
	r.send([]byte("one")).Close()
	r.waitLines(1)
	r.send([]byte("two")).Close()
	jobs, lines := r.waitLines(2)
	if len(lines) != 2 || !strings.HasSuffix(lines[0], "failed: 5 Access is denied.") || !strings.HasSuffix(lines[1], "printed") {
		t.Fatalf("log %q, want failed then printed", lines)
	}
	if len(jobs) != 1 || string(jobs[0].data) != "two" {
		t.Fatalf("jobs %v, want job two printed", jobs)
	}
}

// line returns the regexp for one §2.5 line from client c.
func line(c net.Conn, bytesPart, result string) *regexp.Regexp {
	return regexp.MustCompile("^" + regexp.QuoteMeta(fixedNow().Format(time.RFC3339)+" "+c.LocalAddr().String()+" ") +
		bytesPart + regexp.QuoteMeta(` "`+printerName+`" `+result) + "$")
}

func oneLine(t *testing.T, lines []string, re *regexp.Regexp) {
	t.Helper()
	if len(lines) != 1 || !re.MatchString(lines[0]) {
		t.Fatalf("log %q, want one line matching %s", lines, re)
	}
}

// TC-LOG-001, 003, 004, 006, 007
func TestREQ_LOG_001_LineFormatPerResult(t *testing.T) {
	oldLocal := time.Local
	time.Local = time.FixedZone("UTC+7", 7*3600) // CI runs in UTC; a UTC-formatting build must not pass
	defer func() { time.Local = oldLocal }()
	t.Run("TC-LOG-001", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		c := r.send([]byte("hello"))
		c.Close()
		_, lines := r.waitLines(1)
		oneLine(t, lines, line(c, "5B", "printed"))
	})
	t.Run("TC-LOG-003", func(t *testing.T) {
		r := newRig(t, func(r *rig) { r.srv.Idle = 5 * time.Second })
		defer r.close()
		c := r.send(bytes.Repeat([]byte{'x'}, 1025))
		defer c.Close()
		_, lines := r.waitLines(1)
		oneLine(t, lines, line(c, "([0-9]+)B", "dropped: too large"))
		n, _ := strconv.Atoi(line(c, "([0-9]+)B", "dropped: too large").FindStringSubmatch(lines[0])[1])
		if n <= 1024 {
			t.Fatalf("too-large line reports %d bytes, want > 1024", n)
		}
	})
	t.Run("TC-LOG-004", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		c := r.send([]byte("hi"))
		defer c.Close()
		_, lines := r.waitLines(1)
		oneLine(t, lines, line(c, "2B", "dropped: idle"))
	})
	t.Run("TC-LOG-006", func(t *testing.T) {
		r := newRig(t, func(r *rig) { r.failNext = []error{&spool.Error{Code: 5, Msg: "Access is denied."}} })
		defer r.close()
		c := r.send([]byte("x"))
		c.Close()
		_, lines := r.waitLines(1)
		oneLine(t, lines, line(c, "1B", "failed: 5 Access is denied."))
	})
	t.Run("TC-LOG-007", func(t *testing.T) {
		r := newRig(t, nil)
		defer r.close()
		r.send([]byte("a")).Close()
		r.dial().Close()
		r.send([]byte("b")).Close()
		r.waitLines(3)
		time.Sleep(200 * time.Millisecond) // a fourth line would show up here
		_, lines := r.snapshot()
		if len(lines) != 3 {
			t.Fatalf("log %q, want exactly 3 lines for 3 connections", lines)
		}
	})
}

// TC-LOG-010
func TestREQ_LOG_001_ConcurrentConnectionsWriteWholeLines(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	var wg sync.WaitGroup
	for i := 0; i < 20; i++ {
		wg.Add(1)
		go func(i int) {
			defer wg.Done()
			c, err := net.DialTimeout("tcp", r.srv.ListenAddr(), 2*time.Second)
			if err != nil {
				return
			}
			if i%2 == 0 { // half the lines come from the receiving goroutines (empty), half from printing
				c.Write([]byte(strconv.Itoa(i)))
			}
			c.Close()
		}(i)
	}
	wg.Wait()
	_, lines := r.waitLines(20)
	if len(lines) != 20 {
		t.Fatalf("%d lines for 20 connections", len(lines))
	}
	re := regexp.MustCompile(`^\S+ 127\.0\.0\.1:[0-9]+ [0-9]+B "` + regexp.QuoteMeta(printerName) + `" (printed|empty)$`)
	for _, l := range lines {
		if !re.MatchString(l) {
			t.Fatalf("line %q is not one whole §2.5 line (interleaved writes?)", l)
		}
	}
}

// TC-LOG-013
func TestREQ_LOG_001_PrinterNameLoggedVerbatim(t *testing.T) {
	r := newRig(t, func(r *rig) { r.srv.Printer = `\\srv\Canon "A4"` })
	defer r.close()
	r.send([]byte("x")).Close()
	_, lines := r.waitLines(1)
	if len(lines) != 1 || !strings.HasSuffix(lines[0], ` 1B "\\srv\Canon "A4"" printed`) {
		t.Fatalf("log %q, want the printer value verbatim between the quotes (no Go escaping)", lines)
	}
}
