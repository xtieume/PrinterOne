package receiver

import (
	"net"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/allow"
)

// F-004 receiver tests (spec docs/specs/2026-11-07-f004-allowlist.md §2): Server.Allow decides per connection.
// TC ids: .testcases/runs/F-004-r1 and testcases.md.

const rejectedLine = `0B "` + printerName + `" rejected: not allowed`

// readCountListener counts Read calls made by the server on the connections it accepts.
type readCountListener struct {
	net.Listener
	mu    sync.Mutex
	reads int
}

type readCountConn struct {
	net.Conn
	l *readCountListener
}

func (c *readCountConn) Read(p []byte) (int, error) {
	c.l.mu.Lock()
	c.l.reads++
	c.l.mu.Unlock()
	return c.Conn.Read(p)
}

func (l *readCountListener) Accept() (net.Conn, error) {
	c, err := l.Listener.Accept()
	if err != nil {
		return nil, err
	}
	return &readCountConn{Conn: c, l: l}, nil
}

func (l *readCountListener) count() int {
	l.mu.Lock()
	defer l.mu.Unlock()
	return l.reads
}

// listenOn returns a listener on network/addr, or skips when the host has no such loopback.
func listenOn(t *testing.T, network, addr string) net.Listener {
	t.Helper()
	l, err := net.Listen(network, addr)
	if err != nil {
		t.Skipf("host cannot listen on %s %s: %v", network, addr, err)
	}
	return l
}

func parseList(t *testing.T, items ...string) allow.List {
	t.Helper()
	l, bad, err := allow.Parse(items)
	if err != nil {
		t.Fatalf("allow.Parse(%q): item %d: %v", items, bad, err)
	}
	return l
}

// ---- REQ-ALLOW-001 ----

// TC-ALLOW-010
func TestREQ_ALLOW_001_NilAllowPrints(t *testing.T) {
	r := newRig(t, nil)
	defer r.close()
	r.send([]byte("job")).Close()
	jobs, lines := r.waitLines(1)
	if len(jobs) != 1 || string(jobs[0].data) != "job" {
		t.Fatalf("jobs %v, want `job` printed (Allow nil = every client)", jobs)
	}
	lastLineEnds(t, lines, `3B "`+printerName+`" printed`)
	if strings.Contains(strings.Join(lines, "\n"), "rejected") {
		t.Fatalf("log %q mentions a rejection", lines)
	}
	// the nil allow.List (no `allow` key) must allow too when wired as Allow
	var none allow.List
	if !none.Allows(net.ParseIP("127.0.0.1")) {
		t.Fatalf("nil allow.List refuses 127.0.0.1")
	}
}

// ---- REQ-ALLOW-002 ----

// TC-ALLOW-011
func TestREQ_ALLOW_002_AllowedClientPrintedAndAskedWithItsIP(t *testing.T) {
	var mu sync.Mutex
	var asked []string
	r := newRig(t, func(r *rig) {
		r.srv.Allow = func(ip net.IP) bool {
			mu.Lock()
			asked = append(asked, ip.String())
			mu.Unlock()
			return true
		}
	})
	defer r.close()
	r.send([]byte("job")).Close()
	jobs, lines := r.waitLines(1)
	mu.Lock()
	defer mu.Unlock()
	if len(asked) != 1 || asked[0] != "127.0.0.1" {
		t.Fatalf("Allow called with %q, want exactly once with 127.0.0.1", asked)
	}
	if len(jobs) != 1 || string(jobs[0].data) != "job" {
		t.Fatalf("jobs %v, want `job` printed", jobs)
	}
	lastLineEnds(t, lines, `3B "`+printerName+`" printed`)
}

// ---- REQ-ALLOW-003 ----

// TC-ALLOW-012
func TestREQ_ALLOW_003_RejectedWithoutReadOrPrint(t *testing.T) {
	rl := &readCountListener{Listener: listenOn(t, "tcp", "127.0.0.1:0")}
	r := newRig(t, func(r *rig) {
		r.srv.Listener = rl
		r.srv.Idle = 5 * time.Second
		r.srv.Allow = func(net.IP) bool { return false }
	})
	defer r.close()
	c := r.send([]byte("secret job bytes"))
	defer c.Close()
	start := time.Now()
	if !serverClosed(c, time.Second) {
		t.Fatalf("server did not close the rejected connection within 1 s")
	}
	if d := time.Since(start); d > time.Second {
		t.Fatalf("rejected connection closed after %v, want at once", d)
	}
	r.waitLines(1)
	time.Sleep(100 * time.Millisecond) // a late Print or second line would show up here
	jobs, lines := r.snapshot()
	if len(jobs) != 0 {
		t.Fatalf("jobs %v, want nothing printed", jobs)
	}
	lastLineEnds(t, lines, rejectedLine)
	if !strings.Contains(lines[0], " 127.0.0.1:") {
		t.Fatalf("log line %q does not name the client ip:port", lines[0])
	}
	if n := rl.count(); n != 0 {
		t.Fatalf("server called Read %d time(s) on the rejected connection, want 0", n)
	}
}

// TC-ALLOW-013
func TestREQ_ALLOW_003_RejectedClientHoldsNoSlot(t *testing.T) {
	var mu sync.Mutex
	calls := 0
	r := newRig(t, func(r *rig) {
		r.srv.Slots = 1
		r.srv.Idle = 5 * time.Second
		r.srv.Allow = func(net.IP) bool { // first client refused, later ones allowed
			mu.Lock()
			defer mu.Unlock()
			calls++
			return calls > 1
		}
	})
	defer r.close()
	refused := r.send([]byte("x")) // stays open: the server must not keep its slot
	defer refused.Close()
	time.Sleep(100 * time.Millisecond)
	r.send([]byte("allowed")).Close()
	jobs := r.waitJobsFor(1, time.Second)
	if len(jobs) != 1 || string(jobs[0].data) != "allowed" {
		t.Fatalf("jobs within 1 s %v, want `allowed` printed while the refused client is still open (1 slot)", jobs)
	}
}

func (r *rig) waitJobsFor(n int, d time.Duration) []job {
	deadline := time.Now().Add(d)
	for {
		jobs, _ := r.snapshot()
		if len(jobs) >= n || time.Now().After(deadline) {
			return jobs
		}
		time.Sleep(10 * time.Millisecond)
	}
}

// TC-ALLOW-014
func TestREQ_ALLOW_003_ManyRejectedThenAllowed(t *testing.T) {
	var mu sync.Mutex
	open := true
	r := newRig(t, func(r *rig) {
		r.srv.Slots = 2
		r.srv.Idle = 5 * time.Second
		r.srv.Allow = func(net.IP) bool { mu.Lock(); defer mu.Unlock(); return !open }
	})
	defer r.close()
	var held []net.Conn
	for i := 0; i < 5; i++ {
		held = append(held, r.send([]byte("x")))
	}
	defer func() {
		for _, c := range held {
			c.Close()
		}
	}()
	_, lines := r.waitLines(5)
	mu.Lock()
	open = false
	mu.Unlock()
	if len(lines) != 5 {
		t.Fatalf("log %q, want 5 rejection lines", lines)
	}
	for _, l := range lines {
		if !strings.HasSuffix(l, rejectedLine) {
			t.Fatalf("log line %q, want suffix %q", l, rejectedLine)
		}
	}
	r.send([]byte("ok")).Close()
	if jobs := r.waitJobsFor(1, time.Second); len(jobs) != 1 || string(jobs[0].data) != "ok" {
		t.Fatalf("jobs %v, want `ok` printed after 5 rejected clients that stay open (2 slots)", jobs)
	}
}

// ---- REQ-ALLOW-004 ----

// TC-ALLOW-015
func TestREQ_ALLOW_004_IPv6ClientMatchedByIPv6List(t *testing.T) {
	probe := listenOn(t, "tcp6", "[::1]:0")
	probe.Close()
	for _, c := range []struct {
		items  []string
		printd bool
	}{
		{[]string{"::1"}, true},
		{[]string{"::/0"}, true},
		{[]string{"127.0.0.1", "0.0.0.0/0"}, false},
		{[]string{"fd00::/8"}, false},
	} {
		list := parseList(t, c.items...)
		r := newRig(t, func(r *rig) {
			r.srv.Listener = listenOn(t, "tcp6", "[::1]:0")
			r.srv.Allow = list.Allows
		})
		conn, err := net.DialTimeout("tcp6", r.srv.ListenAddr(), 2*time.Second)
		if err != nil {
			r.close()
			t.Fatalf("dial %s: %v", r.srv.ListenAddr(), err)
		}
		conn.Write([]byte("v6"))
		conn.Close()
		jobs, lines := r.waitLines(1)
		r.close()
		if c.printd && (len(jobs) != 1 || len(lines) != 1 || !strings.HasSuffix(lines[0], `2B "`+printerName+`" printed`)) {
			t.Errorf("list %q, client ::1: jobs %v log %q; want printed", c.items, jobs, lines)
		}
		if !c.printd && (len(jobs) != 0 || len(lines) != 1 || !strings.HasSuffix(lines[0], rejectedLine)) {
			t.Errorf("list %q, client ::1: jobs %v log %q; want rejected", c.items, jobs, lines)
		}
	}
}

// TC-ALLOW-016
func TestREQ_ALLOW_004_MappedIPv4OnDualStackSocket(t *testing.T) {
	probe := listenOn(t, "tcp6", "[::1]:0")
	probe.Close()
	for _, c := range []struct {
		items  []string
		printd bool
	}{
		{[]string{"127.0.0.1"}, true},
		{[]string{"127.0.0.0/8"}, true},
		{[]string{"10.0.0.0/8"}, false},
		{[]string{"::1"}, false},
	} {
		list := parseList(t, c.items...)
		var mu sync.Mutex
		var seen net.IP
		r := newRig(t, func(r *rig) {
			r.srv.Listener = listenOn(t, "tcp", "[::]:0") // dual-stack: IPv4 clients arrive as ::ffff:a.b.c.d
			r.srv.Allow = func(ip net.IP) bool {
				mu.Lock()
				seen = ip
				mu.Unlock()
				return list.Allows(ip)
			}
		})
		port := r.srv.Listener.Addr().(*net.TCPAddr).Port
		conn, err := net.DialTimeout("tcp4", "127.0.0.1:"+strconv.Itoa(port), 2*time.Second)
		if err != nil {
			r.close()
			t.Skipf("dual-stack socket does not take IPv4 here: %v", err)
		}
		conn.Write([]byte("v4"))
		conn.Close()
		jobs, lines := r.waitLines(1)
		r.close()
		mu.Lock()
		got := seen
		mu.Unlock()
		if got == nil || got.To4() == nil || !got.To4().Equal(net.IPv4(127, 0, 0, 1).To4()) {
			t.Errorf("Allow saw %v, want 127.0.0.1 (in either form)", got)
		}
		if c.printd && (len(jobs) != 1 || len(lines) != 1 || !strings.HasSuffix(lines[0], `2B "`+printerName+`" printed`)) {
			t.Errorf("list %q, client 127.0.0.1 via [::]: jobs %v log %q; want printed", c.items, jobs, lines)
		}
		if !c.printd && (len(jobs) != 0 || len(lines) != 1 || !strings.HasSuffix(lines[0], rejectedLine)) {
			t.Errorf("list %q, client 127.0.0.1 via [::]: jobs %v log %q; want rejected", c.items, jobs, lines)
		}
	}
}

// fakeRemoteListener makes every accepted conn report remote as its RemoteAddr (LocalAddr stays loopback).
type fakeRemoteListener struct {
	net.Listener
	remote net.Addr
}

type fakeRemoteConn struct {
	net.Conn
	remote net.Addr
}

func (c *fakeRemoteConn) RemoteAddr() net.Addr { return c.remote }

func (l *fakeRemoteListener) Accept() (net.Conn, error) {
	c, err := l.Listener.Accept()
	if err != nil {
		return nil, err
	}
	return &fakeRemoteConn{Conn: c, remote: l.remote}, nil
}

// TC-ALLOW-040
func TestREQ_ALLOW_003_ChecksRemoteNotLocalAddress(t *testing.T) {
	remote := &net.TCPAddr{IP: net.ParseIP("10.0.0.5"), Port: 40000}
	for _, c := range []struct {
		items  []string
		printd bool
	}{
		{[]string{"10.0.0.5"}, true},
		{[]string{"127.0.0.1"}, false},
	} {
		list := parseList(t, c.items...)
		r := newRig(t, func(r *rig) {
			r.srv.Listener = &fakeRemoteListener{Listener: listenOn(t, "tcp", "127.0.0.1:0"), remote: remote}
			r.srv.Allow = list.Allows
		})
		r.send([]byte("job")).Close()
		jobs, lines := r.waitLines(1)
		r.close()
		if c.printd && (len(jobs) != 1 || len(lines) != 1 || !strings.HasSuffix(lines[0], `3B "`+printerName+`" printed`)) {
			t.Errorf("list %q, remote 10.0.0.5 on local 127.0.0.1: jobs %v log %q; want printed", c.items, jobs, lines)
		}
		if !c.printd && (len(jobs) != 0 || len(lines) != 1 || !strings.HasSuffix(lines[0], rejectedLine) ||
			!strings.Contains(lines[0], " 10.0.0.5:40000 ")) {
			t.Errorf("list %q, remote 10.0.0.5 on local 127.0.0.1: jobs %v log %q; want rejected, naming 10.0.0.5:40000", c.items, jobs, lines)
		}
	}
}

// TC-ALLOW-041
func TestREQ_ALLOW_003_RejectedAtOnceWhileSlotsBusy(t *testing.T) {
	var mu sync.Mutex
	calls := 0
	r := newRig(t, func(r *rig) {
		r.srv.Slots = 1
		r.srv.Idle = 5 * time.Second
		r.srv.Allow = func(net.IP) bool { // first client allowed (holds the only slot), later ones refused
			mu.Lock()
			defer mu.Unlock()
			calls++
			return calls == 1
		}
	})
	defer r.close()
	busy := r.send([]byte("x")) // allowed, stays open: the only slot is taken
	defer busy.Close()
	time.Sleep(100 * time.Millisecond)
	c := r.send([]byte("y"))
	defer c.Close()
	if !serverClosed(c, time.Second) {
		t.Fatalf("refused client not closed within 1 s while the only slot was busy")
	}
	_, lines := r.waitLines(1)
	if len(lines) != 1 || !strings.HasSuffix(lines[0], rejectedLine) {
		t.Fatalf("log %q, want one rejection line while the allowed client is still open", lines)
	}
}
