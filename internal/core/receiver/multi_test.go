package receiver

import (
	"net"
	"sync"
	"testing"
	"time"
)

// F-003 receiver tests (spec docs/specs/2026-11-07-f003-multi-printer.md §2). TC ids: .testcases/runs/F-003-r1 and
// testcases.md.

// TC-MULTI-027
func TestREQ_MULTI_008_SlotPoolSharedByServers(t *testing.T) {
	pool := NewSlotPool(1)
	shared := func(r *rig) { r.srv.Pool = pool; r.srv.Slots = 8; r.srv.Idle = 5 * time.Second }
	one := newRig(t, shared)
	defer one.close()
	two := newRig(t, shared)
	defer two.close()

	a := one.send([]byte("a")) // holds the only slot while open
	// With a shared pool a port takes the slot only for AcceptWait before handing it back, so `a` can sit in the
	// backlog that long while the other port holds the slot; wait until it has surely been accepted.
	time.Sleep(2*AcceptWait + 100*time.Millisecond)
	b := two.send([]byte("b"))
	b.Close()
	time.Sleep(400 * time.Millisecond)
	if jobs, lines := two.snapshot(); len(jobs) != 0 || len(lines) != 0 {
		a.Close()
		t.Fatalf("server 2 received a job while server 1 held the only shared slot: jobs=%v logs=%q", jobs, lines)
	}
	a.Close()
	jobsA, _ := one.waitLines(1)
	jobsB, _ := two.waitLines(1)
	if len(jobsA) != 1 || string(jobsA[0].data) != "a" || len(jobsB) != 1 || string(jobsB[0].data) != "b" {
		t.Fatalf("after the slot frees: server 1 jobs %v, server 2 jobs %v; want `a` and `b`", jobsA, jobsB)
	}
}

// deadlineCountListener counts accepted connections and keeps the *net.TCPListener methods (SetDeadline).
type deadlineCountListener struct {
	*net.TCPListener
	mu       sync.Mutex
	accepted int
}

func (l *deadlineCountListener) Accept() (net.Conn, error) {
	c, err := l.TCPListener.Accept()
	if err == nil {
		l.mu.Lock()
		l.accepted++
		l.mu.Unlock()
	}
	return c, err
}

func (l *deadlineCountListener) count() int { l.mu.Lock(); defer l.mu.Unlock(); return l.accepted }

func newDeadlineCountListener(t *testing.T) *deadlineCountListener {
	t.Helper()
	base, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	return &deadlineCountListener{TCPListener: base.(*net.TCPListener)}
}

// TC-RECV-030
func TestREQ_RECV_007_PoolServerWaitsInBacklogUntilSlotFrees(t *testing.T) {
	listener := newDeadlineCountListener(t)
	r := newRig(t, func(r *rig) { r.srv.Pool = NewSlotPool(1); r.srv.Idle = 5 * time.Second; r.srv.Listener = listener })
	defer r.close()
	a := r.dial()
	time.Sleep(100 * time.Millisecond)
	r.send([]byte("B")).Close()
	time.Sleep(400 * time.Millisecond)
	if got := listener.count(); got != 1 {
		a.Close()
		t.Fatalf("accepted %d connections while the shared pool has one slot, want 1", got)
	}
	a.Close()
	if jobs := r.waitJobs(1); len(jobs) != 1 || string(jobs[0].data) != "B" {
		t.Fatalf("after the slot frees: jobs %v, want B", jobs)
	}
}

// TC-MULTI-048
func TestREQ_MULTI_008_IdlePortsDoNotStarveBusyPort(t *testing.T) {
	pool := NewSlotPool(SlotsLegacy)
	shared := func(r *rig) { r.srv.Pool = pool; r.srv.Idle = 5 * time.Second }
	idle1 := newRig(t, shared)
	defer idle1.close()
	idle2 := newRig(t, shared)
	defer idle2.close()
	busy := newRig(t, shared)
	defer busy.close()
	time.Sleep(300 * time.Millisecond) // both idle ports are waiting for clients
	for i := 0; i < 3; i++ {
		busy.send([]byte{'a' + byte(i)}).Close()
		if jobs := busy.waitJobsFor(i+1, 2*time.Second); len(jobs) != i+1 {
			t.Fatalf("job %d on the third port not printed within 2 s while two idle ports share the 2 legacy slots", i+1)
		}
	}
}
