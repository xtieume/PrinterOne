package receiver

import (
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
	time.Sleep(100 * time.Millisecond)
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
