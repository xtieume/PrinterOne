package cli

import (
	"strconv"
	"strings"
	"testing"
	"time"
)

// F-004 tests through `printerone serve` (spec docs/specs/2026-11-07-f004-allowlist.md §2). TC ids:
// .testcases/runs/F-004-r1 and testcases.md. Two-printer cases also need F-003 (several printers).

const notAllowed = `rejected: not allowed`

// allowEntry is one printer entry; allowJSON is "" (no key) or `,"allow":[...]`.
func allowEntry(name, printer string, port int, allowJSON string) string {
	return `{"name":"` + name + `","printer":"` + printer + `","port":` + strconv.Itoa(port) + allowJSON + `}`
}

func allowConfig(entries ...string) string {
	return `{"version":2,"printers":[` + strings.Join(entries, ",") + `]}`
}

// startupLines returns the stdout lines written before Ready that name 0.0.0.0:<port>, joined by "\n".
func startupLines(t *testing.T, out string, port int) string {
	t.Helper()
	var got []string
	for _, l := range strings.Split(out, "\n") {
		if strings.Contains(l, "0.0.0.0:"+strconv.Itoa(port)) {
			got = append(got, l)
		}
	}
	if len(got) == 0 {
		t.Fatalf("no start-up line names 0.0.0.0:%d in %q", port, out)
	}
	return strings.Join(got, "\n")
}

// showsList reports whether one of lines has `allow:` followed by every item.
func showsList(lines string, items ...string) bool {
	for _, l := range strings.Split(lines, "\n") {
		i := strings.Index(l, "allow:")
		if i < 0 {
			continue
		}
		ok := true
		for _, it := range items {
			ok = ok && strings.Contains(l[i:], it)
		}
		if ok {
			return true
		}
	}
	return false
}

// ---- REQ-ALLOW-001 / REQ-ALLOW-002 ----

// TC-ALLOW-017
func TestREQ_ALLOW_002_ServeListedClientPrinted(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, allowConfig(allowEntry("front", printerName, r.port, `,"allow":["10.0.0.5","127.0.0.0/8"]`)))
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.send("first").Close()
	if docs := r.sp.waitDone(1, 3*time.Second); len(docs) != 1 || string(docs[0].data) != "first" {
		r.stop()
		t.Fatalf("documents %+v, want `first` printed (127.0.0.1 is inside 127.0.0.0/8)", docs)
	}
	// S1:33: a changed list needs a restart, so rewriting the file while serving changes nothing
	r.write(r.cfg, allowConfig(allowEntry("front", printerName, r.port, `,"allow":["10.0.0.5"]`)))
	r.send("second").Close()
	docs := r.sp.waitDone(2, 3*time.Second)
	lines := r.waitConnLines(2)
	r.stop()
	if len(docs) != 2 || string(docs[1].data) != "second" {
		t.Fatalf("documents %+v, want `second` printed too (no reload without restart)", docs)
	}
	if strings.Contains(strings.Join(lines, "\n"), notAllowed) {
		t.Fatalf("connection lines %q mention a rejection", lines)
	}
	if l := startupLines(t, r.stdout.String(), r.port); !showsList(l, "10.0.0.5", "127.0.0.0/8") {
		t.Fatalf("start-up lines %q do not show the allowlist", l)
	}
	// TC-ALLOW-038: after a restart the rewritten list (10.0.0.5 only) applies
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("third")
	serverClosedWithin(c, time.Second)
	c.Close()
	lines = r.waitConnLines(3)
	r.stop()
	if docs := r.sp.done(); len(docs) != 2 || len(lines) != 3 || !strings.HasSuffix(lines[2], notAllowed) {
		t.Fatalf("after restart: documents %d (want 2), connection lines %q (want the 3rd rejected)", len(docs), lines)
	}
}

// ---- REQ-ALLOW-003 ----

// TC-ALLOW-024
func TestREQ_ALLOW_003_ServeRejectsUnlistedClient(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, allowConfig(allowEntry("front", printerName, r.port, `,"allow":["10.9.9.9","192.168.0.0/16"]`)))
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("secret")
	defer c.Close()
	closed := serverClosedWithin(c, time.Second)
	lines := r.waitConnLines(1)
	time.Sleep(200 * time.Millisecond)
	r.stop()
	if !closed {
		t.Fatalf("the server did not close the rejected connection within 1 s")
	}
	r.sp.mu.Lock()
	started, opened := r.sp.startDoc, len(r.sp.opens)
	r.sp.mu.Unlock()
	if started != 0 || opened != 0 {
		t.Fatalf("OpenPrinter %d / StartDocPrinter %d call(s), want nothing printed", opened, started)
	}
	want := `0B "` + printerName + `" ` + notAllowed
	if len(lines) != 1 || !strings.HasSuffix(lines[0], want) {
		t.Fatalf("connection lines %q, want exactly one ending %q", lines, want)
	}
}

// ---- REQ-ALLOW-005 / REQ-ALLOW-006 ----

// TC-ALLOW-032
func TestREQ_ALLOW_005_ServeEmptyAllowExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, allowConfig(allowEntry("front", printerName, r.port, `,"allow":[]`)))
	if code := r.run("serve", "--config", r.cfg); code != 2 {
		t.Fatalf("exit %d, want 2; stderr %q", code, r.stderr.String())
	}
	oneStderrLine(t, r.stderr.String(), r.cfg, "printers[0].allow")
	if strings.Contains(r.stderr.String(), "printers[0].allow[") {
		t.Fatalf("stderr %q names an item, want the list field", r.stderr.String())
	}
	r.noReady()
}

// TC-ALLOW-033
func TestREQ_ALLOW_006_ServeInvalidItemExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, allowConfig(allowEntry("front", printerName, r.port, `,"allow":["10.0.0.5","10.0.0.0/33"]`)))
	if code := r.run("serve", "--config", r.cfg); code != 2 {
		t.Fatalf("exit %d, want 2; stderr %q", code, r.stderr.String())
	}
	oneStderrLine(t, r.stderr.String(), r.cfg, "printers[0].allow[1]")
	r.noReady()
}

// ---- REQ-ALLOW-007 ----

// TC-ALLOW-034
func TestREQ_ALLOW_007_ListAppliesOnlyToItsPort(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	p2 := freePort(t)
	r.write(r.cfg, allowConfig(allowEntry("front", "A", r.port, `,"allow":["10.9.9.9"]`), allowEntry("back", "B", p2, ``)))
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	sendPort(t, r.port, "to-A").Close()
	sendPort(t, p2, "to-B").Close()
	docs := r.sp.waitDone(1, 3*time.Second)
	lines := r.waitConnLines(2)
	time.Sleep(200 * time.Millisecond)
	docs = r.sp.done()
	r.stop()
	if len(docs) != 1 || docs[0].printer != "B" || string(docs[0].data) != "to-B" {
		t.Fatalf("documents %+v, want only `to-B` on B (front's list must not apply to back's port)", docs)
	}
	var rejected []string
	for _, l := range lines {
		if strings.HasSuffix(l, notAllowed) {
			rejected = append(rejected, l)
		}
	}
	if len(rejected) != 1 || !strings.HasSuffix(rejected[0], `0B "A" `+notAllowed) {
		t.Fatalf("connection lines %q, want one `0B \"A\" %s`", lines, notAllowed)
	}
}

// ---- REQ-ALLOW-008 ----

// TC-ALLOW-035
func TestREQ_ALLOW_008_StartUpLinesShowEachList(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	p2 := freePort(t)
	r.write(r.cfg, allowConfig(allowEntry("front", "A", r.port, `,"allow":["192.168.1.0/24","10.0.0.5"]`),
		allowEntry("back", "B", p2, ``)))
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	out := r.stdout.String()
	r.stop()
	front, back := startupLines(t, out, r.port), startupLines(t, out, p2)
	if !showsList(front, "192.168.1.0/24", "10.0.0.5") || strings.Contains(front, "allow: all") {
		t.Fatalf("front start-up lines %q, want `allow:` followed by 192.168.1.0/24 and 10.0.0.5", front)
	}
	if !strings.Contains(back, "allow: all") || strings.Contains(back, "10.0.0.5") {
		t.Fatalf("back start-up lines %q, want `allow: all` and no other printer's list", back)
	}
}

// TC-ALLOW-036
func TestREQ_ALLOW_008_SinglePrinterAllowAll(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg) // C-std: one printer, no allow key
	r.waitReady()
	out := r.stdout.String()
	r.stop()
	if l := startupLines(t, out, r.port); !strings.Contains(l, "allow: all") {
		t.Fatalf("start-up lines %q, want `allow: all`", l)
	}
}

// serverClosedWithin reports whether the server closed c (EOF or reset) within d, as opposed to a read timeout.
func serverClosedWithin(c interface {
	Read([]byte) (int, error)
	SetReadDeadline(time.Time) error
}, d time.Duration) bool {
	c.SetReadDeadline(time.Now().Add(d))
	buf := make([]byte, 64)
	for {
		_, err := c.Read(buf)
		if err == nil {
			continue
		}
		if ne, ok := err.(interface{ Timeout() bool }); ok && ne.Timeout() {
			return false
		}
		return true
	}
}
