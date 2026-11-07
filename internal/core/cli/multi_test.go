package cli

import (
	"errors"
	"net"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

// F-003 tests through `printerone serve` and `printerone service install|uninstall` (spec
// docs/specs/2026-11-07-f003-multi-printer.md). TC ids: .testcases/runs/F-003-r1 and testcases.md.
// Contract: Env.Ready is called once, after every port listens.

var multiNames = []string{"front", "back", "third"}

// multiJSON is a config with one entry per port: names front/back/third, printers[i] on ports[i].
func multiJSON(printers []string, ports []int) string {
	parts := make([]string, len(ports))
	for i := range ports {
		parts[i] = `{"name":"` + multiNames[i] + `","printer":"` + printers[i] + `","port":` + strconv.Itoa(ports[i]) + `}`
	}
	return `{"version":2,"printers":[` + strings.Join(parts, ",") + `]}`
}

// multiRig is a serve rig whose config has one entry per printer, on r.port and fresh free ports.
func multiRig(t *testing.T, edition string, printers ...string) (*rig, []int) {
	r := newRig(t, edition)
	ports := []int{r.port}
	for len(ports) < len(printers) {
		ports = append(ports, freePort(t))
	}
	r.write(r.cfg, multiJSON(printers, ports))
	return r, ports
}

func dialPort(t *testing.T, port int) *net.TCPConn {
	t.Helper()
	c, err := net.DialTimeout("tcp", "127.0.0.1:"+strconv.Itoa(port), 2*time.Second)
	if err != nil {
		t.Fatalf("dial port %d: %v", port, err)
	}
	return c.(*net.TCPConn)
}

func sendPort(t *testing.T, port int, data string) *net.TCPConn {
	t.Helper()
	c := dialPort(t, port)
	if _, err := c.Write([]byte(data)); err != nil {
		t.Fatal(err)
	}
	return c
}

func refused(port int) bool {
	c, err := net.DialTimeout("tcp", "127.0.0.1:"+strconv.Itoa(port), 500*time.Millisecond)
	if err != nil {
		return true
	}
	c.Close()
	return false
}

func docsOn(docs []doc, printer string) []string {
	var out []string
	for _, d := range docs {
		if d.printer == printer {
			out = append(out, string(d.data))
		}
	}
	return out
}

func (f *fakeSpool) writes() int {
	n := 0
	for {
		select {
		case <-f.writing:
			n++
		default:
			return n
		}
	}
}

// ---- REQ-MULTI-002 ----

// TC-MULTI-011
func TestREQ_MULTI_002_ServeInvalidEntryExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, multiJSON([]string{"A", "B"}, []int{r.port, 65536}))
	if code := r.run("serve", "--config", r.cfg); code != 2 {
		t.Fatalf("exit %d, want 2; stderr %q", code, r.stderr.String())
	}
	oneStderrLine(t, r.stderr.String(), r.cfg, "printers[1].port")
	r.noReady()
}

// ---- REQ-MULTI-005 ----

// TC-MULTI-021
func TestREQ_MULTI_005_EachPortPrintsOnItsPrinter(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	sendPort(t, ports[0], "to-A").Close()
	sendPort(t, ports[1], "to-B").Close()
	docs := r.sp.waitDone(2, 3*time.Second)
	a, b := docsOn(docs, "A"), docsOn(docs, "B")
	if len(docs) != 2 || len(a) != 1 || a[0] != "to-A" || len(b) != 1 || b[0] != "to-B" {
		t.Fatalf("documents %+v; want exactly `to-A` on A and `to-B` on B", docs)
	}
	r.stop()
}

// ---- REQ-MULTI-006 ----

// TC-MULTI-022
func TestREQ_MULTI_006_DifferentPrintersPrintAtOnce(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	gate := make(chan struct{})
	r.sp.gate = gate
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	sendPort(t, ports[0], "job-A").Close()
	sendPort(t, ports[1], "job-B").Close()
	n := 0
	for deadline := time.Now().Add(2 * time.Second); n < 2 && time.Now().Before(deadline); time.Sleep(10 * time.Millisecond) {
		n += r.sp.writes()
	}
	close(gate)
	if n < 2 {
		r.stop()
		t.Fatalf("%d WritePrinter call(s) in progress while the first was blocked; want 2 (printers A and B at once)", n)
	}
	if docs := r.sp.waitDone(2, 3*time.Second); len(docs) != 2 {
		t.Fatalf("documents %+v, want 2", docs)
	}
	r.stop()
}

// TC-MULTI-023
func TestREQ_MULTI_006_SameWindowsPrinterOneAfterAnother(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "A")
	defer r.cleanup()
	gate := make(chan struct{})
	r.sp.gate = gate
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	sendPort(t, ports[0], "one").Close()
	sendPort(t, ports[1], "two").Close()
	time.Sleep(500 * time.Millisecond)
	n := r.sp.writes()
	close(gate)
	docs := r.sp.waitDone(2, 3*time.Second)
	r.sp.mu.Lock()
	maxOpen := r.sp.maxOpen
	r.sp.mu.Unlock()
	if n != 1 || len(docs) != 2 || maxOpen != 1 {
		t.Fatalf("writes in progress %d (want 1), documents %d (want 2), max open documents %d (want 1)", n, len(docs), maxOpen)
	}
	r.stop()
}

// ---- REQ-MULTI-007 ----

func holdPort(t *testing.T, port int) net.Listener {
	t.Helper()
	l, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(port))
	if err != nil {
		t.Fatal(err)
	}
	return l
}

// TC-MULTI-024
func TestREQ_MULTI_007_SecondPortBusyNothingListens(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	busy := holdPort(t, ports[1])
	defer busy.Close()
	if code := r.run("serve", "--config", r.cfg); code != 1 {
		t.Fatalf("exit %d, want 1; stderr %q", code, r.stderr.String())
	}
	if !strings.Contains(r.stderr.String(), strconv.Itoa(ports[1])) {
		t.Fatalf("stderr %q does not name busy port %d", r.stderr.String(), ports[1])
	}
	r.noReady()
	l, err := net.Listen("tcp", "0.0.0.0:"+strconv.Itoa(ports[0]))
	if err != nil {
		t.Fatalf("port %d still bound after the exit: %v", ports[0], err)
	}
	l.Close()
}

// TC-MULTI-025
func TestREQ_MULTI_007_FirstPortBusyNamesIt(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	busy := holdPort(t, ports[0])
	defer busy.Close()
	code := r.run("serve", "--config", r.cfg)
	s := r.stderr.String()
	if code != 1 || !strings.Contains(s, strconv.Itoa(ports[0])) || strings.Contains(s, strconv.Itoa(ports[1])) {
		t.Fatalf("exit %d, stderr %q; want 1 naming %d and not %d", code, s, ports[0], ports[1])
	}
	r.noReady()
}

// ---- REQ-MULTI-008 ----

// TC-MULTI-028
func TestREQ_MULTI_008_LegacySlotsSharedAcrossPorts(t *testing.T) {
	r, ports := multiRig(t, "legacy", "A", "B", "C")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	held1 := sendPort(t, ports[0], "x")
	held2 := sendPort(t, ports[1], "x")
	defer held2.Close()
	time.Sleep(100 * time.Millisecond)
	sendPort(t, ports[2], "c").Close()
	time.Sleep(500 * time.Millisecond)
	if got := docsOn(r.sp.done(), "C"); len(got) != 0 {
		held1.Close()
		t.Fatalf("job on the third port printed while both legacy slots were held on other ports: %q", got)
	}
	held1.Close()
	deadline := time.Now().Add(3 * time.Second)
	for len(docsOn(r.sp.done(), "C")) == 0 && time.Now().Before(deadline) {
		time.Sleep(10 * time.Millisecond)
	}
	if got := docsOn(r.sp.done(), "C"); len(got) != 1 || got[0] != "c" {
		t.Fatalf("after a slot frees: documents on C %q, want [c]", got)
	}
	held2.Close()
	r.stop()
}

// TC-MULTI-029
func TestREQ_MULTI_008_JobLimitOnSecondPort(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.sp.countAll = true
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := dialPort(t, ports[1])
	chunk := make([]byte, 1<<20)
	for left := 268435457; left > 0; {
		n := len(chunk)
		if n > left {
			n = left
		}
		if _, err := c.Write(chunk[:n]); err != nil {
			break
		}
		left -= n
	}
	c.CloseWrite()
	lines := r.waitConnLines(1)
	c.Close()
	if len(lines) != 1 || !strings.HasSuffix(lines[0], `"B" dropped: too large`) {
		t.Fatalf("connection lines %q, want one ending `\"B\" dropped: too large`", lines)
	}
	r.sp.mu.Lock()
	n := r.sp.startDoc
	r.sp.mu.Unlock()
	if n != 0 {
		t.Fatalf("StartDocPrinter called %d times for a dropped job", n)
	}
	r.stop()
}

// ---- REQ-MULTI-009 ----

// TC-MULTI-030
func TestREQ_MULTI_009_StartUpLinesListEveryPrinter(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.stop()
	logFile := readFileOrEmpty(filepath.Join(r.pd, "PrinterOne", "logs", "printerone.log"))
	for i, p := range []string{"A", "B"} {
		want := multiNames[i] + " -> " + p + " on 0.0.0.0:" + strconv.Itoa(ports[i])
		if strings.Count(r.stdout.String(), want) != 1 || !strings.Contains(logFile, want) {
			t.Fatalf("want %q once on stdout and in printerone.log\nstdout %q\nlog %q", want, r.stdout.String(), logFile)
		}
	}
}

// TC-MULTI-031
func TestREQ_MULTI_009_ConnectionLineNamesPortPrinter(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	sendPort(t, ports[1], "x").Close()
	first := r.waitConnLines(1)
	sendPort(t, ports[0], "y").Close()
	lines := r.waitConnLines(2)
	r.stop()
	if len(first) != 1 || len(lines) != 2 || !strings.HasSuffix(lines[0], `1B "B" printed`) || !strings.HasSuffix(lines[1], `1B "A" printed`) {
		t.Fatalf("connection lines %q; want the p2 job on \"B\", then the p1 job on \"A\"", lines)
	}
}

// ---- REQ-MULTI-010 ----

// TC-MULTI-033
func TestREQ_MULTI_010_CtrlCDropsReceivingOnEveryPort(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c1 := sendPort(t, ports[0], "x")
	defer c1.Close()
	c2 := sendPort(t, ports[1], "x")
	defer c2.Close()
	time.Sleep(200 * time.Millisecond)
	r.stop()
	lines := r.connLines()
	a, b := 0, 0
	for _, l := range lines {
		if strings.HasSuffix(l, `"A" dropped: shutdown`) {
			a++
		}
		if strings.HasSuffix(l, `"B" dropped: shutdown`) {
			b++
		}
	}
	if a != 1 || b != 1 || len(r.sp.done()) != 0 {
		t.Fatalf("connection lines %q, documents %d; want one `\"A\" dropped: shutdown`, one `\"B\" dropped: shutdown`, nothing printed", lines, len(r.sp.done()))
	}
	if !refused(ports[0]) || !refused(ports[1]) {
		t.Fatalf("a port still accepts connections after Ctrl+C (p1 refused=%v, p2 refused=%v)", refused(ports[0]), refused(ports[1]))
	}
}

// TC-MULTI-034
func TestREQ_MULTI_010_CtrlCPrintsReceivedJobOnSecondPort(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	gate := make(chan struct{})
	r.sp.gate = gate
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	sendPort(t, ports[1], "hello").Close()
	select {
	case <-r.sp.writing:
	case <-time.After(3 * time.Second):
		close(gate)
		t.Fatal("job on the second port never reached WritePrinter")
	}
	r.sig <- os.Interrupt
	if code, ok := r.waitExit(300 * time.Millisecond); ok {
		close(gate)
		t.Fatalf("Run returned %d while the job on the second port was still being written", code)
	}
	close(gate)
	code, ok := r.waitExit(5 * time.Second)
	b := docsOn(r.sp.done(), "B")
	if !ok || code != 0 || len(b) != 1 || b[0] != "hello" {
		t.Fatalf("returned=%v code=%d documents on B %q; want exit 0 with `hello` printed", ok, code, b)
	}
}

// ---- REQ-MULTI-011 (fake SCM + netsh, service_test.go) ----

func (r *svcRig) multiConfig(name string, ports ...int) string {
	printers := make([]string, len(ports))
	for i := range printers {
		printers[i] = printerName
	}
	return r.writeFile(filepath.Join(r.dir, name, "config.json"), multiJSON(printers, ports))
}

func withPrefix(calls []string, prefix string) []string {
	var out []string
	for _, c := range calls {
		if strings.HasPrefix(c, prefix) {
			out = append(out, c)
		}
	}
	return out
}

// TC-MULTI-036
func TestREQ_MULTI_011_VistaInstallOneRuleEveryPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	code := r.run("service", "install", "--config", r.multiConfig("new", 9100, 9101))
	adds := withPrefix(r.netsh.list(), "netsh advfirewall firewall add")
	if code != 0 || len(adds) != 1 || adds[0] != advAdd+"9100,9101" {
		t.Fatalf("exit %d; add calls %q, want exactly [%q]\n%s", code, adds, advAdd+"9100,9101", r.dump())
	}
}

// TC-MULTI-037
func TestREQ_MULTI_011_XPInstallOpensEveryPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	code := r.run("service", "install", "--config", r.multiConfig("new", 9100, 9101))
	calls := r.netsh.list()
	adds := withPrefix(calls, "netsh firewall add portopening")
	want := []string{"netsh firewall add portopening protocol=TCP port=9100 name=PrinterOne", "netsh firewall add portopening protocol=TCP port=9101 name=PrinterOne"}
	if code != 0 || strings.Join(adds, "\n") != strings.Join(want, "\n") || len(withPrefix(calls, "netsh advfirewall")) != 0 {
		t.Fatalf("exit %d; add calls %q, want %q and no advfirewall\n%s", code, adds, want, r.dump())
	}
}

// TC-MULTI-038
func TestREQ_MULTI_011_XPUninstallClosesEveryPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	r.scm.exists, r.scm.settings = true, wantSettings(r.multiConfig("old", 9102, 9103))
	r.netsh.openPorts = map[int]bool{9102: true, 9103: true}
	code := r.run("service", "uninstall")
	calls := r.netsh.list()
	show := indexOf(calls, "netsh firewall show portopening")
	d1 := indexOf(calls, "netsh firewall delete portopening protocol=TCP port=9102")
	d2 := indexOf(calls, "netsh firewall delete portopening protocol=TCP port=9103")
	if code != 0 || show < 0 || d1 < show || d2 < show || len(r.netsh.openPorts) != 0 {
		t.Fatalf("exit %d; netsh %q; want show, then delete 9102 and 9103\n%s", code, calls, r.dump())
	}
}

// TC-MULTI-039
func TestREQ_MULTI_011_XPInstallClosesEveryOldPortFirst(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	r.scm.exists, r.scm.settings = true, wantSettings(r.multiConfig("old", 9102, 9103))
	r.netsh.openPorts = map[int]bool{9102: true, 9103: true}
	code := r.run("service", "install", "--config", r.multiConfig("new", 9100, 9101))
	calls := r.netsh.list()
	d1 := indexOf(calls, "netsh firewall delete portopening protocol=TCP port=9102")
	d2 := indexOf(calls, "netsh firewall delete portopening protocol=TCP port=9103")
	a1 := indexOf(calls, "netsh firewall add portopening protocol=TCP port=9100 name=PrinterOne")
	a2 := indexOf(calls, "netsh firewall add portopening protocol=TCP port=9101 name=PrinterOne")
	if code != 0 || d1 < 0 || d2 < 0 || a1 < 0 || a2 < 0 || d1 > a1 || d2 > a1 {
		t.Fatalf("exit %d; netsh %q; want deletes of 9102 and 9103 before adds of 9100 and 9101\n%s", code, calls, r.dump())
	}
}

// TC-MULTI-040
func TestREQ_MULTI_011_XPRollbackClosesEveryPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	r.scm.startErr = errors.New("1053 The service did not respond to the start or control request in a timely fashion.")
	code := r.run("service", "install", "--config", r.multiConfig("new", 9100, 9101))
	calls := r.netsh.list()
	opened := len(withPrefix(calls, "netsh firewall add portopening")) == 2
	if code != 1 || !opened || len(r.netsh.openPorts) != 0 {
		t.Fatalf("exit %d; netsh %q; open ports after rollback %v; want exit 1, both ports opened then all closed\n%s", code, calls, r.netsh.openPorts, r.dump())
	}
}

// TC-MULTI-041
func TestREQ_MULTI_011_VistaUninstallRemovesRule(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists, r.scm.settings = true, wantSettings(r.multiConfig("old", 9100, 9101))
	r.netsh.ruleExists = true
	code := r.run("service", "uninstall")
	if got, want := r.netsh.list(), []string{advShow, advDelete}; code != 0 || strings.Join(got, "\n") != strings.Join(want, "\n") {
		t.Fatalf("exit %d; netsh %q, want exactly %q\n%s", code, got, want, r.dump())
	}
}

// TC-MULTI-043
func TestREQ_MULTI_011_XPAddFailureClosesOpenedPorts(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	const out = "An unexpected error occurred."
	r.netsh.reply = func(args []string) netshReply {
		if strings.Join(args, " ") == "firewall add portopening protocol=TCP port=9101 name=PrinterOne" {
			return netshReply{out: out + "\r\n", err: errExit1}
		}
		return r.netsh.statefulReply(args)
	}
	code := r.run("service", "install", "--config", r.multiConfig("new", 9100, 9101))
	calls := r.netsh.list()
	if code != 1 || !strings.Contains(r.stderr.String(), out) || indexOf(calls, "netsh firewall add portopening protocol=TCP port=9100 name=PrinterOne") < 0 ||
		r.netsh.openPorts[9100] || r.scm.exists {
		t.Fatalf("exit %d; open ports %v; service exists %v; want exit 1 with the netsh output, 9100 opened then closed, service removed\n%s",
			code, r.netsh.openPorts, r.scm.exists, r.dump())
	}
}

// TC-MULTI-045
func TestREQ_MULTI_002_ServiceInstallInvalidEntryChangesNothing(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	cfg := r.writeFile(filepath.Join(r.dir, "bad", "config.json"), multiJSON([]string{"A", "B"}, []int{9100, 65536}))
	code := r.run("service", "install", "--config", cfg)
	if code != 2 || !strings.Contains(r.stderr.String(), "printers[1].port") || len(r.scm.mutations()) != 0 || len(r.netsh.list()) != 0 {
		t.Fatalf("exit %d; want 2 naming printers[1].port with no SCM or netsh change\n%s", code, r.dump())
	}
}

// TC-MULTI-044
func TestREQ_MULTI_008_MainSlotsSharedAcrossPorts(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	var held []*net.TCPConn
	for i := 0; i < 8; i++ {
		held = append(held, sendPort(t, ports[i%2], "x"))
	}
	defer func() {
		for _, c := range held {
			c.Close()
		}
	}()
	time.Sleep(200 * time.Millisecond)
	sendPort(t, ports[1], "ninth").Close()
	time.Sleep(500 * time.Millisecond)
	if got := docsOn(r.sp.done(), "B"); len(got) != 0 {
		t.Fatalf("documents on B %q while 8 shared slots were held; want none", got)
	}
	held[0].Close() // a p1 connection
	deadline := time.Now().Add(3 * time.Second)
	for len(docsOn(r.sp.done(), "B")) == 0 && time.Now().Before(deadline) {
		time.Sleep(10 * time.Millisecond)
	}
	if got := docsOn(r.sp.done(), "B"); len(got) != 1 || got[0] != "ninth" {
		t.Fatalf("after a slot frees: documents on B %q, want [ninth]", got)
	}
	for _, c := range held {
		c.Close()
	}
	r.stop()
}

// TC-MULTI-046
func TestREQ_MULTI_008_JobLimitIsPerJobNotPerProcess(t *testing.T) {
	r, ports := multiRig(t, "main", "A", "B")
	defer r.cleanup()
	r.sp.countAll = true
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	const size = 150 << 20 // two at once: 300 MiB in flight, above 256 MB
	done := make(chan struct{}, 2)
	for _, p := range ports {
		c := dialPort(t, p)
		go func(c *net.TCPConn) {
			chunk := make([]byte, 1<<20)
			for left := size; left > 0; left -= len(chunk) {
				if _, err := c.Write(chunk); err != nil {
					break
				}
			}
			c.CloseWrite()
			done <- struct{}{}
		}(c)
		defer c.Close()
	}
	<-done
	<-done
	lines := r.waitConnLinesFor(2, 10*time.Second)
	if len(lines) != 2 || !strings.HasSuffix(lines[0], " printed") || !strings.HasSuffix(lines[1], " printed") {
		t.Fatalf("connection lines %q; want both 150 MiB jobs printed", lines)
	}
	r.stop()
}

// TC-MULTI-047
func TestREQ_MULTI_009_SinglePrinterStartUpLine(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.cfg, `{"version":2,"printers":[{"printer":"`+printerName+`","port":`+strconv.Itoa(r.port)+`}]}`)
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	r.stop()
	want := "default -> " + printerName + " on 0.0.0.0:" + strconv.Itoa(r.port)
	if strings.Count(r.stdout.String(), want) != 1 {
		t.Fatalf("stdout %q, want %q once", r.stdout.String(), want)
	}
}
