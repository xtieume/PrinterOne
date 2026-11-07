package cli

import (
	"bytes"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/service"
)

// Tests of `printerone service install|uninstall|run` (spec docs/specs/2026-11-07-f008-service.md), driven
// through Run with a fake SCM, a fake netsh and a fake clock. TC ids: .testcases/runs/F-008-r1 and testcases.md.

const svcExe = `C:\Program Files\PrinterOne\printerone.exe`

// ---- fake clock: Sleep advances Now; a run past 10 minutes of fake time is a wait loop that never ends ----

type fakeClock struct {
	mu  sync.Mutex
	now time.Time
}

func newFakeClock() *fakeClock { return &fakeClock{now: time.Date(2026, 11, 7, 9, 0, 0, 0, time.UTC)} }

func (c *fakeClock) Now() time.Time {
	c.mu.Lock()
	defer c.mu.Unlock()
	return c.now
}

func (c *fakeClock) Sleep(d time.Duration) {
	c.mu.Lock()
	defer c.mu.Unlock()
	if d <= 0 {
		d = time.Millisecond
	}
	c.now = c.now.Add(d)
	if c.now.Sub(newFakeClock().now) > 10*time.Minute {
		panic("fake clock: waited more than 10 minutes of fake time (no timeout?)")
	}
}

// ---- fake SCM ----

type fakeSCM struct {
	mu        sync.Mutex
	clock     *fakeClock
	admin     bool
	exists    bool
	settings  service.Settings
	state     service.State
	calls     []string      // create, update, start, stop, delete — in order, with the state seen at the call
	runAfter  time.Duration // fake time from Start to Running; < 0 never
	stopAfter time.Duration // fake time from Stop to Stopped
	since     time.Time
	startErr  error
	deleteErr error
	created   []service.Settings
	updated   []service.Settings
}

func newFakeSCM(c *fakeClock) *fakeSCM {
	return &fakeSCM{clock: c, admin: true, state: service.Stopped, runAfter: 2 * time.Second, stopAfter: 2 * time.Second}
}

func (f *fakeSCM) record(s string) {
	f.calls = append(f.calls, s)
}

func (f *fakeSCM) IsAdmin() bool { return f.admin }
func (f *fakeSCM) Exists(name string) (bool, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	return f.exists && name == service.Name, nil
}
func (f *fakeSCM) Create(name string, s service.Settings) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.record("create")
	f.created = append(f.created, s)
	f.exists, f.settings, f.state = true, s, service.Stopped
	return nil
}
func (f *fakeSCM) Update(name string, s service.Settings) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.record("update")
	f.updated = append(f.updated, s)
	f.settings = s
	return nil
}
func (f *fakeSCM) Config(name string) (service.Settings, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	if !f.exists {
		return service.Settings{}, &service.Error{Code: 1, Msg: "1060 The specified service does not exist as an installed service."}
	}
	return f.settings, nil
}
func (f *fakeSCM) Start(name string) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.record("start:" + stateName(f.current()))
	if f.startErr != nil {
		return f.startErr
	}
	f.state, f.since = service.StartPending, f.clock.Now()
	return nil
}
func (f *fakeSCM) Stop(name string) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.record("stop:" + stateName(f.current()))
	f.state, f.since = service.StopPending, f.clock.Now()
	return nil
}
func (f *fakeSCM) Delete(name string) error {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.record("delete:" + stateName(f.current()))
	if f.deleteErr != nil {
		return f.deleteErr
	}
	f.exists = false
	return nil
}
func (f *fakeSCM) State(name string) (service.State, error) {
	f.mu.Lock()
	defer f.mu.Unlock()
	if !f.exists {
		return 0, &service.Error{Code: 1, Msg: "1060 The specified service does not exist as an installed service."}
	}
	return f.current(), nil
}

// current advances pending states by the fake clock.
func (f *fakeSCM) current() service.State {
	elapsed := f.clock.Now().Sub(f.since)
	if f.state == service.StartPending && f.runAfter >= 0 && elapsed >= f.runAfter {
		f.state = service.Running
	}
	if f.state == service.StopPending && elapsed >= f.stopAfter {
		f.state = service.Stopped
	}
	return f.state
}

func stateName(s service.State) string {
	switch s {
	case service.Stopped:
		return "stopped"
	case service.StartPending:
		return "start-pending"
	case service.StopPending:
		return "stop-pending"
	case service.Running:
		return "running"
	}
	return fmt.Sprintf("state-%d", s)
}

func (f *fakeSCM) callList() []string {
	f.mu.Lock()
	defer f.mu.Unlock()
	return append([]string(nil), f.calls...)
}

// mutations returns every call that changes the machine.
func (f *fakeSCM) mutations() []string { return f.callList() }

func (f *fakeSCM) has(prefix string) bool {
	for _, c := range f.callList() {
		if strings.HasPrefix(c, prefix) {
			return true
		}
	}
	return false
}

// ---- fake netsh ----

type netshReply struct {
	out string
	err error
}

type fakeNetsh struct {
	mu    sync.Mutex
	calls []string                       // "netsh arg arg ..."
	reply func(args []string) netshReply // nil: defaultNetsh
}

var errExit1 = errors.New("exit status 1")

func defaultNetsh(args []string) netshReply {
	if len(args) > 2 && args[2] == "delete" {
		return netshReply{out: "\r\nDeleted 1 rule(s).\r\nOk.\r\n"}
	}
	return netshReply{out: "Ok.\r\n"}
}

func noRulesMatch(args []string) netshReply {
	if len(args) > 2 && args[2] == "delete" {
		return netshReply{out: "\r\nNo rules match the specified criteria.\r\n\r\n", err: errExit1}
	}
	return netshReply{out: "Ok.\r\n"}
}

func (n *fakeNetsh) run(name string, args ...string) ([]byte, error) {
	n.mu.Lock()
	defer n.mu.Unlock()
	n.calls = append(n.calls, name+" "+strings.Join(args, " "))
	r := n.reply
	if r == nil {
		r = defaultNetsh
	}
	got := r(args)
	return []byte(got.out), got.err
}

func (n *fakeNetsh) list() []string {
	n.mu.Lock()
	defer n.mu.Unlock()
	return append([]string(nil), n.calls...)
}

// ---- rig ----

type svcRig struct {
	t              *testing.T
	dir            string // temp dir: config, exe folder, data dir
	clock          *fakeClock
	scm            *fakeSCM
	netsh          *fakeNetsh
	stdout, stderr bytes.Buffer
	env            Env
}

func newSvcRig(t *testing.T) *svcRig {
	dir, err := ioutil.TempDir("", "printerone-svc")
	if err != nil {
		t.Fatal(err)
	}
	r := &svcRig{t: t, dir: dir, clock: newFakeClock(), netsh: &fakeNetsh{}}
	r.scm = newFakeSCM(r.clock)
	pd := filepath.Join(dir, "ProgramData")
	r.env = Env{Edition: "main", ExeDir: filepath.Join(dir, "exe"), ExePath: svcExe, Spool: newFakeSpool(),
		Getenv: func(k string) string {
			if k == "ProgramData" {
				return pd
			}
			return ""
		},
		Stdout: &r.stdout, Stderr: &r.stderr,
		Service: service.Deps{SCM: r.scm, Netsh: r.netsh.run, OSMajor: 10, Sleep: r.clock.Sleep, Now: r.clock.Now}}
	return r
}

func (r *svcRig) cleanup() { os.RemoveAll(r.dir) }

// config writes a valid one-printer config for port into the rig dir (or path when given) and returns its path.
func (r *svcRig) config(port int) string {
	return r.writeFile(filepath.Join(r.dir, "config.json"), configJSON(port))
}

func (r *svcRig) writeFile(path, body string) string {
	if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		r.t.Fatal(err)
	}
	if err := ioutil.WriteFile(path, []byte(body), 0644); err != nil {
		r.t.Fatal(err)
	}
	return path
}

func (r *svcRig) run(args ...string) int {
	r.t.Helper()
	return Run(args, r.env)
}

func (r *svcRig) dump() string {
	return fmt.Sprintf("stdout %q\nstderr %q\nscm %q\nnetsh %q", r.stdout.String(), r.stderr.String(), r.scm.callList(), r.netsh.list())
}

func cmdLine(cfg string) string { return `"` + svcExe + `" service run --config "` + cfg + `"` }

func wantSettings(cfg string) service.Settings {
	return service.Settings{DisplayName: "PrinterOne print server", Description: "Receives RAW print jobs over TCP and prints them",
		BinaryPath: cmdLine(cfg), StartType: 2, Account: "LocalSystem"}
}

const (
	advDelete = "netsh advfirewall firewall delete rule name=PrinterOne"
	advAdd    = "netsh advfirewall firewall add rule name=PrinterOne dir=in action=allow protocol=TCP localport="
	okInstall = "service PrinterOne installed and running"
	okRemove  = "service PrinterOne removed"
)

func indexOf(list []string, s string) int {
	for i, v := range list {
		if v == s {
			return i
		}
	}
	return -1
}

func lastIndexOf(list []string, s string) int {
	for i := len(list) - 1; i >= 0; i-- {
		if list[i] == s {
			return i
		}
	}
	return -1
}

// ---- REQ-SVC-001 ----

// TC-SVC-001
func TestREQ_SVC_001_InstallWithoutAdminChangesNothing(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.admin = false
	code := r.run("service", "install", "--config", r.config(9101))
	if code != 1 || !strings.Contains(r.stderr.String(), "run as administrator") || len(r.scm.mutations()) != 0 || len(r.netsh.list()) != 0 {
		t.Fatalf("exit %d; want 1, stderr `run as administrator`, no SCM change, no netsh\n%s", code, r.dump())
	}
}

// TC-SVC-002
func TestREQ_SVC_001_AdminCheckedBeforeConfig(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.admin = false
	code := r.run("service", "install", "--config", filepath.Join(r.dir, "missing.json"))
	if code != 1 || !strings.Contains(r.stderr.String(), "run as administrator") {
		t.Fatalf("exit %d; want 1 and `run as administrator` (admin is checked before the config)\n%s", code, r.dump())
	}
}

// TC-SVC-003
func TestREQ_SVC_001_UninstallWithoutAdminChangesNothing(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.admin = false
	r.scm.exists, r.scm.state, r.scm.settings = true, service.Running, wantSettings(r.config(9101))
	code := r.run("service", "uninstall")
	if code != 1 || !strings.Contains(r.stderr.String(), "run as administrator") || len(r.scm.mutations()) != 0 || len(r.netsh.list()) != 0 {
		t.Fatalf("exit %d; want 1, stderr `run as administrator`, no SCM change, no netsh\n%s", code, r.dump())
	}
}

// ---- REQ-SVC-002 ----

// TC-SVC-004, TC-SVC-005
func TestREQ_SVC_002_InvalidConfigExits2ChangesNothing(t *testing.T) {
	for _, c := range []struct{ tc, body string }{
		{"TC-SVC-004", `{"version":2,`},
		{"TC-SVC-005", ""}, // no file
	} {
		r := newSvcRig(t)
		cfg := filepath.Join(r.dir, "config.json")
		r.writeFile(filepath.Join(r.env.ExeDir, "config.json"), configJSON(9101)) // a fall-through would find this
		if c.body != "" {
			r.writeFile(cfg, c.body)
		}
		code := r.run("service", "install", "--config", cfg)
		if code != 2 || !strings.Contains(r.stderr.String(), cfg) || len(r.scm.mutations()) != 0 || len(r.netsh.list()) != 0 {
			t.Errorf("%s: exit %d; want 2, stderr naming %s, no SCM change, no netsh\n%s", c.tc, code, cfg, r.dump())
		}
		r.cleanup()
	}
}

// TC-SVC-006
func TestREQ_SVC_002_ConfigLookedUpLikeServe(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	cfg := r.writeFile(filepath.Join(r.env.ExeDir, "config.json"), configJSON(9101))
	code := r.run("service", "install")
	if code != 0 || len(r.scm.created) != 1 || r.scm.created[0].BinaryPath != cmdLine(cfg) {
		t.Fatalf("exit %d; want 0 and a service whose command line names %s\ncreated %+v\n%s", code, cfg, r.scm.created, r.dump())
	}
}

// ---- REQ-SVC-003 ----

// TC-SVC-007
func TestREQ_SVC_003_CreatesServiceWithSettings(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	cfg := r.config(9101)
	code := r.run("service", "install", "--config", cfg)
	if code != 0 || len(r.scm.created) != 1 || r.scm.created[0] != wantSettings(cfg) {
		t.Fatalf("exit %d; created %+v\nwant one Create with %+v\n%s", code, r.scm.created, wantSettings(cfg), r.dump())
	}
}

// TC-SVC-008
func TestREQ_SVC_003_RelativeConfigBecomesAbsolute(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.config(9101)
	old, err := os.Getwd()
	if err != nil {
		t.Fatal(err)
	}
	defer os.Chdir(old)
	if err := os.Chdir(r.dir); err != nil {
		t.Fatal(err)
	}
	cwd, _ := os.Getwd()
	code := r.run("service", "install", "--config", "config.json")
	want := cmdLine(filepath.Join(cwd, "config.json"))
	if code != 0 || len(r.scm.created) != 1 || r.scm.created[0].BinaryPath != want {
		t.Fatalf("exit %d; created %+v\nwant command line %s\n%s", code, r.scm.created, want, r.dump())
	}
}

// ---- REQ-SVC-004 ----

// TC-SVC-009
func TestREQ_SVC_004_ExistingServiceUpdated(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists = true
	r.scm.settings = service.Settings{DisplayName: "old", BinaryPath: `"C:\old.exe" service run --config "C:\old.json"`, StartType: 3, Account: "LocalSystem"}
	cfg := r.config(9101)
	code := r.run("service", "install", "--config", cfg)
	if code != 0 || len(r.scm.created) != 0 || r.scm.has("delete") || len(r.scm.updated) != 1 || r.scm.updated[0] != wantSettings(cfg) {
		t.Fatalf("exit %d; updated %+v; want exit 0, one Update with %+v, no Create, no Delete\n%s", code, r.scm.updated, wantSettings(cfg), r.dump())
	}
}

// ---- REQ-SVC-005 ----

// TC-SVC-010
func TestREQ_SVC_005_RuleReplacedDeleteThenAdd(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	code := r.run("service", "install", "--config", r.config(9101))
	calls := r.netsh.list()
	adds := 0
	for _, c := range calls {
		if strings.HasPrefix(c, "netsh advfirewall firewall add") {
			adds++
		}
	}
	del, add := indexOf(calls, advDelete), indexOf(calls, advAdd+"9101")
	if code != 0 || del < 0 || add < 0 || del > add || adds != 1 {
		t.Fatalf("exit %d; netsh %q; want one delete name=PrinterOne, then exactly one add for 9101\n%s", code, calls, r.dump())
	}
}

// ---- REQ-SVC-006 ----

// TC-SVC-011
func TestREQ_SVC_006_StartsAndWaitsForRunning(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	start := r.clock.Now()
	code := r.run("service", "install", "--config", r.config(9101))
	waited := r.clock.Now().Sub(start)
	if code != 0 || strings.TrimSpace(r.stdout.String()) != okInstall || !r.scm.has("start") || waited < 2*time.Second {
		t.Fatalf("exit %d, waited %v of fake time; want 0, stdout %q, Start called and a wait until Running (2 s)\n%s", code, waited, okInstall, r.dump())
	}
}

// TC-SVC-012
func TestREQ_SVC_006_RunningServiceRestarted(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists, r.scm.state = true, service.Running
	r.scm.settings = wantSettings(filepath.Join(r.dir, "config.json"))
	code := r.run("service", "install", "--config", r.config(9101))
	calls := r.scm.callList()
	stop, start := indexOf(calls, "stop:running"), indexOf(calls, "start:stopped")
	if code != 0 || stop < 0 || start < stop {
		t.Fatalf("exit %d; scm %q; want Stop of the running service, then Start once it reports stopped\n%s", code, calls, r.dump())
	}
}

// TC-SVC-013
func TestREQ_SVC_006_RunningWithin30sSucceeds(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.runAfter = 29 * time.Second
	code := r.run("service", "install", "--config", r.config(9101))
	if code != 0 || strings.TrimSpace(r.stdout.String()) != okInstall {
		t.Fatalf("exit %d; want 0 and %q when Running comes after 29 s\n%s", code, okInstall, r.dump())
	}
}

// ---- REQ-SVC-007 ----

// TC-SVC-014
func TestREQ_SVC_007_RuleAddFailureRollsBack(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	const out = "The requested operation requires elevation."
	r.netsh.reply = func(args []string) netshReply {
		if len(args) > 2 && args[2] == "add" {
			return netshReply{out: out + "\r\n", err: errExit1}
		}
		return noRulesMatch(args)
	}
	code := r.run("service", "install", "--config", r.config(9101))
	if code != 1 || !r.scm.has("delete") || r.scm.has("start") || !strings.Contains(r.stderr.String(), out) {
		t.Fatalf("exit %d; want 1, the created service deleted, no Start, stderr with the netsh output\n%s", code, r.dump())
	}
}

// TC-SVC-015
func TestREQ_SVC_007_StartFailureRollsBack(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.startErr = &service.Error{Code: 1, Msg: "1053 The service did not respond to the start or control request in a timely fashion."}
	code := r.run("service", "install", "--config", r.config(9101))
	calls := r.netsh.list()
	if code != 1 || !r.scm.has("delete") || lastIndexOf(calls, advDelete) < indexOf(calls, advAdd+"9101") || indexOf(calls, advAdd+"9101") < 0 ||
		!strings.Contains(r.stderr.String(), "1053") {
		t.Fatalf("exit %d; want 1, service deleted, rule deleted after its add, stderr with 1053\n%s", code, r.dump())
	}
}

// TC-SVC-016
func TestREQ_SVC_007_ExistingServiceNotDeletedOnFailure(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists = true
	r.scm.settings = wantSettings(filepath.Join(r.dir, "config.json"))
	r.scm.startErr = &service.Error{Code: 1, Msg: "1053 The service did not respond to the start or control request in a timely fashion."}
	code := r.run("service", "install", "--config", r.config(9101))
	calls := r.netsh.list()
	if code != 1 || r.scm.has("delete") || lastIndexOf(calls, advDelete) < indexOf(calls, advAdd+"9101") {
		t.Fatalf("exit %d; want 1, the pre-existing service kept, the rule added this run removed\n%s", code, r.dump())
	}
}

// TC-SVC-017
func TestREQ_SVC_007_StartTimeoutRollsBack(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.runAfter = -1
	start := r.clock.Now()
	code := r.run("service", "install", "--config", r.config(9101))
	waited := r.clock.Now().Sub(start)
	if code != 1 || !r.scm.has("delete") || waited < 30*time.Second || waited > 32*time.Second {
		t.Fatalf("exit %d, gave up after %v of fake time; want 1, a 30 s wait (one poll of slack), service deleted\n%s", code, waited, r.dump())
	}
}

// ---- REQ-SVC-008 ----

// TC-SVC-018
func TestREQ_SVC_008_UninstallStopsWaitsDeletesRemovesRule(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists, r.scm.state, r.scm.settings = true, service.Running, wantSettings(r.config(9101))
	code := r.run("service", "uninstall")
	calls := r.scm.callList()
	if code != 0 || indexOf(calls, "stop:running") < 0 || indexOf(calls, "delete:stopped") < 0 || indexOf(r.netsh.list(), advDelete) < 0 ||
		strings.TrimSpace(r.stdout.String()) != okRemove {
		t.Fatalf("exit %d; scm %q; want Stop, Delete only once stopped, rule deleted, stdout %q\n%s", code, calls, okRemove, r.dump())
	}
}

// TC-SVC-041
func TestREQ_SVC_008_StopTimeoutGivesUpAfter30s(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists, r.scm.state, r.scm.settings = true, service.Running, wantSettings(r.config(9101))
	r.scm.stopAfter = time.Hour
	start := r.clock.Now()
	code := r.run("service", "uninstall")
	waited := r.clock.Now().Sub(start)
	if code != 1 || r.scm.has("delete") || waited < 30*time.Second || waited > 32*time.Second {
		t.Fatalf("exit %d, gave up after %v of fake time; want 1 after a 30 s wait and no Delete of a service still stopping\n%s", code, waited, r.dump())
	}
}

// TC-SVC-019
func TestREQ_SVC_008_NothingInstalledExits0(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.netsh.reply = noRulesMatch
	code := r.run("service", "uninstall")
	if code != 0 || strings.TrimSpace(r.stdout.String()) != okRemove || r.scm.has("stop") || r.scm.has("delete") || indexOf(r.netsh.list(), advDelete) < 0 {
		t.Fatalf("exit %d; want 0, %q, no Stop/Delete, the rule delete still tried\n%s", code, okRemove, r.dump())
	}
}

// TC-SVC-020
func TestREQ_SVC_008_DeleteFailureExits1(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.scm.exists, r.scm.settings = true, wantSettings(r.config(9101))
	r.scm.deleteErr = &service.Error{Code: 1, Msg: "5 Access is denied."}
	code := r.run("service", "uninstall")
	if code != 1 || !strings.Contains(r.stderr.String(), "Access is denied") || strings.Contains(r.stdout.String(), okRemove) {
		t.Fatalf("exit %d; want 1, stderr with the Windows error, no success line\n%s", code, r.dump())
	}
}

// ---- REQ-SVC-009 ----

// TC-SVC-022
func TestREQ_SVC_009_ConsoleRunExits2(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	called := false
	r.env.Service.Dispatch = func(name string, main func(service.Session)) error {
		called = true
		return service.ErrNotService
	}
	code := r.run("service", "run", "--config", r.config(9101))
	if code != 2 || strings.TrimSpace(r.stderr.String()) != "service run is started by Windows; use serve to run in a console" || !called {
		t.Fatalf("exit %d, dispatcher called %v; want 2 and the §2.3 message\n%s", code, called, r.dump())
	}
}

// ---- REQ-SVC-014 ----

// TC-SVC-032
func TestREQ_SVC_014_VistaUsesAdvfirewall(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 6
	cfg := r.config(9101)
	code := r.run("service", "install", "--config", cfg)
	install := r.netsh.list()
	want := []string{advDelete, advAdd + "9101"}
	if code != 0 || strings.Join(install, "\n") != strings.Join(want, "\n") {
		t.Fatalf("install exit %d; netsh %q, want exactly %q\n%s", code, install, want, r.dump())
	}
	r.netsh.calls = nil
	code = r.run("service", "uninstall")
	if got := r.netsh.list(); code != 0 || strings.Join(got, "\n") != advDelete {
		t.Fatalf("uninstall exit %d; netsh %q, want exactly %q\n%s", code, got, advDelete, r.dump())
	}
}

// ---- REQ-SVC-015 ----

// TC-SVC-033
func TestREQ_SVC_015_XPInstallOpensPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	code := r.run("service", "install", "--config", r.config(9101))
	calls := r.netsh.list()
	want := "netsh firewall add portopening protocol=TCP port=9101 name=PrinterOne"
	adv := false
	for _, c := range calls {
		adv = adv || strings.Contains(c, "advfirewall")
	}
	if code != 0 || indexOf(calls, want) < 0 || adv {
		t.Fatalf("exit %d; netsh %q; want %q and no advfirewall call\n%s", code, calls, want, r.dump())
	}
}

// TC-SVC-034
func TestREQ_SVC_015_XPUninstallClosesConfiguredPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	r.scm.exists, r.scm.settings = true, wantSettings(r.config(9102))
	code := r.run("service", "uninstall")
	want := "netsh firewall delete portopening protocol=TCP port=9102"
	if calls := r.netsh.list(); code != 0 || indexOf(calls, want) < 0 {
		t.Fatalf("exit %d; netsh %q; want %q\n%s", code, calls, want, r.dump())
	}
}

// TC-SVC-035
func TestREQ_SVC_015_XPUnreadableConfigWarnsExits0(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	gone := filepath.Join(r.dir, "gone", "config.json")
	r.scm.exists, r.scm.settings = true, wantSettings(gone)
	code := r.run("service", "uninstall")
	if code != 0 || !strings.Contains(r.stderr.String(), gone) || !r.scm.has("delete") {
		t.Fatalf("exit %d; want 0, a stderr warning naming %s, the service deleted\n%s", code, gone, r.dump())
	}
}

// TC-SVC-036
func TestREQ_SVC_015_XPInstallClosesOldPort(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.env.Service.OSMajor = 5
	old := r.writeFile(filepath.Join(r.dir, "old", "config.json"), configJSON(9102))
	r.scm.exists, r.scm.settings = true, wantSettings(old)
	code := r.run("service", "install", "--config", r.config(9101))
	calls := r.netsh.list()
	del := indexOf(calls, "netsh firewall delete portopening protocol=TCP port=9102")
	add := indexOf(calls, "netsh firewall add portopening protocol=TCP port=9101 name=PrinterOne")
	if code != 0 || del < 0 || add < del {
		t.Fatalf("exit %d; netsh %q; want the old port 9102 closed before 9101 is opened\n%s", code, calls, r.dump())
	}
}

// ---- REQ-SVC-016 ----

// TC-SVC-037
func TestREQ_SVC_016_NetshFailureReportsOutput(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	const out = "An unexpected error occurred."
	r.netsh.reply = func(args []string) netshReply { return netshReply{out: out + "\r\n", err: errExit1} }
	code := r.run("service", "uninstall")
	if code != 1 || !strings.Contains(r.stderr.String(), out) {
		t.Fatalf("exit %d; want 1 and stderr with the netsh output %q\n%s", code, out, r.dump())
	}
}

// TC-SVC-038
func TestREQ_SVC_016_NoRulesMatchIsNotFailure(t *testing.T) {
	r := newSvcRig(t)
	defer r.cleanup()
	r.netsh.reply = noRulesMatch
	code := r.run("service", "install", "--config", r.config(9101))
	if code != 0 || indexOf(r.netsh.list(), advAdd+"9101") < 0 || strings.TrimSpace(r.stdout.String()) != okInstall {
		t.Fatalf("exit %d; want 0: `No rules match` on removal is not a failure\n%s", code, r.dump())
	}
}
