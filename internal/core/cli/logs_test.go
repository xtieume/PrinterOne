package cli

import (
	"encoding/json"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

// Tests of the log file in `serve` and the `logs` command (spec docs/specs/2026-11-07-f002-logs.md §2).
// TC ids: .testcases/runs/F-002-r1/testcase/testcases.md and testcases.md.

func (r *rig) defaultLog() string { return filepath.Join(r.pd, "PrinterOne", "logs", "printerone.log") }

func jsonStr(s string) string { b, _ := json.Marshal(s); return string(b) }

// cfgLog writes C-std plus the given `log` section and returns its path.
func (r *rig) cfgLog(logJSON string) string {
	body := strings.TrimSuffix(configJSON(r.port), "}") + `,"log":` + logJSON + `}`
	return r.write(filepath.Join(r.dir, "cfglog.json"), body)
}

func readFile(t *testing.T, path string) string {
	t.Helper()
	b, err := ioutil.ReadFile(path)
	if err != nil {
		t.Fatalf("read %s: %v", path, err)
	}
	return string(b)
}

// serveOnce runs serve with the config, waits until listening, optionally sends one job and waits for its line,
// then stops with Ctrl+C.
func (r *rig) serveOnce(cfg string, job string) {
	r.t.Helper()
	r.start("serve", "--config", cfg)
	r.waitReady()
	if job != "" {
		c := r.send(job)
		c.Close()
		if l := r.waitConnLines(1); len(l) != 1 {
			r.t.Fatalf("no connection line on stdout %q", r.stdout.String())
		}
	}
	r.stop()
}

// ---- REQ-LOGF-001 / 002 ----

// TC-LOGF-001
func TestREQ_LOGF_001_ServeWritesToDataDirLogs(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.serveOnce(r.cfg, "")
	if b, err := ioutil.ReadFile(r.defaultLog()); err != nil || !strings.Contains(string(b), "Listening on") {
		t.Fatalf("%s: %q, %v; want the start-up lines", r.defaultLog(), b, err)
	}
}

// TC-LOGF-005
func TestREQ_LOGF_002_ServeWritesToLogDir(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "L")
	r.serveOnce(r.cfgLog(`{"dir":`+jsonStr(dir)+`}`), "")
	b, err := ioutil.ReadFile(filepath.Join(dir, "printerone.log"))
	if err != nil || !strings.Contains(string(b), "Listening on") {
		t.Fatalf("log.dir file: %q, %v; want the start-up lines", b, err)
	}
	if _, err := os.Stat(r.defaultLog()); err == nil {
		t.Fatalf("%s written although log.dir is set", r.defaultLog())
	}
}

// ---- REQ-LOGF-003 ----

// TC-LOGF-017
func TestREQ_LOGF_003_ServeInvalidKeepDaysExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	code := r.run("serve", "--config", r.cfgLog(`{"keep_days":0}`))
	r.noReady()
	if code != 2 || !strings.Contains(r.stderr.String(), "log.keep_days") {
		t.Fatalf("exit %d, stderr %q; want 2 naming log.keep_days", code, r.stderr.String())
	}
}

// ---- REQ-LOGF-004 ----

// TC-LOGF-018
func TestREQ_LOGF_004_SameLinesInFileAndConsole(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.serveOnce(r.cfg, "hello")
	conn := r.connLines()[0]
	if !strings.HasSuffix(conn, ` 5B "`+printerName+`" printed`) {
		t.Fatalf("connection line %q", conn)
	}
	file := readFileOrEmpty(r.defaultLog())
	for _, want := range []string{"Listening on", "Printer: " + printerName, conn} {
		if !strings.Contains(r.stdout.String(), want) || !strings.Contains(file, want) {
			t.Fatalf("%q missing; stdout %q, file %q", want, r.stdout.String(), file)
		}
	}
}

func readFileOrEmpty(path string) string { b, _ := ioutil.ReadFile(path); return string(b) }

// TC-LOGF-019
func TestREQ_LOGF_004_ServeAppendsToExistingFile(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.write(r.defaultLog(), "earlier line\n")
	r.serveOnce(r.cfg, "")
	if got := readFile(t, r.defaultLog()); !strings.HasPrefix(got, "earlier line\n") || !strings.Contains(got, "Listening on") {
		t.Fatalf("file %q; want earlier line kept and start-up lines appended", got)
	}
}

// ---- REQ-LOGF-006 ----

// TC-LOGF-024
func TestREQ_LOGF_006_FileReadableWhileServing(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	f, err := os.Open(r.defaultLog())
	if err != nil {
		r.stop()
		t.Fatalf("open while serving: %v", err)
	}
	defer f.Close()
	b, _ := ioutil.ReadAll(f)
	c := r.send("hello")
	c.Close()
	lines := r.waitConnLines(1)
	r.stop()
	if !strings.Contains(string(b), "Listening on") || len(lines) != 1 || !strings.Contains(readFileOrEmpty(r.defaultLog()), lines[0]) {
		t.Fatalf("read while open %q; file after job %q; want start-up lines, then the connection line", b, readFileOrEmpty(r.defaultLog()))
	}
}

// ---- REQ-LOGF-008 ----

// TC-LOGF-036
func TestREQ_LOGF_008_ServeUsesConfiguredKeepDays(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "L")
	now := time.Now()
	old := filepath.Join(dir, "printerone-"+now.AddDate(0, 0, -2).Format("2006-01-02")+".log")
	kept := filepath.Join(dir, "printerone-"+now.AddDate(0, 0, -1).Format("2006-01-02")+".log")
	r.write(old, "x\n")
	r.write(kept, "x\n")
	r.serveOnce(r.cfgLog(`{"dir":`+jsonStr(dir)+`,"keep_days":1}`), "")
	if _, err := os.Stat(old); err == nil {
		t.Fatalf("%s kept; keep_days 1 deletes it at start-up", old)
	}
	if _, err := os.Stat(kept); err != nil {
		t.Fatalf("%s deleted; keep_days 1 keeps yesterday", kept)
	}
}

// ---- REQ-LOGF-009 ----

// TC-LOGF-038
func TestREQ_LOGF_009_ServeKeepsServingWhenFolderCannotBeCreated(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	blocker := r.write(filepath.Join(r.dir, "f"), "a file, not a folder\n")
	dir := filepath.Join(blocker, "sub")
	r.serveOnce(r.cfgLog(`{"dir":`+jsonStr(dir)+`}`), "hello")
	if docs := r.sp.waitDone(1, 3*time.Second); len(docs) != 1 {
		t.Fatalf("%d jobs printed; want 1", len(docs))
	}
	warn := strings.TrimSpace(r.stderr.String())
	if warn == "" || strings.Count(warn, "\n") != 0 || !strings.Contains(warn, dir) {
		t.Fatalf("stderr %q; want exactly one warning naming %s", r.stderr.String(), dir)
	}
}

// ---- REQ-LOGF-010 ----

// TC-LOGF-040
func TestREQ_LOGF_010_LogsWithoutConfigPrintsDefaultPath(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	code := r.run("logs")
	if code != 0 || r.stdout.String() != r.defaultLog()+"\n" {
		t.Fatalf("exit %d, stdout %q; want 0 and %q", code, r.stdout.String(), r.defaultLog()+"\n")
	}
}

// TC-LOGF-041
func TestREQ_LOGF_010_LogsWithConfigLogDir(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "L")
	code := r.run("logs", "--config", r.cfgLog(`{"dir":`+jsonStr(dir)+`}`))
	if want := filepath.Join(dir, "printerone.log") + "\n"; code != 0 || r.stdout.String() != want {
		t.Fatalf("exit %d, stdout %q; want 0 and %q", code, r.stdout.String(), want)
	}
}

// TC-LOGF-042
func TestREQ_LOGF_010_LogsIgnoresInvalidPrinters(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "L")
	cfg := r.write(filepath.Join(r.dir, "np.json"), `{"version":2,"printers":[],"log":{"dir":`+jsonStr(dir)+`}}`)
	code := r.run("logs", "--config", cfg)
	if want := filepath.Join(dir, "printerone.log") + "\n"; code != 0 || r.stdout.String() != want {
		t.Fatalf("exit %d, stdout %q, stderr %q; want 0 and %q", code, r.stdout.String(), r.stderr.String(), want)
	}
}

// TC-LOGF-043
func TestREQ_LOGF_010_LogsInvalidLogSectionExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	code := r.run("logs", "--config", r.cfgLog(`{"keep_days":0}`))
	if code != 2 || r.stdout.String() != "" || !strings.Contains(r.stderr.String(), "log.keep_days") {
		t.Fatalf("exit %d, stdout %q, stderr %q; want 2 naming log.keep_days", code, r.stdout.String(), r.stderr.String())
	}
}

// TC-LOGF-044
func TestREQ_LOGF_010_LogsFollowsLookupOrder(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "E")
	r.write(filepath.Join(r.exe, "config.json"), `{"version":2,"printers":[],"log":{"dir":`+jsonStr(dir)+`}}`)
	code := r.run("logs")
	if want := filepath.Join(dir, "printerone.log") + "\n"; code != 0 || r.stdout.String() != want {
		t.Fatalf("exit %d, stdout %q; want 0 and %q", code, r.stdout.String(), want)
	}
}

// TC-LOGF-045
func TestREQ_LOGF_010_LogsBadArgumentsPrintUsage(t *testing.T) {
	for _, args := range [][]string{{"logs", "extra"}, {"logs", "--config"}} {
		r := newRig(t, "main")
		code := r.run(args...)
		r.cleanup()
		if code != 2 || !strings.Contains(r.stderr.String(), "usage:") {
			t.Fatalf("%q: exit %d, stderr %q; want 2 with usage", args, code, r.stderr.String())
		}
	}
}

// TC-LOGF-054
func TestREQ_LOGF_010_LogsPrintsAbsolutePathForRelativeDir(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	cfg := r.write(filepath.Join(r.dir, "cfgrel", "c.json"), `{"version":2,"printers":[],"log":{"dir":"rel"}}`)
	code := r.run("logs", "--config", cfg)
	if want := filepath.Join(r.dir, "cfgrel", "rel", "printerone.log") + "\n"; code != 0 || r.stdout.String() != want {
		t.Fatalf("exit %d, stdout %q; want 0 and %q", code, r.stdout.String(), want)
	}
}

// TC-LOGF-046
func TestREQ_LOGF_010_LogsMissingConfigFileExits2(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	code := r.run("logs", "--config", filepath.Join(r.dir, "nope.json"))
	if code != 2 || r.stdout.String() != "" {
		t.Fatalf("exit %d, stdout %q; want 2 and nothing printed", code, r.stdout.String())
	}
}

// ---- REQ-LOGF-011 ----

// TC-LOGF-047
func TestREQ_LOGF_011_ServeFirstLineNamesLogFile(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.serveOnce(r.cfg, "")
	want := "Log file: " + r.defaultLog()
	first := strings.SplitN(r.stdout.String(), "\n", 2)[0]
	fileFirst := strings.SplitN(readFileOrEmpty(r.defaultLog()), "\n", 2)[0]
	if first != want || fileFirst != want {
		t.Fatalf("first stdout line %q, first file line %q; want %q", first, fileFirst, want)
	}
}

// TC-LOGF-060
func TestREQ_LOGF_004_FileGetsLineWhenStdoutFails(t *testing.T) {
	r := newRig(t, "main")
	defer r.cleanup()
	r.env.Stdout = failingWriter{}
	r.start("serve", "--config", r.cfg)
	r.waitReady()
	c := r.send("hello")
	c.Close()
	time.Sleep(100 * time.Millisecond)
	r.stop()
	if got := readFileOrEmpty(r.defaultLog()); !strings.Contains(got, "Listening on") || !strings.Contains(got, " printed") {
		t.Fatalf("log file missing startup or connection line: %q", got)
	}
}

type failingWriter struct{}

func (failingWriter) Write([]byte) (int, error) { return 0, os.ErrPermission }
