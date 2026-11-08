package service

import (
	"encoding/json"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

// Tests of `service run` writing the F-002 log file (spec docs/specs/2026-11-07-f002-logs.md §2).
// TC ids: .testcases/runs/F-002-r1/testcase/testcases.md and testcases.md.

func (r *runRig) writeConfigWithLog(logJSON string) {
	r.t.Helper()
	body := `{"version":2,"printers":[{"name":"default","printer":"` + testPrinter + `","port":` + strconv.Itoa(r.port) + `}],"log":` + logJSON + `}`
	if err := ioutil.WriteFile(r.opts.ConfigPath, []byte(body), 0644); err != nil {
		r.t.Fatal(err)
	}
}

func quote(s string) string { b, _ := json.Marshal(s); return string(b) }

func readOrEmpty(path string) string { b, _ := ioutil.ReadFile(path); return string(b) }

func waitContains(path, want string, d time.Duration) string {
	for deadline := time.Now().Add(d); time.Now().Before(deadline); time.Sleep(20 * time.Millisecond) {
		if s := readOrEmpty(path); strings.Contains(s, want) {
			return s
		}
	}
	return readOrEmpty(path)
}

// TC-LOGF-021, TC-LOGF-050
func TestREQ_LOGF_005_ServiceRunWritesToLogDir(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "S")
	r.writeConfigWithLog(`{"dir":` + quote(dir) + `}`)
	r.start()
	r.waitReady()
	r.send("hello")
	path := filepath.Join(dir, "printerone.log")
	got := waitContains(path, `5B "`+testPrinter+`" printed`, 3*time.Second)
	if !strings.Contains(got, "Listening on") || !strings.Contains(got, `5B "`+testPrinter+`" printed`) {
		t.Fatalf("%s = %q; want start-up lines and the connection line", path, got)
	}
	if strings.Contains(r.log(), "Listening on") {
		t.Fatalf("default log %q has start-up lines although log.dir is set", r.log())
	}
}

// TC-LOGF-022
func TestREQ_LOGF_005_InvalidKeepDaysLoggedToDefaultFolder(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	bad := filepath.Join(r.dir, "S")
	r.writeConfigWithLog(`{"dir":` + quote(bad) + `,"keep_days":0}`)
	r.start()
	select {
	case <-r.done:
	case <-time.After(5 * time.Second):
		t.Fatal("Run still running with an invalid log section")
	}
	if _, err := os.Stat(bad); err == nil || r.last() != (report{Stopped, 2}) || !strings.Contains(r.log(), "log.keep_days") {
		t.Fatalf("reports %v, default log %q, %s created %v; want Stopped 2, the error naming log.keep_days in the default log", r.session.list(), r.log(), bad, err == nil)
	}
}

// TC-LOGF-055
func TestREQ_LOGF_005_PrinterErrorLoggedInValidLogDir(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "S")
	body := `{"version":2,"printers":[],"log":{"dir":` + quote(dir) + `}}`
	if err := ioutil.WriteFile(r.opts.ConfigPath, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	r.start()
	select {
	case <-r.done:
	case <-time.After(5 * time.Second):
		t.Fatal("Run still running with an invalid printer list")
	}
	got := readOrEmpty(filepath.Join(dir, "printerone.log"))
	if r.last() != (report{Stopped, 2}) || !strings.Contains(got, "printers") {
		t.Fatalf("reports %v, %s/printerone.log %q; want Stopped 2 and the error there", r.session.list(), dir, got)
	}
}

// TC-LOGF-048
func TestREQ_LOGF_011_ServiceRunFirstLineNamesLogFile(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	dir := filepath.Join(r.dir, "S")
	r.writeConfigWithLog(`{"dir":` + quote(dir) + `}`)
	r.start()
	r.waitReady()
	path := filepath.Join(dir, "printerone.log")
	want := "Log file: " + path
	if first := strings.SplitN(readOrEmpty(path), "\n", 2)[0]; first != want {
		t.Fatalf("first line %q; want %q", first, want)
	}
}
