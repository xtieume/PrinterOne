package config

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// Tests of the optional `log` section (spec docs/specs/2026-11-07-f002-logs.md §2).
// TC ids: .testcases/runs/F-002-r1/testcase/testcases.md and testcases.md.

const stdPrinters = `"printers":[{"name":"front","printer":"P","port":9100}]`

func cfgWithLog(logJSON string) string {
	return `{"version":2,` + stdPrinters + `,"log":` + logJSON + `}`
}

func jsonString(s string) string {
	b, _ := json.Marshal(s)
	return string(b)
}

func loadLogCfg(t *testing.T, body string) (*Config, error, string) {
	t.Helper()
	dir := tempDir(t)
	p := writeFile(t, filepath.Join(dir, "c.json"), body)
	c, err := Load(Lookup{Flag: p})
	return c, err, dir
}

// TC-LOGF-004
func TestREQ_LOGF_001_KeepDaysOnlyKeepsDefaultFolder(t *testing.T) {
	c, err, dir := loadLogCfg(t, cfgWithLog(`{"keep_days":7}`))
	defer os.RemoveAll(dir)
	if err != nil || c.Log.Dir != "" || c.Log.KeepDays != 7 {
		t.Fatalf("Load = %+v, %v; want Log{Dir:\"\", KeepDays:7}", c, err)
	}
}

// TC-LOGF-006
func TestREQ_LOGF_002_LoadReturnsConfiguredDir(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	want := filepath.Join(dir, "L")
	p := writeFile(t, filepath.Join(dir, "c.json"), cfgWithLog(`{"dir":`+jsonString(want)+`}`))
	c, err := Load(Lookup{Flag: p})
	if err != nil || c.Log.Dir != want || c.Log.KeepDays != 30 {
		t.Fatalf("Load = %+v, %v; want Log{Dir:%q, KeepDays:30}", c, err, want)
	}
}

// TC-LOGF-007
func TestREQ_LOGF_002_RelativeDirResolvesAgainstConfigFolder(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	p := writeFile(t, filepath.Join(dir, "cfgdir", "c.json"), cfgWithLog(`{"dir":"rel"}`))
	c, err := Load(Lookup{Flag: p})
	want := filepath.Join(dir, "cfgdir", "rel")
	if err != nil || c.Log.Dir != want {
		t.Fatalf("Load = %+v, %v; want Log.Dir %q", c, err, want)
	}
}

// TC-LOGF-008
func TestREQ_LOGF_003_MissingLogSectionIsValid(t *testing.T) {
	c, err, dir := loadLogCfg(t, `{"version":2,`+stdPrinters+`}`)
	defer os.RemoveAll(dir)
	if err != nil || c.Log.Dir != "" || c.Log.KeepDays != 30 {
		t.Fatalf("Load = %+v, %v; want no error and Log{Dir:\"\", KeepDays:30}", c, err)
	}
}

// TC-LOGF-009 … TC-LOGF-012, TC-LOGF-014 … TC-LOGF-016
func TestREQ_LOGF_003_InvalidLogFieldNamesField(t *testing.T) {
	cases := []struct{ tc, log, field string }{
		{"TC-LOGF-009", `{"dir":""}`, "log.dir"},
		{"TC-LOGF-010", `{"dir":5}`, "log.dir"},
		{"TC-LOGF-011", `{"keep_days":0}`, "log.keep_days"},
		{"TC-LOGF-012", `{"keep_days":3651}`, "log.keep_days"},
		{"TC-LOGF-014", `{"keep_days":1.5}`, "log.keep_days"},
		{"TC-LOGF-015", `{"keep_days":"30"}`, "log.keep_days"},
		{"TC-LOGF-016", `"x"`, "log"},
		{"TC-LOGF-016", `[]`, "log"},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) {
			_, err, dir := loadLogCfg(t, cfgWithLog(c.log))
			defer os.RemoveAll(dir)
			e, ok := err.(*Error)
			if !ok || !strings.Contains(e.Msg, "invalid field "+c.field) {
				t.Fatalf("log %s: Load error %v; want *Error naming %s", c.log, err, c.field)
			}
		})
	}
}

// TC-LOGF-013
func TestREQ_LOGF_003_KeepDaysBoundsValid(t *testing.T) {
	for _, n := range []int{1, 3650} {
		c, err, dir := loadLogCfg(t, cfgWithLog(`{"keep_days":`+jsonInt(n)+`}`))
		os.RemoveAll(dir)
		if err != nil || c.Log.KeepDays != n {
			t.Fatalf("keep_days %d: Load = %+v, %v; want valid", n, c, err)
		}
	}
}

func jsonInt(n int) string { b, _ := json.Marshal(n); return string(b) }

// ---- LoadLog (`printerone logs`, REQ-LOGF-010) ----

// TC-LOGF-042 (config level)
func TestREQ_LOGF_010_LoadLogIgnoresPrinters(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	want := filepath.Join(dir, "L")
	p := writeFile(t, filepath.Join(dir, "c.json"), `{"version":2,"printers":[],"log":{"dir":`+jsonString(want)+`}}`)
	l, err := LoadLog(Lookup{Flag: p})
	if err != nil || l.Dir != want || l.KeepDays != 30 {
		t.Fatalf("LoadLog = %+v, %v; want Dir %q KeepDays 30", l, err, want)
	}
}

// TC-LOGF-040 (config level)
func TestREQ_LOGF_010_LoadLogWithoutConfigReturnsDefaults(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	l, err := LoadLog(Lookup{ExeDir: filepath.Join(dir, "exe"), DataDir: filepath.Join(dir, "pd")})
	if err != nil || l.Dir != "" || l.KeepDays != 30 {
		t.Fatalf("LoadLog = %+v, %v; want defaults, no error", l, err)
	}
}
