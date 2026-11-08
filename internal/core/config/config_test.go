package config

import (
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"testing"
)

// Go 1.10: no t.TempDir / t.Cleanup (rule 02-testing).
func tempDir(t *testing.T) string {
	t.Helper()
	d, err := ioutil.TempDir("", "cfgtest")
	if err != nil {
		t.Fatal(err)
	}
	return d
}

func writeFile(t *testing.T, path, body string) string {
	t.Helper()
	if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		t.Fatal(err)
	}
	if err := ioutil.WriteFile(path, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	return path
}

func cfgWithPort(port int) string {
	return `{"version":2,"printers":[{"name":"default","printer":"P","port":` + strconv.Itoa(port) + `}]}`
}

func mustLoad(t *testing.T, l Lookup) *Config {
	t.Helper()
	c, err := Load(l)
	if err != nil {
		t.Fatalf("Load(%+v) error = %v, want nil", l, err)
	}
	if c == nil || len(c.Printers) != 1 {
		t.Fatalf("Load(%+v) = %+v, want one printer", l, c)
	}
	return c
}

func mustFail(t *testing.T, l Lookup) *Error {
	t.Helper()
	c, err := Load(l)
	if err == nil {
		t.Fatalf("Load(%+v) = %+v, nil; want *config.Error", l, c)
	}
	e, ok := err.(*Error)
	if !ok {
		t.Fatalf("Load error type %T (%v), want *config.Error", err, err)
	}
	return e
}

// TC-CFG-001
func TestREQ_CFG_001_FlagWinsOverExeAndDataDir(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	flag := writeFile(t, filepath.Join(d, "flag.json"), cfgWithPort(1111))
	writeFile(t, filepath.Join(d, "exe", "config.json"), cfgWithPort(2222))
	writeFile(t, filepath.Join(d, "data", "config.json"), cfgWithPort(3333))
	c := mustLoad(t, Lookup{Flag: flag, ExeDir: filepath.Join(d, "exe"), DataDir: filepath.Join(d, "data")})
	if c.Printers[0].Port != 1111 || c.Path != flag {
		t.Fatalf("got port %d path %q, want 1111 %q", c.Printers[0].Port, c.Path, flag)
	}
}

// TC-CFG-002
func TestREQ_CFG_001_ExeDirWinsOverDataDir(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	exe := writeFile(t, filepath.Join(d, "exe", "config.json"), cfgWithPort(2222))
	writeFile(t, filepath.Join(d, "data", "config.json"), cfgWithPort(3333))
	c := mustLoad(t, Lookup{ExeDir: filepath.Join(d, "exe"), DataDir: filepath.Join(d, "data")})
	if c.Printers[0].Port != 2222 || c.Path != exe {
		t.Fatalf("got port %d path %q, want 2222 %q", c.Printers[0].Port, c.Path, exe)
	}
}

// TC-CFG-003
func TestREQ_CFG_001_DataDirUsedWhenNothingElse(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	os.MkdirAll(filepath.Join(d, "exe"), 0755)
	data := writeFile(t, filepath.Join(d, "data", "config.json"), cfgWithPort(3333))
	c := mustLoad(t, Lookup{ExeDir: filepath.Join(d, "exe"), DataDir: filepath.Join(d, "data")})
	if c.Printers[0].Port != 3333 || c.Path != data {
		t.Fatalf("got port %d path %q, want 3333 %q", c.Printers[0].Port, c.Path, data)
	}
}

// TC-CFG-004
func TestREQ_CFG_001_MissingFlagFileDoesNotFallThrough(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	writeFile(t, filepath.Join(d, "exe", "config.json"), cfgWithPort(2222))
	missing := filepath.Join(d, "nope.json")
	e := mustFail(t, Lookup{Flag: missing, ExeDir: filepath.Join(d, "exe")})
	if !strings.Contains(e.Msg, missing) {
		t.Fatalf("Msg %q does not name %q", e.Msg, missing)
	}
}

// TC-CFG-038
func TestREQ_CFG_001_InvalidExeDirFileDoesNotFallThrough(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	exe := writeFile(t, filepath.Join(d, "exe", "config.json"), "{bad")
	writeFile(t, filepath.Join(d, "data", "config.json"), cfgWithPort(3333))
	e := mustFail(t, Lookup{ExeDir: filepath.Join(d, "exe"), DataDir: filepath.Join(d, "data")})
	if !strings.Contains(e.Msg, exe) {
		t.Fatalf("Msg %q does not name the exe-folder file %q", e.Msg, exe)
	}
}

// envOf is a case-insensitive getenv, like Windows.
func envOf(m map[string]string) func(string) string {
	return func(k string) string {
		for n, v := range m {
			if strings.EqualFold(n, k) {
				return v
			}
		}
		return ""
	}
}

// TC-CFG-042, TC-CFG-044
func TestREQ_CFG_001_UnreadableOrEmptyExeDirFileDoesNotFallThrough(t *testing.T) {
	for _, tc := range []string{"TC-CFG-042", "TC-CFG-044"} {
		t.Run(tc, func(t *testing.T) {
			d := tempDir(t)
			defer os.RemoveAll(d)
			exe := filepath.Join(d, "exe", "config.json")
			if tc == "TC-CFG-042" {
				os.MkdirAll(exe, 0755) // exists but cannot be read as a file
			} else {
				writeFile(t, exe, "") // 0 bytes
			}
			writeFile(t, filepath.Join(d, "data", "config.json"), cfgWithPort(3333))
			e := mustFail(t, Lookup{ExeDir: filepath.Join(d, "exe"), DataDir: filepath.Join(d, "data")})
			if !strings.Contains(e.Msg, exe) {
				t.Fatalf("Msg %q does not name the exe-folder file %q", e.Msg, exe)
			}
		})
	}
}

// TC-CFG-007
func TestREQ_CFG_002_DataDirFromProgramData(t *testing.T) {
	got := DataDir(envOf(map[string]string{"ProgramData": "/pd", "ALLUSERSPROFILE": "/au"}))
	if want := filepath.Join("/pd", "PrinterOne"); got != want {
		t.Fatalf("DataDir = %q, want %q", got, want)
	}
}

// TC-CFG-008
func TestREQ_CFG_002_DataDirXPFallback(t *testing.T) {
	got := DataDir(envOf(map[string]string{"ALLUSERSPROFILE": "/au"}))
	if want := filepath.Join("/au", "Application Data", "PrinterOne"); got != want {
		t.Fatalf("DataDir = %q, want %q", got, want)
	}
}

// TC-CFG-009..013, TC-CFG-043
func TestREQ_CFG_003_ValidFileLoadsWithDefaults(t *testing.T) {
	cases := []struct {
		tc, body string
		want     Printer
	}{
		{"TC-CFG-009", `{"version":2,"printers":[{"name":"front","printer":"Canon MF220 Series","port":9101}]}`,
			Printer{"front", "Canon MF220 Series", 9101}},
		{"TC-CFG-010", `{"version":2,"printers":[{"printer":"P","port":9101}]}`, Printer{"default", "P", 9101}},
		{"TC-CFG-011", `{"version":2,"printers":[{"name":"","printer":"P","port":9101}]}`, Printer{"default", "P", 9101}},
		{"TC-CFG-012", `{"version":2,"printers":[{"name":"n","printer":"P"}]}`, Printer{"n", "P", 9100}},
		{"TC-CFG-013", `{"version":2,"theme":"dark","printers":[{"name":"n","printer":"P","port":9101,"color":true}]}`,
			Printer{"n", "P", 9101}},
		{"TC-CFG-043", `{"version":2,"printers":[{"name":"n","printer":"P","Port":"x","Printer":"B"}]}`,
			Printer{"n", "P", 9100}}, // keys are case-sensitive: "Port"/"Printer" are unknown keys
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) {
			d := tempDir(t)
			defer os.RemoveAll(d)
			f := writeFile(t, filepath.Join(d, "c.json"), c.body)
			got := mustLoad(t, Lookup{Flag: f})
			if got.Printers[0] != c.want || got.Path != f {
				t.Fatalf("got %+v path %q, want %+v %q", got.Printers[0], got.Path, c.want, f)
			}
		})
	}
}

func TestREQ_CFG_003_LeadingUTF8BOMIgnored(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	f := writeFile(t, filepath.Join(d, "bom.json"), "\xef\xbb\xbf"+cfgWithPort(9101))
	c := mustLoad(t, Lookup{Flag: f})
	if c.Printers[0].Port != 9101 {
		t.Fatalf("port = %d, want 9101", c.Printers[0].Port)
	}
}

func TestREQ_CFG_005_TopLevelMustBeObject(t *testing.T) {
	for _, body := range []string{`[]`, `"x"`, `3`} {
		d := tempDir(t)
		f := writeFile(t, filepath.Join(d, "bad.json"), body)
		e := mustFail(t, Lookup{Flag: f})
		if !strings.Contains(e.Msg, f) || !strings.Contains(e.Msg, "top level must be a JSON object") {
			t.Fatalf("error %q, want path and top level message", e.Msg)
		}
		os.RemoveAll(d)
	}
}

// TC-CFG-015..017
func TestREQ_CFG_004_UnreadableOrBadJSONNamesPathAndPosition(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)

	t.Run("TC-CFG-015", func(t *testing.T) {
		dir := filepath.Join(d, "dir")
		os.MkdirAll(dir, 0755)
		_, readErr := ioutil.ReadFile(dir)
		pe, ok := readErr.(*os.PathError)
		if !ok {
			t.Fatalf("reading a directory gave %T %v, test needs an OS error", readErr, readErr)
		}
		e := mustFail(t, Lookup{Flag: dir})
		if !strings.Contains(e.Msg, dir) || !strings.Contains(e.Msg, pe.Err.Error()) {
			t.Fatalf("Msg %q, want path %q and OS error %q", e.Msg, dir, pe.Err.Error())
		}
	})
	t.Run("TC-CFG-016", func(t *testing.T) {
		f := writeFile(t, filepath.Join(d, "bad.json"), "{\n  \"version\": 2,\n  \"printers\": [,]\n}")
		e := mustFail(t, Lookup{Flag: f})
		if !strings.Contains(e.Msg, f) || !regexp.MustCompile(`\b3:1[5-7]\b`).MatchString(e.Msg) {
			t.Fatalf("Msg %q, want path %q and line:column 3:15..3:17 (the comma, 0- or 1-based, or just after it)", e.Msg, f)
		}
	})
	t.Run("TC-CFG-017", func(t *testing.T) {
		f := writeFile(t, filepath.Join(d, "trunc.json"), "{\n  \"version\": 2,") // no trailing newline: ends on line 2
		e := mustFail(t, Lookup{Flag: f})
		if !strings.Contains(e.Msg, f) || !regexp.MustCompile(`\b2:[0-9]+\b`).MatchString(e.Msg) {
			t.Fatalf("Msg %q, want path %q and line:column 2:<col>", e.Msg, f)
		}
	})
}

// TC-CFG-014
func TestREQ_CFG_005_PortBoundsAccepted(t *testing.T) {
	for _, p := range []int{1, 65535} {
		d := tempDir(t)
		f := writeFile(t, filepath.Join(d, "c.json"), cfgWithPort(p))
		c := mustLoad(t, Lookup{Flag: f})
		os.RemoveAll(d)
		if c.Printers[0].Port != p {
			t.Fatalf("port %d loaded as %d", p, c.Printers[0].Port)
		}
	}
}

// TC-CFG-019..034, TC-CFG-039..041
func TestREQ_CFG_005_InvalidFieldNamesField(t *testing.T) {
	entry := `{"printer":"P","port":9100}`
	wrap := func(e string) string { return `{"version":2,"printers":[` + e + `]}` }
	cases := []struct{ tc, body, field string }{
		{"TC-CFG-019", `{"printers":[` + entry + `]}`, "version"},
		{"TC-CFG-020", `{"version":1,"printers":[` + entry + `]}`, "version"},
		{"TC-CFG-021", `{"version":"2","printers":[` + entry + `]}`, "version"},
		{"TC-CFG-039", `{"version":3,"printers":[` + entry + `]}`, "version"},
		{"TC-CFG-022", `{"version":2}`, "printers"},
		{"TC-CFG-023", `{"version":2,"printers":{"printer":"P"}}`, "printers"},
		{"TC-CFG-024", `{"version":2,"printers":[]}`, "printers"},
		// TC-CFG-025 (two printers) moved to F-003: several printers are valid there (TC-MULTI-*).
		{"TC-CFG-026", wrap(`"Canon"`), "printers[0]"},
		{"TC-CFG-040", wrap(`null`), "printers[0]"},
		{"TC-CFG-027", wrap(`{"port":9100}`), "printers[0].printer"},
		{"TC-CFG-028", wrap(`{"printer":"","port":9100}`), "printers[0].printer"},
		{"TC-CFG-029", wrap(`{"printer":5,"port":9100}`), "printers[0].printer"},
		{"TC-CFG-030", wrap(`{"printer":"P","port":"9100"}`), "printers[0].port"},
		{"TC-CFG-031", wrap(`{"printer":"P","port":9100.5}`), "printers[0].port"},
		{"TC-CFG-032", wrap(`{"printer":"P","port":0}`), "printers[0].port"},
		{"TC-CFG-033", wrap(`{"printer":"P","port":65536}`), "printers[0].port"},
		{"TC-CFG-041", wrap(`{"printer":"P","port":74636}`), "printers[0].port"},
		{"TC-CFG-034", wrap(`{"printer":"P","port":-1}`), "printers[0].port"},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) {
			d := tempDir(t)
			defer os.RemoveAll(d)
			f := writeFile(t, filepath.Join(d, "c.json"), c.body)
			e := mustFail(t, Lookup{Flag: f})
			if !strings.Contains(e.Msg, f) || !namesField(e.Msg, c.field) {
				t.Fatalf("Msg %q, want path %q and exactly field %q", e.Msg, f, c.field)
			}
		})
	}
}

// namesField reports whether msg names field and not a longer field path that starts with it
// ("printers" must not match a message about "printers[0].port").
func namesField(msg, field string) bool {
	i := strings.Index(msg, field)
	for i >= 0 {
		end := i + len(field)
		if end == len(msg) || (msg[end] != '[' && msg[end] != '.') {
			return true
		}
		j := strings.Index(msg[end:], field)
		if j < 0 {
			return false
		}
		i = end + j
	}
	return false
}

// TC-CFG-036
func TestREQ_CFG_006_NoConfigNamesDataDirAndHint(t *testing.T) {
	d := tempDir(t)
	defer os.RemoveAll(d)
	exe, data := filepath.Join(d, "exe"), filepath.Join(d, "data")
	os.MkdirAll(exe, 0755)
	e := mustFail(t, Lookup{ExeDir: exe, DataDir: data})
	if !strings.Contains(e.Msg, data) || !strings.Contains(e.Msg, "printerone printers") {
		t.Fatalf("Msg %q, want data dir %q and hint `printerone printers`", e.Msg, data)
	}
}
