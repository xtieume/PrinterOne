package webui

// F-005 cases added by the second-pass review (TC-UI-076 … TC-UI-089 except 083).

import (
	"errors"
	"net"
	"net/http"
	"net/http/httptest"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/config"
)

// TC-UI-076
func TestREQ_UI_013_GetNeverChangesState(t *testing.T) {
	sink := newJobSink(t)
	defer sink.ln.Close()
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, sink.cfg())
	before := readFile(t, u.path)
	for _, path := range []string{
		"/printers?token=" + testToken + "&action=remove&index=0",
		"/settings?token=" + testToken + "&keep_days=9&ui_port=9180",
		"/test?token=" + testToken + "&name=A",
	} {
		get(u.srv, path)
	}
	if readFile(t, u.path) != before || u.applied.n() != 0 {
		t.Fatal("a GET with form fields changed the config")
	}
	if job, ok := sink.wait(500 * time.Millisecond); ok {
		t.Fatalf("a GET sent a test job: %q", job)
	}
	// the pages themselves must still render, so the red/green signal is not a 404 everywhere
	if rec := get(u.srv, "/printers"); rec.Code != 200 || !strings.Contains(rec.Body.String(), "9101") {
		t.Fatalf("GET /printers = %d", rec.Code)
	}
}

// TC-UI-077
func TestREQ_UI_012_SettingsValidateUI(t *testing.T) {
	cases := []struct {
		kv    []string
		field string
	}{
		{[]string{"keep_days", "30", "ui_port", "9180", "remote", "on", "password", ""}, "ui.password"},
		{[]string{"keep_days", "30", "ui_port", "9180", "remote", "on", "password", "1234567"}, "ui.password"},
		{[]string{"keep_days", "30", "ui_port", "9101"}, "ui.port"},
	}
	for _, c := range cases {
		dir, done := tmpDir(t)
		u := newFileUI(t, dir, oneCfg)
		before := readFile(t, u.path)
		rec := postForm(u.srv, "/settings", vals(append([]string{"token", testToken}, c.kv...)...))
		refused(t, u, before, rec, c.field)
		done()
	}
}

// TC-UI-078
func TestREQ_UI_007_PrinterPortEqualToUIPortRefused(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9180"))
	refused(t, u, before, rec, "ui.port")
}

// TC-UI-079
func TestREQ_UI_012_UntickedRemoteWritesFalse(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101}],`+
		`"ui":{"remote":true,"password":"s3cretpass"}}`)
	rec := postForm(u.srv, "/settings", vals("token", testToken, "keep_days", "30", "ui_port", "9180", "password", ""))
	mustRedirectPrefix(t, rec, "/settings")
	ui, _ := rawCfg(t, u.path)["ui"].(map[string]interface{})
	if ui == nil || ui["remote"] != false {
		t.Fatalf("unticked remote saved as %s", readFile(t, u.path))
	}
}

func mustRedirectPrefix(t *testing.T, rec *httptest.ResponseRecorder, prefix string) {
	if rec.Code != http.StatusSeeOther || !strings.HasPrefix(rec.Header().Get("Location"), prefix) {
		t.Fatalf("want 303 to %s…, got %d %q", prefix, rec.Code, rec.Header().Get("Location"))
	}
}

// TC-UI-080
func TestREQ_UI_008_ConcurrentSavesLeaveValidFile(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	var wg sync.WaitGroup
	codes := make([]int, 10)
	for i := 0; i < 10; i++ {
		wg.Add(1)
		go func(i int) {
			defer wg.Done()
			rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
				"printer", "Canon", "port", strconv.Itoa(9111+i)))
			codes[i] = rec.Code
		}(i)
	}
	wg.Wait()
	for i, c := range codes {
		if c != http.StatusSeeOther {
			t.Errorf("save %d answered %d, want 303", i, c)
		}
	}
	c := loadCfg(t, u.path)
	if p := c.Printers[0].Port; p < 9111 || p > 9120 {
		t.Fatalf("final port %d not one of the saves", p)
	}
	mustOnlyConfig(t, dir)
}

// TC-UI-081
func TestREQ_UI_012_SettingsShowCurrentValues(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	logs := filepath.Join(dir, "logsX")
	u := newFileUI(t, dir, `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101}],`+
		`"log":{"dir":"`+filepath.ToSlash(logs)+`","keep_days":12},"ui":{"port":9182}}`)
	body := get(u.srv, "/settings").Body.String()
	for _, w := range []string{`value="12"`, `value="9182"`, filepath.ToSlash(logs)} {
		if !strings.Contains(filepath.ToSlash(body), w) {
			t.Errorf("settings page lacks %q", w)
		}
	}
}

// TC-UI-082
func TestREQ_UI_009_ApplyFailureShownAfterSave(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, oneCfg)
	s := New(Deps{Config: loadCfg(t, p), ConfigPath: p, Token: testToken,
		Apply: func(*config.Config) error { return errors.New("port 9105 is already in use") }})
	rec := postForm(s, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9105"))
	if rec.Code == http.StatusSeeOther || !strings.Contains(rec.Body.String(), "port 9105 is already in use") {
		t.Fatalf("Apply failure answered %d without the error on the page", rec.Code)
	}
	if loadCfg(t, p).Printers[0].Port != 9105 {
		t.Fatal("the valid save was not kept")
	}
}

// TC-UI-085
func TestREQ_UI_006_InstalledFailureKeepsTypedName(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, oneCfg)
	s := New(Deps{Config: loadCfg(t, p), ConfigPath: p, Token: testToken,
		Installed: func() ([]string, error) { return nil, errors.New("spooler down") }})
	rec := get(s, "/printers")
	if rec.Code != 200 || !strings.Contains(rec.Body.String(), `name="printer_text"`) {
		t.Fatalf("GET /printers with spooler down = %d, want 200 with a typed-name input", rec.Code)
	}
}

// TC-UI-086
func TestREQ_UI_010_NothingListeningShowsError(t *testing.T) {
	ln, err := net.Listen("tcp4", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	port := ln.Addr().(*net.TCPAddr).Port
	ln.Close()
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, `{"version":2,"printers":[{"name":"A","printer":"Canon","port":`+strconv.Itoa(port)+`}]}`)
	start := time.Now()
	rec := postForm(u.srv, "/test", vals("token", testToken, "name", "A"))
	if time.Since(start) > 5*time.Second || rec.Code == http.StatusSeeOther || rec.Code >= 500 ||
		!strings.Contains(rec.Body.String(), `class="error"`) {
		t.Fatalf("test to a closed port = %d after %v, want an error page within 5 s", rec.Code, time.Since(start))
	}
}

// TC-UI-089
func TestREQ_UI_011_RefreshWithoutJS(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	logLines(t, lp, 3)
	body := get(New(Deps{Token: testToken, LogPath: lp}), "/log").Body.String()
	if !strings.Contains(body, `href="/log"`) && !strings.Contains(body, `action="/log"`) {
		t.Fatal("log page has no link or form to refresh /log")
	}
}
