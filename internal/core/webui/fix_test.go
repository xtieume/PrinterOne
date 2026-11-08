package webui

// F-005 review findings, 2026-11-07 (TC-UI-092 … TC-UI-102). Go 1.10: no t.TempDir/t.Cleanup.

import (
	"bufio"
	"encoding/json"
	"net"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/allow"
	"github.com/xtieume/PrinterOne/internal/core/config"
)

// listenUI is a UI over <dir>/config.json whose Apply drives real Listeners over the recording fake.
func listenUI(t *testing.T, dir, body string) (*Server, string, *listenRec) {
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, body)
	cfg := loadCfg(t, p)
	r := &listenRec{}
	l := &Listeners{Listen: r.listen}
	if err := l.Apply(cfg); err != nil {
		t.Fatal(err)
	}
	r.take()
	return New(Deps{Config: cfg, ConfigPath: p, Token: testToken, Apply: l.Apply}), p, r
}

func rawPrinter(t *testing.T, path string, i int) map[string]interface{} {
	ps, _ := rawCfg(t, path)["printers"].([]interface{})
	if i >= len(ps) {
		t.Fatalf("saved config has %d printers, want index %d", len(ps), i)
	}
	return ps[i].(map[string]interface{})
}

// TC-UI-092
func TestREQ_UI_008_AddKeepsOtherPrintersAllowAndUnknownKeys(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	s, p, r := listenUI(t, dir, `{"version":2,"printers":[`+
		`{"name":"A","printer":"Canon","port":9101,"allow":["10.0.0.5"],"note":"keep-a"},`+
		`{"name":"B","printer":"HP","port":9102,"allow":["192.168.0.0/24"]}]}`)
	mustRedirect(t, postForm(s, "/printers", vals("token", testToken, "action", "add", "name", "C",
		"printer", "Brother", "port", "9103")), "/printers")
	a, b := rawPrinter(t, p, 0), rawPrinter(t, p, 1)
	if !reflect.DeepEqual(a["allow"], []interface{}{"10.0.0.5"}) || a["note"] != "keep-a" ||
		!reflect.DeepEqual(b["allow"], []interface{}{"192.168.0.0/24"}) {
		t.Fatalf("after add: A %v, B %v; want both allowlists and A's note kept", a, b)
	}
	if got := r.take(); !sameSet(got, []string{"start Brother@9103 allow all"}) {
		t.Fatalf("listener events %q; want only C started (A and B keep their allowlists)", got)
	}
}

// TC-UI-093
func TestREQ_UI_008_RemoveKeepsSettingsWithTheirPrinter(t *testing.T) {
	cases := []struct {
		name, body string
		wantAllow  interface{}
	}{
		{"allow must not move onto B", `{"version":2,"printers":[` +
			`{"name":"A","printer":"Canon","port":9101,"allow":["10.0.0.5"]},` +
			`{"name":"B","printer":"HP","port":9102}]}`, nil},
		{"B must keep its allow", `{"version":2,"printers":[` +
			`{"name":"A","printer":"Canon","port":9101},` +
			`{"name":"B","printer":"HP","port":9102,"allow":["192.168.0.0/24"]}]}`, []interface{}{"192.168.0.0/24"}},
	}
	for _, c := range cases {
		dir, done := tmpDir(t)
		s, p, r := listenUI(t, dir, c.body)
		mustRedirect(t, postForm(s, "/printers", vals("token", testToken, "action", "remove", "index", "0")), "/printers")
		b := rawPrinter(t, p, 0)
		if b["name"] != "B" || !reflect.DeepEqual(b["allow"], c.wantAllow) {
			done()
			t.Fatalf("%s: after removing A the first entry is %v; want B with allow %v", c.name, b, c.wantAllow)
		}
		for _, e := range r.take() {
			if strings.Contains(e, "HP@9102") {
				done()
				t.Fatalf("%s: B's listener restarted (%q)", c.name, e)
			}
		}
		done()
	}
}

// readEvents sends every "data:" line of an SSE body to a channel.
func readEvents(body *bufio.Reader) chan string {
	ch := make(chan string, 16)
	go func() {
		for {
			line, err := body.ReadString('\n')
			if err != nil {
				close(ch)
				return
			}
			ch <- strings.TrimRight(line, "\r\n")
		}
	}()
	return ch
}

func waitData(ch chan string, want string, d time.Duration) (string, bool) {
	deadline := time.After(d)
	seen := ""
	for {
		select {
		case l, ok := <-ch:
			if !ok {
				return seen + "<EOF>", false
			}
			seen += l + "|"
			if l == "data: "+want {
				return seen, true
			}
		case <-deadline:
			return seen, false
		}
	}
}

// TC-UI-094
func TestREQ_UI_016_StreamWithPollUnsetDeliversNewLine(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	writeFile(t, lp, "old\n")
	ts := httptest.NewServer(New(Deps{LogPath: lp, Token: testToken}))
	defer ts.Close()
	resp, err := http.Get(ts.URL + "/api/log/stream")
	if err != nil {
		t.Fatal(err)
	}
	ch := readEvents(bufio.NewReader(resp.Body))
	if l := <-ch; l != ": connected" {
		resp.Body.Close()
		t.Fatalf("first line %q", l)
	}
	f, _ := os.OpenFile(lp, os.O_APPEND|os.O_WRONLY, 0644)
	f.WriteString("fresh-line\n")
	f.Close()
	seen, ok := waitData(ch, "fresh-line", 3*time.Second)
	resp.Body.Close()
	if !ok {
		t.Fatalf("stream with Poll unset sent %q; want data: fresh-line", seen)
	}
}

// TC-UI-095
func TestREQ_UI_016_StreamResumesFromLastEventID(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	writeFile(t, lp, "old\nmissed-line\n")
	ts := httptest.NewServer(New(Deps{LogPath: lp, Token: testToken, Poll: 20 * time.Millisecond}))
	defer ts.Close()
	req, _ := http.NewRequest("GET", ts.URL+"/api/log/stream", nil)
	req.Header.Set("Last-Event-ID", "4") // the browser reconnects after the server dropped it at offset 4
	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		t.Fatal(err)
	}
	seen, ok := waitData(readEvents(bufio.NewReader(resp.Body)), "missed-line", 3*time.Second)
	resp.Body.Close()
	if !ok {
		t.Fatalf("reconnect with Last-Event-ID 4 sent %q; want the line written while disconnected", seen)
	}
}

// TC-UI-096
func TestREQ_UI_008_SaveTempFileNameHasNoPattern(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, oneCfg)
	seen := ""
	s := New(Deps{Config: loadCfg(t, p), ConfigPath: p, Token: testToken,
		Load: func(path string) (*config.Config, error) {
			seen = path
			return config.Load(config.Lookup{Flag: path})
		}})
	mustRedirect(t, postForm(s, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9105")), "/printers")
	if seen == "" || strings.Contains(filepath.Base(seen), "*") || filepath.Dir(seen) != dir {
		t.Fatalf("candidate validated at %q; want a temp file in %s without '*' (Go 1.10 TempFile has no pattern)", seen, dir)
	}
	if loadCfg(t, p).Printers[0].Port != 9105 {
		t.Fatal("save not written")
	}
	mustOnlyConfig(t, dir)
}

// TC-UI-097
func TestREQ_UI_016_InvalidSaveAnswers400WithFieldMessage(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	// a plain urlencoded post (what app.js now sends) carries the token: the valid save goes through
	ok := newFileUI(t, dir, oneCfg)
	mustRedirect(t, postForm(ok.srv, "/settings", vals("token", testToken, "keep_days", "30", "ui_port", "9180")), "/settings?saved=1")
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "0"))
	if rec.Code != http.StatusBadRequest {
		t.Fatalf("invalid save answered %d; want 400 so app.js sees the failure", rec.Code)
	}
	refused(t, u, before, rec, "printers[0].port")
	rec = postForm(u.srv, "/settings", vals("token", testToken, "keep_days", "-1", "ui_port", "9180"))
	if rec.Code != http.StatusBadRequest {
		t.Fatalf("invalid settings save answered %d; want 400", rec.Code)
	}
	refused(t, u, before, rec, "log.keep_days")
}

// TC-UI-098
func TestREQ_UI_004_WrongLengthPasswordRefused(t *testing.T) {
	s := New(Deps{Token: testToken, UI: config.UI{Port: 9180, Remote: true, Password: "s3cretpass"}})
	for _, pass := range []string{"s3cretpas", "s3cretpass1", ""} {
		if rec := getAuth(s, "/", "admin", pass); rec.Code != http.StatusUnauthorized {
			t.Fatalf("password %q answered %d; want 401", pass, rec.Code)
		}
	}
	if rec := getAuth(s, "/", "admin1", "s3cretpass"); rec.Code != http.StatusUnauthorized {
		t.Fatalf("user admin1 answered %d; want 401", rec.Code)
	}
	if rec := getAuth(s, "/", "admin", "s3cretpass"); rec.Code != 200 {
		t.Fatalf("right password answered %d", rec.Code)
	}
}

// TC-UI-099
func TestREQ_UI_018_APITestErrorIsJSONMessage(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	rec := postJSON(u.srv, "/api/test", `{"name":"nope"}`, testToken)
	var v map[string]string
	if rec.Code != 400 || json.Unmarshal(rec.Body.Bytes(), &v) != nil || v["error"] != "unknown printer" || v["field"] != "name" {
		t.Fatalf("/api/test unknown printer = %d %q; want 400 {\"error\":\"unknown printer\",\"field\":\"name\"}", rec.Code, rec.Body.String())
	}
}

// TC-UI-100
func TestREQ_UI_007_OversizedPostRefused(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	pad := strings.Repeat("x", 2<<20)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9105", "pad", pad))
	if rec.Code != 413 && rec.Code != 400 {
		t.Fatalf("2 MB form answered %d; want 413 or 400", rec.Code)
	}
	rec = postJSON(u.srv, "/api/printers", strings.Repeat(" ", 2<<20)+`[{"name":"A","printer":"Canon","port":9105}]`, testToken)
	if rec.Code != 413 && rec.Code != 400 {
		t.Fatalf("2 MB API body answered %d; want 413 or 400", rec.Code)
	}
	if readFile(t, u.path) != before || u.applied.n() != 0 {
		t.Fatal("an oversized POST was processed")
	}
}

// TC-UI-101
func TestREQ_UI_012_FirstSaveNeedsNoRestart(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json") // does not exist yet
	s := New(Deps{Config: &config.Config{Path: p}, ConfigPath: p, Token: testToken, UI: config.UI{Port: config.DefaultUIPort}})
	mustRedirect(t, postForm(s, "/printers", vals("token", testToken, "action", "add", "name", "A",
		"printer", "Canon", "port", "9101")), "/printers")
	if body := get(s, "/printers").Body.String(); strings.Contains(body, "restart") {
		t.Fatal("first save with no config file claims a UI restart is needed")
	}
}

// TC-UI-102
func TestREQ_UI_009_PortSwapInOneSave(t *testing.T) {
	held := map[int]bool{}
	l := &Listeners{Listen: func(p config.Printer, _ allow.List) (func(), error) {
		if held[p.Port] {
			return nil, &net.OpError{Op: "listen", Err: os.ErrExist}
		}
		held[p.Port] = true
		return func() { delete(held, p.Port) }, nil
	}}
	A := config.Printer{Name: "A", Printer: "Canon", Port: 9101}
	B := config.Printer{Name: "B", Printer: "HP", Port: 9102}
	if err := l.Apply(literal(t, []config.Printer{A, B}, nil)); err != nil {
		t.Fatal(err)
	}
	A.Port, B.Port = 9102, 9101
	if err := l.Apply(literal(t, []config.Printer{A, B}, nil)); err != nil {
		t.Fatalf("port swap: %v", err)
	}
	if !held[9101] || !held[9102] || len(l.running) != 2 {
		t.Fatalf("after swap held %v, running %d; want both ports listening", held, len(l.running))
	}
}
