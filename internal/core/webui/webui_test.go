package webui

// F-005 web UI and API (spec docs/specs/2026-11-07-f005-web-ui.md). Each test names its TC-UI-* case
// (.testcases/runs/F-005-r1/testcase/testcases.md; why-records in testcases.md).

import (
	"bufio"
	"encoding/json"
	"errors"
	"io/ioutil"
	"net"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/allow"
	"github.com/xtieume/PrinterOne/internal/core/config"
)

// --- REQ-UI-001 / 003 / 004: binding, browser, remote ---

func startUI(t *testing.T, ui config.UI, h http.Handler, open func(string) error) (string, func()) {
	addr, stop, err := Start(ui, h, open)
	if err != nil {
		t.Fatalf("Start: %v", err)
	}
	if stop == nil {
		t.Fatal("Start returned a nil stop")
	}
	return addr, stop
}

func hostOf(t *testing.T, addr string) (string, string) {
	h, p, err := net.SplitHostPort(addr)
	if err != nil {
		t.Fatalf("Start returned address %q: %v", addr, err)
	}
	return h, p
}

// TC-UI-003
func TestREQ_UI_001_BindsLoopbackAndServes(t *testing.T) {
	addr, stop := startUI(t, config.UI{Port: 0}, New(Deps{Token: testToken}), nil)
	defer stop()
	if h, _ := hostOf(t, addr); h != "127.0.0.1" {
		t.Fatalf("bound to %s, want 127.0.0.1 only", addr)
	}
	resp, err := (&http.Client{Timeout: 5 * time.Second}).Get("http://" + addr + "/")
	if err != nil {
		t.Fatal(err)
	}
	b, _ := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if resp.StatusCode != 200 || !strings.Contains(string(b), "PrinterOne") {
		t.Fatalf("GET / = %d %.200q, want 200 with PrinterOne", resp.StatusCode, b)
	}
}

// TC-UI-012
func TestREQ_UI_003_BrowserCommandIsRundll32(t *testing.T) {
	u := "http://127.0.0.1:9180/"
	want := []string{"rundll32", "url.dll,FileProtocolHandler", u}
	if got := BrowserCommand(u); !reflect.DeepEqual(got, want) {
		t.Fatalf("BrowserCommand = %q, want %q", got, want)
	}
}

// TC-UI-013
func TestREQ_UI_003_OpenerGetsURLAfterBind(t *testing.T) {
	if got := URL(9180); got != "http://127.0.0.1:9180/" {
		t.Fatalf("URL(9180) = %q", got)
	}
	var mu sync.Mutex
	var urls []string
	open := func(u string) error {
		mu.Lock()
		urls = append(urls, u)
		mu.Unlock()
		// the UI must already answer when the browser is told to open it
		resp, err := (&http.Client{Timeout: 5 * time.Second}).Get(u)
		if err != nil {
			return err
		}
		resp.Body.Close()
		return nil
	}
	addr, stop := startUI(t, config.UI{Port: 0}, New(Deps{Token: testToken}), open)
	defer stop()
	_, port := hostOf(t, addr)
	time.Sleep(100 * time.Millisecond)
	mu.Lock()
	defer mu.Unlock()
	if len(urls) != 1 || urls[0] != "http://127.0.0.1:"+port+"/" {
		t.Fatalf("opener calls %q, want exactly [http://127.0.0.1:%s/]", urls, port)
	}
}

// TC-UI-071
func TestREQ_UI_003_OpenerFailureKeepsUI(t *testing.T) {
	addr, stop := startUI(t, config.UI{Port: 0}, New(Deps{Token: testToken}),
		func(string) error { return errors.New("no browser") })
	defer stop()
	resp, err := (&http.Client{Timeout: 5 * time.Second}).Get("http://" + addr + "/")
	if err != nil {
		t.Fatalf("UI not serving after opener failure: %v", err)
	}
	resp.Body.Close()
	if resp.StatusCode != 200 {
		t.Fatalf("GET / = %d", resp.StatusCode)
	}
}

// TC-UI-014
func TestREQ_UI_003_NoConfigEmptyList(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	s := New(Deps{ConfigPath: filepath.Join(dir, "config.json"), Token: testToken})
	rec := get(s, "/")
	if rec.Code != 200 || !strings.Contains(rec.Body.String(), `href="/printers"`) {
		t.Fatalf("GET / with no config = %d, want 200 linking to /printers", rec.Code)
	}
	rec = get(s, "/api/printers")
	if rec.Code != 200 || strings.TrimSpace(rec.Body.String()) != "[]" {
		t.Fatalf("GET /api/printers with no config = %d %q, want 200 []", rec.Code, rec.Body.String())
	}
}

// TC-UI-015
func TestREQ_UI_003_FirstSaveWritesConfigAndStartsReceivers(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json")
	a := &applyRec{}
	s := New(Deps{ConfigPath: p, Token: testToken, Apply: a.apply})
	rec := postForm(s, "/printers", vals("token", testToken, "action", "add", "name", "A", "printer", "Canon", "port", "9101"))
	mustRedirect(t, rec, "/printers")
	c := loadCfg(t, p)
	if len(c.Printers) != 1 || c.Printers[0].Name != "A" || c.Printers[0].Printer != "Canon" || c.Printers[0].Port != 9101 {
		t.Fatalf("saved printers %+v", c.Printers)
	}
	if a.n() != 1 {
		t.Fatalf("Apply called %d times, want 1", a.n())
	}
}

// TC-UI-019
func TestREQ_UI_004_RemoteBindsAllInterfaces(t *testing.T) {
	ui := config.UI{Port: 0, Remote: true, Password: "s3cretpass"}
	addr, stop := startUI(t, ui, New(Deps{Token: testToken, UI: ui}), nil)
	defer stop()
	if h, _ := hostOf(t, addr); h != "0.0.0.0" {
		t.Fatalf("remote bound to %s, want 0.0.0.0", addr)
	}
}

// TC-UI-020, TC-UI-021
func TestREQ_UI_004_RemoteNeedsBasicAuth(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, oneCfg)
	before := readFile(t, p)
	s := New(Deps{Config: loadCfg(t, p), ConfigPath: p, Token: testToken, UI: config.UI{Port: 9180, Remote: true, Password: "s3cretpass"}})
	type req struct{ method, path, body, ctype string }
	form := vals("token", testToken, "action", "edit", "index", "0", "name", "A", "printer", "Canon", "port", "9105").Encode()
	reqs := []req{{"GET", "/", "", ""}, {"GET", "/printers", "", ""}, {"GET", "/test", "", ""}, {"GET", "/log", "", ""},
		{"GET", "/settings", "", ""}, {"GET", "/base.css", "", ""}, {"GET", "/api/status", "", ""},
		{"GET", "/api/printers", "", ""}, {"GET", "/api/installed-printers", "", ""}, {"GET", "/api/log", "", ""},
		{"POST", "/printers", form, "application/x-www-form-urlencoded"},
		{"POST", "/api/printers", `[{"name":"A","printer":"Canon","port":9105}]`, "application/json"}}
	for _, q := range reqs {
		for _, c := range [][2]string{{"", ""}, {"admin", "wrong"}, {"root", "s3cretpass"}} {
			rec := httptest.NewRecorder()
			r := httptest.NewRequest(q.method, q.path, strings.NewReader(q.body))
			if q.ctype != "" {
				r.Header.Set("Content-Type", q.ctype)
				r.Header.Set(TokenHeader, testToken)
			}
			if c[0] != "" {
				r.SetBasicAuth(c[0], c[1])
			}
			s.ServeHTTP(rec, r)
			if rec.Code != 401 || !strings.HasPrefix(rec.Header().Get("WWW-Authenticate"), "Basic") {
				t.Errorf("%s %s as %q/%q = %d %q, want 401 with Basic challenge", q.method, q.path, c[0], c[1], rec.Code, rec.Header().Get("WWW-Authenticate"))
			}
		}
	}
	if readFile(t, p) != before {
		t.Fatal("an unauthenticated POST changed the config")
	}
	if rec := getAuth(s, "/", "admin", "s3cretpass"); rec.Code != 200 {
		t.Fatalf("GET / as admin = %d, want 200", rec.Code)
	}
}

// TC-UI-022
func TestREQ_UI_004_LocalAsksNoPassword(t *testing.T) {
	s := New(Deps{Token: testToken, UI: config.UI{Port: 9180, Password: "s3cretpass"}})
	rec := get(s, "/")
	if rec.Code != 200 || !strings.Contains(rec.Body.String(), "PrinterOne") {
		t.Fatalf("GET / without remote = %d, want 200 page", rec.Code)
	}
}

// TC-UI-075
func TestREQ_UI_016_StreamNeedsAuthWhenRemote(t *testing.T) {
	s := New(Deps{Token: testToken, UI: config.UI{Port: 9180, Remote: true, Password: "s3cretpass"}})
	if rec := get(s, "/api/log/stream"); rec.Code != 401 {
		t.Fatalf("GET /api/log/stream without auth = %d, want 401", rec.Code)
	}
}

// --- REQ-UI-005 status ---

func statusDeps(t *testing.T, dir string) Deps {
	c := literal(t, []config.Printer{{Name: "PrnAlpha", Printer: "Canon", Port: 9101}, {Name: "PrnBeta", Printer: "HP", Port: 9102}},
		[][]string{{"10.0.0.0/8"}, nil})
	return Deps{Config: c, Token: testToken, LogPath: filepath.Join(dir, "printerone.log"),
		UI: config.UI{Port: 9180},
		Stats: func(name string) (int, string) {
			if name == "PrnAlpha" {
				return 4711, "printed"
			}
			return 0, ""
		},
		Installed: func() ([]string, error) { return []string{"Canon", "HP"}, nil }}
}

// TC-UI-023
func TestREQ_UI_005_StatusShowsEveryPrinter(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	d := statusDeps(t, dir)
	body := get(New(d), "/").Body.String()
	for _, want := range []string{d.LogPath, "http://127.0.0.1:9180/"} {
		if !strings.Contains(body, want) {
			t.Errorf("status page lacks %q", want)
		}
	}
	rowChecks := map[string][]string{
		"PrnAlpha": {"Canon", "9101", "10.0.0.0/8", "4711", "printed"},
		"PrnBeta":  {"HP", "9102", "allow: all"},
	}
	for name, wants := range rowChecks {
		row := rowOf(body, name)
		if row == "" {
			t.Errorf("status page has no table row for %s", name)
			continue
		}
		for _, w := range wants {
			if !strings.Contains(row, w) {
				t.Errorf("row of %s lacks %q", name, w)
			}
		}
	}
}

// rowOf returns the <tr>…</tr> holding name, or "".
func rowOf(body, name string) string {
	i := strings.Index(body, name)
	if i < 0 {
		return ""
	}
	start := strings.LastIndex(body[:i], "<tr")
	end := strings.Index(body[i:], "</tr>")
	if start < 0 || end < 0 {
		return ""
	}
	return body[start : i+end]
}

// --- REQ-UI-006 / 007 / 008: printer forms ---

// TC-UI-024, TC-UI-025, TC-UI-026, TC-UI-027
func TestREQ_UI_006_AddEditRemoveWithPlainForms(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	page := get(u.srv, "/printers").Body.String()
	if !strings.Contains(page, `<form method="post" action="/printers"`) {
		t.Fatalf("printers page has no plain POST form")
	}
	for _, opt := range []string{">Canon<", ">HP<"} {
		if !strings.Contains(page, opt) {
			t.Errorf("installed printer %s not offered", opt)
		}
	}
	tok := pageToken(t, u.srv, "/printers")

	rec := postForm(u.srv, "/printers", vals("token", tok, "action", "add", "name", "B", "printer", "HP", "port", "9102",
		"allow", "10.0.0.5\r\n\r\n192.168.0.0/24"))
	mustRedirect(t, rec, "/printers")
	c := loadCfg(t, u.path)
	if len(c.Printers) != 2 || c.Printers[1].Name != "B" || c.Printers[1].Printer != "HP" || c.Printers[1].Port != 9102 {
		t.Fatalf("after add: %+v", c.Printers)
	}
	ps := rawCfg(t, u.path)["printers"].([]interface{})
	if got := ps[1].(map[string]interface{})["allow"]; !reflect.DeepEqual(got, []interface{}{"10.0.0.5", "192.168.0.0/24"}) {
		t.Fatalf("after add: allow %v", got)
	}
	if !strings.Contains(get(u.srv, "/printers").Body.String(), "9102") {
		t.Fatal("printers page does not list the added printer")
	}

	rec = postForm(u.srv, "/printers", vals("token", tok, "action", "edit", "index", "0", "name", "A", "printer", "Canon",
		"printer_text", "Canon MF220", "port", "9105"))
	mustRedirect(t, rec, "/printers")
	c = loadCfg(t, u.path)
	if c.Printers[0].Port != 9105 || c.Printers[0].Printer != "Canon MF220" || len(c.Printers) != 2 {
		t.Fatalf("after edit: %+v", c.Printers)
	}

	rec = postForm(u.srv, "/printers", vals("token", tok, "action", "remove", "index", "1"))
	mustRedirect(t, rec, "/printers")
	c = loadCfg(t, u.path)
	if len(c.Printers) != 1 || c.Printers[0].Name != "A" {
		t.Fatalf("after remove: %+v", c.Printers)
	}
	if u.applied.n() != 3 {
		t.Fatalf("Apply called %d times, want 3", u.applied.n())
	}
}

// TC-UI-072
func TestREQ_UI_006_BadIndexWritesNothing(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	for _, v := range [][]string{
		{"action", "edit", "index", "5", "name", "A", "printer", "Canon", "port", "9101"},
		{"action", "remove", "index", "-1"},
		{"action", "remove", "index", "x"},
		{"action", "bogus", "index", "0"},
	} {
		rec := postForm(u.srv, "/printers", vals(append([]string{"token", testToken}, v...)...))
		if rec.Code == http.StatusSeeOther || rec.Code == 403 || rec.Code >= 500 {
			t.Errorf("%v: status %d, want a refusal page", v, rec.Code)
		}
	}
	if readFile(t, u.path) != before || u.applied.n() != 0 {
		t.Fatal("bad index changed the config or applied it")
	}
	if !strings.Contains(get(u.srv, "/printers").Body.String(), "9101") {
		t.Fatal("printers page broken after bad posts")
	}
}

// refused checks an invalid form: not redirected, loader message next to field, nothing written.
func refused(t *testing.T, u *fileUI, before string, rec *httptest.ResponseRecorder, field string) {
	if rec.Code == http.StatusSeeOther || rec.Code == 403 {
		t.Fatalf("invalid form answered %d", rec.Code)
	}
	body := rec.Body.String()
	marker := `<span class="error" data-field="` + field + `">`
	i := strings.Index(body, marker)
	if i < 0 {
		t.Fatalf("no error element for %s in page", field)
	}
	if !strings.Contains(body[i:], "invalid field "+field+"</span>") {
		t.Fatalf("error element for %s does not hold the loader message", field)
	}
	if readFile(t, u.path) != before {
		t.Fatal("config written although the form was invalid")
	}
	mustOnlyConfig(t, u.dir)
	if u.applied.n() != 0 {
		t.Fatal("Apply called for an invalid form")
	}
}

// TC-UI-028, TC-UI-031
func TestREQ_UI_007_BadPortRefusedNextToField(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "70000"))
	refused(t, u, before, rec, "printers[0].port")
	if !strings.Contains(rec.Body.String(), u.path+": invalid field printers[0].port") || strings.Contains(rec.Body.String(), ".tmp") {
		t.Fatal("message does not name the real config path (or leaks the temp file)")
	}
}

// TC-UI-029 (F-003 rule)
func TestREQ_UI_007_DuplicatePortRefused(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "add", "name", "B", "printer", "HP", "port", "9101"))
	refused(t, u, before, rec, "printers[1].port")
}

// TC-UI-030 (F-004 rule)
func TestREQ_UI_007_BadAllowItemRefused(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9101", "allow", "not-an-ip"))
	refused(t, u, before, rec, "printers[0].allow[0]")
}

// TC-UI-032
func TestREQ_UI_007_RemovingLastPrinterRefused(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "remove", "index", "0"))
	refused(t, u, before, rec, "printers")
}

// TC-UI-033, TC-UI-034
func TestREQ_UI_008_KeepsUnknownKeysAtomically(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, `{"version":2,"extra":{"keep":1},"log":{"keep_days":5,"x":1},`+
		`"printers":[{"name":"A","printer":"Canon","port":9101,"note":"n"}]}`)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9105"))
	mustRedirect(t, rec, "/printers")
	m := rawCfg(t, u.path)
	p0 := m["printers"].([]interface{})[0].(map[string]interface{})
	lg, _ := m["log"].(map[string]interface{})
	ex, _ := m["extra"].(map[string]interface{})
	if ex == nil || ex["keep"] != 1.0 || p0["note"] != "n" || p0["port"] != 9105.0 || lg == nil || lg["x"] != 1.0 || lg["keep_days"] != 5.0 {
		t.Fatalf("saved %s", readFile(t, u.path))
	}
	mustOnlyConfig(t, dir)
}

// TC-UI-073
func TestREQ_UI_008_WriteFailureKeepsOriginal(t *testing.T) {
	if os.Geteuid() == 0 || os.PathSeparator == '\\' {
		t.Skip("a read-only folder does not stop root or Windows; covered on a non-root Unix run")
	}
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	if err := os.Chmod(dir, 0500); err != nil {
		t.Fatal(err)
	}
	defer os.Chmod(dir, 0700)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0", "name", "A",
		"printer", "Canon", "port", "9105"))
	if rec.Code == http.StatusSeeOther || rec.Code == 403 || !strings.Contains(rec.Body.String(), `class="error"`) {
		t.Fatalf("failed write answered %d without an error on the page", rec.Code)
	}
	if readFile(t, u.path) != before || u.applied.n() != 0 {
		t.Fatal("original config changed or applied after a failed write")
	}
}

// --- REQ-UI-009 listeners ---

type listenRec struct {
	mu     sync.Mutex
	events []string
	fail   map[int]bool
}

func (r *listenRec) listen(p config.Printer, a allow.List) (func(), error) {
	r.mu.Lock()
	defer r.mu.Unlock()
	key := p.Printer + "@" + strconv.Itoa(p.Port) + " allow " + a.String()
	if r.fail[p.Port] {
		return nil, errors.New("port busy")
	}
	r.events = append(r.events, "start "+key)
	return func() {
		r.mu.Lock()
		r.events = append(r.events, "stop "+key)
		r.mu.Unlock()
	}, nil
}

func (r *listenRec) take() []string {
	r.mu.Lock()
	defer r.mu.Unlock()
	e := r.events
	r.events = nil
	return e
}

func sameSet(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	m := map[string]int{}
	for _, s := range a {
		m[s]++
	}
	for _, s := range b {
		m[s]--
	}
	for _, v := range m {
		if v != 0 {
			return false
		}
	}
	return true
}

func applyExpect(t *testing.T, l *Listeners, r *listenRec, c *config.Config, step string, want ...string) {
	if err := l.Apply(c); err != nil {
		t.Fatalf("%s: Apply: %v", step, err)
	}
	if got := r.take(); !sameSet(got, want) {
		t.Fatalf("%s: listener events %q, want %q", step, got, want)
	}
}

// TC-UI-035, TC-UI-036, TC-UI-037, TC-UI-038
func TestREQ_UI_009_RestartsOnlyChangedListeners(t *testing.T) {
	r := &listenRec{}
	l := &Listeners{Listen: r.listen}
	A := config.Printer{Name: "A", Printer: "Canon", Port: 9101}
	B := config.Printer{Name: "B", Printer: "HP", Port: 9102}
	C := config.Printer{Name: "C", Printer: "Brother", Port: 9103}
	applyExpect(t, l, r, literal(t, []config.Printer{A, B, C}, nil), "first apply",
		"start Canon@9101 allow all", "start HP@9102 allow all", "start Brother@9103 allow all")

	B2 := B
	B2.Port = 9104
	applyExpect(t, l, r, literal(t, []config.Printer{A, B2, C}, nil), "port change",
		"stop HP@9102 allow all", "start HP@9104 allow all")

	C2 := C
	C2.Printer = "HP2"
	applyExpect(t, l, r, literal(t, []config.Printer{A, B2, C2}, nil), "printer change",
		"stop Brother@9103 allow all", "start HP2@9103 allow all")

	applyExpect(t, l, r, literal(t, []config.Printer{A, B2, C2}, [][]string{{"10.0.0.0/8"}}), "allow change",
		"stop Canon@9101 allow all", "start Canon@9101 allow 10.0.0.0/8")

	B3 := B2
	B3.Name = "B2"
	applyExpect(t, l, r, literal(t, []config.Printer{A, B3, C2}, [][]string{{"10.0.0.0/8"}}), "name-only change")

	D := config.Printer{Name: "D", Printer: "Epson", Port: 9106}
	applyExpect(t, l, r, literal(t, []config.Printer{A, B3, C2, D}, [][]string{{"10.0.0.0/8"}}), "add",
		"start Epson@9106 allow all")

	applyExpect(t, l, r, literal(t, []config.Printer{A, B3, D}, [][]string{{"10.0.0.0/8"}}), "remove",
		"stop HP2@9103 allow all")
}

// TC-UI-074
func TestREQ_UI_009_FailedStartKeepsOthersAndRetries(t *testing.T) {
	r := &listenRec{fail: map[int]bool{}}
	l := &Listeners{Listen: r.listen}
	A := config.Printer{Name: "A", Printer: "Canon", Port: 9101}
	B := config.Printer{Name: "B", Printer: "HP", Port: 9102}
	applyExpect(t, l, r, literal(t, []config.Printer{A, B}, nil), "first apply",
		"start Canon@9101 allow all", "start HP@9102 allow all")
	B2 := B
	B2.Port = 9104
	r.fail[9104] = true
	c := literal(t, []config.Printer{A, B2}, nil)
	if err := l.Apply(c); err == nil {
		t.Fatal("Apply hid the start failure")
	}
	if got := r.take(); !sameSet(got, []string{"stop HP@9102 allow all"}) {
		t.Fatalf("failed apply events %q, want only B's old listener stopped", got)
	}
	r.fail[9104] = false
	applyExpect(t, l, r, c, "retry", "start HP@9104 allow all")
}

// --- REQ-UI-010 test page ---

// TC-UI-040
func TestREQ_UI_010_TestPageSendsThroughPort(t *testing.T) {
	sink := newJobSink(t)
	defer sink.ln.Close()
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, strings.Replace(sink.cfg(), `"name":"A"`, `"name":"PrnAlpha"`, 1))
	tok := pageToken(t, u.srv, "/test")
	mustRedirect(t, postForm(u.srv, "/test", vals("token", tok, "name", "PrnAlpha")), "/test")
	job, ok := sink.wait(3 * time.Second)
	if !ok || !strings.Contains(job, "PrnAlpha") || !strings.Contains(job, "2026-11-07") {
		t.Fatalf("test job %q (received %v), want the printer name and the date", job, ok)
	}
}

// TC-UI-041
func TestREQ_UI_010_UnknownPrinterSendsNothing(t *testing.T) {
	sink := newJobSink(t)
	defer sink.ln.Close()
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, sink.cfg())
	rec := postForm(u.srv, "/test", vals("token", testToken, "name", "Z"))
	if rec.Code == http.StatusSeeOther || rec.Code == 403 || !strings.Contains(rec.Body.String(), `class="error"`) {
		t.Fatalf("unknown printer answered %d without an error", rec.Code)
	}
	if job, ok := sink.wait(500 * time.Millisecond); ok {
		t.Fatalf("a job was sent for an unknown printer: %q", job)
	}
}

// --- REQ-UI-011 log page ---

// TC-UI-043
func TestREQ_UI_011_LastTwoHundredLinesNewestLast(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	logLines(t, lp, 250)
	body := get(New(Deps{Token: testToken, LogPath: lp}), "/log").Body.String()
	i, j := strings.Index(body, "line-051;"), strings.Index(body, "line-250;")
	if strings.Contains(body, "line-050;") || i < 0 || j < 0 || i > j {
		t.Fatalf("log page: has 050=%v, 051 at %d, 250 at %d; want lines 051..250, newest last",
			strings.Contains(body, "line-050;"), i, j)
	}
}

// TC-UI-044
func TestREQ_UI_011_MissingLogFile(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	rec := get(New(Deps{Token: testToken, LogPath: filepath.Join(dir, "none.log")}), "/log")
	if rec.Code != 200 || !strings.Contains(rec.Body.String(), "PrinterOne") {
		t.Fatalf("GET /log with no file = %d, want 200 page", rec.Code)
	}
}

// --- REQ-UI-012 settings ---

// TC-UI-045, TC-UI-047
func TestREQ_UI_012_SettingsSaveLogAndUI(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	logs := filepath.Join(dir, "logs2")
	tok := pageToken(t, u.srv, "/settings")
	rec := postForm(u.srv, "/settings", vals("token", tok, "log_dir", logs, "keep_days", "10", "ui_port", "9181",
		"remote", "on", "password", "s3cretpass"))
	if rec.Code != http.StatusSeeOther || !strings.HasPrefix(rec.Header().Get("Location"), "/settings") {
		t.Fatalf("settings save = %d Location %q, want 303 to /settings", rec.Code, rec.Header().Get("Location"))
	}
	m := rawCfg(t, u.path)
	lg, _ := m["log"].(map[string]interface{})
	ui, _ := m["ui"].(map[string]interface{})
	if lg == nil || ui == nil || lg["dir"] != logs || lg["keep_days"] != 10.0 || ui["port"] != 9181.0 ||
		ui["remote"] != true || ui["password"] != "s3cretpass" {
		t.Fatalf("saved %s", readFile(t, u.path))
	}
	if !strings.Contains(get(u.srv, rec.Header().Get("Location")).Body.String(), "restart") {
		t.Fatal("settings page does not say a restart is needed for the UI change")
	}
	// control: a save that leaves ui.* as stored does not ask for a restart
	rec = postForm(u.srv, "/settings", vals("token", tok, "log_dir", logs, "keep_days", "11", "ui_port", "9181",
		"remote", "on", "password", ""))
	if rec.Code != http.StatusSeeOther {
		t.Fatalf("control save = %d", rec.Code)
	}
	if strings.Contains(get(u.srv, rec.Header().Get("Location")).Body.String(), "restart") {
		t.Fatal("restart hint shown although ui.* did not change")
	}
}

// TC-UI-046
func TestREQ_UI_012_PasswordNeverSentBack(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101}],"ui":{"password":"s3cretpass"}}`)
	s := New(Deps{Config: loadCfg(t, p), ConfigPath: p, Token: testToken, UI: config.UI{Port: 9180, Password: "s3cretpass"}})
	if !strings.Contains(get(s, "/settings").Body.String(), `name="password"`) {
		t.Fatal("settings page has no password field")
	}
	for _, path := range []string{"/settings", "/api/status", "/api/printers", "/"} {
		rec := get(s, path)
		if rec.Code != 200 {
			t.Errorf("GET %s = %d", path, rec.Code)
		}
		if strings.Contains(rec.Body.String(), "s3cretpass") {
			t.Errorf("GET %s sends the password back", path)
		}
	}
}

// TC-UI-048
func TestREQ_UI_012_BadKeepDaysRefused(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postForm(u.srv, "/settings", vals("token", testToken, "keep_days", "0", "ui_port", "9180"))
	refused(t, u, before, rec, "log.keep_days")
}

// TC-UI-049
func TestREQ_UI_012_EmptyPasswordKeepsStored(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101}],`+
		`"ui":{"remote":true,"password":"s3cretpass"}}`)
	rec := postForm(u.srv, "/settings", vals("token", testToken, "keep_days", "20", "ui_port", "9180", "remote", "on", "password", ""))
	if rec.Code != http.StatusSeeOther {
		t.Fatalf("settings save = %d, want 303", rec.Code)
	}
	ui, _ := rawCfg(t, u.path)["ui"].(map[string]interface{})
	if ui == nil || ui["password"] != "s3cretpass" {
		t.Fatalf("empty password field changed the stored one: %s", readFile(t, u.path))
	}
}

// --- REQ-UI-013 token ---

// TC-UI-050, TC-UI-051
func TestREQ_UI_013_PostWithoutTokenIs403(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	forms := map[string][]string{
		"/printers": {"action", "add", "name", "B", "printer", "HP", "port", "9102"},
		"/settings": {"keep_days", "9", "ui_port", "9180"},
		"/test":     {"name", "A"},
	}
	for path, kv := range forms {
		for _, tok := range []string{"", "bad"} {
			v := vals(kv...)
			if tok != "" {
				v.Set("token", tok)
			}
			if rec := postForm(u.srv, path, v); rec.Code != 403 {
				t.Errorf("POST %s token %q = %d, want 403", path, tok, rec.Code)
			}
		}
	}
	for _, path := range []string{"/api/printers", "/api/test"} {
		for _, tok := range []string{"", "bad"} {
			if rec := postJSON(u.srv, path, `{"name":"A"}`, tok); rec.Code != 403 {
				t.Errorf("POST %s header %q = %d, want 403", path, tok, rec.Code)
			}
		}
	}
	if readFile(t, u.path) != before || u.applied.n() != 0 {
		t.Fatal("a refused POST changed the config")
	}
}

// TC-UI-052
func TestREQ_UI_013_TokenRandomPerProcess(t *testing.T) {
	a, b := New(Deps{}).Token(), New(Deps{}).Token()
	if len(a) < 16 || len(b) < 16 || a == b {
		t.Fatalf("tokens %q and %q: want two different random tokens of at least 16 chars", a, b)
	}
	if got := New(Deps{Token: testToken}).Token(); got != testToken {
		t.Fatalf("Token() = %q, want the injected %q", got, testToken)
	}
}

// --- REQ-UI-014 escaping ---

// TC-UI-053, TC-UI-054
func TestREQ_UI_014_HostileValuesEscaped(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	writeFile(t, lp, "<script>log()</script>\n")
	c := literal(t, []config.Printer{{Name: `<script>alert(1)</script>"'&`, Printer: `"><b>x`, Port: 9101}}, nil)
	s := New(Deps{Config: c, Token: testToken, LogPath: lp})
	for _, path := range []string{"/", "/printers"} {
		body := get(s, path).Body.String()
		if strings.Contains(body, "<script>alert(1)") || strings.Contains(body, `"><b>x`) {
			t.Errorf("GET %s renders a hostile value raw", path)
		}
		if !strings.Contains(body, "&lt;script&gt;alert(1)&lt;/script&gt;") {
			t.Errorf("GET %s does not show the escaped printer name", path)
		}
	}
	body := get(s, "/log").Body.String()
	if strings.Contains(body, "<script>log()") || !strings.Contains(body, "&lt;script&gt;log()") {
		t.Error("log page does not escape the log line")
	}
}

// TC-UI-055
func TestREQ_UI_014_EchoedInputEscaped(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	rec := postForm(u.srv, "/printers", vals("token", testToken, "action", "edit", "index", "0",
		"name", `"><script>evil()</script>`, "printer", "Canon", "port", "bad"))
	body := rec.Body.String()
	if strings.Contains(body, "<script>evil()") || !strings.Contains(body, "&lt;script&gt;evil()") {
		t.Fatalf("invalid form echoes the input raw or not at all (status %d)", rec.Code)
	}
}

// --- REQ-UI-016 live log ---

// TC-UI-059
func TestREQ_UI_016_StreamSendsNewLines(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	writeFile(t, lp, "old-1\nold-2\nold-3\n")
	ts := httptest.NewServer(New(Deps{Token: testToken, LogPath: lp, Poll: 50 * time.Millisecond}))
	defer ts.Close()
	resp, err := (&http.Client{Timeout: 5 * time.Second}).Get(ts.URL + "/api/log/stream")
	if err != nil {
		t.Fatal(err)
	}
	defer resp.Body.Close()
	if !strings.HasPrefix(resp.Header.Get("Content-Type"), "text/event-stream") {
		t.Fatalf("stream Content-Type %q, want text/event-stream", resp.Header.Get("Content-Type"))
	}
	lines := make(chan string, 16)
	go func() {
		sc := bufio.NewScanner(resp.Body)
		for sc.Scan() {
			lines <- sc.Text()
		}
		close(lines)
	}()
	time.Sleep(200 * time.Millisecond)
	f, err := os.OpenFile(lp, os.O_APPEND|os.O_WRONLY, 0644)
	if err != nil {
		t.Fatal(err)
	}
	f.WriteString("new-line-1\nnew-line-2\n")
	f.Close()
	deadline := time.After(3 * time.Second)
	want := []string{"data: new-line-1", "data: new-line-2"}
	for {
		select {
		case l, ok := <-lines:
			if !ok {
				t.Fatal("stream ended before the new line")
			}
			if strings.HasPrefix(l, "data:") && strings.Contains(l, "old-") {
				t.Fatalf("stream replayed an old line: %q", l)
			}
			if strings.HasPrefix(l, "data:") {
				if l != want[0] {
					t.Fatalf("event %q, want %q (one event per line)", l, want[0])
				}
				want = want[1:]
				if len(want) == 0 {
					return
				}
			}
		case <-deadline:
			t.Fatalf("missing events %q within 3 s", want)
		}
	}
}

// --- REQ-UI-018 JSON API ---

// TC-UI-062
func TestREQ_UI_018_ReadEndpoints(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	d := statusDeps(t, dir)
	s := New(d)

	rec := get(s, "/api/status")
	var st struct {
		Printers []struct {
			Name       string   `json:"name"`
			Printer    string   `json:"printer"`
			Port       int      `json:"port"`
			Allow      []string `json:"allow"`
			JobsToday  int      `json:"jobs_today"`
			LastResult string   `json:"last_result"`
		} `json:"printers"`
		UIURL   string `json:"ui_url"`
		LogPath string `json:"log_path"`
	}
	if !strings.HasPrefix(rec.Header().Get("Content-Type"), "application/json") || json.Unmarshal(rec.Body.Bytes(), &st) != nil {
		t.Fatalf("/api/status not JSON: %d %q", rec.Code, rec.Body.String())
	}
	if len(st.Printers) != 2 || st.Printers[0].Name != "PrnAlpha" || st.Printers[0].Port != 9101 || st.Printers[0].JobsToday != 4711 ||
		st.Printers[0].LastResult != "printed" || !reflect.DeepEqual(st.Printers[0].Allow, []string{"10.0.0.0/8"}) ||
		st.Printers[1].Allow != nil || st.UIURL != "http://127.0.0.1:9180/" || st.LogPath != d.LogPath {
		t.Fatalf("/api/status = %+v", st)
	}

	var ps []map[string]interface{}
	if err := json.Unmarshal(get(s, "/api/printers").Body.Bytes(), &ps); err != nil || len(ps) != 2 {
		t.Fatalf("/api/printers: %v %v", ps, err)
	}
	if ps[0]["name"] != "PrnAlpha" || ps[0]["printer"] != "Canon" || ps[0]["port"] != 9101.0 {
		t.Fatalf("/api/printers[0] = %v", ps[0])
	}
	if v, ok := ps[1]["allow"]; !ok || v != nil {
		t.Fatalf("/api/printers[1].allow = %v (present %v), want null", v, ok)
	}

	var inst []string
	if err := json.Unmarshal(get(s, "/api/installed-printers").Body.Bytes(), &inst); err != nil || !reflect.DeepEqual(inst, []string{"Canon", "HP"}) {
		t.Fatalf("/api/installed-printers = %v %v", inst, err)
	}
}

// TC-UI-063, TC-UI-064
func TestREQ_UI_018_PostPrintersValidatesAndSaves(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, oneCfg)
	before := readFile(t, u.path)
	rec := postJSON(u.srv, "/api/printers", `[{"name":"A","printer":"Canon","port":70000}]`, testToken)
	var e struct{ Error, Field string }
	if rec.Code != 400 || json.Unmarshal(rec.Body.Bytes(), &e) != nil || e.Field != "printers[0].port" ||
		!strings.HasSuffix(e.Error, "invalid field printers[0].port") {
		t.Fatalf("invalid list = %d %q, want 400 with error and field printers[0].port", rec.Code, rec.Body.String())
	}
	if readFile(t, u.path) != before {
		t.Fatal("invalid API save wrote the config")
	}
	rec = postJSON(u.srv, "/api/printers", `[{"name":"A","printer":"Canon","port":9105}]`, testToken)
	if rec.Code != 200 || loadCfg(t, u.path).Printers[0].Port != 9105 || u.applied.n() != 1 {
		t.Fatalf("valid API save = %d, file %s, applied %d", rec.Code, readFile(t, u.path), u.applied.n())
	}
}

// TC-UI-065
func TestREQ_UI_018_LogLinesParam(t *testing.T) {
	dir, done := tmpDir(t)
	defer done()
	lp := filepath.Join(dir, "printerone.log")
	logLines(t, lp, 250)
	s := New(Deps{Token: testToken, LogPath: lp})
	count := func(q string) (int, []string) {
		rec := get(s, "/api/log"+q)
		var r struct{ Lines []string }
		if json.Unmarshal(rec.Body.Bytes(), &r) != nil {
			return rec.Code, nil
		}
		return rec.Code, r.Lines
	}
	if code, l := count("?lines=5"); code != 200 || len(l) != 5 || !strings.HasPrefix(l[4], "line-250;") || !strings.HasPrefix(l[0], "line-246;") {
		t.Errorf("lines=5 → %d %q", code, l)
	}
	if code, l := count("?lines=1000"); code != 200 || len(l) != 250 {
		t.Errorf("lines=1000 → %d, %d lines, want all 250", code, len(l))
	}
	if code, l := count(""); code != 200 || len(l) != 200 || !strings.HasPrefix(l[0], "line-051;") {
		t.Errorf("no lines → %d, %d lines, want the last 200", code, len(l))
	}
	for _, q := range []string{"?lines=1001", "?lines=0", "?lines=-1", "?lines=abc"} {
		rec := get(s, "/api/log"+q)
		var e struct{ Error, Field string }
		if rec.Code != 400 || json.Unmarshal(rec.Body.Bytes(), &e) != nil || e.Field != "lines" || e.Error == "" {
			t.Errorf("%s → %d %q, want 400 field lines", q, rec.Code, rec.Body.String())
		}
	}
}

// TC-UI-066
func TestREQ_UI_018_PostTestSendsJob(t *testing.T) {
	sink := newJobSink(t)
	defer sink.ln.Close()
	dir, done := tmpDir(t)
	defer done()
	u := newFileUI(t, dir, sink.cfg())
	if rec := postJSON(u.srv, "/api/test", `{"name":"A"}`, testToken); rec.Code != 200 {
		t.Fatalf("POST /api/test = %d %q", rec.Code, rec.Body.String())
	}
	if job, ok := sink.wait(3 * time.Second); !ok || !strings.Contains(job, "2026-11-07") {
		t.Fatalf("no test job received (%q)", job)
	}
}
