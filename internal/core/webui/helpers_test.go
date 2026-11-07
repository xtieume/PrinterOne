package webui

// Shared helpers for the F-005 tests (spec docs/specs/2026-11-07-f005-web-ui.md). Go 1.10: no t.TempDir/t.Cleanup.

import (
	"encoding/json"
	"io/ioutil"
	"net"
	"net/http"
	"net/http/httptest"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/allow"
	"github.com/xtieume/PrinterOne/internal/core/config"
)

const testToken = "tok123"

const oneCfg = `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101}]}`

func tmpDir(t *testing.T) (string, func()) {
	dir, err := ioutil.TempDir("", "webui")
	if err != nil {
		t.Fatal(err)
	}
	return dir, func() { os.RemoveAll(dir) }
}

func writeFile(t *testing.T, path, body string) {
	if err := ioutil.WriteFile(path, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
}

func readFile(t *testing.T, path string) string {
	b, err := ioutil.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

func loadCfg(t *testing.T, path string) *config.Config {
	c, err := config.Load(config.Lookup{Flag: path})
	if err != nil {
		t.Fatalf("config.Load(%s): %v", path, err)
	}
	return c
}

// applyRec records every Apply call.
type applyRec struct {
	mu    sync.Mutex
	calls []*config.Config
}

func (a *applyRec) apply(c *config.Config) error {
	a.mu.Lock()
	a.calls = append(a.calls, c)
	a.mu.Unlock()
	return nil
}

func (a *applyRec) n() int {
	a.mu.Lock()
	defer a.mu.Unlock()
	return len(a.calls)
}

// fileUI is a UI over a config file at <dir>/config.json holding body.
type fileUI struct {
	dir, path string
	srv       *Server
	applied   *applyRec
}

func newFileUI(t *testing.T, dir, body string) *fileUI {
	p := filepath.Join(dir, "config.json")
	writeFile(t, p, body)
	a := &applyRec{}
	d := Deps{Config: loadCfg(t, p), ConfigPath: p, Token: testToken, Apply: a.apply,
		Installed: func() ([]string, error) { return []string{"Canon", "HP"}, nil },
		Now:       func() time.Time { return time.Date(2026, 11, 7, 10, 20, 30, 0, time.Local) }}
	return &fileUI{dir: dir, path: p, srv: New(d), applied: a}
}

func get(h http.Handler, path string) *httptest.ResponseRecorder {
	rec := httptest.NewRecorder()
	h.ServeHTTP(rec, httptest.NewRequest("GET", path, nil))
	return rec
}

func getAuth(h http.Handler, path, user, pass string) *httptest.ResponseRecorder {
	rec := httptest.NewRecorder()
	r := httptest.NewRequest("GET", path, nil)
	if user != "" || pass != "" {
		r.SetBasicAuth(user, pass)
	}
	h.ServeHTTP(rec, r)
	return rec
}

func postForm(h http.Handler, path string, v url.Values) *httptest.ResponseRecorder {
	rec := httptest.NewRecorder()
	r := httptest.NewRequest("POST", path, strings.NewReader(v.Encode()))
	r.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	h.ServeHTTP(rec, r)
	return rec
}

func postJSON(h http.Handler, path, body, token string) *httptest.ResponseRecorder {
	rec := httptest.NewRecorder()
	r := httptest.NewRequest("POST", path, strings.NewReader(body))
	r.Header.Set("Content-Type", "application/json")
	if token != "" {
		r.Header.Set(TokenHeader, token)
	}
	h.ServeHTTP(rec, r)
	return rec
}

var tokenRe = regexp.MustCompile(`name="token"\s+value="([^"]+)"`)

// pageToken reads the hidden token from a page, as a browser without JS would submit it.
func pageToken(t *testing.T, h http.Handler, path string) string {
	rec := get(h, path)
	m := tokenRe.FindStringSubmatch(rec.Body.String())
	if rec.Code != 200 || m == nil {
		t.Fatalf("GET %s: status %d, no hidden token input in page", path, rec.Code)
	}
	return m[1]
}

func vals(kv ...string) url.Values {
	v := url.Values{}
	for i := 0; i+1 < len(kv); i += 2 {
		v.Set(kv[i], kv[i+1])
	}
	return v
}

// mustRedirect checks a successful PRG answer.
func mustRedirect(t *testing.T, rec *httptest.ResponseRecorder, to string) {
	if rec.Code != http.StatusSeeOther || rec.Header().Get("Location") != to {
		t.Fatalf("want 303 to %s, got %d Location %q body %.300s", to, rec.Code, rec.Header().Get("Location"), rec.Body.String())
	}
}

// rawCfg decodes the saved file generically.
func rawCfg(t *testing.T, path string) map[string]interface{} {
	var m map[string]interface{}
	if err := json.Unmarshal([]byte(readFile(t, path)), &m); err != nil {
		t.Fatalf("saved config is not JSON: %v", err)
	}
	return m
}

func dirNames(t *testing.T, dir string) []string {
	fis, err := ioutil.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	var names []string
	for _, fi := range fis {
		names = append(names, fi.Name())
	}
	sort.Strings(names)
	return names
}

func mustOnlyConfig(t *testing.T, dir string) {
	if n := dirNames(t, dir); len(n) != 1 || n[0] != "config.json" {
		t.Fatalf("config folder holds %v, want only config.json", n)
	}
}

// literal builds a config without a file; allows[i] are Printers[i]'s items (nil: every client).
func literal(t *testing.T, ps []config.Printer, allows [][]string) *config.Config {
	c := &config.Config{Printers: ps, Allow: make([]allow.List, len(ps))}
	for i := range ps {
		if i < len(allows) && allows[i] != nil {
			l, _, err := allow.Parse(allows[i])
			if err != nil {
				t.Fatal(err)
			}
			c.Allow[i] = l
		}
	}
	return c
}

// jobSink is a TCP listener on 127.0.0.1 that collects the bytes of each connection.
type jobSink struct {
	ln   net.Listener
	jobs chan string
}

func newJobSink(t *testing.T) *jobSink {
	ln, err := net.Listen("tcp4", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	s := &jobSink{ln: ln, jobs: make(chan string, 4)}
	go func() {
		for {
			c, err := ln.Accept()
			if err != nil {
				return
			}
			c.SetReadDeadline(time.Now().Add(3 * time.Second))
			b, _ := ioutil.ReadAll(c)
			c.Close()
			s.jobs <- string(b)
		}
	}()
	return s
}

func (s *jobSink) port() int { return s.ln.Addr().(*net.TCPAddr).Port }

func (s *jobSink) cfg() string {
	return `{"version":2,"printers":[{"name":"A","printer":"Canon","port":` + strconv.Itoa(s.port()) + `}]}`
}

func (s *jobSink) wait(d time.Duration) (string, bool) {
	select {
	case j := <-s.jobs:
		return j, true
	case <-time.After(d):
		return "", false
	}
}

func logLines(t *testing.T, path string, n int) {
	var b strings.Builder
	for i := 1; i <= n; i++ {
		b.WriteString("line-" + pad3(i) + "; job\n")
	}
	writeFile(t, path, b.String())
}

func pad3(i int) string {
	s := strconv.Itoa(i)
	for len(s) < 3 {
		s = "0" + s
	}
	return s
}
