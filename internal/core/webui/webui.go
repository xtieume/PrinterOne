// Package webui is the browser UI and JSON API served by PrinterOne (spec docs/specs/2026-11-07-f005-web-ui.md).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code): no embed, so static/ is compiled in through
// assets_gen.go, produced by `go generate` (gen/main.go).
//
// Routes. Pages (html/template, every action a plain <form method="post">): GET / (status), GET+POST /printers,
// GET+POST /test, GET /log, GET+POST /settings. Assets: GET /base.css, /modern.css, /app.js. API: GET /api/status,
// GET /api/printers, GET /api/installed-printers, POST /api/printers, POST /api/test, GET /api/log?lines=N,
// GET /api/log/stream (SSE).
//
// Markup contract the tests rely on: every form carries <input type="hidden" name="token" value="…">; a field error
// is rendered right after its input as <span class="error" data-field="<config field>">message</span>.
//
// Form fields. POST /printers: token, action (add | edit | remove), index (edit, remove), name, printer (chosen
// from the installed list), printer_text (typed; wins when non-empty), port, allow (one item per line, blank lines
// ignored). POST /test: token, name. POST /settings: token, log_dir, keep_days, ui_port, remote ("on"), password
// (empty keeps the stored one). A valid POST answers 303 to its page; an invalid one re-renders the page (200)
// with the loader's message next to the field and writes nothing.
//
// API writes need header X-PrinterOne-Token. JSON: /api/status {"printers":[{"name","printer","port","allow",
// "jobs_today","last_result"}],"ui_url","log_path"}; /api/printers and POST /api/printers body [{"name","printer",
// "port","allow"}] (allow null: every client); /api/installed-printers ["…"]; POST /api/test {"name"};
// /api/log {"lines":["…"]} (lines default 200, 1–1000). Errors: 400 {"error":"…","field":"…"}, 403 bad token,
// 401 missing Basic auth when remote.
package webui

//go:generate go run gen/main.go ../../../static assets_gen.go

import (
	"crypto/rand"
	"encoding/hex"
	"fmt"
	"net"
	"net/http"
	"os/exec"
	"strconv"
	"sync"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/allow"
	"github.com/xtieume/PrinterOne/internal/core/config"
)

// LogPageLines is how many log lines /log shows; MaxAPILogLines bounds /api/log?lines=N.
const (
	LogPageLines   = 200
	MaxAPILogLines = 1000
)

// TokenHeader carries the per-process token on API writes.
const TokenHeader = "X-PrinterOne-Token"

// BasicUser is the only user name accepted with ui.remote.
const BasicUser = "admin"

// Deps is everything the UI takes from the process, so tests can drive it.
type Deps struct {
	Config     *config.Config // current config; nil: none yet (empty printer list)
	ConfigPath string         // file a save writes (Config.Path, or <data dir>\config.json when there is none)
	// Load validates a written candidate exactly like start-up; nil: config.Load(config.Lookup{Flag: path}).
	Load      func(path string) (*config.Config, error)
	Installed func() ([]string, error) // installed Windows printers (spool.List in production)
	// Stats gives a printer's jobs today and last result (log result text); nil: 0 and "".
	Stats   func(name string) (jobsToday int, lastResult string)
	LogPath string // the current log file
	// Apply is called after a valid printer save with the new config (Listeners.Apply in production).
	Apply func(c *config.Config) error
	Token string    // per-process form / API token; "" → New makes a random one
	UI    config.UI // the UI settings the process started with (Basic auth when Remote)
	Now   func() time.Time
	Poll  time.Duration // SSE log poll interval; 0 → 500 ms
}

// Server is the UI's http.Handler.
type Server struct {
	d               Deps
	mu              sync.Mutex
	token           string
	restartRequired bool
}

// New returns the UI handler for d.
func New(d Deps) *Server {
	if d.Token == "" {
		b := make([]byte, 24)
		if _, err := rand.Read(b); err == nil {
			d.Token = hex.EncodeToString(b)
		} else {
			d.Token = fmt.Sprintf("%d", time.Now().UnixNano())
		}
	}
	return &Server{d: d, token: d.Token}
}

// Token is the per-process token every form and API write must carry.
func (s *Server) Token() string {
	return s.token
}

// ServeHTTP serves every route listed in the package comment.
func (s *Server) ServeHTTP(w http.ResponseWriter, r *http.Request) { s.serveHTTP(w, r) }

// URL is the address the UI is opened at: http://127.0.0.1:<port>/ (also with remote access on).
func URL(port int) string {
	return fmt.Sprintf("http://127.0.0.1:%d/", port)
}

// Start binds the UI on 127.0.0.1:<ui.Port>, or 0.0.0.0:<ui.Port> (tcp4) with ui.Remote, and serves h in the
// background. Port 0 picks a free port (tests). open, non-nil only for `printerone` with no arguments, is called
// once with URL(<bound port>) after the bind. addr is the bound address; stop closes the listener.
func Start(ui config.UI, h http.Handler, open func(url string) error) (addr string, stop func(), err error) {
	host := "127.0.0.1"
	if ui.Remote {
		host = "0.0.0.0"
	}
	ln, err := net.Listen("tcp4", net.JoinHostPort(host, strconv.Itoa(ui.Port)))
	if err != nil {
		return "", nil, err
	}
	server := &http.Server{Handler: h, ReadHeaderTimeout: 5 * time.Second}
	go server.Serve(ln)
	addr = ln.Addr().String()
	if open != nil {
		_, port, _ := net.SplitHostPort(addr)
		_ = open(URL(mustPort(port)))
	}
	return addr, func() { server.Close() }, nil
}

func mustPort(v string) int { n, _ := strconv.Atoi(v); return n }

// BrowserCommand is the command that opens url in the default browser on XP to 11:
// rundll32 url.dll,FileProtocolHandler <url>.
func BrowserCommand(url string) []string {
	return []string{"rundll32", "url.dll,FileProtocolHandler", url}
}

// OpenBrowser runs BrowserCommand(url) without waiting for it.
func OpenBrowser(url string) error {
	args := BrowserCommand(url)
	return exec.Command(args[0], args[1:]...).Start()
}

// ListenFunc starts one printer's receiver; stop shuts it down gracefully (F-001 shutdown: jobs already received
// are still printed).
type ListenFunc func(p config.Printer, a allow.List) (stop func(), err error)

// Listeners keeps one running receiver per printer and restarts, on Apply, only those whose port, Windows printer
// or allowlist changed (a name-only change keeps the running listener).
type Listeners struct {
	Listen  ListenFunc
	running map[string]listener
	mu      sync.Mutex
}

// Apply makes the running listeners match c: stops removed and changed ones, starts new and changed ones, leaves
// the others running. A start error is returned after the rest were applied.
func (l *Listeners) Apply(c *config.Config) error {
	l.mu.Lock()
	defer l.mu.Unlock()
	if l.Listen == nil {
		return nil
	}
	if c == nil {
		c = &config.Config{}
	}
	if l.running == nil {
		l.running = make(map[string]listener)
	}
	next := make(map[string]listener)
	stopped := make(map[string]bool)
	var first error
	for i, p := range c.Printers {
		var a allow.List
		if i < len(c.Allow) {
			a = c.Allow[i]
		}
		key := listenerKey(p, a)
		found := ""
		for name, old := range l.running {
			if old.key == key {
				found = name
				break
			}
		}
		if found != "" {
			next[p.Name] = l.running[found]
			continue
		}
		old, hadOld := l.running[p.Name]
		if hadOld {
			old.stop()
			stopped[p.Name] = true
		}
		stop, err := l.Listen(p, a)
		if err != nil {
			if first == nil {
				first = err
			}
			continue
		}
		next[p.Name] = listener{key: key, stop: stop}
	}
	for name, old := range l.running {
		kept := false
		for _, current := range next {
			if current.key == old.key {
				kept = true
				break
			}
		}
		if !kept && !stopped[name] {
			old.stop()
		}
	}
	l.running = next
	return first
}

type listener struct {
	key  string
	stop func()
}

func listenerKey(p config.Printer, a allow.List) string {
	return fmt.Sprintf("%d\x00%s\x00%s", p.Port, p.Printer, a.String())
}
