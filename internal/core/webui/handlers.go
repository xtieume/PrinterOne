package webui

import (
	"bufio"
	"encoding/json"
	"fmt"
	"html/template"
	"io"
	"io/ioutil"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/config"
)

type pageData struct {
	Title, Token, Error, Field, LogPath, UIURL string
	Printers                                   []printerRow
	Installed                                  []string
	Lines                                      []string
	LogDir                                     string
	KeepDays, UIPort                           int
	Remote                                     bool
	Inputs                                     map[string]string
}
type printerRow struct {
	Name, Printer, Allow, AllowForm, Last string
	Port, Jobs                            int
}

func (s *Server) serveHTTP(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("X-Frame-Options", "DENY")
	if s.d.UI.Remote {
		user, pass, ok := r.BasicAuth()
		if !ok || user != BasicUser || pass != s.d.UI.Password {
			w.Header().Set("WWW-Authenticate", `Basic realm="PrinterOne"`)
			http.Error(w, "authentication required", http.StatusUnauthorized)
			return
		}
	} else if s.d.UI.Port != 0 && localRemote(r.RemoteAddr) {
		port := s.d.UI.Port
		if port == 0 {
			port = config.DefaultUIPort
		}
		host := r.Host
		if host != net.JoinHostPort("127.0.0.1", strconv.Itoa(port)) && host != net.JoinHostPort("localhost", strconv.Itoa(port)) {
			http.Error(w, "invalid Host", http.StatusForbidden)
			return
		}
	}
	if strings.HasPrefix(r.URL.Path, "/api/") && r.Method == http.MethodPost {
		if r.Header.Get(TokenHeader) != s.token {
			writeJSON(w, http.StatusForbidden, map[string]string{"error": "invalid token"})
			return
		}
	} else if r.Method == http.MethodPost && r.URL.Path != "/" {
		if r.ParseForm() != nil || r.Form.Get("token") != s.token {
			http.Error(w, "invalid token", http.StatusForbidden)
			return
		}
	}
	switch r.URL.Path {
	case "/":
		if r.Method == "GET" {
			s.render(w, "status", s.data("Status"))
		} else {
			http.NotFound(w, r)
		}
	case "/printers":
		if r.Method == "GET" {
			s.printersPage(w)
		} else if r.Method == "POST" {
			s.savePrinters(w, r)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/test":
		if r.Method == "GET" {
			s.render(w, "test", s.data("Test"))
		} else if r.Method == "POST" {
			s.testPrinter(w, r)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/log":
		if r.Method == "GET" {
			d := s.data("Log")
			d.Lines, _ = readTail(s.d.LogPath, LogPageLines)
			s.render(w, "log", d)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/settings":
		if r.Method == "GET" {
			s.render(w, "settings", s.data("Settings"))
		} else if r.Method == "POST" {
			s.saveSettings(w, r)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/api/status":
		if r.Method == "GET" {
			writeJSON(w, 200, s.statusJSON())
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/api/printers":
		if r.Method == "GET" {
			writeJSON(w, 200, s.printerJSON())
		} else if r.Method == "POST" {
			s.apiSavePrinters(w, r)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/api/installed-printers":
		if r.Method == "GET" {
			names, err := s.installed()
			if err != nil {
				writeJSON(w, 500, map[string]string{"error": err.Error()})
			} else {
				writeJSON(w, 200, names)
			}
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/api/test":
		if r.Method == "POST" {
			s.apiTest(w, r)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/api/log":
		if r.Method == "GET" {
			n := 200
			if v := r.URL.Query().Get("lines"); v != "" {
				x, e := strconv.Atoi(v)
				if e != nil || x < 1 || x > MaxAPILogLines {
					writeJSON(w, 400, map[string]string{"error": "invalid lines", "field": "lines"})
					return
				}
				n = x
			}
			lines, _ := readTail(s.d.LogPath, n)
			writeJSON(w, 200, map[string][]string{"lines": lines})
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/api/log/stream":
		if r.Method == "GET" {
			s.stream(w, r)
		} else {
			http.Error(w, "method not allowed", 405)
		}
	case "/base.css", "/modern.css", "/app.js":
		s.asset(w, r.URL.Path)
	default:
		http.NotFound(w, r)
	}
}
func localRemote(v string) bool {
	h, _, err := net.SplitHostPort(v)
	if err != nil {
		return false
	}
	ip := net.ParseIP(h)
	return ip != nil && ip.IsLoopback()
}
func (s *Server) data(title string) pageData {
	d := pageData{Title: title, Token: s.token, UIURL: URL(s.d.UI.Port), LogPath: s.d.LogPath, KeepDays: 30, UIPort: s.d.UI.Port, Remote: s.d.UI.Remote}
	cfg, restart := s.configSnapshot()
	if cfg != nil {
		d.LogDir = cfg.Log.Dir
		if cfg.Log.KeepDays != 0 {
			d.KeepDays = cfg.Log.KeepDays
		}
		if u, err := config.UIOf(cfg); err == nil {
			d.UIPort = u.Port
			d.Remote = u.Remote
		}
	}
	if restart {
		d.Error = "UI settings changed; restart PrinterOne for them to take effect."
	}
	for i, p := range s.printers() {
		jobs, last := 0, ""
		if s.d.Stats != nil {
			jobs, last = s.d.Stats(p.Name)
		}
		allowForm := ""
		if cfg != nil && i < len(cfg.Allow) && cfg.Allow[i] != nil {
			allowForm = strings.Replace(cfg.Allow[i].String(), ", ", "\n", -1)
		}
		d.Printers = append(d.Printers, printerRow{Name: p.Name, Printer: p.Printer, Port: p.Port, Allow: s.allowString(i), AllowForm: allowForm, Jobs: jobs, Last: last})
	}
	return d
}
func (s *Server) printers() []config.Printer {
	cfg, _ := s.configSnapshot()
	if cfg == nil {
		return nil
	}
	return cfg.Printers
}
func (s *Server) allowString(i int) string {
	cfg, _ := s.configSnapshot()
	if cfg == nil || i >= len(cfg.Allow) || cfg.Allow[i] == nil {
		return "allow: all"
	}
	return cfg.Allow[i].String()
}

func (s *Server) configSnapshot() (*config.Config, bool) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.d.Config == nil {
		return nil, s.restartRequired
	}
	c := *s.d.Config
	c.Printers = append([]config.Printer(nil), s.d.Config.Printers...)
	c.Allow = append(s.d.Config.Allow[:0:0], s.d.Config.Allow...)
	return &c, s.restartRequired
}
func (s *Server) installed() ([]string, error) {
	if s.d.Installed == nil {
		return nil, nil
	}
	return s.d.Installed()
}
func (s *Server) printersPage(w http.ResponseWriter) {
	d := s.data("Printers")
	d.Installed, _ = s.installed()
	s.render(w, "printers", d)
}
func (s *Server) render(w http.ResponseWriter, name string, d pageData) {
	body := assets[name+".html"]
	if body == "" {
		http.Error(w, "template missing", 500)
		return
	}
	t, e := template.New(name).Parse(body)
	if e != nil {
		http.Error(w, e.Error(), 500)
		return
	}
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	if e = t.Execute(w, d); e != nil {
		return
	}
}
func (s *Server) asset(w http.ResponseWriter, path string) {
	name := strings.TrimPrefix(path, "/")
	body, ok := assets[name]
	if !ok {
		http.NotFound(w, nil)
		return
	}
	if strings.HasSuffix(name, ".css") {
		w.Header().Set("Content-Type", "text/css; charset=utf-8")
	} else {
		w.Header().Set("Content-Type", "text/javascript; charset=utf-8")
	}
	io.WriteString(w, body)
}
func (s *Server) statusJSON() map[string]interface{} {
	cfg, _ := s.configSnapshot()
	rows := []map[string]interface{}{}
	for i, p := range s.printers() {
		j, l := 0, ""
		if s.d.Stats != nil {
			j, l = s.d.Stats(p.Name)
		}
		var a interface{}
		if cfg != nil && i < len(cfg.Allow) && cfg.Allow[i] != nil {
			parts := strings.Split(cfg.Allow[i].String(), ", ")
			a = parts
		}
		rows = append(rows, map[string]interface{}{"name": p.Name, "printer": p.Printer, "port": p.Port, "allow": a, "jobs_today": j, "last_result": l})
	}
	return map[string]interface{}{"printers": rows, "ui_url": URL(s.d.UI.Port), "log_path": s.d.LogPath}
}
func (s *Server) printerJSON() []map[string]interface{} {
	cfg, _ := s.configSnapshot()
	a := []map[string]interface{}{}
	for i, p := range s.printers() {
		var allowValue interface{}
		if cfg != nil && i < len(cfg.Allow) && cfg.Allow[i] != nil {
			allowValue = strings.Split(cfg.Allow[i].String(), ", ")
		}
		a = append(a, map[string]interface{}{"name": p.Name, "printer": p.Printer, "port": p.Port, "allow": allowValue})
	}
	return a
}
func writeJSON(w http.ResponseWriter, status int, v interface{}) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(status)
	json.NewEncoder(w).Encode(v)
}
func (s *Server) savePrinters(w http.ResponseWriter, r *http.Request) {
	if e := r.ParseForm(); e != nil {
		s.formErrorValues(w, "printers", "printers", e.Error(), r.Form)
		return
	}
	p := s.printers()
	idx, _ := strconv.Atoi(r.Form.Get("index"))
	action := r.Form.Get("action")
	if action == "add" {
		p = append(p, config.Printer{})
		idx = len(p) - 1
	}
	if idx < 0 || idx >= len(p) || (action != "add" && action != "edit" && action != "remove") {
		s.formErrorValues(w, "printers", "printers", "invalid action or printer index", r.Form)
		return
	}
	if action == "remove" {
		p = append(p[:idx], p[idx+1:]...)
	} else {
		printer := r.Form.Get("printer")
		if r.Form.Get("printer_text") != "" {
			printer = r.Form.Get("printer_text")
		}
		p[idx] = config.Printer{Name: r.Form.Get("name"), Printer: printer}
		p[idx].Port, _ = strconv.Atoi(r.Form.Get("port"))
	}
	raw := s.readRoot()
	old, _ := raw["printers"].([]interface{})
	parr := []interface{}{}
	for i, v := range p {
		var m map[string]interface{}
		if action != "add" && i < len(old) {
			m, _ = old[i].(map[string]interface{})
		}
		if m == nil {
			m = map[string]interface{}{}
		}
		m["name"], m["printer"], m["port"] = v.Name, v.Printer, v.Port
		parr = append(parr, m)
	}
	raw["printers"] = parr
	_, hasAllow := r.Form["allow"]
	if action != "remove" && hasAllow {
		items := []string{}
		for _, v := range strings.Split(strings.Replace(r.Form.Get("allow"), "\r", "", -1), "\n") {
			if strings.TrimSpace(v) != "" {
				items = append(items, strings.TrimSpace(v))
			}
		}
		if len(items) > 0 {
			parr[idx].(map[string]interface{})["allow"] = items
		} else {
			delete(parr[idx].(map[string]interface{}), "allow")
		}
	}
	if err := s.persist(raw); err != nil {
		field, msg := errParts(err)
		s.formErrorValues(w, field, "printers", msg, r.Form)
		return
	}
	http.Redirect(w, r, "/printers", 303)
}
func (s *Server) apiSavePrinters(w http.ResponseWriter, r *http.Request) {
	var p []config.Printer
	var wire []struct {
		Name    string   `json:"name"`
		Printer string   `json:"printer"`
		Port    int      `json:"port"`
		Allow   []string `json:"allow"`
	}
	if json.NewDecoder(r.Body).Decode(&wire) != nil {
		writeJSON(w, 400, map[string]string{"error": "invalid JSON", "field": "printers"})
		return
	}
	raw := s.readRoot()
	old, _ := raw["printers"].([]interface{})
	parr := []interface{}{}
	for i, v := range wire {
		p = append(p, config.Printer{Name: v.Name, Printer: v.Printer, Port: v.Port})
		var m map[string]interface{}
		if i < len(old) {
			m, _ = old[i].(map[string]interface{})
		}
		if m == nil {
			m = map[string]interface{}{}
		}
		m["name"], m["printer"], m["port"] = v.Name, v.Printer, v.Port
		if v.Allow != nil {
			m["allow"] = v.Allow
		} else {
			delete(m, "allow")
		}
		parr = append(parr, m)
	}
	raw["printers"] = parr
	if err := s.persist(raw); err != nil {
		f, msg := errParts(err)
		writeJSON(w, 400, map[string]string{"error": msg, "field": f})
		return
	}
	writeJSON(w, 200, s.printerJSON())
}
func (s *Server) testPrinter(w http.ResponseWriter, r *http.Request) {
	name := r.Form.Get("name")
	cfg, _ := s.configSnapshot()
	for _, p := range s.printers() {
		if p.Name == name {
			if i := indexPrinter(s.printers(), name); cfg != nil && i < len(cfg.Allow) && cfg.Allow[i] != nil && !cfg.Allow[i].Allows(net.ParseIP("127.0.0.1")) {
				s.formError(w, "name", "test", "test rejected by allowlist")
				return
			}
			payload := fmt.Sprintf("PrinterOne test page\r\nPrinter: %s\r\nTime: %s\r\n", p.Name, s.now().Format(time.RFC3339))
			c, e := net.DialTimeout("tcp4", fmt.Sprintf("127.0.0.1:%d", p.Port), 2*time.Second)
			if e == nil {
				c.SetWriteDeadline(time.Now().Add(3 * time.Second))
				_, e = io.WriteString(c, payload)
				if e == nil {
					if x, ok := c.(*net.TCPConn); ok {
						x.CloseWrite()
					}
				}
				c.Close()
			}
			if e != nil {
				s.formError(w, "name", "test", e.Error())
				return
			}
			http.Redirect(w, r, "/test", 303)
			return
		}
	}
	s.formError(w, "name", "test", "unknown printer")
}
func indexPrinter(p []config.Printer, n string) int {
	for i, v := range p {
		if v.Name == n {
			return i
		}
	}
	return -1
}
func (s *Server) apiTest(w http.ResponseWriter, r *http.Request) {
	var v struct {
		Name string `json:"name"`
	}
	if json.NewDecoder(r.Body).Decode(&v) != nil {
		writeJSON(w, 400, map[string]string{"error": "invalid JSON", "field": "name"})
		return
	}
	r.Form = map[string][]string{"name": {v.Name}}
	rec := &capture{header: make(http.Header)}
	s.testPrinter(rec, r)
	if rec.status == 303 {
		writeJSON(w, 200, map[string]string{"status": "sent"})
	} else {
		writeJSON(w, 400, map[string]string{"error": rec.body.String(), "field": "name"})
	}
}

type capture struct {
	header http.Header
	status int
	body   strings.Builder
}

func (c *capture) Header() http.Header         { return c.header }
func (c *capture) WriteHeader(n int)           { c.status = n }
func (c *capture) Write(b []byte) (int, error) { return c.body.Write(b) }
func (s *Server) saveSettings(w http.ResponseWriter, r *http.Request) {
	raw := s.readRoot()
	var root map[string]interface{}
	b, _ := json.Marshal(raw)
	json.Unmarshal(b, &root)
	lg, _ := root["log"].(map[string]interface{})
	if lg == nil {
		lg = map[string]interface{}{}
	}
	lg["keep_days"], _ = strconv.Atoi(r.Form.Get("keep_days"))
	if v := r.Form.Get("log_dir"); v != "" {
		lg["dir"] = v
	}
	root["log"] = lg
	ui, _ := root["ui"].(map[string]interface{})
	if ui == nil {
		ui = map[string]interface{}{}
	}
	port, _ := strconv.Atoi(r.Form.Get("ui_port"))
	ui["port"] = port
	ui["remote"] = r.Form.Get("remote") == "on"
	if p := r.Form.Get("password"); p != "" {
		ui["password"] = p
	}
	root["ui"] = ui
	if err := s.persist(root); err != nil {
		f, msg := errParts(err)
		s.formErrorValues(w, f, "settings", msg, r.Form)
		return
	}
	http.Redirect(w, r, "/settings?saved=1", 303)
}
func (s *Server) now() time.Time {
	if s.d.Now != nil {
		return s.d.Now()
	}
	return time.Now()
}
func (s *Server) formError(w http.ResponseWriter, field, page, msg string) {
	s.formErrorValues(w, field, page, msg, nil)
}
func (s *Server) formErrorValues(w http.ResponseWriter, field, page, msg string, values map[string][]string) {
	d := s.data(strings.Title(page))
	d.Field = field
	d.Error = msg
	d.Inputs = make(map[string]string)
	for key, list := range values {
		if key != "password" && len(list) > 0 {
			d.Inputs[key] = list[0]
		}
	}
	if page == "printers" {
		d.Installed, _ = s.installed()
	}
	s.render(w, page, d)
}
func (s *Server) readRoot() map[string]interface{} {
	m := map[string]interface{}{"version": 2}
	if s.d.ConfigPath != "" {
		b, e := ioutil.ReadFile(s.d.ConfigPath)
		if e == nil {
			json.Unmarshal(b, &m)
		}
	}
	if m["printers"] == nil {
		m["printers"] = []interface{}{}
	}
	return m
}
func (s *Server) persist(raw map[string]interface{}) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	oldUI, _ := config.UIOf(s.d.Config)
	path := s.d.ConfigPath
	if path == "" && s.d.Config != nil {
		path = s.d.Config.Path
	}
	if path == "" {
		return fmt.Errorf("config path is not set")
	}
	b, e := json.MarshalIndent(raw, "", "  ")
	if e != nil {
		return e
	}
	dir := filepath.Dir(path)
	f, e := ioutil.TempFile(dir, "config-*.tmp")
	if e != nil {
		return e
	}
	tmp := f.Name()
	ok := false
	defer func() {
		f.Close()
		if !ok {
			os.Remove(tmp)
		}
	}()
	if _, e = f.Write(b); e != nil {
		return e
	}
	if e = f.Sync(); e != nil {
		return e
	}
	if e = f.Close(); e != nil {
		return e
	}
	candidate, e := config.Load(config.Lookup{Flag: tmp})
	if e != nil {
		msg := strings.Replace(e.Error(), tmp, path, 1)
		return fmt.Errorf("%s", msg)
	}
	ui, e := config.UIOf(candidate)
	if e != nil {
		msg := strings.Replace(e.Error(), tmp, path, 1)
		return fmt.Errorf("%s", msg)
	}
	candidate.Path = path
	if e = os.Rename(tmp, path); e != nil {
		return e
	}
	ok = true
	s.d.Config = candidate
	s.restartRequired = ui.Port != oldUI.Port || ui.Remote != oldUI.Remote || ui.Password != oldUI.Password
	if s.d.Apply != nil {
		if e = s.d.Apply(candidate); e != nil {
			return fmt.Errorf("%v", e)
		}
	}
	return nil
}
func errParts(err error) (string, string) {
	msg := err.Error()
	field := "printers"
	if i := strings.Index(msg, "invalid field "); i >= 0 {
		field = msg[i+14:]
		if j := strings.IndexAny(field, " \n"); j >= 0 {
			field = field[:j]
		}
	}
	return field, msg
}
func readTail(path string, n int) ([]string, error) {
	if path == "" {
		return []string{}, nil
	}
	f, e := os.Open(path)
	if e != nil {
		if os.IsNotExist(e) {
			return []string{}, nil
		}
		return nil, e
	}
	defer func() { f.Close() }()
	out := []string{}
	sc := bufio.NewScanner(f)
	sc.Buffer(make([]byte, 4096), 1024*1024)
	for sc.Scan() {
		out = append(out, sc.Text())
		if len(out) > n {
			out = out[1:]
		}
	}
	return out, sc.Err()
}

// LogStats returns today's connection count and the newest result for a printer from the current log file.
func LogStats(path, printer string, now time.Time) (int, string) {
	f, err := os.Open(path)
	if err != nil {
		return 0, ""
	}
	defer f.Close()
	marker := ` "` + strings.Replace(printer, `"`, `\"`, -1) + `" `
	today, last := now.Format("2006-01-02"), ""
	count := 0
	sc := bufio.NewScanner(f)
	for sc.Scan() {
		line := sc.Text()
		if !strings.Contains(line, marker) {
			continue
		}
		i := strings.Index(line, marker)
		prefix := strings.Fields(line[:i])
		if len(prefix) < 1 {
			continue
		}
		result := strings.TrimSpace(line[i+len(marker):])
		last = result
		if strings.HasPrefix(prefix[0], today) {
			count++
		}
	}
	return count, last
}
func (s *Server) stream(w http.ResponseWriter, r *http.Request) {
	f, e := os.Open(s.d.LogPath)
	if e != nil {
		http.Error(w, "log unavailable", 500)
		return
	}
	defer func() { f.Close() }()
	f.Seek(0, io.SeekEnd)
	w.Header().Set("Content-Type", "text/event-stream")
	fl, ok := w.(http.Flusher)
	if !ok {
		return
	}
	fmt.Fprint(w, ": connected\n\n")
	fl.Flush()
	pos, _ := f.Seek(0, io.SeekCurrent)
	ticker := time.NewTicker(s.d.Poll)
	if s.d.Poll == 0 {
		ticker.Reset(500 * time.Millisecond)
	}
	defer ticker.Stop()
	for {
		select {
		case <-r.Context().Done():
			return
		case <-ticker.C:
			st, e := os.Stat(s.d.LogPath)
			if e != nil {
				continue
			}
			oldInfo, statErr := f.Stat()
			if statErr == nil && !os.SameFile(oldInfo, st) {
				f.Close()
				f, e = os.Open(s.d.LogPath)
				if e != nil {
					continue
				}
				pos = 0
			} else if st.Size() < pos {
				f.Close()
				f, e = os.Open(s.d.LogPath)
				if e != nil {
					continue
				}
				pos = 0
			}
			f.Seek(pos, io.SeekStart)
			sc := bufio.NewScanner(f)
			for sc.Scan() {
				fmt.Fprintf(w, "data: %s\n\n", sc.Text())
				fl.Flush()
			}
			pos, _ = f.Seek(0, io.SeekCurrent)
		}
	}
}
