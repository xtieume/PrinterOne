// The `ui` section of config.json (spec docs/specs/2026-11-07-f005-web-ui.md §2.1).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).

package config

import (
	"encoding/json"
	"io/ioutil"
	"strconv"
	"unicode/utf8"
)

// DefaultUIPort is the web UI port when `ui.port` is absent.
const DefaultUIPort = 9180

// MinUIPasswordLen is the shortest `ui.password` accepted when `ui.remote` is true.
const MinUIPasswordLen = 8

// UI is the optional `ui` section: {"port": 9180, "remote": false, "password": "…"}.
type UI struct {
	Port     int    // 1–65535, never equal to a printer port; DefaultUIPort when absent
	Remote   bool   // true: bind 0.0.0.0 and require HTTP Basic admin / Password
	Password string // checked (>= MinUIPasswordLen) only when Remote is true
}

// UIOf reads and validates the `ui` section of c's file (c.Path) against c's printer ports. A nil c (no config
// yet) gives the defaults. Every failure is an *Error naming the field: `ui` (not an object), `ui.port` (not an
// integer, outside 1–65535, or equal to a printer port), `ui.remote` (not a boolean), `ui.password` (not a string,
// or shorter than MinUIPasswordLen / missing while remote is true). The caller exits 2 on it, as for Load.
func UIOf(c *Config) (UI, error) {
	u := UI{Port: DefaultUIPort}
	if c == nil || c.Path == "" {
		return u, nil
	}
	b, err := ioutil.ReadFile(c.Path)
	if err != nil {
		return UI{}, fileError(c.Path, err)
	}
	if len(b) >= 3 && b[0] == 0xef && b[1] == 0xbb && b[2] == 0xbf {
		b = b[3:]
	}
	var root map[string]json.RawMessage
	if err = json.Unmarshal(b, &root); err != nil {
		return UI{}, jsonError(c.Path, b, err)
	}
	raw, ok := root["ui"]
	obj := map[string]json.RawMessage{}
	if ok && (json.Unmarshal(raw, &obj) != nil || obj == nil) {
		return UI{}, fieldError(c.Path, "ui")
	}
	if v, ok := obj["port"]; ok {
		n, parseErr := strconv.Atoi(string(v))
		if parseErr != nil {
			return UI{}, fieldError(c.Path, "ui.port")
		}
		u.Port = n
	}
	if u.Port < 1 || u.Port > 65535 {
		return UI{}, fieldError(c.Path, "ui.port")
	}
	for _, p := range c.Printers {
		if p.Port == u.Port {
			return UI{}, fieldError(c.Path, "ui.port")
		}
	}
	if v, ok := obj["remote"]; ok {
		if string(v) != "true" && string(v) != "false" {
			return UI{}, fieldError(c.Path, "ui.remote")
		}
		u.Remote = string(v) == "true"
	}
	if v, ok := obj["password"]; ok {
		if json.Unmarshal(v, &u.Password) != nil || len(v) == 0 || v[0] != '"' {
			return UI{}, fieldError(c.Path, "ui.password")
		}
	}
	if u.Remote && (!utf8.ValidString(u.Password) || utf8.RuneCountInString(u.Password) < MinUIPasswordLen) {
		return UI{}, fieldError(c.Path, "ui.password")
	}
	return u, nil
}
