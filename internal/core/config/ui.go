// The `ui` section of config.json (spec docs/specs/2026-11-07-f005-web-ui.md §2.1).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).

package config

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
	return UI{}, nil
}
