// Package config finds, reads and validates config.json (spec docs/specs/2026-11-07-f001-core.md §2.1).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package config

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"

	"github.com/xtieume/PrinterOne/internal/core/allow"
	"github.com/xtieume/PrinterOne/internal/core/logfile"
)

// DefaultPort is used when a printer entry has no port.
const DefaultPort = 9100

// Printer is one shared printer: the Windows printer name and the TCP port it is served on.
type Printer struct {
	Name    string
	Printer string
	Port    int
}

// Log is the optional `log` section (spec docs/specs/2026-11-07-f002-logs.md §2).
type Log struct {
	Dir      string // absolute log folder; "" when log.dir is absent (then <data dir>\logs)
	KeepDays int    // 1–3650; 30 when log.keep_days is absent
}

// Config is a loaded, validated config file.
type Config struct {
	Path     string // the file it was read from
	Printers []Printer
	Allow    []allow.List // Allow[i] is Printers[i]'s `allow` list (F-004); nil entry: every client
	Log      Log
}

// Lookup holds the places Load searches, in order: Flag (the --config value, "" when absent), ExeDir, DataDir.
type Lookup struct {
	Flag    string
	ExeDir  string
	DataDir string
}

// Error is a config problem; Msg is the single line printed to stderr and the program exits with code 2.
type Error struct {
	Msg string
}

func (e *Error) Error() string { return e.Msg }

// DataDir returns %ProgramData%\PrinterOne, or %ALLUSERSPROFILE%\Application Data\PrinterOne when
// %ProgramData% is not set (Windows XP).
func DataDir(getenv func(string) string) string {
	if getenv == nil {
		getenv = os.Getenv
	}
	if p := getenv("ProgramData"); p != "" {
		return filepath.Join(p, "PrinterOne")
	}
	return filepath.Join(getenv("ALLUSERSPROFILE"), "Application Data", "PrinterOne")
}

// Load finds the config by the lookup order and validates it. Every failure is an *Error.
func Load(l Lookup) (*Config, error) {
	if l.Flag != "" {
		return loadFile(l.Flag)
	}
	if l.ExeDir != "" {
		p := filepath.Join(l.ExeDir, "config.json")
		if _, err := os.Stat(p); err == nil {
			return loadFile(p)
		} else if !os.IsNotExist(err) {
			return nil, fileError(p, err)
		}
	}
	if l.DataDir != "" {
		p := filepath.Join(l.DataDir, "config.json")
		if _, err := os.Stat(p); err == nil {
			return loadFile(p)
		} else if !os.IsNotExist(err) {
			return nil, fileError(p, err)
		}
	}
	return nil, &Error{Msg: fmt.Sprintf("no config found in %s; run printerone printers", l.DataDir)}
}

// LoadLog finds the config by the lookup order and reads only its `log` section, for `printerone logs`: the printer
// list is not validated. With no config file it returns the defaults. An invalid log section is an *Error.
func LoadLog(l Lookup) (Log, error) {
	path := ""
	if l.Flag != "" {
		path = l.Flag
	} else if l.ExeDir != "" {
		p := filepath.Join(l.ExeDir, "config.json")
		if _, err := os.Stat(p); err == nil {
			path = p
		} else if !os.IsNotExist(err) {
			return Log{}, fileError(p, err)
		}
	}
	if path == "" && l.DataDir != "" {
		p := filepath.Join(l.DataDir, "config.json")
		if _, err := os.Stat(p); err == nil {
			path = p
		} else if !os.IsNotExist(err) {
			return Log{}, fileError(p, err)
		}
	}
	if path == "" {
		return Log{KeepDays: logfile.DefaultKeepDays}, nil
	}
	b, err := ioutil.ReadFile(path)
	if err != nil {
		return Log{}, fileError(path, err)
	}
	if len(b) >= 3 && b[0] == 0xef && b[1] == 0xbb && b[2] == 0xbf {
		b = b[3:]
	}
	var root map[string]json.RawMessage
	if err = json.Unmarshal(b, &root); err != nil {
		return Log{}, jsonError(path, b, err)
	}
	if root == nil {
		return Log{}, &Error{Msg: fmt.Sprintf("%s: top level must be a JSON object", path)}
	}
	return parseLog(root, path)
}

func loadFile(path string) (*Config, error) {
	b, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, fileError(path, err)
	}
	if len(b) >= 3 && b[0] == 0xef && b[1] == 0xbb && b[2] == 0xbf {
		b = b[3:]
	}
	var top interface{}
	if err := json.Unmarshal(b, &top); err != nil {
		return nil, jsonError(path, b, err)
	}
	if _, ok := top.(map[string]interface{}); !ok {
		return nil, &Error{Msg: fmt.Sprintf("%s: top level must be a JSON object", path)}
	}
	var root map[string]json.RawMessage
	if err := json.Unmarshal(b, &root); err != nil {
		return nil, jsonError(path, b, err)
	}
	if string(root["version"]) != "2" {
		return nil, fieldError(path, "version")
	}
	var entries []json.RawMessage
	if len(root["printers"]) == 0 || json.Unmarshal(root["printers"], &entries) != nil || len(entries) < 1 || len(entries) > 32 {
		return nil, fieldError(path, "printers")
	}
	printers := make([]Printer, 0, len(entries))
	var cfgAllow []allow.List
	ports, names := make(map[int]bool), make(map[string]bool)
	for i, raw := range entries {
		prefix := fmt.Sprintf("printers[%d]", i)
		var entry map[string]json.RawMessage
		if json.Unmarshal(raw, &entry) != nil || entry == nil {
			return nil, fieldError(path, prefix)
		}
		var printer string
		if len(entry["printer"]) == 0 || json.Unmarshal(entry["printer"], &printer) != nil || printer == "" {
			return nil, fieldError(path, prefix+".printer")
		}
		name := "default"
		if len(entries) > 1 {
			name = fmt.Sprintf("printer%d", i+1)
		}
		if value, ok := entry["name"]; ok {
			var n string
			if json.Unmarshal(value, &n) == nil && n != "" {
				name = n
			}
		}
		port := DefaultPort
		if value, ok := entry["port"]; ok {
			n, e := strconv.Atoi(string(value))
			if e != nil || n < 1 || n > 65535 {
				return nil, fieldError(path, prefix+".port")
			}
			port = n
		}
		printers = append(printers, Printer{Name: name, Printer: printer, Port: port})
		if allowRaw, ok := entry["allow"]; ok {
			var rawItems []json.RawMessage
			if json.Unmarshal(allowRaw, &rawItems) != nil || len(rawItems) == 0 {
				return nil, fieldError(path, prefix+".allow")
			}
			items := make([]string, len(rawItems))
			for j, value := range rawItems {
				if json.Unmarshal(value, &items[j]) != nil {
					return nil, fieldError(path, fmt.Sprintf("%s.allow[%d]", prefix, j))
				}
			}
			list, bad, e := allow.Parse(items)
			if e != nil {
				return nil, fieldError(path, fmt.Sprintf("%s.allow[%d]", prefix, bad))
			}
			cfgAllow = append(cfgAllow, list)
		} else {
			cfgAllow = append(cfgAllow, nil)
		}
	}
	// Validate every entry's fields before reporting a duplicate across entries.
	for i, p := range printers {
		if ports[p.Port] {
			return nil, fieldError(path, fmt.Sprintf("printers[%d].port", i))
		}
		ports[p.Port] = true
		if names[p.Name] {
			return nil, fieldError(path, fmt.Sprintf("printers[%d].name", i))
		}
		names[p.Name] = true
	}
	lg, err := parseLog(root, path)
	if err != nil {
		return nil, err
	}
	return &Config{Path: path, Log: lg, Printers: printers, Allow: cfgAllow}, nil
}

func fileError(path string, err error) error { return &Error{Msg: fmt.Sprintf("%s: %v", path, err)} }
func fieldError(path, field string) error {
	return &Error{Msg: fmt.Sprintf("%s: invalid field %s", path, field)}
}
func jsonError(path string, data []byte, err error) error {
	line, col := 1, 1
	if se, ok := err.(*json.SyntaxError); ok {
		offset := int(se.Offset)
		if offset > len(data) {
			offset = len(data)
		}
		for _, c := range data[:offset] {
			if c == '\n' {
				line++
				col = 1
			} else {
				col++
			}
		}
	}
	return &Error{Msg: fmt.Sprintf("%s: invalid JSON at %d:%d", path, line, col)}
}

func parseLog(root map[string]json.RawMessage, path string) (Log, error) {
	l := Log{KeepDays: logfile.DefaultKeepDays}
	raw, ok := root["log"]
	if !ok {
		return l, nil
	}
	var obj map[string]json.RawMessage
	if json.Unmarshal(raw, &obj) != nil || obj == nil {
		return Log{}, fieldError(path, "log")
	}
	if v, ok := obj["dir"]; ok {
		if json.Unmarshal(v, &l.Dir) != nil || l.Dir == "" {
			return Log{}, fieldError(path, "log.dir")
		}
		if !filepath.IsAbs(l.Dir) && !isWindowsAbs(l.Dir) {
			l.Dir = filepath.Join(filepath.Dir(path), l.Dir)
		}
		if !isWindowsAbs(l.Dir) {
			l.Dir, _ = filepath.Abs(l.Dir)
		}
	}
	if v, ok := obj["keep_days"]; ok {
		n, e := strconv.Atoi(string(v))
		if e != nil || n < 1 || n > 3650 {
			return Log{}, fieldError(path, "log.keep_days")
		}
		l.KeepDays = n
	}
	return l, nil
}
func isWindowsAbs(p string) bool {
	return len(p) >= 3 && ((p[0] >= 'A' && p[0] <= 'Z') || (p[0] >= 'a' && p[0] <= 'z')) && p[1] == ':' && (p[2] == '\\' || p[2] == '/')
}
