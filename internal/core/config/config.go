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
func LoadLog(l Lookup) (Log, error) { return Log{}, nil }

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
	if len(root["printers"]) == 0 || json.Unmarshal(root["printers"], &entries) != nil || len(entries) != 1 {
		return nil, fieldError(path, "printers")
	}
	var entry map[string]json.RawMessage
	if json.Unmarshal(entries[0], &entry) != nil || entry == nil {
		return nil, fieldError(path, "printers[0]")
	}
	var printer string
	if len(entry["printer"]) == 0 || json.Unmarshal(entry["printer"], &printer) != nil || printer == "" {
		return nil, fieldError(path, "printers[0].printer")
	}
	name := "default"
	if raw, ok := entry["name"]; ok {
		var n string
		if json.Unmarshal(raw, &n) == nil && n != "" {
			name = n
		}
	}
	port := DefaultPort
	if raw, ok := entry["port"]; ok {
		v, err := strconv.Atoi(string(raw))
		if err != nil || v < 1 || v > 65535 {
			return nil, fieldError(path, "printers[0].port")
		}
		port = v
	}
	return &Config{Path: path, Printers: []Printer{{Name: name, Printer: printer, Port: port}}}, nil
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
