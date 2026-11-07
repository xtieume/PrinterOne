// Package logfile writes the log lines of `serve` and `service run` to one shared, daily rotated file
// (spec docs/specs/2026-11-07-f002-logs.md §2).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package logfile

import (
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

const FileName = "printerone.log"
const DefaultKeepDays = 30

func DefaultDir(dataDir string) string { return filepath.Join(dataDir, "logs") }

func Path(dir string) string {
	return filepath.Join(dir, FileName)
}

type Writer struct {
	Dir      string
	KeepDays int
	Now      func() time.Time
	Warn     io.Writer
	mu       sync.Mutex
	warned   bool
}

var locksMu sync.Mutex
var locks = make(map[string]*sync.Mutex)

func pathLock(path string) *sync.Mutex {
	locksMu.Lock()
	defer locksMu.Unlock()
	m := locks[path]
	if m == nil {
		m = &sync.Mutex{}
		locks[path] = m
	}
	return m
}

func (w *Writer) Start() {
	now := w.now()
	if w.Dir == "" {
		return
	}
	if err := os.MkdirAll(w.Dir, 0755); err != nil {
		w.warn(Path(w.Dir), err)
		return
	}
	l := pathLock(Path(w.Dir))
	l.Lock()
	defer l.Unlock()
	w.clean(now)
}

func (w *Writer) Write(p []byte) (int, error) {
	if w.Dir == "" {
		return len(p), nil
	}
	now := w.now()
	path := Path(w.Dir)
	l := pathLock(path)
	l.Lock()
	defer l.Unlock()
	if err := os.MkdirAll(w.Dir, 0755); err != nil {
		w.warn(path, err)
		return len(p), nil
	}
	if fi, err := os.Stat(path); err == nil && localDate(fi.ModTime()) != localDate(now) {
		oldDate := localDate(fi.ModTime()).Format("2006-01-02")
		rotated := filepath.Join(w.Dir, "printerone-"+oldDate+".log")
		if err = os.Rename(path, rotated); err == nil {
			w.clean(now)
		}
	}
	f, err := os.OpenFile(path, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
	if err == nil {
		_, err = f.Write(p)
		if closeErr := f.Close(); err == nil {
			err = closeErr
		}
	}
	if err == nil {
		err = os.Chtimes(path, now, now)
	}
	if err != nil {
		w.warn(path, err)
	}
	return len(p), nil
}

func (w *Writer) now() time.Time {
	if w.Now != nil {
		return w.Now()
	}
	return time.Now()
}
func localDate(t time.Time) time.Time {
	y, m, d := t.In(time.Local).Date()
	return time.Date(y, m, d, 0, 0, 0, 0, time.Local)
}
func (w *Writer) keepDays() int {
	if w.KeepDays < 1 {
		return DefaultKeepDays
	}
	return w.KeepDays
}
func (w *Writer) clean(now time.Time) {
	entries, err := ioutil.ReadDir(w.Dir)
	if err != nil {
		return
	}
	cutoff := localDate(now).AddDate(0, 0, -w.keepDays())
	for _, entry := range entries {
		name := entry.Name()
		if !strings.HasPrefix(name, "printerone-") || !strings.HasSuffix(name, ".log") {
			continue
		}
		dateText := strings.TrimSuffix(strings.TrimPrefix(name, "printerone-"), ".log")
		date, err := time.ParseInLocation("2006-01-02", dateText, time.Local)
		if err != nil || date.Format("2006-01-02") != dateText {
			continue
		}
		if localDate(date).Before(cutoff) {
			_ = os.Remove(filepath.Join(w.Dir, name))
		}
	}
}
func (w *Writer) warn(path string, err error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.warned || w.Warn == nil {
		return
	}
	w.warned = true
	io.WriteString(w.Warn, path+": "+err.Error()+"\n")
}
