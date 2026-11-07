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
var datesMu sync.Mutex
var dates = make(map[string]time.Time)

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
	if w.Dir == "" {
		return
	}
	if err := os.MkdirAll(w.Dir, 0755); err != nil {
		return
	}
	w.clean(w.now())
}

func (w *Writer) Write(p []byte) (int, error) {
	if w.Dir == "" {
		return len(p), nil
	}
	path := Path(w.Dir)
	if err := os.MkdirAll(w.Dir, 0755); err != nil {
		w.warn(path, err) // no folder means the line cannot be appended: a real write failure
		return len(p), nil
	}
	// Rotation is serialized between processes; ordinary O_APPEND line writes remain lock-free.
	now := time.Time{}
	if lock, ok := acquireRotationLock(path); ok {
		now = w.now()
		w.rotate(path, now)
		_ = os.Remove(lock)
	} else {
		now = w.now()
	}
	f, err := os.OpenFile(path, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
	if err == nil {
		_, err = f.Write(p)
		if ce := f.Close(); err == nil {
			err = ce
		}
	}
	if err != nil {
		w.warn(path, err)
	} else {
		datesMu.Lock()
		d := dates[path]
		if d.IsZero() {
			dates[path] = localDate(now)
		}
		datesMu.Unlock()
	}
	return len(p), nil
}

func acquireRotationLock(path string) (string, bool) {
	lock := path + ".lock"
	for i := 0; i < 100; i++ {
		f, err := os.OpenFile(lock, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
		if err == nil {
			f.Close()
			return lock, true
		}
		if fi, e := os.Stat(lock); e == nil && time.Since(fi.ModTime()) > 10*time.Second {
			_ = os.Remove(lock)
			continue
		}
		time.Sleep(10 * time.Millisecond)
	}
	return "", false
}

func (w *Writer) rotate(path string, now time.Time) {
	fi, err := os.Stat(path)
	if err != nil {
		return
	}
	datesMu.Lock()
	fileDate := dates[path]
	datesMu.Unlock()
	if fileDate.IsZero() {
		fileDate = localDate(fi.ModTime())
	}
	today := localDate(now)
	if !today.After(fileDate) {
		return
	}
	target := filepath.Join(w.Dir, "printerone-"+fileDate.Format("2006-01-02")+".log")
	if existing, e := os.Stat(target); e == nil && existing.IsDir() {
		return
	}
	if _, e := os.Stat(target); e == nil {
		src, e := os.Open(path)
		if e != nil {
			return
		}
		defer src.Close()
		dst, e := os.OpenFile(target, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
		if e != nil {
			return
		}
		_, e = io.Copy(dst, src)
		ce := dst.Close()
		if e != nil || ce != nil {
			return
		}
		if os.Remove(path) != nil {
			return
		}
	} else if e = os.Rename(path, target); e != nil {
		return
	}
	datesMu.Lock()
	delete(dates, path)
	datesMu.Unlock()
	w.clean(now)
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
		if localDate(date).Before(cutoff) && !entry.IsDir() {
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
