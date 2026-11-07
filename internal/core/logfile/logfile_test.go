package logfile

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"sync"
	"testing"
	"time"
)

// Tests of the log file writer (spec docs/specs/2026-11-07-f002-logs.md §2), with a fake clock.
// TC ids: .testcases/runs/F-002-r1/testcase/testcases.md and testcases.md.

func TestMain(m *testing.M) {
	time.Local = time.FixedZone("UTC+7", 7*3600)
	if dir := os.Getenv("LOGF_CHILD_DIR"); dir != "" {
		childWrite(dir, os.Getenv("LOGF_CHILD_NAME"))
		os.Exit(0)
	}
	os.Exit(m.Run())
}

// childWrite is the body of a child process in TC-LOGF-052.
func childWrite(dir, name string) {
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	pad := strings.Repeat("x", 100)
	for i := 0; i < 200; i++ {
		w.Write([]byte(fmt.Sprintf("%s-%d-%s\n", name, i, pad)))
	}
}

// checkWholeLines asserts the file holds exactly 2 × 200 intact lines written by A and B.
func checkWholeLines(t *testing.T, path string) {
	t.Helper()
	body := read(t, path)
	lines := strings.Split(strings.TrimSuffix(body, "\n"), "\n")
	if len(lines) != 400 {
		t.Fatalf("%d lines; want 400", len(lines))
	}
	re := regexp.MustCompile(`^([AB])-(\d+)-x{100}$`)
	seen := map[string]bool{}
	for _, l := range lines {
		m := re.FindStringSubmatch(l)
		if m == nil {
			t.Fatalf("broken line %q", l)
		}
		seen[m[1]+"-"+m[2]] = true
	}
	if len(seen) != 400 {
		t.Fatalf("%d distinct lines; want 400", len(seen))
	}
}

type clock struct {
	mu sync.Mutex
	t  time.Time
}

func (c *clock) now() time.Time {
	c.mu.Lock()
	defer c.mu.Unlock()
	return c.t
}

func (c *clock) set(t time.Time) {
	c.mu.Lock()
	c.t = t
	c.mu.Unlock()
}

func at(y int, m time.Month, d, h, min, s int) time.Time {
	return time.Date(y, m, d, h, min, s, 0, time.Local)
}

func tempDir(t *testing.T) string {
	t.Helper()
	d, err := ioutil.TempDir("", "logftest")
	if err != nil {
		t.Fatal(err)
	}
	return d
}

func read(t *testing.T, path string) string {
	t.Helper()
	b, err := ioutil.ReadFile(path)
	if err != nil {
		t.Fatalf("read %s: %v", path, err)
	}
	return string(b)
}

func touch(t *testing.T, path, body string, mtime time.Time) {
	t.Helper()
	if err := ioutil.WriteFile(path, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.Chtimes(path, mtime, mtime); err != nil {
		t.Fatal(err)
	}
}

func exists(path string) bool { _, err := os.Stat(path); return err == nil }

func names(t *testing.T, dir string) []string {
	t.Helper()
	fs, err := ioutil.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	var out []string
	for _, f := range fs {
		out = append(out, f.Name())
	}
	return out
}

func mustWrite(t *testing.T, w *Writer, s string) {
	t.Helper()
	n, err := w.Write([]byte(s))
	if n != len(s) || err != nil {
		t.Fatalf("Write(%q) = %d, %v; want %d, nil", s, n, err, len(s))
	}
}

// ---- REQ-LOGF-001 ----

// TC-LOGF-002
func TestREQ_LOGF_001_DefaultDirAndPath(t *testing.T) {
	d := filepath.Join("x", "PrinterOne")
	if got := DefaultDir(d); got != filepath.Join(d, "logs") {
		t.Fatalf("DefaultDir(%q) = %q; want %q", d, got, filepath.Join(d, "logs"))
	}
	if got := Path(filepath.Join(d, "logs")); got != filepath.Join(d, "logs", "printerone.log") {
		t.Fatalf("Path = %q; want %q", got, filepath.Join(d, "logs", "printerone.log"))
	}
}

// TC-LOGF-003
func TestREQ_LOGF_001_StartCreatesNestedFolder(t *testing.T) {
	root := tempDir(t)
	defer os.RemoveAll(root)
	dir := filepath.Join(root, "a", "b")
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	if fi, err := os.Stat(dir); err != nil || !fi.IsDir() {
		t.Fatalf("Start did not create %s (err %v)", dir, err)
	}
}

// ---- REQ-LOGF-004 ----

// TC-LOGF-020
func TestREQ_LOGF_004_WritesAppendInOrder(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "a\n")
	mustWrite(t, w, "b\n")
	if got := read(t, filepath.Join(dir, FileName)); got != "a\nb\n" {
		t.Fatalf("file %q; want %q", got, "a\nb\n")
	}
}

// ---- REQ-LOGF-006 ----

// TC-LOGF-023
func TestREQ_LOGF_006_TwoWritersNeverSplitLines(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	a := &Writer{Dir: dir, Now: c.now}
	b := &Writer{Dir: dir, Now: c.now}
	a.Start()
	b.Start()
	pad := strings.Repeat("x", 100)
	gate := make(chan struct{})
	var wg sync.WaitGroup
	for _, p := range []struct {
		name string
		w    *Writer
	}{{"A", a}, {"B", b}} {
		wg.Add(1)
		go func(name string, w *Writer) {
			defer wg.Done()
			<-gate
			for i := 0; i < 200; i++ {
				w.Write([]byte(fmt.Sprintf("%s-%d-%s\n", name, i, pad)))
				runtime.Gosched()
			}
		}(p.name, p.w)
	}
	close(gate)
	wg.Wait()
	checkWholeLines(t, filepath.Join(dir, FileName))
}

// TC-LOGF-052
func TestREQ_LOGF_006_TwoProcessesNeverSplitLines(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	var cmds []*exec.Cmd
	for _, name := range []string{"A", "B"} {
		cmd := exec.Command(os.Args[0], "-test.run=^$")
		cmd.Env = append(os.Environ(), "LOGF_CHILD_DIR="+dir, "LOGF_CHILD_NAME="+name)
		if err := cmd.Start(); err != nil {
			t.Fatal(err)
		}
		cmds = append(cmds, cmd)
	}
	for _, cmd := range cmds {
		if err := cmd.Wait(); err != nil {
			t.Fatalf("child: %v", err)
		}
	}
	checkWholeLines(t, filepath.Join(dir, FileName))
}

// ---- REQ-LOGF-007 ----

// TC-LOGF-026
func TestREQ_LOGF_007_LocalMidnightRotates(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 23, 59, 59)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "one\n")
	c.set(at(2026, 11, 8, 0, 0, 0))
	mustWrite(t, w, "two\n")
	rotated := filepath.Join(dir, "printerone-2026-11-07.log")
	if !exists(rotated) || read(t, rotated) != "one\n" || read(t, filepath.Join(dir, FileName)) != "two\n" {
		t.Fatalf("files %v; want printerone-2026-11-07.log = one, printerone.log = two", names(t, dir))
	}
}

// TC-LOGF-027
func TestREQ_LOGF_007_SameLocalDayAcrossUTCMidnightDoesNotRotate(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 5, 0, 0)} // 2026-11-06 22:00 UTC
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "one\n")
	c.set(at(2026, 11, 7, 8, 0, 0)) // 2026-11-07 01:00 UTC
	mustWrite(t, w, "two\n")
	if n := names(t, dir); len(n) != 1 || read(t, filepath.Join(dir, FileName)) != "one\ntwo\n" {
		t.Fatalf("files %v; want only printerone.log with both lines", n)
	}
}

// TC-LOGF-028
func TestREQ_LOGF_007_GapOfDaysNamesOldLinesDate(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "one\n")
	c.set(at(2026, 11, 10, 10, 0, 0))
	mustWrite(t, w, "two\n")
	rotated := filepath.Join(dir, "printerone-2026-11-07.log")
	if !exists(rotated) || read(t, rotated) != "one\n" || exists(filepath.Join(dir, "printerone-2026-11-09.log")) || exists(filepath.Join(dir, "printerone-2026-11-10.log")) {
		t.Fatalf("files %v; want printerone-2026-11-07.log holding line 1 only", names(t, dir))
	}
}

// TC-LOGF-029
func TestREQ_LOGF_007_RestartOnLaterDayRotatesOldFile(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	touch(t, filepath.Join(dir, FileName), "old\n", at(2026, 11, 5, 12, 0, 0))
	c := &clock{t: at(2026, 11, 7, 9, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "new\n")
	rotated := filepath.Join(dir, "printerone-2026-11-05.log")
	if !exists(rotated) || read(t, rotated) != "old\n" || read(t, filepath.Join(dir, FileName)) != "new\n" {
		t.Fatalf("files %v; want printerone-2026-11-05.log = old, printerone.log = new", names(t, dir))
	}
}

// TC-LOGF-030
func TestREQ_LOGF_007_RestartSameDayDoesNotRotate(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	touch(t, filepath.Join(dir, FileName), "old\n", at(2026, 11, 7, 8, 0, 0))
	c := &clock{t: at(2026, 11, 7, 9, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "new\n")
	if n := names(t, dir); len(n) != 1 || read(t, filepath.Join(dir, FileName)) != "old\nnew\n" {
		t.Fatalf("files %v, printerone.log %q; want only printerone.log = old+new", n, read(t, filepath.Join(dir, FileName)))
	}
}

// TC-LOGF-031
func TestREQ_LOGF_007_TwoWritersCrossMidnightRotateOnce(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 23, 0, 0)}
	a := &Writer{Dir: dir, Now: c.now}
	b := &Writer{Dir: dir, Now: c.now}
	a.Start()
	b.Start()
	mustWrite(t, a, "a1\n")
	mustWrite(t, b, "b1\n")
	c.set(at(2026, 11, 8, 0, 1, 0))
	mustWrite(t, a, "a2\n")
	mustWrite(t, b, "b2\n")
	n := names(t, dir)
	rotated := filepath.Join(dir, "printerone-2026-11-07.log")
	if len(n) != 2 || !exists(rotated) || read(t, rotated) != "a1\nb1\n" || read(t, filepath.Join(dir, FileName)) != "a2\nb2\n" {
		t.Fatalf("files %v; want printerone-2026-11-07.log = a1 b1 and printerone.log = a2 b2", n)
	}
}

// TC-LOGF-032
func TestREQ_LOGF_007_FailedRenameKeepsLogging(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	if err := os.MkdirAll(filepath.Join(dir, "printerone-2026-11-07.log", "x"), 0755); err != nil {
		t.Fatal(err)
	}
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "one\n")
	c.set(at(2026, 11, 8, 10, 0, 0))
	mustWrite(t, w, "two\n")
	if got := read(t, filepath.Join(dir, FileName)); !strings.HasSuffix(got, "two\n") {
		t.Fatalf("printerone.log %q; want the line appended although the rename failed", got)
	}
}

// TC-LOGF-051
func TestREQ_LOGF_007_TwoWritersRotateAtSameInstant(t *testing.T) {
	for round := 0; round < 20; round++ {
		dir := tempDir(t)
		c := &clock{t: at(2026, 11, 7, 23, 0, 0)}
		a := &Writer{Dir: dir, Now: c.now}
		b := &Writer{Dir: dir, Now: c.now}
		a.Start()
		b.Start()
		mustWrite(t, a, "a1\n")
		mustWrite(t, b, "b1\n")
		c.set(at(2026, 11, 8, 0, 1, 0))
		gate := make(chan struct{})
		var wg sync.WaitGroup
		for _, p := range []struct {
			w    *Writer
			line string
		}{{a, "a2\n"}, {b, "b2\n"}} {
			wg.Add(1)
			go func(w *Writer, line string) {
				defer wg.Done()
				<-gate
				w.Write([]byte(line))
			}(p.w, p.line)
		}
		close(gate)
		wg.Wait()
		n := names(t, dir)
		rot := readOrEmpty(filepath.Join(dir, "printerone-2026-11-07.log"))
		cur := readOrEmpty(filepath.Join(dir, FileName))
		os.RemoveAll(dir)
		if len(n) != 2 || rot != "a1\nb1\n" || !(cur == "a2\nb2\n" || cur == "b2\na2\n") {
			t.Fatalf("round %d: files %v, rotated %q, current %q; want one rotation and no line lost", round, n, rot, cur)
		}
	}
}

func readOrEmpty(path string) string { b, _ := ioutil.ReadFile(path); return string(b) }

// TC-LOGF-053
func TestREQ_LOGF_007_RotationWithOpenReaderLosesNothing(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "one\n")
	reader, err := os.Open(filepath.Join(dir, FileName))
	if err != nil {
		t.Fatalf("open for reading: %v", err)
	}
	defer reader.Close()
	c.set(at(2026, 11, 8, 10, 0, 0))
	mustWrite(t, w, "two\n")
	cur := readOrEmpty(filepath.Join(dir, FileName))
	all := readOrEmpty(filepath.Join(dir, "printerone-2026-11-07.log")) + cur
	if !strings.HasSuffix(cur, "two\n") || !strings.Contains(all, "one\n") {
		t.Fatalf("files %v, current %q; want two appended and one kept", names(t, dir), cur)
	}
}

// ---- REQ-LOGF-008 ----

// TC-LOGF-033
func TestREQ_LOGF_008_StartCleansWithDefault30Days(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	now := time.Now()
	touch(t, filepath.Join(dir, "printerone-2026-10-07.log"), "x\n", now)
	touch(t, filepath.Join(dir, "printerone-2026-10-08.log"), "x\n", now)
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	if exists(filepath.Join(dir, "printerone-2026-10-07.log")) || !exists(filepath.Join(dir, "printerone-2026-10-08.log")) {
		t.Fatalf("files %v; want -10-07 deleted and -10-08 kept (cut-off 2026-10-08)", names(t, dir))
	}
}

// TC-LOGF-034
func TestREQ_LOGF_008_OtherFilesNeverTouched(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	old := at(2000, 1, 1, 0, 0, 0)
	keep := []string{FileName, "printerone-foo.log", "printerone-2026-13-01.log", "notes.txt", "printerone-2020-01-01.txt"}
	for _, n := range keep {
		touch(t, filepath.Join(dir, n), "x\n", old)
	}
	touch(t, filepath.Join(dir, FileName), "x\n", at(2026, 11, 7, 9, 0, 0)) // today: no rotation in play
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, KeepDays: 1, Now: c.now}
	w.Start()
	for _, n := range keep {
		if !exists(filepath.Join(dir, n)) {
			t.Fatalf("%s deleted; files left %v", n, names(t, dir))
		}
	}
}

// TC-LOGF-035
func TestREQ_LOGF_008_CleanUpAlsoAtRotation(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	old := filepath.Join(dir, "printerone-2026-11-05.log")
	touch(t, old, "x\n", time.Now())
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, KeepDays: 2, Now: c.now}
	w.Start()
	if !exists(old) {
		t.Fatalf("Start deleted %s; with keep 2 the cut-off is 2026-11-05", old)
	}
	mustWrite(t, w, "a\n")
	c.set(at(2026, 11, 8, 10, 0, 0))
	mustWrite(t, w, "b\n")
	if exists(old) || !exists(filepath.Join(dir, "printerone-2026-11-07.log")) {
		t.Fatalf("files %v; want -11-05 deleted at rotation and -11-07 created", names(t, dir))
	}
}

// ---- REQ-LOGF-009 ----

// TC-LOGF-037
func TestREQ_LOGF_009_FailingWritesWarnOnce(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	path := filepath.Join(dir, FileName)
	if err := os.Mkdir(path, 0755); err != nil {
		t.Fatal(err)
	}
	_, oserr := os.OpenFile(path, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
	if oserr == nil {
		t.Fatal("opening a directory for append succeeded; the test needs a failing open")
	}
	osText := oserr.Error()
	if pe, ok := oserr.(*os.PathError); ok {
		osText = pe.Err.Error()
	}
	var warn bytes.Buffer
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now, Warn: &warn}
	w.Start()
	for i := 0; i < 3; i++ {
		mustWrite(t, w, "x\n")
	}
	got := warn.String()
	if strings.Count(got, "\n") != 1 || !strings.Contains(got, path) || !strings.Contains(got, osText) {
		t.Fatalf("warnings %q; want exactly one line naming %s and %q", got, path, osText)
	}
}

// TC-LOGF-039
func TestREQ_LOGF_009_NilWarnDoesNotPanic(t *testing.T) {
	dir := tempDir(t)
	defer os.RemoveAll(dir)
	if err := os.Mkdir(filepath.Join(dir, FileName), 0755); err != nil {
		t.Fatal(err)
	}
	c := &clock{t: at(2026, 11, 7, 10, 0, 0)}
	w := &Writer{Dir: dir, Now: c.now}
	w.Start()
	mustWrite(t, w, "x\n")
}
