//go:build e2e && windows
// +build e2e,windows

// Package e2e is the windows-e2e registry check (spec §2.6): a real spooler, a "Generic / Text Only" printer on a
// local port that writes to a file, and the real printerone.exe. Run on windows-latest:
// go test -tags e2e ./internal/core/e2e/...
// It also covers what only real Windows can show: winsock error codes (reset, port in use) and EnumPrinters.
package e2e

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"net"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"syscall"
	"testing"
	"time"
)

const printerName = "PrinterOne E2E"

var (
	exe       string // built cmd/printerone
	legacyExe string // built cmd/printerone-legacy
	outFile   string // the printer's local file port
)

func powershell(script string) error {
	out, err := exec.Command("powershell", "-NoProfile", "-NonInteractive", "-Command", script).CombinedOutput()
	if err != nil {
		return fmt.Errorf("powershell %q: %v\n%s", script, err, out)
	}
	return nil
}

func psQuote(s string) string { return "'" + strings.Replace(s, "'", "''", -1) + "'" }

func TestMain(m *testing.M) {
	dir, err := ioutil.TempDir("", "printerone-e2e")
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	outFile = filepath.Join(dir, "out.prn")
	exe = filepath.Join(dir, "printerone.exe")
	legacyExe = filepath.Join(dir, "printerone-legacy.exe")
	code := 1
	if err := setup(); err != nil {
		fmt.Fprintln(os.Stderr, "e2e setup:", err)
	} else {
		code = m.Run()
	}
	powershell("Remove-Printer -Name " + psQuote(printerName))
	powershell("Remove-PrinterPort -Name " + psQuote(outFile))
	os.RemoveAll(dir)
	os.Exit(code)
}

func setup() error {
	// A local port named after a file path makes the spooler write each job to that file.
	for _, s := range []string{
		"Add-PrinterDriver -Name 'Generic / Text Only'",
		"Add-PrinterPort -Name " + psQuote(outFile),
		"Add-Printer -Name " + psQuote(printerName) + " -DriverName 'Generic / Text Only' -PortName " + psQuote(outFile),
	} {
		if err := powershell(s); err != nil {
			return err
		}
	}
	for out, pkg := range map[string]string{exe: "./cmd/printerone", legacyExe: "./cmd/printerone-legacy"} {
		build := exec.Command("go", "build", "-o", out, pkg)
		build.Dir = filepath.Join("..", "..", "..") // repo root from internal/core/e2e
		if b, err := build.CombinedOutput(); err != nil {
			return fmt.Errorf("go build %s: %v\n%s", pkg, err, b)
		}
	}
	return nil
}

type syncBuf struct {
	mu sync.Mutex
	b  bytes.Buffer
}

func (s *syncBuf) Write(p []byte) (int, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.b.Write(p)
}
func (s *syncBuf) String() string {
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.b.String()
}

func freePort(t *testing.T) int {
	t.Helper()
	l, err := net.Listen("tcp", "0.0.0.0:0")
	if err != nil {
		t.Fatal(err)
	}
	defer l.Close()
	return l.Addr().(*net.TCPAddr).Port
}

func configJSON(printer string, port int) string {
	return `{"version":2,"printers":[{"name":"default","printer":"` + printer + `","port":` + strconv.Itoa(port) + `}]}`
}

func writeConfig(t *testing.T, port int) string { return writeConfigFor(t, printerName, port) }

func writeConfigFor(t *testing.T, printer string, port int) string {
	t.Helper()
	f, err := ioutil.TempFile("", "printerone-e2e-cfg")
	if err != nil {
		t.Fatal(err)
	}
	f.WriteString(configJSON(printer, port))
	f.Close()
	return f.Name()
}

type server struct {
	cmd            *exec.Cmd
	stdout, stderr syncBuf
	addr           string
}

// startServe runs `printerone serve --config` and waits until the port accepts.
func startServe(t *testing.T) *server {
	t.Helper()
	port := freePort(t)
	return startCmd(t, port, exec.Command(exe, "serve", "--config", writeConfig(t, port)))
}

// startCmd starts cmd (in its own process group, so it can get CTRL_BREAK) and waits until port accepts.
func startCmd(t *testing.T, port int, cmd *exec.Cmd) *server {
	t.Helper()
	s := &server{addr: "127.0.0.1:" + strconv.Itoa(port), cmd: cmd}
	s.cmd.Stdout, s.cmd.Stderr = &s.stdout, &s.stderr
	s.cmd.SysProcAttr = &syscall.SysProcAttr{CreationFlags: syscall.CREATE_NEW_PROCESS_GROUP}
	if err := s.cmd.Start(); err != nil {
		t.Fatal(err)
	}
	for deadline := time.Now().Add(30 * time.Second); ; time.Sleep(200 * time.Millisecond) {
		c, err := net.DialTimeout("tcp", s.addr, time.Second)
		if err == nil {
			c.Close() // logs one `empty` line
			return s
		}
		if time.Now().After(deadline) {
			s.cmd.Process.Kill()
			t.Fatalf("serve not listening on %s after 30 s: %v\nstdout %s\nstderr %s", s.addr, err, s.stdout.String(), s.stderr.String())
		}
	}
}

func (s *server) stop() { s.cmd.Process.Kill(); s.cmd.Wait() }

func exitCode(err error) int {
	if err == nil {
		return 0
	}
	if ee, ok := err.(*exec.ExitError); ok {
		return ee.Sys().(syscall.WaitStatus).ExitStatus()
	}
	return -1
}

// TC-E2E-001
func TestREQ_E2E_001_BytesArriveInPrinterFileUnchanged(t *testing.T) {
	s := startServe(t)
	defer s.stop()
	conn, err := net.DialTimeout("tcp", s.addr, 5*time.Second)
	if err != nil {
		t.Fatal(err)
	}
	want := []byte("PrinterOne e2e\r\n")
	for i := 0; i < 256; i++ {
		want = append(want, byte(i))
	}
	want = append(want, []byte("\r\nend\f")...)
	if _, err := conn.Write(want); err != nil {
		t.Fatal(err)
	}
	conn.Close()

	var got []byte
	for deadline := time.Now().Add(60 * time.Second); time.Now().Before(deadline); time.Sleep(500 * time.Millisecond) {
		if got, err = ioutil.ReadFile(outFile); err == nil && len(got) >= len(want) {
			break
		}
	}
	if !bytes.Equal(got, want) {
		t.Fatalf("output file has %d bytes, want %d identical bytes\ngot  % x\nwant % x\nstdout %s\nstderr %s",
			len(got), len(want), got, want, s.stdout.String(), s.stderr.String())
	}
}

// TC-E2E-002
func TestREQ_SPOOL_003_WindowsPrintersListsRealPrinter(t *testing.T) {
	var stdout, stderr bytes.Buffer
	cmd := exec.Command(exe, "printers")
	cmd.Stdout, cmd.Stderr = &stdout, &stderr
	code := exitCode(cmd.Run())
	found := false
	for _, l := range strings.Split(strings.Replace(stdout.String(), "\r\n", "\n", -1), "\n") {
		found = found || l == printerName
	}
	if code != 0 || !found {
		t.Fatalf("printers: exit %d, stdout %q, stderr %q; want 0 and a line %q", code, stdout.String(), stderr.String(), printerName)
	}
}

// TC-E2E-003
func TestREQ_RECV_008_WindowsBusyPortExits1(t *testing.T) {
	occ, err := net.Listen("tcp", "0.0.0.0:0")
	if err != nil {
		t.Fatal(err)
	}
	defer occ.Close()
	port := occ.Addr().(*net.TCPAddr).Port
	var stdout, stderr bytes.Buffer
	cmd := exec.Command(exe, "serve", "--config", writeConfig(t, port))
	cmd.Stdout, cmd.Stderr = &stdout, &stderr
	if err := cmd.Start(); err != nil {
		t.Fatal(err)
	}
	done := make(chan error, 1)
	go func() { done <- cmd.Wait() }()
	select {
	case err := <-done:
		if code := exitCode(err); code != 1 || !strings.Contains(stderr.String(), strconv.Itoa(port)) {
			t.Fatalf("exit %d, stderr %q; want 1 and the port %d (WSAEADDRINUSE mapped)", code, stderr.String(), port)
		}
	case <-time.After(10 * time.Second):
		cmd.Process.Kill()
		t.Fatalf("serve still running on a busy port; stdout %q", stdout.String())
	}
	c, err := net.DialTimeout("tcp", "127.0.0.1:"+strconv.Itoa(port), time.Second)
	if err != nil {
		t.Fatalf("occupant no longer accepts: %v", err)
	}
	c.Close()
}

// TC-E2E-004
func TestREQ_RECV_005_WindowsResetLoggedAsReset(t *testing.T) {
	s := startServe(t)
	defer s.stop()
	conn, err := net.DialTimeout("tcp", s.addr, 5*time.Second)
	if err != nil {
		t.Fatal(err)
	}
	conn.Write([]byte("partial"))
	time.Sleep(300 * time.Millisecond)
	conn.(*net.TCPConn).SetLinger(0)
	conn.Close() // RST: WSAECONNRESET on the server
	for deadline := time.Now().Add(10 * time.Second); time.Now().Before(deadline); time.Sleep(200 * time.Millisecond) {
		if strings.Contains(s.stdout.String(), "dropped: reset") {
			return
		}
	}
	t.Fatalf("no `dropped: reset` line after a reset; stdout %q", s.stdout.String())
}

// TC-E2E-005
func TestREQ_CLI_001_WindowsCtrlBreakExits0(t *testing.T) {
	s := startServe(t)
	defer s.stop()
	gen := syscall.NewLazyDLL("kernel32.dll").NewProc("GenerateConsoleCtrlEvent")
	if r, _, err := gen.Call(1 /* CTRL_BREAK_EVENT */, uintptr(s.cmd.Process.Pid)); r == 0 {
		t.Fatalf("GenerateConsoleCtrlEvent: %v", err)
	}
	done := make(chan error, 1)
	go func() { done <- s.cmd.Wait() }()
	select {
	case err := <-done:
		if code := exitCode(err); code != 0 {
			t.Fatalf("exit %d after Ctrl+Break, want 0 (signal not wired to the shutdown path?); stderr %q", code, s.stderr.String())
		}
	case <-time.After(15 * time.Second):
		t.Fatal("serve still running 15 s after Ctrl+Break")
	}
}

// TC-E2E-006
func TestREQ_CFG_001_WindowsConfigNextToExeFound(t *testing.T) {
	dir, err := ioutil.TempDir("", "printerone-exedir")
	if err != nil {
		t.Fatal(err)
	}
	defer os.RemoveAll(dir)
	b, err := ioutil.ReadFile(exe)
	if err != nil {
		t.Fatal(err)
	}
	copyExe := filepath.Join(dir, "printerone.exe")
	if err := ioutil.WriteFile(copyExe, b, 0755); err != nil {
		t.Fatal(err)
	}
	port := freePort(t)
	if err := ioutil.WriteFile(filepath.Join(dir, "config.json"), []byte(configJSON(printerName, port)), 0644); err != nil {
		t.Fatal(err)
	}
	cmd := exec.Command(copyExe, "serve")
	cmd.Dir = os.TempDir() // working directory is not the exe folder
	s := startCmd(t, port, cmd)
	s.stop()
}

// TC-E2E-007
func TestREQ_CLI_002_WindowsExesNameTheirEdition(t *testing.T) {
	for bin, ed := range map[string]string{exe: "(main)", legacyExe: "(legacy)"} {
		out, err := exec.Command(bin, "version").Output()
		if err != nil || !strings.HasSuffix(strings.TrimRight(string(out), "\r\n"), ed) {
			t.Fatalf("%s version: %v, stdout %q; want `PrinterOne <v> %s`", filepath.Base(bin), err, out, ed)
		}
	}
}

// TC-E2E-008
func TestREQ_LOG_001_WindowsFailedLineFromRealSpooler(t *testing.T) {
	port := freePort(t)
	s := startCmd(t, port, exec.Command(exe, "serve", "--config", writeConfigFor(t, "PrinterOne E2E Missing", port)))
	defer s.stop()
	conn, err := net.DialTimeout("tcp", s.addr, 5*time.Second)
	if err != nil {
		t.Fatal(err)
	}
	conn.Write([]byte("x"))
	conn.Close()
	want := "failed: 1801 The printer name is invalid."
	for deadline := time.Now().Add(10 * time.Second); time.Now().Before(deadline); time.Sleep(200 * time.Millisecond) {
		out := s.stdout.String()
		if !strings.Contains(out, "failed:") {
			continue
		}
		var failed []string
		for _, l := range strings.Split(out, "\n") {
			if strings.Contains(l, " 1B ") {
				failed = append(failed, l)
			}
		}
		if len(failed) != 1 || !strings.HasSuffix(failed[0], want) || strings.Contains(out, want+"\r") {
			t.Fatalf("stdout %q: want exactly one line for the job, ending %q, with no trailing CR from FormatMessage", out, want)
		}
		return
	}
	t.Fatalf("no `failed:` line within 10 s; stdout %q", s.stdout.String())
}
