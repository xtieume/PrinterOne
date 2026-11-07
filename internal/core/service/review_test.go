package service

import (
	"errors"
	"strings"
	"testing"
	"time"
)

type reviewSCM struct {
	admin                     bool
	exists                    bool
	state                     State
	code                      uint32
	specific                  uint32
	states                    []State
	deleted, stopped, started int
	order                     []string
}

func (s *reviewSCM) IsAdmin() bool                 { return s.admin }
func (s *reviewSCM) Exists(string) (bool, error)   { return s.exists, nil }
func (s *reviewSCM) Create(string, Settings) error { s.exists = true; return nil }
func (s *reviewSCM) Update(string, Settings) error { return nil }
func (s *reviewSCM) Config(string) (Settings, error) {
	return Settings{BinaryPath: `"C:\\PrinterOne.exe" service run --config "C:\\config.json"`}, nil
}
func (s *reviewSCM) Start(string) error { s.started++; return nil }
func (s *reviewSCM) Stop(string) error {
	s.stopped++
	s.state = Stopped
	s.order = append(s.order, "stop")
	return nil
}
func (s *reviewSCM) Delete(string) error {
	s.deleted++
	s.exists = false
	s.order = append(s.order, "delete")
	return nil
}
func (s *reviewSCM) State(string) (State, uint32, uint32, error) {
	if len(s.states) > 0 {
		v := s.states[0]
		s.states = s.states[1:]
		return v, s.code, s.specific, nil
	}
	return s.state, s.code, s.specific, nil
}

func TestREQ_SVC_014_InstallRemovalChecksRuleExistence(t *testing.T) {
	scm := &reviewSCM{admin: true, states: []State{Running, Stopped, Running}}
	var calls []string
	d := Deps{SCM: scm, OSMajor: 10, Sleep: func(time.Duration) {}, Netsh: func(_ string, args ...string) ([]byte, error) {
		calls = append(calls, strings.Join(args, " "))
		if strings.HasPrefix(calls[len(calls)-1], "advfirewall firewall show rule") {
			return []byte("localized output"), nil
		}
		if strings.HasPrefix(calls[len(calls)-1], "advfirewall firewall delete rule") {
			return nil, errors.New("localized failure")
		}
		return nil, nil
	}}
	_ = Install(InstallOptions{ExePath: `C:\PrinterOne.exe`, ConfigPath: `C:\config.json`, Ports: []int{9100}}, d)
	if len(calls) < 2 || !strings.HasPrefix(calls[0], "advfirewall firewall show rule name=PrinterOne") {
		t.Fatalf("calls %v; want show rule before delete", calls)
	}
}

func TestREQ_SVC_014_RemovalSkipsDeleteWhenShowFails(t *testing.T) {
	scm := &reviewSCM{admin: true}
	var deleted bool
	d := Deps{SCM: scm, OSMajor: 10, Netsh: func(_ string, args ...string) ([]byte, error) {
		if strings.Contains(strings.Join(args, " "), "show rule") {
			return []byte("localized message"), errors.New("localized command failure")
		}
		deleted = true
		return nil, nil
	}}
	if _, err := Uninstall(d); err != nil {
		t.Fatal(err)
	}
	if deleted {
		t.Fatal("delete called after show reported that no rule exists")
	}
}

func TestREQ_SVC_006_StoppedAfterStartFailsImmediatelyWithExitCode(t *testing.T) {
	scm := &reviewSCM{admin: true, states: []State{Stopped, Stopped}, code: 1066, specific: 2}
	d := Deps{SCM: scm, OSMajor: 10, Sleep: func(time.Duration) { t.Fatal("waited after service stopped") }}
	err := startAndWait(d)
	if err == nil || !strings.Contains(err.Error(), "1066") || !strings.Contains(err.Error(), "2") {
		t.Fatalf("error %v; want Win32 1066 and service-specific 2", err)
	}
}

func TestREQ_SVC_006_StartPendingServiceWaitsThenRestarts(t *testing.T) {
	scm := &reviewSCM{admin: true, exists: true, states: []State{StartPending, Running, Stopped, Running}}
	d := Deps{SCM: scm, OSMajor: 10, Sleep: func(time.Duration) {}, Netsh: func(_ string, args ...string) ([]byte, error) {
		if strings.Contains(strings.Join(args, " "), "show rule") {
			return nil, errors.New("not present")
		}
		return nil, nil
	}}
	if err := Install(InstallOptions{ExePath: `C:\PrinterOne.exe`, ConfigPath: `C:\config.json`, Ports: []int{9100}}, d); err != nil {
		t.Fatal(err)
	}
	if scm.stopped != 1 || scm.started != 1 {
		t.Fatalf("Stop calls %d, Start calls %d; want one restart", scm.stopped, scm.started)
	}
}

func TestREQ_SVC_008_StopPendingUninstallWaitsWithoutResendingStop(t *testing.T) {
	scm := &reviewSCM{admin: true, exists: true, states: []State{StopPending, Stopped}}
	d := Deps{SCM: scm, OSMajor: 10, Sleep: func(time.Duration) {}, Netsh: func(_ string, args ...string) ([]byte, error) {
		return nil, errors.New("not present")
	}}
	if _, err := Uninstall(d); err != nil {
		t.Fatal(err)
	}
	if scm.stopped != 0 || scm.deleted != 1 {
		t.Fatalf("Stop calls %d, Delete calls %d; want wait then delete", scm.stopped, scm.deleted)
	}
}

func TestREQ_SVC_015_XPRemovalMatchesPortNumber(t *testing.T) {
	var deleted []string
	d := Deps{OSMajor: 5, Netsh: func(_ string, args ...string) ([]byte, error) {
		if len(args) > 1 && args[1] == "show" {
			return []byte("Port 9100\nPort 19100\n"), nil
		}
		deleted = append(deleted, strings.Join(args, " "))
		return nil, nil
	}}
	if err := removeXPRules(d, []int{910}); err != nil {
		t.Fatal(err)
	}
	if len(deleted) != 0 {
		t.Fatalf("deleted %v; port 910 is absent", deleted)
	}
}

func TestREQ_SVC_007_RollbackStopsCreatedServiceBeforeDelete(t *testing.T) {
	scm := &reviewSCM{admin: true, exists: true, state: Running}
	d := Deps{SCM: scm, Sleep: func(time.Duration) {}}
	rollbackInstall(d, true, nil, errors.New("start failed"))
	if len(scm.order) != 2 || scm.order[0] != "stop" || scm.order[1] != "delete" {
		t.Fatalf("SCM order %v; want stop then delete", scm.order)
	}
}
