package config

// F-005 `ui` section (spec docs/specs/2026-11-07-f005-web-ui.md §2.1). Cases: .testcases/runs/F-005-r1 TC-UI-*.

import (
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// uiLoad writes body to a temp config.json, loads it and returns UIOf's result.
func uiLoad(t *testing.T, body string) (UI, error) {
	dir, err := ioutil.TempDir("", "ui-cfg")
	if err != nil {
		t.Fatal(err)
	}
	defer os.RemoveAll(dir)
	p := filepath.Join(dir, "config.json")
	if err := ioutil.WriteFile(p, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	c, err := Load(Lookup{Flag: p})
	if err != nil {
		t.Fatalf("Load: %v", err)
	}
	return UIOf(c)
}

func uiCfg(ui string) string {
	s := `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101}]`
	if ui != "" {
		s += `,"ui":` + ui
	}
	return s + `}`
}

func uiFail(t *testing.T, body, field string) {
	_, err := uiLoad(t, body)
	ce, ok := err.(*Error)
	if !ok {
		t.Fatalf("%s: want *Error naming %s, got %v", body, field, err)
	}
	if !strings.HasSuffix(ce.Msg, "invalid field "+field) {
		t.Fatalf("%s: want message ending %q, got %q", body, "invalid field "+field, ce.Msg)
	}
}

func uiOK(t *testing.T, body string) UI {
	u, err := uiLoad(t, body)
	if err != nil {
		t.Fatalf("%s: unexpected error %v", body, err)
	}
	return u
}

// TC-UI-001
func TestREQ_UI_001_DefaultPort9180RemoteOff(t *testing.T) {
	u := uiOK(t, uiCfg(""))
	if u.Port != 9180 || u.Remote || DefaultUIPort != 9180 {
		t.Fatalf("got %+v (DefaultUIPort %d), want port 9180 remote off", u, DefaultUIPort)
	}
	if u, err := UIOf(nil); err != nil || u.Port != 9180 {
		t.Fatalf("UIOf(nil) = %+v, %v; want port 9180", u, err)
	}
}

// TC-UI-002
func TestREQ_UI_001_ConfiguredPortUsed(t *testing.T) {
	if u := uiOK(t, uiCfg(`{"port":9181}`)); u.Port != 9181 {
		t.Fatalf("got %+v, want port 9181", u)
	}
}

// TC-UI-007
func TestREQ_UI_002_PortEqualToPrinterPort(t *testing.T) {
	uiFail(t, uiCfg(`{"port":9101}`), "ui.port")
}

// TC-UI-008
func TestREQ_UI_002_PortRangeBoundaries(t *testing.T) {
	uiFail(t, uiCfg(`{"port":0}`), "ui.port")
	uiFail(t, uiCfg(`{"port":65536}`), "ui.port")
	if u := uiOK(t, uiCfg(`{"port":1}`)); u.Port != 1 {
		t.Fatalf("port 1: got %+v", u)
	}
	if u := uiOK(t, uiCfg(`{"port":65535}`)); u.Port != 65535 {
		t.Fatalf("port 65535: got %+v", u)
	}
}

// TC-UI-009
func TestREQ_UI_002_PortWrongType(t *testing.T) {
	uiFail(t, uiCfg(`{"port":"9180"}`), "ui.port")
	uiFail(t, uiCfg(`{"port":9180.5}`), "ui.port")
	uiFail(t, uiCfg(`{"port":null}`), "ui.port")
}

// TC-UI-010
func TestREQ_UI_002_SectionAndRemoteWrongType(t *testing.T) {
	uiFail(t, uiCfg(`5`), "ui")
	uiFail(t, uiCfg(`{"remote":"yes"}`), "ui.remote")
}

// TC-UI-011 (needs F-003: two printers)
func TestREQ_UI_002_PortEqualToSecondPrinterPort(t *testing.T) {
	uiFail(t, `{"version":2,"printers":[{"name":"A","printer":"Canon","port":9101},`+
		`{"name":"B","printer":"HP","port":9102}],"ui":{"port":9102}}`, "ui.port")
}

// TC-UI-017
func TestREQ_UI_004_RemotePasswordLength(t *testing.T) {
	uiFail(t, uiCfg(`{"remote":true,"password":"1234567"}`), "ui.password")
	uiFail(t, uiCfg(`{"remote":true}`), "ui.password")
	u := uiOK(t, uiCfg(`{"remote":true,"password":"12345678"}`))
	if !u.Remote || u.Password != "12345678" {
		t.Fatalf("got %+v, want remote with password 12345678", u)
	}
}

// TC-UI-018
func TestREQ_UI_004_ShortPasswordAcceptedWhenRemoteOff(t *testing.T) {
	u := uiOK(t, uiCfg(`{"password":"x"}`))
	if u.Remote || u.Password != "x" || u.Port != 9180 {
		t.Fatalf("got %+v, want remote off, password x, port 9180", u)
	}
}

// TC-UI-087
func TestREQ_UI_004_PasswordLengthInCharacters(t *testing.T) {
	uiFail(t, uiCfg(`{"remote":true,"password":"ééééééé"}`), "ui.password")
}
