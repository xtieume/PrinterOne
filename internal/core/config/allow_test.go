package config

import (
	"net"
	"os"
	"path/filepath"
	"testing"
)

// F-004 config tests (spec docs/specs/2026-11-07-f004-allowlist.md §2). TC ids: .testcases/runs/F-004-r1 and
// testcases.md. Config.Allow[i] is Printers[i]'s list; a nil entry allows everyone.

func allowCfg(allowJSON string) string {
	return `{"version":2,"printers":[{"name":"labels","printer":"P","port":9101` + allowJSON + `}]}`
}

// allowCfg2 is two printers; entry 1 carries allowJSON (needs F-003 for two entries).
func allowCfg2(allowJSON string) string {
	return `{"version":2,"printers":[{"name":"a","printer":"P","port":9101,"allow":["10.0.0.5"]},` +
		`{"name":"b","printer":"P","port":9102` + allowJSON + `}]}`
}

func loadAllow(t *testing.T, body string) *Config {
	t.Helper()
	d := tempDir(t)
	defer os.RemoveAll(d)
	c, err := Load(Lookup{Flag: writeFile(t, filepath.Join(d, "c.json"), body)})
	if err != nil {
		t.Fatalf("Load(%s) error = %v, want nil", body, err)
	}
	if len(c.Allow) != len(c.Printers) {
		t.Fatalf("Load(%s): %d allow lists for %d printers, want one per printer", body, len(c.Allow), len(c.Printers))
	}
	return c
}

// ---- REQ-ALLOW-001 ----

// TC-ALLOW-001
func TestREQ_ALLOW_001_NoAllowKeyLoadsNilList(t *testing.T) {
	c := loadAllow(t, allowCfg(``))
	if c.Allow[0] != nil {
		t.Fatalf("Allow[0] = %v, want nil (every client)", c.Allow[0])
	}
}

// ---- REQ-ALLOW-002 / REQ-ALLOW-004 ----

// TC-ALLOW-009
func TestREQ_ALLOW_002_ConfigListIsParsed(t *testing.T) {
	c := loadAllow(t, allowCfg(`,"allow":["192.168.1.0/24","10.0.0.5","fd00::/8"]`))
	l := c.Allow[0]
	for ip, want := range map[string]bool{"192.168.1.77": true, "10.0.0.5": true, "fd00::9": true,
		"10.0.0.6": false, "192.168.2.1": false, "fe80::1": false} {
		if got := l.Allows(net.ParseIP(ip)); got != want {
			t.Errorf("Allow[0].Allows(%s) = %v, want %v", ip, got, want)
		}
	}
	if c.Printers[0].Port != 9101 || c.Printers[0].Name != "labels" {
		t.Errorf("Printers[0] = %+v; the allow key changed the entry", c.Printers[0])
	}
}

// ---- REQ-ALLOW-005 ----

func TestREQ_ALLOW_005_EmptyOrWrongTypeAllowNamesField(t *testing.T) {
	cases := []struct{ tc, allow string }{
		{"TC-ALLOW-021", `[]`},
		{"TC-ALLOW-022/null", `null`},
		{"TC-ALLOW-022/string", `"10.0.0.5"`},
		{"TC-ALLOW-022/number", `5`},
		{"TC-ALLOW-022/object", `{"ip":"10.0.0.5"}`},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) {
			failMulti(t, allowCfg(`,"allow":`+c.allow), "printers[0].allow")
		})
	}
	t.Run("TC-ALLOW-023", func(t *testing.T) { failMulti(t, allowCfg2(`,"allow":[]`), "printers[1].allow") })
}

// ---- REQ-ALLOW-006 ----

func TestREQ_ALLOW_006_InvalidItemNamesIndex(t *testing.T) {
	cases := []struct{ tc, allow, field string }{
		{"TC-ALLOW-026/text", `["nope"]`, "printers[0].allow[0]"},
		{"TC-ALLOW-026/second", `["10.0.0.5","10.0.0.256"]`, "printers[0].allow[1]"},
		{"TC-ALLOW-026/prefix", `["10.0.0.5","fd00::/8","10.0.0.0/33"]`, "printers[0].allow[2]"},
		{"TC-ALLOW-026/hostname", `["printer.example"]`, "printers[0].allow[0]"},
		{"TC-ALLOW-026/empty", `["10.0.0.5",""]`, "printers[0].allow[1]"},
		{"TC-ALLOW-027", `["10.0.0.5",5]`, "printers[0].allow[1]"},
		{"TC-ALLOW-028", `["10.0.0.5","x","y"]`, "printers[0].allow[1]"},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) { failMulti(t, allowCfg(`,"allow":`+c.allow), c.field) })
	}
	t.Run("TC-ALLOW-029", func(t *testing.T) {
		failMulti(t, allowCfg2(`,"allow":["10.0.0.5","bad"]`), "printers[1].allow[1]")
	})
}

// ---- REQ-ALLOW-007 ----

// TC-ALLOW-030
func TestREQ_ALLOW_007_EachEntryKeepsItsOwnList(t *testing.T) {
	c := loadAllow(t, allowCfg2(``))
	if len(c.Printers) != 2 {
		t.Fatalf("%d printers, want 2", len(c.Printers))
	}
	if c.Allow[0] == nil || !c.Allow[0].Allows(net.ParseIP("10.0.0.5")) || c.Allow[0].Allows(net.ParseIP("10.0.0.6")) {
		t.Errorf("Allow[0] = %v, want exactly 10.0.0.5", c.Allow[0])
	}
	if c.Allow[1] != nil {
		t.Errorf("Allow[1] = %v, want nil: entry 1 has no allow key", c.Allow[1])
	}
}

// TC-ALLOW-042: the list sits on entry 1 only; Allow is indexed like Printers, not by "entries with a list".
func TestREQ_ALLOW_007_ListOnLaterEntryOnly(t *testing.T) {
	c := loadAllow(t, `{"version":2,"printers":[{"name":"a","printer":"P","port":9101},`+
		`{"name":"b","printer":"P","port":9102,"allow":["10.0.0.5"]}]}`)
	if c.Allow[0] != nil {
		t.Errorf("Allow[0] = %v, want nil: entry 0 has no allow key", c.Allow[0])
	}
	if c.Allow[1] == nil || !c.Allow[1].Allows(net.ParseIP("10.0.0.5")) || c.Allow[1].Allows(net.ParseIP("10.0.0.6")) {
		t.Errorf("Allow[1] = %v, want exactly 10.0.0.5", c.Allow[1])
	}
}
