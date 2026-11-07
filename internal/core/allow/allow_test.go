package allow

import (
	"net"
	"strings"
	"testing"
)

// F-004 matcher tests (spec docs/specs/2026-11-07-f004-allowlist.md §2). TC ids: .testcases/runs/F-004-r1 and
// testcases.md. Go 1.10: net.ParseIP / net.ParseCIDR only.

func mustParse(t *testing.T, items ...string) List {
	t.Helper()
	l, bad, err := Parse(items)
	if err != nil {
		t.Fatalf("Parse(%q) = item %d error %v, want nil", items, bad, err)
	}
	if len(l) == 0 {
		t.Fatalf("Parse(%q) = empty List, want %d entries", items, len(items))
	}
	return l
}

// expect checks Allows for every ip; want is the expected answer for all of them.
func expect(t *testing.T, l List, want bool, ips ...string) {
	t.Helper()
	for _, s := range ips {
		ip := net.ParseIP(s)
		if ip == nil {
			t.Fatalf("bad test ip %q", s)
		}
		if got := l.Allows(ip); got != want {
			t.Errorf("List %v Allows(%s) = %v, want %v", l, s, got, want)
		}
	}
}

// ---- REQ-ALLOW-001 ----

// TC-ALLOW-002
func TestREQ_ALLOW_001_NilListAllowsEveryone(t *testing.T) {
	var l List
	expect(t, l, true, "127.0.0.1", "10.9.9.9", "::1", "fd12::1", "::ffff:192.168.1.20")
	if s := l.String(); s != "all" {
		t.Errorf("nil List String() = %q, want %q", s, "all")
	}
}

// ---- REQ-ALLOW-002 ----

func TestREQ_ALLOW_002_ListedAddressOrBlockMatches(t *testing.T) {
	t.Run("TC-ALLOW-003", func(t *testing.T) {
		l := mustParse(t, "10.0.0.5")
		expect(t, l, true, "10.0.0.5")
		expect(t, l, false, "10.0.0.6", "10.0.0.4", "10.0.0.50")
	})
	t.Run("TC-ALLOW-004", func(t *testing.T) {
		l := mustParse(t, "192.168.1.0/24")
		expect(t, l, true, "192.168.1.0", "192.168.1.20", "192.168.1.255")
		expect(t, l, false, "192.168.0.255", "192.168.2.0", "192.169.1.20")
	})
	t.Run("TC-ALLOW-005", func(t *testing.T) {
		l := mustParse(t, "192.168.1.0/24", "10.0.0.5")
		expect(t, l, true, "192.168.1.77", "10.0.0.5")
		expect(t, l, false, "10.0.0.6", "192.168.2.1")
	})
	t.Run("TC-ALLOW-006", func(t *testing.T) {
		l := mustParse(t, "10.0.0.5/32")
		expect(t, l, true, "10.0.0.5")
		expect(t, l, false, "10.0.0.4", "10.0.0.6")
		all := mustParse(t, "0.0.0.0/0")
		expect(t, all, true, "1.2.3.4", "255.255.255.255", "127.0.0.1")
		expect(t, all, false, "::1") // not the same as no list
		if str := all.String(); !strings.Contains(str, "0.0.0.0/0") {
			t.Errorf("String() of [0.0.0.0/0] = %q", str)
		}
	})
	t.Run("TC-ALLOW-007", func(t *testing.T) { // Claude decision: host bits set = the block
		l := mustParse(t, "192.168.1.5/24")
		expect(t, l, true, "192.168.1.200", "192.168.1.5")
		expect(t, l, false, "192.168.2.5")
	})
	t.Run("TC-ALLOW-008", func(t *testing.T) { // Claude decision: duplicates are accepted
		l := mustParse(t, "10.0.0.5", "10.0.0.5", "10.0.0.0/8")
		expect(t, l, true, "10.0.0.5", "10.200.0.1")
	})
}

// ---- REQ-ALLOW-004 ----

func TestREQ_ALLOW_004_IPv6AndMappedIPv4(t *testing.T) {
	t.Run("TC-ALLOW-018", func(t *testing.T) {
		l := mustParse(t, "fd00::/8", "2001:db8::7")
		expect(t, l, true, "fd12::1", "fdff:ffff::1", "2001:db8::7")
		expect(t, l, false, "fe80::1", "fc00::1", "2001:db8::8", "::1")
	})
	t.Run("TC-ALLOW-019", func(t *testing.T) {
		l := mustParse(t, "10.0.0.5", "192.168.1.0/24")
		mapped := net.ParseIP("::ffff:10.0.0.5") // 16-byte form, as a dual-stack socket reports it
		if len(mapped) != net.IPv6len {
			t.Fatalf("test setup: mapped ip has %d bytes", len(mapped))
		}
		if !l.Allows(mapped) {
			t.Errorf("Allows(::ffff:10.0.0.5) = false, want true (matched as 10.0.0.5)")
		}
		expect(t, l, true, "::ffff:192.168.1.9")
		expect(t, l, false, "::ffff:10.0.0.6", "::ffff:192.168.2.9")
	})
	t.Run("TC-ALLOW-020", func(t *testing.T) {
		v4 := mustParse(t, "10.0.0.5", "0.0.0.0/0")
		expect(t, v4, false, "::a00:5", "::1", "fd00::5") // IPv4-compatible (not mapped) and real IPv6
		v6 := mustParse(t, "::/0")
		expect(t, v6, false, "10.0.0.5", "127.0.0.1")
		expect(t, v6, true, "::1", "fd00::5")
	})
}

// TC-ALLOW-043: no client IP must never pass a list (fail closed).
func TestREQ_ALLOW_004_NilIPNotAllowedByAList(t *testing.T) {
	for _, items := range [][]string{{"10.0.0.5"}, {"0.0.0.0/0"}, {"::/0"}} {
		if mustParse(t, items...).Allows(nil) {
			t.Errorf("List %q Allows(nil) = true, want false", items)
		}
	}
}

// TC-ALLOW-044: a mapped item in the list matches the plain IPv4 client.
func TestREQ_ALLOW_004_MappedItemMatchesIPv4Client(t *testing.T) {
	l := mustParse(t, "::ffff:10.0.0.5")
	expect(t, l, true, "10.0.0.5", "::ffff:10.0.0.5")
	expect(t, l, false, "10.0.0.6")
}

// ---- REQ-ALLOW-006 ----

// TC-ALLOW-025
func TestREQ_ALLOW_006_ParseNamesFirstInvalidItem(t *testing.T) {
	invalid := []string{"nope", "10.0.0.256", "10.0.0.0/33", "fd00::/129", "", "printer.example", " 10.0.0.5",
		"10.0.0.5 ", "fe80::1%eth0", "1.2.3.4/24/1", "10.0.0.5/", "10.0.0", "/24"}
	for _, bad := range invalid {
		_, idx, err := Parse([]string{"10.0.0.5", bad, "also-bad"})
		if err == nil || idx != 1 {
			t.Errorf("Parse([10.0.0.5 %q also-bad]) = index %d, err %v; want index 1 and an error", bad, idx, err)
		}
	}
	if _, idx, err := Parse([]string{"x"}); err == nil || idx != 0 {
		t.Errorf("Parse([x]) = index %d, err %v; want index 0 and an error", idx, err)
	}
	valid := []string{"10.0.0.5", "::1", "fd00::/8", "192.168.1.0/24", "2001:DB8::1", "::ffff:10.0.0.5"}
	if _, idx, err := Parse(valid); err != nil {
		t.Errorf("Parse(%q) = index %d error %v, want nil", valid, idx, err)
	}
}

// TC-ALLOW-039: Go 1.10 net.ParseIP reads `010.0.0.5` as 10.0.0.5 while Go 1.17+ refuses it; both editions must
// refuse it the same way (Claude decision).
func TestREQ_ALLOW_006_LeadingZeroRefusedInBothEditions(t *testing.T) {
	for _, bad := range []string{"010.0.0.5", "10.0.0.05", "010.0.0.0/8", "192.168.001.020", "10.0.0.0/08"} {
		if _, idx, err := Parse([]string{"10.0.0.5", bad}); err == nil || idx != 1 {
			t.Errorf("Parse([10.0.0.5 %q]) = index %d, err %v; want index 1 and an error", bad, idx, err)
		}
	}
}

// ---- REQ-ALLOW-008 ----

// TC-ALLOW-031
func TestREQ_ALLOW_008_StringListsEveryItem(t *testing.T) {
	l := mustParse(t, "192.168.1.0/24", "10.0.0.5", "fd00::/8")
	s := l.String()
	for _, want := range []string{"192.168.1.0/24", "10.0.0.5", "fd00::/8"} {
		if !strings.Contains(s, want) {
			t.Errorf("String() = %q, missing %q", s, want)
		}
	}
	if s == "all" {
		t.Errorf("String() of a non-empty List = %q", s)
	}
}

// TC-ALLOW-045: Go 1.10 reads the embedded IPv4 of `::ffff:010.0.0.5` with its leading zero while Go 1.17+ refuses
// it; both editions must refuse it the same way.
func TestREQ_ALLOW_006_LeadingZeroInEmbeddedIPv4Refused(t *testing.T) {
	for _, bad := range []string{"::ffff:010.0.0.5", "::ffff:10.0.0.05", "::ffff:010.0.0.0/104", "64:ff9b::192.168.001.1"} {
		if !hasLeadingZeroIPv4(strings.Split(bad, "/")[0]) {
			t.Errorf("hasLeadingZeroIPv4(%q) = false, want true (Go 1.10 would accept it)", bad)
		}
		if _, idx, err := Parse([]string{"10.0.0.5", bad}); err == nil || idx != 1 {
			t.Errorf("Parse([10.0.0.5 %q]) = index %d, err %v; want index 1 and an error", bad, idx, err)
		}
	}
	for _, ok := range []string{"::ffff:10.0.0.5", "::ffff:0.0.0.0", "fd00::10", "2001:db8::0", "::ffff:10.0.0.0"} {
		if hasLeadingZeroIPv4(ok) {
			t.Errorf("hasLeadingZeroIPv4(%q) = true, want false", ok)
		}
	}
}
