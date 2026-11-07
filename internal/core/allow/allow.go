// Package allow is a printer's client allowlist (spec docs/specs/2026-11-07-f004-allowlist.md §2).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code): net.ParseCIDR / net.ParseIP, no net/netip.
package allow

import "net"

// List is a parsed `allow` list. A nil List allows every client (no `allow` key).
type List []*net.IPNet

// Parse parses the items of one `allow` list: each an IPv4 / IPv6 address or a CIDR block. On an invalid item it
// returns that item's index and a non-nil error (the first invalid one by index).
func Parse(items []string) (List, int, error) { return nil, 0, nil }

// Allows reports whether a client with this IP may connect. A nil List allows everyone; an IPv4-mapped IPv6 ip
// (::ffff:a.b.c.d) is matched as a.b.c.d.
func (l List) Allows(ip net.IP) bool { return false }

// String is the start-up text after `allow: `: "all" for a nil List, else the items separated by ", ".
func (l List) String() string { return "" }
