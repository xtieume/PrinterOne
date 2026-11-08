// Package allow is a printer's client allowlist (spec docs/specs/2026-11-07-f004-allowlist.md §2).
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code): net.ParseCIDR / net.ParseIP, no net/netip.
package allow

import (
	"fmt"
	"net"
	"strconv"
	"strings"
)

// List is a parsed `allow` list. A nil List allows every client (no `allow` key).
type List []*net.IPNet

// Parse parses the items of one `allow` list: each an IPv4 / IPv6 address or a CIDR block. On an invalid item it
// returns that item's index and a non-nil error (the first invalid one by index).
func Parse(items []string) (List, int, error) {
	list := make(List, 0, len(items))
	for i, item := range items {
		if item == "" || strings.TrimSpace(item) != item {
			return nil, i, fmt.Errorf("invalid IP or CIDR")
		}
		address := item
		if strings.Contains(item, "/") {
			parts := strings.Split(item, "/")
			if len(parts) != 2 || parts[1] == "" {
				return nil, i, fmt.Errorf("invalid CIDR")
			}
			if len(parts[1]) > 1 && parts[1][0] == '0' {
				return nil, i, fmt.Errorf("invalid prefix")
			}
			address = parts[0]
			if hasLeadingZeroIPv4(address) {
				return nil, i, fmt.Errorf("invalid IPv4 address")
			}
			ip, network, err := net.ParseCIDR(item)
			if err != nil || ip == nil {
				return nil, i, fmt.Errorf("invalid CIDR")
			}
			list = append(list, network)
			continue
		}
		if hasLeadingZeroIPv4(item) {
			return nil, i, fmt.Errorf("invalid IPv4 address")
		}
		ip := net.ParseIP(item)
		if ip == nil {
			return nil, i, fmt.Errorf("invalid IP address")
		}
		if v4 := ip.To4(); v4 != nil {
			list = append(list, &net.IPNet{IP: v4, Mask: net.CIDRMask(32, 32)})
		} else {
			list = append(list, &net.IPNet{IP: ip, Mask: net.CIDRMask(128, 128)})
		}
	}
	return list, 0, nil
}

// Allows reports whether a client with this IP may connect. A nil List allows everyone; an IPv4-mapped IPv6 ip
// (::ffff:a.b.c.d) is matched as a.b.c.d.
func (l List) Allows(ip net.IP) bool {
	if l == nil {
		return true
	}
	if ip == nil {
		return false
	}
	if v4 := ip.To4(); v4 != nil {
		ip = v4
	}
	for _, network := range l {
		candidate := ip
		if network.IP.To4() != nil {
			candidate = ip.To4()
		}
		if candidate != nil && network.Contains(candidate) {
			return true
		}
	}
	return false
}

// String is the start-up text after `allow: `: "all" for a nil List, else the items separated by ", ".
func (l List) String() string {
	if l == nil {
		return "all"
	}
	items := make([]string, 0, len(l))
	for _, n := range l {
		ones, bits := n.Mask.Size()
		if ones == bits {
			items = append(items, n.IP.String())
		} else {
			items = append(items, n.String())
		}
	}
	return strings.Join(items, ", ")
}

// hasLeadingZeroIPv4 reports a leading zero in a dotted IPv4, alone or embedded at the end of an IPv6 address
// (`::ffff:010.0.0.5`): Go 1.10 accepts those, Go 1.17+ refuses them.
func hasLeadingZeroIPv4(s string) bool {
	s = s[strings.LastIndex(s, ":")+1:]
	parts := strings.Split(s, ".")
	if len(parts) != 4 {
		return false
	}
	for _, part := range parts {
		if len(part) > 1 && part[0] == '0' {
			return true
		}
		if _, err := strconv.Atoi(part); err != nil {
			return false
		}
	}
	return false
}
