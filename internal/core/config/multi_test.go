package config

import (
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"testing"
)

// F-003 config tests (spec docs/specs/2026-11-07-f003-multi-printer.md §2). TC ids: .testcases/runs/F-003-r1 and
// testcases.md.

// entries returns n entries {"printer":"P","port":9101…9100+n} without names.
func entries(n int) string {
	parts := make([]string, n)
	for i := range parts {
		parts[i] = `{"printer":"P","port":` + strconv.Itoa(9101+i) + `}`
	}
	return strings.Join(parts, ",")
}

func multiCfg(printers string) string { return `{"version":2,"printers":[` + printers + `]}` }

// loadMulti loads body and fails the test on any error.
func loadMulti(t *testing.T, body string) []Printer {
	t.Helper()
	d := tempDir(t)
	defer os.RemoveAll(d)
	f := writeFile(t, filepath.Join(d, "c.json"), body)
	c, err := Load(Lookup{Flag: f})
	if err != nil {
		t.Fatalf("Load(%s) error = %v, want nil", body, err)
	}
	return c.Printers
}

// failMulti loads body and expects an exit-2 error naming exactly field.
func failMulti(t *testing.T, body, field string) {
	t.Helper()
	d := tempDir(t)
	defer os.RemoveAll(d)
	f := writeFile(t, filepath.Join(d, "c.json"), body)
	e := mustFail(t, Lookup{Flag: f})
	if !strings.Contains(e.Msg, f) || !namesField(e.Msg, field) {
		t.Fatalf("Msg %q, want path %q and exactly field %q", e.Msg, f, field)
	}
}

// ---- REQ-MULTI-001 ----

func TestREQ_MULTI_001_PrinterCount(t *testing.T) {
	t.Run("TC-MULTI-001", func(t *testing.T) {
		got := loadMulti(t, multiCfg(`{"name":"front","printer":"A","port":9101},{"name":"back","printer":"B","port":9102}`))
		want := []Printer{{"front", "A", 9101}, {"back", "B", 9102}}
		if !reflect.DeepEqual(got, want) {
			t.Fatalf("Printers = %+v, want %+v", got, want)
		}
	})
	t.Run("TC-MULTI-002", func(t *testing.T) {
		got := loadMulti(t, multiCfg(entries(32)))
		if len(got) != 32 || got[31].Port != 9132 || got[31].Name != "printer32" {
			t.Fatalf("got %d printers, last %+v; want 32, last {printer32 P 9132}", len(got), got[len(got)-1])
		}
	})
	t.Run("TC-MULTI-003", func(t *testing.T) { failMulti(t, multiCfg(entries(33)), "printers") })
	t.Run("TC-MULTI-004", func(t *testing.T) { failMulti(t, multiCfg(``), "printers") })
	t.Run("TC-MULTI-005", func(t *testing.T) {
		got := loadMulti(t, multiCfg(`{"printer":"P","port":9101}`))
		if !reflect.DeepEqual(got, []Printer{{"default", "P", 9101}}) {
			t.Fatalf("Printers = %+v, want [{default P 9101}]", got)
		}
	})
}

// ---- REQ-MULTI-002 ----

func TestREQ_MULTI_002_InvalidEntryNamesIndex(t *testing.T) {
	e1 := `{"printer":"P","port":9101}`
	cases := []struct{ tc, printers, field string }{
		{"TC-MULTI-006", entries(2) + `,{"printer":"P","port":65536}`, "printers[2].port"},
		{"TC-MULTI-007/string", e1 + `,"Canon"`, "printers[1]"},
		{"TC-MULTI-007/null", e1 + `,null`, "printers[1]"},
		{"TC-MULTI-008/missing", e1 + `,{"port":9102}`, "printers[1].printer"},
		{"TC-MULTI-008/empty", e1 + `,{"printer":"","port":9102}`, "printers[1].printer"},
		{"TC-MULTI-008/number", e1 + `,{"printer":5,"port":9102}`, "printers[1].printer"},
		{"TC-MULTI-009/string", e1 + `,{"printer":"P","port":"9102"}`, "printers[1].port"},
		{"TC-MULTI-009/fraction", e1 + `,{"printer":"P","port":9102.5}`, "printers[1].port"},
		{"TC-MULTI-009/zero", e1 + `,{"printer":"P","port":0}`, "printers[1].port"},
		{"TC-MULTI-010", e1 + `,{"printer":"P","port":0},{"printer":"P","port":70000}`, "printers[1].port"},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) { failMulti(t, multiCfg(c.printers), c.field) })
	}
}

// ---- REQ-MULTI-003 ----

func TestREQ_MULTI_003_DuplicatePortNamesSecond(t *testing.T) {
	cases := []struct{ tc, printers, field string }{
		{"TC-MULTI-012", entries(2) + `,{"printer":"P","port":9101}`, "printers[2].port"},
		{"TC-MULTI-013", `{"printer":"P"},{"printer":"Q","port":9100}`, "printers[1].port"},
		{"TC-MULTI-014", `{"printer":"P","port":9101},{"printer":"P","port":9101},{"printer":"P","port":0}`, "printers[2].port"},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) { failMulti(t, multiCfg(c.printers), c.field) })
	}
}

// ---- REQ-MULTI-004 ----

func TestREQ_MULTI_004_Names(t *testing.T) {
	t.Run("TC-MULTI-015", func(t *testing.T) {
		failMulti(t, multiCfg(`{"name":"front","printer":"P","port":9101},{"name":"front","printer":"Q","port":9102}`), "printers[1].name")
	})
	t.Run("TC-MULTI-016", func(t *testing.T) {
		got := loadMulti(t, multiCfg(`{"name":"","printer":"P","port":9101},{"name":"","printer":"Q","port":9102}`))
		if len(got) != 2 || got[0].Name != "printer1" || got[1].Name != "printer2" {
			t.Fatalf("Printers = %+v, want names printer1, printer2", got)
		}
	})
	t.Run("TC-MULTI-017", func(t *testing.T) {
		got := loadMulti(t, multiCfg(entries(2)))
		if len(got) != 2 || got[0].Name != "printer1" || got[1].Name != "printer2" {
			t.Fatalf("Printers = %+v, want names printer1, printer2", got)
		}
	})
	t.Run("TC-MULTI-018", func(t *testing.T) {
		failMulti(t, multiCfg(`{"name":"printer2","printer":"P","port":9101},{"printer":"Q","port":9102}`), "printers[1].name")
	})
	t.Run("TC-MULTI-019", func(t *testing.T) {
		got := loadMulti(t, multiCfg(`{"name":"default","printer":"P","port":9101},{"printer":"Q","port":9102}`))
		if len(got) != 2 || got[0].Name != "default" || got[1].Name != "printer2" {
			t.Fatalf("Printers = %+v, want names default, printer2", got)
		}
	})
	t.Run("TC-MULTI-020", func(t *testing.T) {
		got := loadMulti(t, multiCfg(`{"name":"a","printer":"A","port":9101},{"name":"b","printer":"A","port":9102}`))
		if len(got) != 2 || got[0].Printer != "A" || got[1].Printer != "A" {
			t.Fatalf("Printers = %+v, want both on printer A", got)
		}
	})
}
