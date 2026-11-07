package spool

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

// fakeAPI records every call in order. fail names the call that returns err; maxWrite caps bytes per Write.
type fakeAPI struct {
	calls    []string
	writes   [][]byte
	accepted []byte // the bytes WritePrinter reported as written, in order
	fail     string
	failAbrt bool // AbortPrinter also fails
	err      error
	maxWrite int
	printers []string
}

// rec records "<name> <arg>" and returns err when name is the call set to fail.
func (f *fakeAPI) rec(name, arg string) error {
	f.calls = append(f.calls, name+" "+arg)
	if name == f.fail {
		return f.err
	}
	return nil
}

func (f *fakeAPI) OpenPrinter(name string) (Handle, error) {
	if err := f.rec("Open", name); err != nil {
		return 0, err
	}
	return 7, nil
}
func (f *fakeAPI) StartDocPrinter(h Handle, docName, datatype string) error {
	f.calls = append(f.calls, "StartDoc "+handle(h)+" "+docName+" "+datatype)
	if f.fail == "StartDoc" {
		return f.err
	}
	return nil
}
func (f *fakeAPI) StartPagePrinter(h Handle) error { return f.rec("StartPage", handle(h)) }
func (f *fakeAPI) WritePrinter(h Handle, p []byte) (int, error) {
	f.calls = append(f.calls, "Write "+handle(h))
	f.writes = append(f.writes, append([]byte(nil), p...))
	if f.fail == "Write" {
		return 0, f.err
	}
	n := len(p)
	if f.maxWrite > 0 && n > f.maxWrite {
		n = f.maxWrite
	}
	f.accepted = append(f.accepted, p[:n]...)
	return n, nil
}
func (f *fakeAPI) EndPagePrinter(h Handle) error { return f.rec("EndPage", handle(h)) }
func (f *fakeAPI) EndDocPrinter(h Handle) error  { return f.rec("EndDoc", handle(h)) }
func (f *fakeAPI) AbortPrinter(h Handle) error {
	f.calls = append(f.calls, "Abort "+handle(h))
	if f.failAbrt {
		return &Error{Code: 6, Msg: "The handle is invalid."}
	}
	return nil
}
func (f *fakeAPI) ClosePrinter(h Handle) error {
	f.calls = append(f.calls, "Close "+handle(h))
	if f.fail == "Close" {
		return f.err
	}
	return nil
}
func (f *fakeAPI) EnumPrinters() ([]string, error) {
	f.calls = append(f.calls, "Enum")
	if f.fail == "Enum" {
		return nil, f.err
	}
	return f.printers, nil
}

func handle(h Handle) string {
	if h == 7 {
		return "7"
	}
	return "?"
}

func (f *fakeAPI) written() []byte {
	var b []byte
	for _, w := range f.writes {
		b = append(b, w...)
	}
	return b
}

func (f *fakeAPI) count(prefix string) int {
	n := 0
	for _, c := range f.calls {
		if strings.HasPrefix(c, prefix) {
			n++
		}
	}
	return n
}

const doc = "PrinterOne job from 10.0.0.5"

// TC-SPOOL-001
func TestREQ_SPOOL_001_CallSequence(t *testing.T) {
	f := &fakeAPI{}
	if err := Print(f, "Canon MF220 Series", doc, []byte("hello")); err != nil {
		t.Fatalf("Print error %v", err)
	}
	want := []string{"Open Canon MF220 Series", "StartDoc 7 " + doc + " RAW", "StartPage 7", "Write 7",
		"EndPage 7", "EndDoc 7", "Close 7"}
	if !reflect.DeepEqual(f.calls, want) {
		t.Fatalf("calls\n %q\nwant\n %q", f.calls, want)
	}
}

// TC-SPOOL-002
func TestREQ_SPOOL_001_BytesUnchanged(t *testing.T) {
	data := make([]byte, 256)
	for i := range data {
		data[i] = byte(i)
	}
	f := &fakeAPI{}
	if err := Print(f, "P", doc, data); err != nil {
		t.Fatalf("Print error %v", err)
	}
	if !bytes.Equal(f.written(), data) {
		t.Fatalf("written %x, want %x", f.written(), data)
	}
}

// TC-SPOOL-003
func TestREQ_SPOOL_001_ShortWritesContinued(t *testing.T) {
	f := &fakeAPI{maxWrite: 3}
	if err := Print(f, "P", doc, []byte("0123456789")); err != nil {
		t.Fatalf("Print error %v", err)
	}
	// Any chunking is fine; what the spooler accepted must be every byte, once, in order.
	if string(f.accepted) != "0123456789" || len(f.writes) < 4 {
		t.Fatalf("spooler accepted %q in %d writes, want 0123456789 in at least 4 writes of at most 3 bytes",
			f.accepted, len(f.writes))
	}
}

// TC-SPOOL-006..011
func TestREQ_SPOOL_002_FailingCallAbortAndCloseRules(t *testing.T) {
	cases := []struct {
		tc, fail string
		code     uint32
		want     []string // calls from the failing one to the end
	}{
		{"TC-SPOOL-006", "Open", 1801, []string{"Open P"}},
		{"TC-SPOOL-007", "StartDoc", 5, []string{"Close 7"}}, // Abort before Close allowed or not: spec open
		{"TC-SPOOL-008", "StartPage", 5, []string{"StartPage 7", "Abort 7", "Close 7"}},
		{"TC-SPOOL-009", "Write", 5, []string{"Write 7", "Abort 7", "Close 7"}},
		{"TC-SPOOL-010", "EndPage", 5, []string{"EndPage 7", "Abort 7", "Close 7"}},
		{"TC-SPOOL-011", "EndDoc", 5, []string{"EndDoc 7", "Abort 7", "Close 7"}},
	}
	for _, c := range cases {
		t.Run(c.tc, func(t *testing.T) {
			f := &fakeAPI{fail: c.fail, err: &Error{Code: c.code, Msg: "Access is denied."}}
			err := Print(f, "P", doc, []byte("0123456789"))
			se, ok := err.(*Error)
			if !ok || se.Code != c.code {
				t.Fatalf("Print error %T %v, want *spool.Error code %d", err, err, c.code)
			}
			if len(f.calls) < len(c.want) || !reflect.DeepEqual(f.calls[len(f.calls)-len(c.want):], c.want) {
				t.Fatalf("calls %q, want them to end with %q", f.calls, c.want)
			}
			if f.count("Close") > 1 || f.count("Abort") > 1 {
				t.Fatalf("calls %q: Close/Abort more than once", f.calls)
			}
			if c.fail == "Open" && len(f.calls) != 1 {
				t.Fatalf("calls %q after a failed OpenPrinter, want none", f.calls)
			}
			if f.count("EndDoc") > 0 && c.fail != "EndDoc" {
				t.Fatalf("calls %q: EndDoc after a failure ends a partial document", f.calls)
			}
		})
	}
}

// TC-SPOOL-012: once EndDocPrinter succeeded the spooler holds the whole job, so a failing ClosePrinter
// does not fail it (spec §2.3, decided 2026-11-07).
func TestREQ_SPOOL_002_CloseFailureAfterEndDocStillPrinted(t *testing.T) {
	f := &fakeAPI{fail: "Close", err: &Error{Code: 6, Msg: "The handle is invalid."}}
	if err := Print(f, "P", doc, []byte("0123456789")); err != nil {
		t.Fatalf("Print error %v, want nil: EndDocPrinter succeeded, the job is printed", err)
	}
	if f.count("Abort") != 0 || f.count("Close") != 1 {
		t.Fatalf("calls %q, want no AbortPrinter and ClosePrinter once", f.calls)
	}
}

// TC-SPOOL-019
func TestREQ_SPOOL_002_CloseCalledEvenWhenAbortFails(t *testing.T) {
	f := &fakeAPI{fail: "Write", err: &Error{Code: 5, Msg: "Access is denied."}, failAbrt: true}
	if err := Print(f, "P", doc, []byte("x")); err == nil {
		t.Fatal("Print error nil, want the job to fail") // which call's code is reported is not fixed by the spec
	}
	if f.count("Close") != 1 || f.calls[len(f.calls)-1] != "Close 7" {
		t.Fatalf("calls %q, want ClosePrinter once, last, although AbortPrinter failed", f.calls)
	}
}

// TC-SPOOL-018
func TestREQ_SPOOL_003_ListReturnsNames(t *testing.T) {
	f := &fakeAPI{printers: []string{"A", "B"}}
	got, err := List(f)
	if err != nil || !reflect.DeepEqual(got, []string{"A", "B"}) {
		t.Fatalf("List = %q, %v; want [A B], nil", got, err)
	}
	f = &fakeAPI{fail: "Enum", err: &Error{Code: 1722, Msg: "The RPC server is unavailable."}}
	if _, err := List(f); err == nil {
		t.Fatal("List error nil, want the EnumPrinters error")
	}
}
