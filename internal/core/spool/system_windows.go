package spool

import (
	"syscall"
	"unsafe"
)

// The winspool.drv entry points, loaded through syscall.NewLazyDLL so the package builds with
// Go 1.10 and without x/sys (rule 01-go-code).
var (
	winspool             = syscall.NewLazyDLL("winspool.drv")
	procOpenPrinter      = winspool.NewProc("OpenPrinterW")
	procStartDocPrinter  = winspool.NewProc("StartDocPrinterW")
	procStartPagePrinter = winspool.NewProc("StartPagePrinter")
	procWritePrinter     = winspool.NewProc("WritePrinter")
	procEndPagePrinter   = winspool.NewProc("EndPagePrinter")
	procEndDocPrinter    = winspool.NewProc("EndDocPrinter")
	procAbortPrinter     = winspool.NewProc("AbortPrinter")
	procClosePrinter     = winspool.NewProc("ClosePrinter")
	procEnumPrinters     = winspool.NewProc("EnumPrintersW")
)

const (
	enumLocal       = 0x00000002 // PRINTER_ENUM_LOCAL
	enumConnections = 0x00000004 // PRINTER_ENUM_CONNECTIONS
)

// docInfo1 mirrors DOC_INFO_1W: three pointers, the same layout on 386 and amd64.
type docInfo1 struct {
	pDocName    *uint16
	pOutputFile *uint16
	pDatatype   *uint16
}

// printerInfo4 mirrors PRINTER_INFO_4W: two pointers and a DWORD, the same layout on 386 and amd64.
type printerInfo4 struct {
	pPrinterName *uint16
	pServerName  *uint16
	attributes   uint32
}

// System returns the winspool API of this machine.
func System() API { return winAPI{} }

type winAPI struct{}

// winError turns the GetLastError value a LazyProc.Call returns into *Error; Errno.Error() is
// the FormatMessage text. A zero Errno (success with a zero return value) keeps the call name.
func winError(op string, e error) *Error {
	if errno, ok := e.(syscall.Errno); ok {
		if errno != 0 {
			return &Error{Code: uint32(errno), Msg: errno.Error()}
		}
		return &Error{Code: 0, Msg: op + " failed"}
	}
	if e != nil {
		return &Error{Code: 0, Msg: op + ": " + e.Error()}
	}
	return &Error{Code: 0, Msg: op + " failed"}
}

// utf16Ptr converts a Go string or wraps the conversion error in *Error.
func utf16Ptr(op, s string) (*uint16, error) {
	p, err := syscall.UTF16PtrFromString(s)
	if err != nil {
		return nil, &Error{Code: 0, Msg: op + ": " + err.Error()}
	}
	return p, nil
}

func (winAPI) OpenPrinter(name string) (Handle, error) {
	p, err := utf16Ptr("printer name", name)
	if err != nil {
		return 0, err
	}
	var h Handle
	r1, _, e := procOpenPrinter.Call(uintptr(unsafe.Pointer(p)), uintptr(unsafe.Pointer(&h)), 0)
	if r1 == 0 {
		return 0, winError("OpenPrinter", e)
	}
	return h, nil
}

func (winAPI) StartDocPrinter(h Handle, docName, datatype string) error {
	dn, err := utf16Ptr("document name", docName)
	if err != nil {
		return err
	}
	dt, err := utf16Ptr("datatype", datatype)
	if err != nil {
		return err
	}
	info := docInfo1{pDocName: dn, pDatatype: dt}
	r1, _, e := procStartDocPrinter.Call(uintptr(h), 1, uintptr(unsafe.Pointer(&info)))
	if r1 == 0 {
		return winError("StartDocPrinter", e)
	}
	return nil
}

// handleProc wraps the BOOL winspool calls that take only a handle.
func handleProc(p *syscall.LazyProc, op string, h Handle) error {
	r1, _, e := p.Call(uintptr(h))
	if r1 == 0 {
		return winError(op, e)
	}
	return nil
}

func (a winAPI) StartPagePrinter(h Handle) error {
	return handleProc(procStartPagePrinter, "StartPagePrinter", h)
}
func (a winAPI) EndPagePrinter(h Handle) error {
	return handleProc(procEndPagePrinter, "EndPagePrinter", h)
}
func (a winAPI) EndDocPrinter(h Handle) error {
	return handleProc(procEndDocPrinter, "EndDocPrinter", h)
}
func (a winAPI) AbortPrinter(h Handle) error { return handleProc(procAbortPrinter, "AbortPrinter", h) }
func (a winAPI) ClosePrinter(h Handle) error { return handleProc(procClosePrinter, "ClosePrinter", h) }

func (winAPI) WritePrinter(h Handle, data []byte) (int, error) {
	var p uintptr
	if len(data) > 0 {
		p = uintptr(unsafe.Pointer(&data[0]))
	}
	var written uint32
	r1, _, e := procWritePrinter.Call(uintptr(h), p, uintptr(len(data)), uintptr(unsafe.Pointer(&written)))
	if r1 == 0 {
		return 0, winError("WritePrinter", e)
	}
	return int(written), nil
}

func (winAPI) EnumPrinters() ([]string, error) {
	flags := uintptr(enumLocal | enumConnections)
	var needed, returned uint32
	r1, _, e := procEnumPrinters.Call(flags, 0, 4, 0, 0,
		uintptr(unsafe.Pointer(&needed)), uintptr(unsafe.Pointer(&returned)))
	if r1 == 0 && needed == 0 {
		return nil, winError("EnumPrinters", e)
	}
	if needed == 0 {
		return []string{}, nil // no printers installed, the call succeeded
	}
	buf := make([]byte, needed)
	r1, _, e = procEnumPrinters.Call(flags, 0, 4, uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf)),
		uintptr(unsafe.Pointer(&needed)), uintptr(unsafe.Pointer(&returned)))
	if r1 == 0 {
		return nil, winError("EnumPrinters", e)
	}
	names := make([]string, 0, returned)
	sz := int(unsafe.Sizeof(printerInfo4{}))
	for i := 0; i < int(returned); i++ {
		info := (*printerInfo4)(unsafe.Pointer(&buf[i*sz]))
		names = append(names, utf16PtrString(info.pPrinterName))
	}
	return names, nil
}

// utf16PtrString decodes a NUL-terminated UTF-16 string from winspool memory.
func utf16PtrString(p *uint16) string {
	if p == nil {
		return ""
	}
	n := 0
	for ptr := unsafe.Pointer(p); *(*uint16)(ptr) != 0; n++ {
		ptr = unsafe.Pointer(uintptr(ptr) + 2)
	}
	return syscall.UTF16ToString((*[1 << 29]uint16)(unsafe.Pointer(p))[:n:n])
}
