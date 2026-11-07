//go:build windows
// +build windows

package service

import (
	"runtime"
	"strconv"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

var (
	advapi32                   = syscall.NewLazyDLL("advapi32.dll")
	ntdll                      = syscall.NewLazyDLL("ntdll.dll")
	procOpenSCManager          = advapi32.NewProc("OpenSCManagerW")
	procCreateService          = advapi32.NewProc("CreateServiceW")
	procOpenService            = advapi32.NewProc("OpenServiceW")
	procChangeServiceConfig    = advapi32.NewProc("ChangeServiceConfigW")
	procChangeServiceConfig2   = advapi32.NewProc("ChangeServiceConfig2W")
	procQueryServiceConfig     = advapi32.NewProc("QueryServiceConfigW")
	procStartService           = advapi32.NewProc("StartServiceW")
	procControlService         = advapi32.NewProc("ControlService")
	procQueryServiceStatus     = advapi32.NewProc("QueryServiceStatus")
	procDeleteService          = advapi32.NewProc("DeleteService")
	procCloseServiceHandle     = advapi32.NewProc("CloseServiceHandle")
	procStartServiceDispatcher = advapi32.NewProc("StartServiceCtrlDispatcherW")
	procRegisterServiceHandler = advapi32.NewProc("RegisterServiceCtrlHandlerExW")
	procSetServiceStatus       = advapi32.NewProc("SetServiceStatus")
	procRtlGetVersion          = ntdll.NewProc("RtlGetVersion")
)

func init() {
	for _, p := range []*syscall.LazyProc{procOpenSCManager, procCreateService, procOpenService, procChangeServiceConfig,
		procChangeServiceConfig2, procQueryServiceConfig, procStartService, procControlService, procQueryServiceStatus,
		procDeleteService, procCloseServiceHandle, procStartServiceDispatcher, procRegisterServiceHandler, procSetServiceStatus, procRtlGetVersion} {
		_ = p.Find()
	}
}

const (
	scManagerConnect                    = 0x0001
	scManagerCreate                     = 0x0002
	scManagerAllAccess                  = 0xF003F
	serviceQueryConfig                  = 0x0001
	serviceChangeConfig                 = 0x0002
	serviceQueryStatus                  = 0x0004
	serviceStart                        = 0x0010
	serviceStop                         = 0x0020
	serviceDelete                       = 0x00010000
	serviceAllAccess                    = 0xF01FF
	serviceWin32OwnProcess              = 0x00000010
	serviceErrorNormal                  = 0x00000001
	serviceConfigDesc                   = 1
	serviceControlStop                  = 1
	serviceControlShutdown              = 5
	serviceAcceptStop                   = 1
	serviceAcceptShutdown               = 4
	serviceStopped                      = 1
	serviceStartPending                 = 2
	serviceStopPending                  = 3
	serviceRunning                      = 4
	errorInsufficientBuffer             = 122
	errorServiceSpecificError           = 1066
	errorFailedServiceControllerConnect = 1063
)

type serviceStatus struct {
	serviceType, currentState, controlsAccepted uint32
	win32ExitCode, serviceSpecificExitCode      uint32
	checkPoint, waitHint                        uint32
}

type serviceTableEntry struct {
	name *uint16
	proc uintptr
}

type queryServiceConfig struct {
	serviceType, startType, errorControl uint32
	binaryPath, loadOrderGroup           *uint16
	tagID                                uint32
	dependencies, startName, displayName *uint16
}

type serviceDescription struct{ description *uint16 }

type osVersionInfo struct {
	size, major, minor, build, platform uint32
	version                             [128]uint16
}

type windowsSCM struct{}

func callProc(p *syscall.LazyProc, args ...uintptr) (uintptr, uintptr, error) {
	if err := p.Find(); err != nil {
		return 0, 0, err
	}
	return p.Call(args...)
}

func System() Deps {
	return Deps{SCM: windowsSCM{}, Netsh: Netsh, OSMajor: osMajorVersion(), Dispatch: windowsDispatch,
		Sleep: time.Sleep, Now: time.Now}
}

func (windowsSCM) IsAdmin() bool {
	h, _, e := callProc(procOpenSCManager, 0, 0, scManagerAllAccess)
	if h == 0 {
		return false
	}
	callProc(procCloseServiceHandle, h)
	_ = e
	return true
}

func (windowsSCM) Exists(name string) (bool, error) {
	h, err := openService(name, serviceQueryStatus)
	if err != nil {
		if isWinErr(err, 1060) {
			return false, nil
		}
		return false, err
	}
	closeHandle(h)
	return true, nil
}

func (windowsSCM) Create(name string, s Settings) error {
	m, err := openManager(scManagerConnect | scManagerCreate)
	if err != nil {
		return err
	}
	defer closeHandle(m)
	n, err := wide(name)
	if err != nil {
		return err
	}
	dn, err := wide(s.DisplayName)
	if err != nil {
		return err
	}
	bin, err := wide(s.BinaryPath)
	if err != nil {
		return err
	}
	acct, err := wide("LocalSystem")
	if err != nil {
		return err
	}
	h, _, callErr := callProc(procCreateService, m, uintptr(unsafe.Pointer(n)), uintptr(unsafe.Pointer(dn)), serviceAllAccess,
		serviceWin32OwnProcess, uintptr(s.StartType), serviceErrorNormal, uintptr(unsafe.Pointer(bin)), 0, 0, 0,
		uintptr(unsafe.Pointer(acct)), 0)
	runtime.KeepAlive(n)
	runtime.KeepAlive(dn)
	runtime.KeepAlive(bin)
	runtime.KeepAlive(acct)
	if h == 0 {
		return winServiceError("CreateService", callErr)
	}
	defer closeHandle(h)
	if err = setDescription(h, s.Description); err != nil {
		if r, _, e := callProc(procDeleteService, h); r == 0 {
			return winServiceError("DeleteService", e)
		}
		return err
	}
	return nil
}

func (windowsSCM) Update(name string, s Settings) error {
	h, err := openService(name, serviceChangeConfig)
	if err != nil {
		return err
	}
	defer closeHandle(h)
	bin, err := wide(s.BinaryPath)
	if err != nil {
		return err
	}
	acct, err := wide("LocalSystem")
	if err != nil {
		return err
	}
	if r, _, e := callProc(procChangeServiceConfig, h, ^uintptr(0), uintptr(s.StartType), ^uintptr(0), uintptr(unsafe.Pointer(bin)), 0, 0, 0, uintptr(unsafe.Pointer(acct)), 0, 0); r == 0 {
		runtime.KeepAlive(bin)
		runtime.KeepAlive(acct)
		return winServiceError("ChangeServiceConfig", e)
	}
	runtime.KeepAlive(bin)
	runtime.KeepAlive(acct)
	dn, err := wide(s.DisplayName)
	if err != nil {
		return err
	}
	if r, _, e := callProc(procChangeServiceConfig, h, ^uintptr(0), ^uintptr(0), ^uintptr(0), 0, 0, 0, 0, 0, 0, uintptr(unsafe.Pointer(dn))); r == 0 {
		runtime.KeepAlive(dn)
		return winServiceError("ChangeServiceConfig", e)
	}
	runtime.KeepAlive(dn)
	return setDescription(h, s.Description)
}

func (windowsSCM) Config(name string) (Settings, error) {
	h, err := openService(name, serviceQueryConfig)
	if err != nil {
		return Settings{}, err
	}
	defer closeHandle(h)
	var needed uint32
	callProc(procQueryServiceConfig, h, 0, 0, uintptr(unsafe.Pointer(&needed)))
	runtime.KeepAlive(&needed)
	if needed == 0 {
		return Settings{}, winServiceError("QueryServiceConfig", syscall.GetLastError())
	}
	buf := make([]byte, needed)
	if r, _, e := callProc(procQueryServiceConfig, h, uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf)), uintptr(unsafe.Pointer(&needed))); r == 0 {
		runtime.KeepAlive(buf)
		runtime.KeepAlive(&needed)
		return Settings{}, winServiceError("QueryServiceConfig", e)
	}
	runtime.KeepAlive(buf)
	runtime.KeepAlive(&needed)
	c := (*queryServiceConfig)(unsafe.Pointer(&buf[0]))
	return Settings{BinaryPath: utf16String(c.binaryPath), DisplayName: utf16String(c.displayName), StartType: c.startType, Account: utf16String(c.startName)}, nil
}

func (windowsSCM) Start(name string) error {
	h, err := openService(name, serviceStart)
	if err != nil {
		return err
	}
	defer closeHandle(h)
	if r, _, e := callProc(procStartService, h, 0, 0); r == 0 {
		return winServiceError("StartService", e)
	}
	return nil
}

func (windowsSCM) Stop(name string) error {
	h, err := openService(name, serviceStop)
	if err != nil {
		return err
	}
	defer closeHandle(h)
	var s serviceStatus
	if r, _, e := callProc(procControlService, h, serviceControlStop, uintptr(unsafe.Pointer(&s))); r == 0 {
		runtime.KeepAlive(&s)
		return winServiceError("ControlService", e)
	}
	runtime.KeepAlive(&s)
	return nil
}

func (windowsSCM) Delete(name string) error {
	h, err := openService(name, serviceDelete)
	if err != nil {
		return err
	}
	defer closeHandle(h)
	if r, _, e := callProc(procDeleteService, h); r == 0 {
		return winServiceError("DeleteService", e)
	}
	return nil
}

func (windowsSCM) State(name string) (State, uint32, uint32, error) {
	h, err := openService(name, serviceQueryStatus)
	if err != nil {
		return 0, 0, 0, err
	}
	defer closeHandle(h)
	var s serviceStatus
	if r, _, e := callProc(procQueryServiceStatus, h, uintptr(unsafe.Pointer(&s))); r == 0 {
		runtime.KeepAlive(&s)
		return 0, 0, 0, winServiceError("QueryServiceStatus", e)
	}
	runtime.KeepAlive(&s)
	return State(s.currentState), s.win32ExitCode, s.serviceSpecificExitCode, nil
}

func openManager(access uintptr) (uintptr, error) {
	h, _, e := callProc(procOpenSCManager, 0, 0, access)
	if h == 0 {
		return 0, winServiceError("OpenSCManager", e)
	}
	return h, nil
}
func openService(name string, access uintptr) (uintptr, error) {
	m, err := openManager(scManagerConnect)
	if err != nil {
		return 0, err
	}
	defer closeHandle(m)
	n, err := wide(name)
	if err != nil {
		return 0, err
	}
	h, _, e := callProc(procOpenService, m, uintptr(unsafe.Pointer(n)), access)
	runtime.KeepAlive(n)
	if h == 0 {
		return 0, winServiceError("OpenService", e)
	}
	return h, nil
}
func closeHandle(h uintptr) { callProc(procCloseServiceHandle, h) }
func setDescription(h uintptr, text string) error {
	p, err := wide(text)
	if err != nil {
		return err
	}
	d := serviceDescription{description: p}
	if r, _, e := callProc(procChangeServiceConfig2, h, serviceConfigDesc, uintptr(unsafe.Pointer(&d))); r == 0 {
		runtime.KeepAlive(&d)
		runtime.KeepAlive(p)
		return winServiceError("ChangeServiceConfig2", e)
	}
	runtime.KeepAlive(&d)
	runtime.KeepAlive(p)
	return nil
}
func wide(s string) (*uint16, error) {
	p, e := syscall.UTF16PtrFromString(s)
	if e != nil {
		return nil, &Error{Code: 1, Msg: e.Error()}
	}
	return p, nil
}
func utf16String(p *uint16) string {
	if p == nil {
		return ""
	}
	n := 0
	for q := unsafe.Pointer(p); *(*uint16)(q) != 0; n++ {
		q = unsafe.Pointer(uintptr(q) + 2)
	}
	return syscall.UTF16ToString((*[1 << 29]uint16)(unsafe.Pointer(p))[:n:n])
}
func winServiceError(op string, err error) error {
	code := uint32(0)
	if e, ok := err.(syscall.Errno); ok {
		code = uint32(e)
	}
	msg := op + " failed"
	if code != 0 {
		msg = op + " (" + strconv.FormatUint(uint64(code), 10) + "): " + syscall.Errno(code).Error()
	}
	return &Error{Code: 1, Msg: msg}
}
func isWinErr(err error, code uintptr) bool {
	e, ok := err.(*Error)
	return ok && strings.HasPrefix(e.Msg, "OpenService ("+strconv.FormatUint(uint64(code), 10)+"):")
}

func osMajorVersion() int {
	v := osVersionInfo{size: uint32(unsafe.Sizeof(osVersionInfo{}))}
	r, _, _ := callProc(procRtlGetVersion, uintptr(unsafe.Pointer(&v)))
	runtime.KeepAlive(&v)
	if r != 0 {
		return 0
	}
	return int(v.major)
}

type windowsSession struct {
	handle   uintptr
	controls chan Control
}

func (s *windowsSession) Controls() <-chan Control { return s.controls }
func (s *windowsSession) Report(state State, code uint32) {
	status := serviceStatus{serviceType: serviceWin32OwnProcess, currentState: uint32(state)}
	if state == Running {
		status.controlsAccepted = serviceAcceptStop | serviceAcceptShutdown
	}
	if state == StartPending {
		status.waitHint = 30000
		status.checkPoint = 1
	}
	if state == StopPending {
		status.waitHint = 30000
		status.checkPoint = 1
	}
	if code != 0 {
		status.win32ExitCode = errorServiceSpecificError
		status.serviceSpecificExitCode = code
	}
	callProc(procSetServiceStatus, s.handle, uintptr(unsafe.Pointer(&status)))
	runtime.KeepAlive(&status)
}
func (s *windowsSession) ReportPending(checkpoint uint32) {
	status := serviceStatus{serviceType: serviceWin32OwnProcess, currentState: serviceStopPending, waitHint: 5000, checkPoint: checkpoint}
	callProc(procSetServiceStatus, s.handle, uintptr(unsafe.Pointer(&status)))
	runtime.KeepAlive(&status)
}

var activeSession *windowsSession

func windowsDispatch(name string, main func(Session)) error {
	n, err := wide(name)
	if err != nil {
		return err
	}
	entry := [2]serviceTableEntry{{name: n}, {}}
	entry[0].proc = syscall.NewCallback(func(argc uintptr, argv uintptr) uintptr {
		s := activeSession
		handlerName, _ := wide(name)
		h, _, _ := callProc(procRegisterServiceHandler, uintptr(unsafe.Pointer(handlerName)), syscall.NewCallback(func(control, eventType, eventData, context uintptr) uintptr {
			if control == serviceControlStop {
				s.controls <- ControlStop
				return 0
			}
			if control == serviceControlShutdown {
				s.controls <- ControlShutdown
				return 0
			}
			if control == 4 {
				return 0
			} // SERVICE_CONTROL_INTERROGATE
			return 120 // ERROR_CALL_NOT_IMPLEMENTED
		}), 0)
		runtime.KeepAlive(handlerName)
		s.handle = h
		main(s)
		return 0
	})
	s := &windowsSession{controls: make(chan Control, 2)}
	activeSession = s
	r, _, e := callProc(procStartServiceDispatcher, uintptr(unsafe.Pointer(&entry[0])))
	runtime.KeepAlive(entry)
	runtime.KeepAlive(n)
	activeSession = nil
	if r == 0 {
		if isErrno(e, errorFailedServiceControllerConnect) {
			return ErrNotService
		}
		return winServiceError("StartServiceCtrlDispatcher", e)
	}
	return nil
}
func isErrno(err error, code uintptr) bool {
	e, ok := err.(syscall.Errno)
	return ok && uintptr(e) == code
}
