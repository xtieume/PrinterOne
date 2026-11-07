// Package receiver accepts RAW print jobs over TCP.
// Core package: Go 1.10 syntax and stdlib only (rule 01-go-code).
package receiver

import (
	"bufio"
	"fmt"
	"io"
	"net"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/xtieume/PrinterOne/internal/core/spool"
)

const (
	MaxJobBytes int64 = 268435456
	IdleTimeout       = 60 * time.Second
	SlotsMain         = 8
	SlotsLegacy       = 2
	AcceptWait        = 200 * time.Millisecond // longest a port holds a shared slot while no client arrives
)

type PrintFunc func(docName string, data []byte) error

type Server struct {
	Addr      string
	Listener  net.Listener
	Printer   string
	Print     PrintFunc
	MaxBytes  int64
	Idle      time.Duration
	Slots     int
	Pool      *SlotPool // shared receive slots (F-003); nil: Slots of its own
	PrintLock *sync.Mutex
	Log       io.Writer
	Now       func() time.Time
	Allow     func(net.IP) bool // F-004: false → closed before any Read, logged `0B … rejected: not allowed`; nil: all

	mu           sync.Mutex
	ln           net.Listener
	conns        map[net.Conn]bool
	stopping     bool
	serveDone    chan struct{}
	stop         chan struct{}
	shutdownOnce sync.Once
	logs         sync.Mutex
}

// SlotPool is a set of receive slots shared by every Server of one process (spec
// docs/specs/2026-11-07-f003-multi-printer.md §2: 8 main / 2 legacy for all ports together). A Server whose Pool
// is set takes its slots from it and ignores Slots.
type SlotPool struct {
	n   int
	sem chan struct{}
}

// NewSlotPool returns a pool of n receive slots.
func NewSlotPool(n int) *SlotPool { return &SlotPool{n: n, sem: make(chan struct{}, n)} }

type PortBusyError struct{ Port int }

func (e *PortBusyError) Error() string { return fmt.Sprintf("port %d is already in use", e.Port) }

func (s *Server) Listen() error {
	ln := s.Listener
	var err error
	if ln == nil {
		ln, err = net.Listen("tcp", s.Addr)
	}
	if err != nil {
		if a, ok := net.ResolveTCPAddr("tcp", s.Addr); ok == nil && a.Port != 0 {
			if isAddressInUse(err) {
				return &PortBusyError{Port: a.Port}
			}
		}
		return err
	}
	s.mu.Lock()
	s.ln = ln
	s.conns = make(map[net.Conn]bool)
	s.serveDone = make(chan struct{})
	s.stop = make(chan struct{})
	s.mu.Unlock()
	return nil
}

func (s *Server) ListenAddr() string {
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.ln == nil {
		return ""
	}
	a := s.ln.Addr().String()
	if ta, ok := s.ln.Addr().(*net.TCPAddr); ok && (ta.IP == nil || ta.IP.IsUnspecified()) {
		return "0.0.0.0:" + strconv.Itoa(ta.Port)
	}
	return a
}

// CloseListener releases a bound listener before Serve has started.
func (s *Server) CloseListener() {
	s.mu.Lock()
	if s.ln != nil {
		s.ln.Close()
		s.ln = nil
	}
	s.mu.Unlock()
}

type printJob struct {
	conn    string
	ip      string
	data    []byte
	release func()
}

func (s *Server) Serve() error {
	s.mu.Lock()
	ln := s.ln
	stop := s.stop
	done := s.serveDone
	s.mu.Unlock()
	if ln == nil {
		return fmt.Errorf("receiver: Listen must be called first")
	}
	jobs := make(chan printJob)
	var handlers sync.WaitGroup
	var printer sync.WaitGroup
	printer.Add(1)
	go func() {
		defer printer.Done()
		for j := range jobs {
			if s.PrintLock != nil {
				s.PrintLock.Lock()
			}
			err := s.Print("PrinterOne job from "+j.ip, j.data)
			if s.PrintLock != nil {
				s.PrintLock.Unlock()
			}
			result := "printed"
			if err != nil {
				if e, ok := err.(*spool.Error); ok {
					result = fmt.Sprintf("failed: %d %s", e.Code, e.Msg)
				} else {
					result = "failed: " + err.Error()
				}
			}
			s.log(j.conn, int64(len(j.data)), result)
			j.release()
		}
	}()
	limit := s.Slots
	if limit <= 0 {
		limit = SlotsMain
	}
	sem := make(chan struct{}, limit)
	if s.Pool != nil {
		sem = s.Pool.sem
	}
	// With a shared Pool a port holds a slot only for AcceptWait at a time, so idle ports cannot starve busy ones
	// (REQ-MULTI-008) while no connection is accepted before its slot is held (REQ-RECV-007).
	var deadline interface{ SetDeadline(time.Time) error }
	if s.Pool != nil {
		deadline, _ = ln.(interface{ SetDeadline(time.Time) error })
	}
	backoff := 5 * time.Millisecond
	for {
		// ponytail: with Allow the check comes before the slot (TC-ALLOW-041: refused at once while slots are busy),
		// so an allowed client may wait accepted for a slot.
		preacquired := false
		if s.Allow == nil {
			select {
			case sem <- struct{}{}:
				preacquired = true
			case <-stop:
				goto acceptedDone
			}
			if deadline != nil {
				deadline.SetDeadline(time.Now().Add(AcceptWait))
			}
		}
		c, err := ln.Accept()
		if err != nil {
			if preacquired {
				<-sem
			}
			select {
			case <-stop:
				goto acceptedDone
			default:
			}
			if ne, ok := err.(net.Error); ok && ne.Timeout() && deadline != nil {
				continue
			}
			time.Sleep(backoff)
			backoff *= 2
			if backoff > time.Second {
				backoff = time.Second
			}
			continue
		}
		backoff = 5 * time.Millisecond
		if s.Allow != nil {
			remote, _ := c.RemoteAddr().(*net.TCPAddr)
			var ip net.IP
			if remote != nil {
				ip = remote.IP
			}
			if !s.Allow(ip) {
				s.log(c.RemoteAddr().String(), 0, "rejected: not allowed")
				c.Close()
				continue
			}
		}
		if !preacquired {
			select {
			case sem <- struct{}{}:
			case <-stop:
				s.log(c.RemoteAddr().String(), 0, "dropped: shutdown")
				c.Close()
				goto acceptedDone
			}
		}
		s.mu.Lock()
		if s.stopping {
			s.mu.Unlock()
			s.log(c.RemoteAddr().String(), 0, "dropped: shutdown")
			c.Close()
			<-sem
			continue
		}
		s.conns[c] = true
		s.mu.Unlock()
		handlers.Add(1)
		go func(c net.Conn) {
			defer handlers.Done()
			release := func() { <-sem }
			deferred := true
			defer func() {
				if deferred {
					release()
				}
			}()
			data, result, received := s.receive(c)
			s.mu.Lock()
			delete(s.conns, c)
			stopping := s.stopping
			s.mu.Unlock()
			if stopping && result != "complete" {
				s.log(c.RemoteAddr().String(), received, "dropped: shutdown")
				return
			}
			if result != "complete" {
				s.log(c.RemoteAddr().String(), received, result)
				return
			}
			ip := c.RemoteAddr().(*net.TCPAddr).IP.String()
			deferred = false
			jobs <- printJob{conn: c.RemoteAddr().String(), ip: ip, data: data, release: release}
		}(c)
	}
acceptedDone:
	s.mu.Lock()
	s.stopping = true
	s.mu.Unlock()
	handlers.Wait()
	close(jobs)
	printer.Wait()
	close(done)
	return nil
}

func (s *Server) receive(c net.Conn) ([]byte, string, int64) {
	defer c.Close()
	max := s.MaxBytes
	if max <= 0 {
		max = MaxJobBytes
	}
	idle := s.Idle
	if idle <= 0 {
		idle = IdleTimeout
	}
	buf := make([]byte, 32*1024)
	data := make([]byte, 0)
	var received int64
	for {
		c.SetReadDeadline(time.Now().Add(idle))
		n, err := c.Read(buf)
		if n > 0 {
			received += int64(n)
			if int64(len(data))+int64(n) > max {
				return data, "dropped: too large", received
			}
			data = append(data, buf[:n]...)
		}
		if err != nil {
			if err == io.EOF {
				if len(data) == 0 {
					return data, "empty", received
				}
				return data, "complete", received
			}
			if ne, ok := err.(net.Error); ok && ne.Timeout() {
				return data, "dropped: idle", received
			}
			if strings.Contains(strings.ToLower(err.Error()), "reset") || strings.Contains(strings.ToLower(err.Error()), "abort") {
				return data, "dropped: reset", received
			}
			return data, "dropped: reset", received
		}
	}
}

func (s *Server) log(addr string, n int64, result string) {
	if s.Log == nil {
		return
	}
	now := time.Now
	if s.Now != nil {
		now = s.Now
	}
	s.logs.Lock()
	defer s.logs.Unlock()
	w := bufio.NewWriter(s.Log)
	fmt.Fprintf(w, "%s %s %dB \"%s\" %s\n", now().Format(time.RFC3339), addr, n, s.Printer, result)
	w.Flush()
}

func (s *Server) Shutdown() {
	s.shutdownOnce.Do(func() {
		s.mu.Lock()
		s.stopping = true
		if s.stop != nil {
			close(s.stop)
		}
		if s.ln != nil {
			s.ln.Close()
		}
		for c := range s.conns {
			c.Close()
		}
		done := s.serveDone
		s.mu.Unlock()
		if done != nil {
			<-done
		}
	})
}

// ShutdownAll shuts every server down at once and returns when all have stopped (REQ-MULTI-010): no port keeps
// accepting new jobs while another finishes its print.
func ShutdownAll(servers []*Server) {
	var wg sync.WaitGroup
	for _, s := range servers {
		wg.Add(1)
		go func(s *Server) { defer wg.Done(); s.Shutdown() }(s)
	}
	wg.Wait()
}
