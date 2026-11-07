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
)

type PrintFunc func(docName string, data []byte) error

type Server struct {
	Addr     string
	Printer  string
	Print    PrintFunc
	MaxBytes int64
	Idle     time.Duration
	Slots    int
	Log      io.Writer
	Now      func() time.Time

	mu           sync.Mutex
	ln           net.Listener
	conns        map[net.Conn]bool
	stopping     bool
	serveDone    chan struct{}
	stop         chan struct{}
	shutdownOnce sync.Once
	logs         sync.Mutex
}

type PortBusyError struct{ Port int }

func (e *PortBusyError) Error() string { return fmt.Sprintf("port %d is already in use", e.Port) }

func (s *Server) Listen() error {
	ln, err := net.Listen("tcp", s.Addr)
	if err != nil {
		if a, ok := net.ResolveTCPAddr("tcp", s.Addr); ok == nil && a.Port != 0 {
			if op, ok := err.(*net.OpError); ok && strings.Contains(strings.ToLower(op.Err.Error()), "address already in use") {
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

type printJob struct {
	conn string
	ip   string
	data []byte
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
			err := s.Print("PrinterOne job from "+j.ip, j.data)
			result := "printed"
			if err != nil {
				if e, ok := err.(*spool.Error); ok {
					result = fmt.Sprintf("failed: %d %s", e.Code, e.Msg)
				} else {
					result = "failed: " + err.Error()
				}
			}
			s.log(j.conn, int64(len(j.data)), result)
		}
	}()
	limit := s.Slots
	if limit <= 0 {
		limit = SlotsMain
	}
	sem := make(chan struct{}, limit)
	for {
		c, err := ln.Accept()
		if err != nil {
			break
		}
		s.mu.Lock()
		if s.stopping {
			s.mu.Unlock()
			c.Close()
			continue
		}
		s.conns[c] = true
		s.mu.Unlock()
		handlers.Add(1)
		go func(c net.Conn) {
			defer handlers.Done()
			select {
			case sem <- struct{}{}:
			case <-stop:
				return
			}
			defer func() { <-sem }()
			data, result := s.receive(c)
			s.mu.Lock()
			delete(s.conns, c)
			stopping := s.stopping
			s.mu.Unlock()
			if stopping {
				s.log(c.RemoteAddr().String(), int64(len(data)), "dropped: shutdown")
				return
			}
			if result != "complete" {
				s.log(c.RemoteAddr().String(), int64(len(data)), result)
				return
			}
			ip := c.RemoteAddr().(*net.TCPAddr).IP.String()
			jobs <- printJob{conn: c.RemoteAddr().String(), ip: ip, data: data}
		}(c)
	}
	s.mu.Lock()
	s.stopping = true
	s.mu.Unlock()
	handlers.Wait()
	close(jobs)
	printer.Wait()
	close(done)
	return nil
}

func (s *Server) receive(c net.Conn) ([]byte, string) {
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
	for {
		c.SetReadDeadline(time.Now().Add(idle))
		n, err := c.Read(buf)
		if n > 0 {
			if int64(len(data))+int64(n) > max {
				data = append(data, buf[:n]...)
				return data, "dropped: too large"
			}
			data = append(data, buf[:n]...)
		}
		if err != nil {
			if err == io.EOF {
				if len(data) == 0 {
					return data, "empty"
				}
				return data, "complete"
			}
			if ne, ok := err.(net.Error); ok && ne.Timeout() {
				return data, "dropped: idle"
			}
			if strings.Contains(strings.ToLower(err.Error()), "reset") || strings.Contains(strings.ToLower(err.Error()), "abort") {
				return data, "dropped: reset"
			}
			return data, "dropped: reset"
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
