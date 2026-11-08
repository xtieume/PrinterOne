package webui

import (
	"net/http/httptest"
	"testing"

	"github.com/xtieume/PrinterOne/internal/core/config"
)

// TC-UI-090
func TestREQ_UI_021_RejectsForeignHostWhenRemoteOff(t *testing.T) {
	s := New(Deps{Token: testToken, UI: config.UI{Port: 9180}})
	r := httptest.NewRequest("GET", "http://printerone.example/", nil)
	r.Host = "printerone.example"
	r.RemoteAddr = "127.0.0.1:1234"
	w := httptest.NewRecorder()
	s.ServeHTTP(w, r)
	if w.Code != 403 {
		t.Fatalf("foreign Host response = %d, want 403", w.Code)
	}
}

// TC-UI-091
func TestREQ_UI_022_AllResponsesDenyFraming(t *testing.T) {
	s := New(Deps{Token: testToken, UI: config.UI{Port: 9180}})
	for _, path := range []string{"/", "/api/status", "/base.css", "/missing"} {
		r := httptest.NewRequest("GET", path, nil)
		w := httptest.NewRecorder()
		s.ServeHTTP(w, r)
		if got := w.Header().Get("X-Frame-Options"); got != "DENY" {
			t.Errorf("GET %s X-Frame-Options = %q, want DENY", path, got)
		}
	}
}
