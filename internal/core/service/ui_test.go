package service

import (
	"io/ioutil"
	"net/http"
	"net/url"
	"regexp"
	"strconv"
	"strings"
	"testing"
	"time"
)

// F-005 review finding, 2026-11-07 (TC-UI-103): a UI save that replaces every listener must not stop the service.

// TC-UI-103
func TestREQ_UI_009_ServiceKeepsRunningWhenUIReplacesListeners(t *testing.T) {
	r := newRunRig(t, "main")
	defer r.cleanup()
	uiPort, newPort := freePort(t), freePort(t)
	body := `{"version":2,"printers":[{"name":"default","printer":"` + testPrinter + `","port":` + strconv.Itoa(r.port) +
		`}],"ui":{"port":` + strconv.Itoa(uiPort) + `}}`
	if err := ioutil.WriteFile(r.opts.ConfigPath, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	r.start()
	r.waitReady()
	base := "http://127.0.0.1:" + strconv.Itoa(uiPort)
	resp, err := http.Get(base + "/printers")
	if err != nil {
		t.Fatal(err)
	}
	page, _ := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	m := regexp.MustCompile(`name="token"\s+value="([^"]+)"`).FindSubmatch(page)
	if m == nil {
		t.Fatalf("no token on /printers: %.300s", page)
	}
	client := &http.Client{CheckRedirect: func(*http.Request, []*http.Request) error { return http.ErrUseLastResponse }}
	form := url.Values{"token": {string(m[1])}, "action": {"edit"}, "index": {"0"}, "name": {"default"},
		"printer": {testPrinter}, "port": {strconv.Itoa(newPort)}}
	resp, err = client.PostForm(base+"/printers", form)
	if err != nil {
		t.Fatal(err)
	}
	resp.Body.Close()
	if resp.StatusCode != http.StatusSeeOther {
		t.Fatalf("save answered %d", resp.StatusCode)
	}
	time.Sleep(300 * time.Millisecond)
	if r.session.saw(Stopped) {
		t.Fatalf("service reported Stopped after the UI replaced its only listener; reports %v", r.session.list())
	}
	r.port = newPort
	r.send("moved")
	if !r.spool.waitPrinted("moved", 5*time.Second) || !strings.Contains(r.log(), strconv.Itoa(newPort)) {
		t.Fatalf("job on the new port %d not printed; log %q", newPort, r.log())
	}
}
