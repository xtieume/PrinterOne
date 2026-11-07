package main

import (
	"bytes"
	"io/ioutil"
	"strconv"
	"testing"
)

// TC-UI-061: the committed assets_gen.go is exactly what the generator makes from static/ now.
func TestREQ_UI_017_GeneratedAssetsMatchStatic(t *testing.T) {
	got, err := generate("../../../../static")
	if err != nil {
		t.Fatal(err)
	}
	want, err := ioutil.ReadFile("../assets_gen.go")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, want) {
		t.Fatalf("assets_gen.go differs from static/: run `go generate ./internal/core/webui` (%d vs %d bytes)", len(want), len(got))
	}
	files, err := ioutil.ReadDir("../../../../static")
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range files {
		if !f.IsDir() && !bytes.Contains(got, []byte(strconv.Quote(f.Name())+":")) {
			t.Errorf("static/%s missing from the generated table", f.Name())
		}
	}
}
