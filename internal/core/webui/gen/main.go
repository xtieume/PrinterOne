// Command gen writes internal/core/webui/assets_gen.go from static/ (Go 1.10 has no embed).
// Usage (from internal/core/webui, via go generate): go run gen/main.go <static dir> <output file>.
package main

import (
	"fmt"
	"io/ioutil"
	"os"
)

// generate returns the gofmt'ed source of assets_gen.go for every regular file directly under dir, sorted by
// name: a header line "// Code generated … DO NOT EDIT.", package webui, and var assets = map[string]string{…}
// with strconv.Quote'd contents.
func generate(dir string) ([]byte, error) {
	return nil, nil
}

func main() {
	if len(os.Args) != 3 {
		fmt.Fprintln(os.Stderr, "usage: gen <static dir> <output file>")
		os.Exit(2)
	}
	src, err := generate(os.Args[1])
	if err == nil {
		err = ioutil.WriteFile(os.Args[2], src, 0644)
	}
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
