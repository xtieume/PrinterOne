package webui

// REQ-UI-015 / TC-UI-084 / TC-UI-088: claims about the shipped text of static/ (spec §2.3), not a source grep.

import (
	"io/ioutil"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

const staticDir = "../../../static"

func staticFile(t *testing.T, name string) string {
	b, err := ioutil.ReadFile(filepath.Join(staticDir, name))
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

var cssComment = regexp.MustCompile(`(?s)/\*.*?\*/`)

// TC-UI-056
func TestREQ_UI_015_BaseCSSOnlyIE6(t *testing.T) {
	css := staticFile(t, "base.css")
	flat := strings.ToLower(cssComment.ReplaceAllString(css, ""))
	flat = strings.Join(strings.Fields(flat), "")
	if flat == "" {
		t.Fatal("base.css is empty")
	}
	for _, bad := range []string{"flex", "grid", "@supports", "var(", "calc(", "rgba(", "hsla(", "opacity",
		"border-radius", "box-shadow", "text-shadow", "transform", "transition", "position:fixed", "min-width",
		"max-width", "background-size", ">", "[", ":not(", ":nth-", ":first-child", "::", ":before", ":after",
		"rem;", "vw", "vh;"} {
		if strings.Contains(flat, bad) {
			t.Errorf("base.css uses %q, which IE6 lacks", bad)
		}
	}
}

// TC-UI-057
func TestREQ_UI_015_ModernCSSInsideSupportsGrid(t *testing.T) {
	css := strings.TrimSpace(cssComment.ReplaceAllString(staticFile(t, "modern.css"), ""))
	const head = "@supports (display: grid) {"
	if !strings.HasPrefix(css, head) {
		t.Fatalf("modern.css does not start with %q", head)
	}
	depth := 0
	for i, c := range css {
		switch c {
		case '{':
			depth++
		case '}':
			depth--
			if depth == 0 && i != len(css)-1 {
				t.Fatalf("modern.css has text after the @supports block at byte %d", i)
			}
		}
	}
	if depth != 0 || strings.TrimSpace(css[len(head):len(css)-1]) == "" {
		t.Fatal("modern.css block is unbalanced or empty")
	}
}

func templates(t *testing.T) map[string]string {
	names, err := filepath.Glob(filepath.Join(staticDir, "*.html"))
	if err != nil || len(names) == 0 {
		t.Fatalf("no templates in static/ (%v)", err)
	}
	m := map[string]string{}
	for _, n := range names {
		m[filepath.Base(n)] = staticFile(t, filepath.Base(n))
	}
	return m
}

var scriptTag = regexp.MustCompile(`(?i)<script\b[^>]*>`)

// TC-UI-058
func TestREQ_UI_015_ScriptsAreModulesAndAssetsLinked(t *testing.T) {
	all := ""
	for name, body := range templates(t) {
		all += body
		for _, tag := range scriptTag.FindAllString(body, -1) {
			if !strings.Contains(tag, `type="module"`) {
				t.Errorf("%s: %s is not type=\"module\"", name, tag)
			}
		}
	}
	for _, ref := range []string{`href="/base.css"`, `href="/modern.css"`, `src="/app.js"`} {
		if !strings.Contains(all, ref) {
			t.Errorf("templates do not reference %s", ref)
		}
	}
}

var inlineJS = regexp.MustCompile(`(?i)\son[a-z]+\s*=|javascript:`)

// TC-UI-088
func TestREQ_UI_015_NoInlineHandlers(t *testing.T) {
	for name, body := range templates(t) {
		if m := inlineJS.FindString(body); m != "" {
			t.Errorf("%s: inline script %q runs in IE", name, m)
		}
	}
}

// TC-UI-084
func TestREQ_UI_014_AppJSInsertsTextOnly(t *testing.T) {
	js := staticFile(t, "app.js")
	if strings.TrimSpace(js) == "" {
		t.Fatal("app.js is empty")
	}
	for _, bad := range []string{"innerHTML", "outerHTML", "insertAdjacentHTML", "document.write"} {
		if strings.Contains(js, bad) {
			t.Errorf("app.js uses %s", bad)
		}
	}
}
