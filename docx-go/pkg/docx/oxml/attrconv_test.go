package oxml

import (
	"fmt"
	"testing"
)

func TestParseIntAttr(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input string
		want  int
	}{
		{"42", 42},
		{" 100 ", 100},
		{"0", 0},
		{"-5", -5},
		{"abc", 0},
		{"", 0},
	}
	for _, tc := range tests {
		t.Run(tc.input, func(t *testing.T) {
			t.Parallel()
			if got := parseIntAttr(tc.input); got != tc.want {
				t.Errorf("parseIntAttr(%q) = %d, want %d", tc.input, got, tc.want)
			}
		})
	}
}

func TestParseBoolAttr(t *testing.T) {
	t.Parallel()
	tests := []struct {
		input string
		want  bool
	}{
		{"true", true},
		{"True", true},
		{"TRUE", true},
		{"1", true},
		{"on", true},
		{"false", false},
		{"0", false},
		{"", false},
		{"off", false},
	}
	for _, tc := range tests {
		t.Run(tc.input, func(t *testing.T) {
			t.Parallel()
			if got := parseBoolAttr(tc.input); got != tc.want {
				t.Errorf("parseBoolAttr(%q) = %v, want %v", tc.input, got, tc.want)
			}
		})
	}
}

func TestFormatIntAttr(t *testing.T) {
	t.Parallel()
	if got := formatIntAttr(42); got != "42" {
		t.Errorf("formatIntAttr(42) = %q, want %q", got, "42")
	}
}

func TestFormatBoolAttr(t *testing.T) {
	t.Parallel()
	if got := formatBoolAttr(true); got != "true" {
		t.Errorf("formatBoolAttr(true) = %q", got)
	}
	if got := formatBoolAttr(false); got != "false" {
		t.Errorf("formatBoolAttr(false) = %q", got)
	}
}

func TestParseInt64Attr(t *testing.T) {
	t.Parallel()
	if got := parseInt64Attr("914400"); got != 914400 {
		t.Errorf("parseInt64Attr(\"914400\") = %d, want 914400", got)
	}
}

func TestFormatInt64Attr(t *testing.T) {
	t.Parallel()
	if got := formatInt64Attr(914400); got != "914400" {
		t.Errorf("formatInt64Attr(914400) = %q", got)
	}
}

func TestMustParseEnum(t *testing.T) {
	t.Parallel()
	fromXml := func(s string) (int, error) {
		m := map[string]int{"a": 1, "b": 2}
		if v, ok := m[s]; ok {
			return v, nil
		}
		return 0, fmt.Errorf("unknown: %s", s)
	}

	if got := mustParseEnum("a", fromXml); got != 1 {
		t.Errorf("mustParseEnum(\"a\") = %d, want 1", got)
	}
	if got := mustParseEnum("unknown", fromXml); got != 0 {
		t.Errorf("mustParseEnum(\"unknown\") = %d, want 0", got)
	}
}

func TestParseOptionalEnum(t *testing.T) {
	t.Parallel()
	fromXml := func(s string) (int, error) {
		if s == "a" {
			return 1, nil
		}
		return 0, fmt.Errorf("unknown: %s", s)
	}

	got := parseOptionalEnum("a", fromXml)
	if got == nil || *got != 1 {
		t.Errorf("parseOptionalEnum(\"a\") = %v, want *1", got)
	}

	got = parseOptionalEnum("unknown", fromXml)
	if got != nil {
		t.Errorf("parseOptionalEnum(\"unknown\") = %v, want nil", got)
	}
}

func TestParseOptionalIntAttr(t *testing.T) {
	t.Parallel()
	got := parseOptionalIntAttr("42")
	if got == nil || *got != 42 {
		t.Errorf("parseOptionalIntAttr(\"42\") = %v, want *42", got)
	}
}
