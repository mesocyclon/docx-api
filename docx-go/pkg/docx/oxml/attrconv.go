package oxml

import (
	"strconv"
	"strings"
)

// --- Attribute conversion helpers used by generated code ---

// parseIntAttr parses a string attribute value into an int.
func parseIntAttr(s string) int {
	v, err := strconv.Atoi(strings.TrimSpace(s))
	if err != nil {
		return 0
	}
	return v
}

// parseInt64Attr parses a string attribute value into an int64.
func parseInt64Attr(s string) int64 {
	v, err := strconv.ParseInt(strings.TrimSpace(s), 10, 64)
	if err != nil {
		return 0
	}
	return v
}

// parseBoolAttr parses an XML boolean attribute value.
// Accepts "true", "1", "on" as true; everything else is false.
func parseBoolAttr(s string) bool {
	s = strings.TrimSpace(strings.ToLower(s))
	return s == "true" || s == "1" || s == "on"
}

// parseOptionalIntAttr parses a string into *int.
func parseOptionalIntAttr(s string) *int {
	v := parseIntAttr(s)
	return &v
}

// formatIntAttr formats an int as a string attribute value.
func formatIntAttr(v int) string {
	return strconv.Itoa(v)
}

// formatInt64Attr formats an int64 as a string attribute value.
func formatInt64Attr(v int64) string {
	return strconv.FormatInt(v, 10)
}

// formatBoolAttr formats a bool as an XML attribute value.
func formatBoolAttr(v bool) string {
	if v {
		return "true"
	}
	return "false"
}

// mustParseEnum parses an XML attribute value using the provided fromXml function.
// If parsing fails, returns the zero value.
func mustParseEnum[T any](s string, fromXml func(string) (T, error)) T {
	v, err := fromXml(s)
	if err != nil {
		var zero T
		return zero
	}
	return v
}

// parseOptionalEnum parses an XML attribute value into a pointer to enum type.
func parseOptionalEnum[T any](s string, fromXml func(string) (T, error)) *T {
	v, err := fromXml(s)
	if err != nil {
		return nil
	}
	return &v
}
