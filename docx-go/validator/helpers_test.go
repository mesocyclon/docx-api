package validator

import (
	"strings"
	"testing"
)

// ===========================================================================
// sortIssues — exhaustive edge cases
// ===========================================================================

func TestSortIssues_Empty(t *testing.T) {
	t.Parallel()
	var issues []Issue
	sortIssues(issues) // must not panic
	if len(issues) != 0 {
		t.Error("sorting empty slice should remain empty")
	}
}

func TestSortIssues_SingleElement(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Warn, Code: "A"}}
	sortIssues(issues)
	if issues[0].Code != "A" {
		t.Error("single element should be unchanged")
	}
}

func TestSortIssues_AlreadySorted(t *testing.T) {
	t.Parallel()
	issues := []Issue{
		{Severity: Fatal, Code: "F"},
		{Severity: Error, Code: "E"},
		{Severity: Warn, Code: "W"},
	}
	sortIssues(issues)
	assertOrder(t, issues, []Severity{Fatal, Error, Warn})
}

func TestSortIssues_ReverseSorted(t *testing.T) {
	t.Parallel()
	issues := []Issue{
		{Severity: Warn, Code: "W"},
		{Severity: Error, Code: "E"},
		{Severity: Fatal, Code: "F"},
	}
	sortIssues(issues)
	assertOrder(t, issues, []Severity{Fatal, Error, Warn})
}

func TestSortIssues_AllSameSeverity(t *testing.T) {
	t.Parallel()
	issues := []Issue{
		{Severity: Error, Code: "A"},
		{Severity: Error, Code: "B"},
		{Severity: Error, Code: "C"},
	}
	sortIssues(issues)
	// All same severity → order should be preserved (stable sort).
	if issues[0].Code != "A" || issues[1].Code != "B" || issues[2].Code != "C" {
		t.Error("stable sort should preserve original order for same-severity items")
	}
}

// Stability test: within each severity group, original order must be preserved.
func TestSortIssues_Stability(t *testing.T) {
	t.Parallel()
	issues := []Issue{
		{Severity: Warn, Code: "W1"},
		{Severity: Fatal, Code: "F1"},
		{Severity: Error, Code: "E1"},
		{Severity: Warn, Code: "W2"},
		{Severity: Fatal, Code: "F2"},
		{Severity: Error, Code: "E2"},
	}
	sortIssues(issues)

	// Fatal group: F1, F2 (in original insertion order)
	if issues[0].Code != "F1" || issues[1].Code != "F2" {
		t.Errorf("Fatal group not stable: %s, %s", issues[0].Code, issues[1].Code)
	}
	// Error group: E1, E2
	if issues[2].Code != "E1" || issues[3].Code != "E2" {
		t.Errorf("Error group not stable: %s, %s", issues[2].Code, issues[3].Code)
	}
	// Warn group: W1, W2
	if issues[4].Code != "W1" || issues[5].Code != "W2" {
		t.Errorf("Warn group not stable: %s, %s", issues[4].Code, issues[5].Code)
	}
}

func TestSortIssues_TwoElements(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name string
		in   []Issue
		want []Severity
	}{
		{"Fatal-Warn", []Issue{{Severity: Fatal}, {Severity: Warn}}, []Severity{Fatal, Warn}},
		{"Warn-Fatal", []Issue{{Severity: Warn}, {Severity: Fatal}}, []Severity{Fatal, Warn}},
		{"Error-Fatal", []Issue{{Severity: Error}, {Severity: Fatal}}, []Severity{Fatal, Error}},
		{"Warn-Error", []Issue{{Severity: Warn}, {Severity: Error}}, []Severity{Error, Warn}},
		{"Same-Warn", []Issue{{Severity: Warn}, {Severity: Warn}}, []Severity{Warn, Warn}},
	}

	for _, tc := range tests {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			sortIssues(tc.in)
			assertOrder(t, tc.in, tc.want)
		})
	}
}

// Large-ish mixed input to exercise the insertion sort under load.
func TestSortIssues_LargeMixed(t *testing.T) {
	t.Parallel()

	sevs := []Severity{Warn, Error, Fatal, Warn, Fatal, Error, Warn, Error, Fatal, Warn}
	issues := make([]Issue, len(sevs))
	for i, s := range sevs {
		issues[i] = Issue{Severity: s}
	}
	sortIssues(issues)

	prev := Fatal
	for i, iss := range issues {
		if iss.Severity > prev {
			t.Errorf("position %d: severity %v after %v — not sorted", i, iss.Severity, prev)
		}
		prev = iss.Severity
	}
}

// ===========================================================================
// Summary — all combinations
// ===========================================================================

func TestSummary_OnlyFatal(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Fatal}}
	s := Summary(issues)
	if !strings.Contains(s, "1 fatal") {
		t.Errorf("expected '1 fatal' in %q", s)
	}
	if strings.Contains(s, "error") || strings.Contains(s, "warning") {
		t.Errorf("should not mention errors/warnings when only fatal: %q", s)
	}
}

func TestSummary_OnlyErrors(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Error}, {Severity: Error}}
	s := Summary(issues)
	if !strings.Contains(s, "2 error") {
		t.Errorf("expected '2 error' in %q", s)
	}
}

func TestSummary_OnlyWarnings(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Warn}, {Severity: Warn}, {Severity: Warn}}
	s := Summary(issues)
	if !strings.Contains(s, "3 warning") {
		t.Errorf("expected '3 warning' in %q", s)
	}
}

func TestSummary_Mixed(t *testing.T) {
	t.Parallel()
	issues := []Issue{
		{Severity: Fatal},
		{Severity: Error},
		{Severity: Error},
		{Severity: Warn},
	}
	s := Summary(issues)
	if !strings.Contains(s, "1 fatal") {
		t.Errorf("missing fatal in %q", s)
	}
	if !strings.Contains(s, "2 error") {
		t.Errorf("missing errors in %q", s)
	}
	if !strings.Contains(s, "1 warning") {
		t.Errorf("missing warning in %q", s)
	}
}

func TestSummary_EmptySlice(t *testing.T) {
	t.Parallel()
	s := Summary([]Issue{})
	if s != "document is valid" {
		t.Errorf("empty slice should be 'document is valid', got %q", s)
	}
}

// ===========================================================================
// Filter — exhaustive
// ===========================================================================

func TestFilter_NoMatch(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Fatal}, {Severity: Error}}
	result := Filter(issues, Warn)
	if len(result) != 0 {
		t.Errorf("no warnings present, expected 0, got %d", len(result))
	}
}

func TestFilter_AllMatch(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Error}, {Severity: Error}, {Severity: Error}}
	result := Filter(issues, Error)
	if len(result) != 3 {
		t.Errorf("all are errors, expected 3, got %d", len(result))
	}
}

func TestFilter_EmptyInput(t *testing.T) {
	t.Parallel()
	result := Filter(nil, Fatal)
	if result != nil {
		t.Errorf("Filter(nil) should return nil, got %v", result)
	}
}

func TestFilter_PreservesOrder(t *testing.T) {
	t.Parallel()
	issues := []Issue{
		{Severity: Warn, Code: "W1"},
		{Severity: Error, Code: "E1"},
		{Severity: Warn, Code: "W2"},
		{Severity: Fatal, Code: "F1"},
		{Severity: Warn, Code: "W3"},
	}
	result := Filter(issues, Warn)
	if len(result) != 3 {
		t.Fatalf("expected 3 warnings, got %d", len(result))
	}
	if result[0].Code != "W1" || result[1].Code != "W2" || result[2].Code != "W3" {
		t.Error("Filter should preserve original order")
	}
}

// ===========================================================================
// HasFatal / HasErrors — detailed
// ===========================================================================

func TestHasFatal_OnlyErrors(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Error}, {Severity: Warn}}
	if HasFatal(issues) {
		t.Error("no fatal issues present, should return false")
	}
}

func TestHasFatal_FatalPresent(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Warn}, {Severity: Fatal}}
	if !HasFatal(issues) {
		t.Error("fatal issue present, should return true")
	}
}

func TestHasErrors_OnlyWarnings(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Warn}, {Severity: Warn}}
	if HasErrors(issues) {
		t.Error("only warnings present, HasErrors should return false")
	}
}

func TestHasErrors_ErrorPresent(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Warn}, {Severity: Error}}
	if !HasErrors(issues) {
		t.Error("error present, HasErrors should return true")
	}
}

func TestHasErrors_FatalCountsAsError(t *testing.T) {
	t.Parallel()
	issues := []Issue{{Severity: Fatal}}
	if !HasErrors(issues) {
		t.Error("Fatal >= Error, so HasErrors should return true")
	}
}

// ===========================================================================
// Severity.String() — boundary values
// ===========================================================================

func TestSeverityString_NegativeValue(t *testing.T) {
	t.Parallel()
	s := Severity(-1)
	if s.String() != "UNKNOWN" {
		t.Errorf("negative severity should be UNKNOWN, got %q", s.String())
	}
}

func TestSeverityString_HighValue(t *testing.T) {
	t.Parallel()
	s := Severity(1000)
	if s.String() != "UNKNOWN" {
		t.Errorf("out-of-range severity should be UNKNOWN, got %q", s.String())
	}
}

// ===========================================================================
// helpers
// ===========================================================================

func assertOrder(t *testing.T, issues []Issue, want []Severity) {
	t.Helper()
	if len(issues) != len(want) {
		t.Fatalf("length mismatch: got %d, want %d", len(issues), len(want))
	}
	for i, iss := range issues {
		if iss.Severity != want[i] {
			t.Errorf("position %d: got %v, want %v", i, iss.Severity, want[i])
		}
	}
}
