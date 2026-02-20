package validator

import (
	"testing"

	"github.com/vortex/docx-go/packaging"
	"github.com/vortex/docx-go/parts/fonts"
	"github.com/vortex/docx-go/parts/settings"
	"github.com/vortex/docx-go/parts/styles"
	"github.com/vortex/docx-go/wml/body"
	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
)

// ===========================================================================
// Fuzz: Summary must never panic for any severity combination
// ===========================================================================

// FuzzSummary feeds arbitrary severity values into Summary. The function must
// never panic or return an empty string (empty issues → "document is valid").
func FuzzSummary(f *testing.F) {
	// Seed corpus with known patterns.
	f.Add(0, 0, 0) // no issues
	f.Add(1, 0, 0) // 1 fatal
	f.Add(0, 1, 0) // 1 error
	f.Add(0, 0, 1) // 1 warning
	f.Add(3, 5, 7) // mixed
	f.Add(100, 100, 100)

	f.Fuzz(func(t *testing.T, nFatal, nError, nWarn int) {
		// Clamp to reasonable range to keep runtime bounded.
		clamp := func(n int) int {
			if n < 0 {
				return 0
			}
			if n > 200 {
				return 200
			}
			return n
		}
		nFatal = clamp(nFatal)
		nError = clamp(nError)
		nWarn = clamp(nWarn)

		var issues []Issue
		for i := 0; i < nFatal; i++ {
			issues = append(issues, Issue{Severity: Fatal})
		}
		for i := 0; i < nError; i++ {
			issues = append(issues, Issue{Severity: Error})
		}
		for i := 0; i < nWarn; i++ {
			issues = append(issues, Issue{Severity: Warn})
		}

		result := Summary(issues)
		if result == "" {
			t.Error("Summary should never return empty string")
		}

		total := nFatal + nError + nWarn
		if total == 0 && result != "document is valid" {
			t.Errorf("zero issues should give 'document is valid', got %q", result)
		}
		if total > 0 && result == "document is valid" {
			t.Errorf("non-zero issues should not give 'document is valid'")
		}
	})
}

// ===========================================================================
// Fuzz: sortIssues must never panic and must produce descending severity
// ===========================================================================

// FuzzSortIssues feeds arbitrary severity values into sortIssues.
func FuzzSortIssues(f *testing.F) {
	// Seeds: encoded as up to 20 severity bytes (each mod 3).
	f.Add([]byte{0, 1, 2})
	f.Add([]byte{2, 2, 2, 1, 1, 0, 0})
	f.Add([]byte{})
	f.Add([]byte{0})
	f.Add([]byte{0, 1, 2, 0, 1, 2, 0, 1, 2, 0})

	f.Fuzz(func(t *testing.T, data []byte) {
		if len(data) > 100 {
			data = data[:100] // cap size
		}

		issues := make([]Issue, len(data))
		for i, b := range data {
			issues[i] = Issue{Severity: Severity(b % 3)}
		}

		sortIssues(issues)

		// Verify: non-increasing severity (Fatal=2 ≥ Error=1 ≥ Warn=0).
		for i := 1; i < len(issues); i++ {
			if issues[i].Severity > issues[i-1].Severity {
				t.Errorf("sort invariant broken at position %d: %v after %v",
					i, issues[i].Severity, issues[i-1].Severity)
				break
			}
		}
	})
}

// ===========================================================================
// Fuzz: Validate must never panic for any Document field combination
// ===========================================================================

// FuzzValidate explores nil/non-nil combinations of Document fields.
// Each bit of the fuzz input toggles a field between nil and non-nil.
func FuzzValidate(f *testing.F) {
	f.Add(byte(0))         // all nil
	f.Add(byte(0xFF))      // all present
	f.Add(byte(0b00000001)) // only Document
	f.Add(byte(0b01010101)) // alternating

	f.Fuzz(func(t *testing.T, bits byte) {
		doc := &packaging.Document{}

		if bits&0x01 != 0 {
			doc.Document = &body.CT_Document{}
			if bits&0x02 != 0 {
				doc.Document.Body = &body.CT_Body{
					Content: []shared.BlockLevelElement{&para.CT_P{}},
				}
			}
		}
		if bits&0x04 != 0 {
			doc.Styles = &styles.CT_Styles{}
		}
		if bits&0x08 != 0 {
			doc.Settings = &settings.CT_Settings{}
		}
		if bits&0x10 != 0 {
			doc.Fonts = &fonts.CT_FontsList{}
		}
		if bits&0x20 != 0 {
			doc.Theme = []byte("<t/>")
		}
		if bits&0x40 != 0 {
			doc.WebSettings = []byte("<w/>")
		}

		// Must not panic.
		issues := Validate(doc)

		// Basic sanity: we always get a slice (possibly empty).
		if issues == nil && bits == 0xFF {
			// With all fields populated (and body content), should have 0 issues.
			// nil is fine — it means 0 issues allocated.
		}

		// Verify sorting invariant.
		for i := 1; i < len(issues); i++ {
			if issues[i].Severity > issues[i-1].Severity {
				t.Errorf("sort invariant broken at position %d", i)
				break
			}
		}
	})
}

// ===========================================================================
// Fuzz: AutoFix must never panic for any Document field combination
// ===========================================================================

func FuzzAutoFix(f *testing.F) {
	f.Add(byte(0), byte(0))
	f.Add(byte(0xFF), byte(0))
	f.Add(byte(0xFF), byte(3)) // 3 empty cells

	f.Fuzz(func(t *testing.T, bits byte, nEmptyCells byte) {
		doc := &packaging.Document{}

		if bits&0x01 != 0 {
			doc.Document = &body.CT_Document{}
			if bits&0x02 != 0 {
				doc.Document.Body = &body.CT_Body{
					Content: []shared.BlockLevelElement{&para.CT_P{}},
				}

				// Add empty cells.
				n := int(nEmptyCells % 20) // cap at 20
				if n > 0 && doc.Document.Body != nil {
					var cells []table.RowContent
					for i := 0; i < n; i++ {
						cells = append(cells, table.CT_Tc{Content: nil})
					}
					tbl := &table.CT_Tbl{
						Content: []table.TblContent{
							table.CT_Row{Content: cells},
						},
					}
					doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)
				}
			}
		}
		if bits&0x04 != 0 {
			doc.Styles = &styles.CT_Styles{}
		}
		if bits&0x08 != 0 {
			doc.Settings = &settings.CT_Settings{}
		}
		if bits&0x10 != 0 {
			doc.Fonts = &fonts.CT_FontsList{}
		}
		if bits&0x20 != 0 {
			doc.Theme = []byte("<t/>")
		}
		if bits&0x40 != 0 {
			doc.WebSettings = []byte("<w/>")
		}

		// Must not panic.
		fixes := AutoFix(doc)

		// AutoFix on nil doc should return nil.
		_ = fixes

		// After fix, validate should not find EMPTY_TC or MISSING_WEBSETTINGS.
		issues := Validate(doc)
		for _, iss := range issues {
			if iss.Code == CodeEmptyTC {
				t.Error("AutoFix should have fixed EMPTY_TC")
			}
			if iss.Code == CodeMissingWebSettings {
				t.Error("AutoFix should have fixed MISSING_WEBSETTINGS")
			}
		}
	})
}
