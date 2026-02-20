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

// ---------------------------------------------------------------------------
// Helpers to construct test documents
// ---------------------------------------------------------------------------

// minimalDoc builds the smallest valid Document that produces zero issues.
func minimalDoc() *packaging.Document {
	return &packaging.Document{
		Document: &body.CT_Document{
			Body: &body.CT_Body{
				Content: []shared.BlockLevelElement{
					&para.CT_P{}, // at least one paragraph
				},
			},
		},
		Styles:      &styles.CT_Styles{},
		Settings:    &settings.CT_Settings{},
		Fonts:       &fonts.CT_FontsList{},
		Theme:       []byte("<a:theme/>"),
		WebSettings: []byte("<w:webSettings/>"),
	}
}

// makeTable builds a simple table from a 2D slice of cell contents.
// nil inner element = empty cell (no paragraphs).
func makeTable(cells [][]shared.BlockLevelElement) *table.CT_Tbl {
	var rows []table.TblContent
	for _, rowCells := range cells {
		var rcs []table.RowContent
		for _, cellContent := range rowCells {
			var content []shared.BlockLevelElement
			if cellContent != nil {
				content = []shared.BlockLevelElement{cellContent}
			}
			rcs = append(rcs, table.CT_Tc{Content: content})
		}
		rows = append(rows, table.CT_Row{Content: rcs})
	}
	return &table.CT_Tbl{Content: rows}
}

// firstCell returns the first CT_Tc in the first row of tbl (for assertions).
func firstCell(tbl *table.CT_Tbl) table.CT_Tc {
	row := tbl.Content[0].(table.CT_Row)
	return row.Content[0].(table.CT_Tc)
}

// secondCell returns the second CT_Tc in the first row of tbl.
func secondCell(tbl *table.CT_Tbl) table.CT_Tc {
	row := tbl.Content[0].(table.CT_Row)
	return row.Content[1].(table.CT_Tc)
}

// ---------------------------------------------------------------------------
// Validate — nil document
// ---------------------------------------------------------------------------

func TestValidate_NilDocument(t *testing.T) {
	t.Parallel()
	issues := Validate(nil)

	if len(issues) != 1 {
		t.Fatalf("expected 1 issue, got %d", len(issues))
	}
	if issues[0].Code != CodeNilDocument {
		t.Errorf("expected code %s, got %s", CodeNilDocument, issues[0].Code)
	}
	if issues[0].Severity != Fatal {
		t.Errorf("expected Fatal severity, got %v", issues[0].Severity)
	}
}

// ---------------------------------------------------------------------------
// Validate — minimal valid document → zero issues
// ---------------------------------------------------------------------------

func TestValidate_MinimalValid(t *testing.T) {
	t.Parallel()
	doc := minimalDoc()
	issues := Validate(doc)

	if len(issues) != 0 {
		t.Errorf("expected 0 issues for minimal valid doc, got %d:", len(issues))
		for _, iss := range issues {
			t.Logf("  %s", iss)
		}
	}
}

// ---------------------------------------------------------------------------
// Validate — required parts missing
// ---------------------------------------------------------------------------

func TestValidate_MissingParts(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name     string
		mutate   func(d *packaging.Document)
		wantCode string
		wantSev  Severity
	}{
		{
			name:     "nil CT_Document",
			mutate:   func(d *packaging.Document) { d.Document = nil },
			wantCode: CodeMissingDocumentBody,
			wantSev:  Fatal,
		},
		{
			name: "nil Body",
			mutate: func(d *packaging.Document) {
				d.Document = &body.CT_Document{Body: nil}
			},
			wantCode: CodeMissingDocumentBody,
			wantSev:  Fatal,
		},
		{
			name:     "nil Styles",
			mutate:   func(d *packaging.Document) { d.Styles = nil },
			wantCode: CodeMissingStyles,
			wantSev:  Error,
		},
		{
			name:     "nil Settings",
			mutate:   func(d *packaging.Document) { d.Settings = nil },
			wantCode: CodeMissingSettings,
			wantSev:  Error,
		},
		{
			name:     "nil Fonts",
			mutate:   func(d *packaging.Document) { d.Fonts = nil },
			wantCode: CodeMissingFonts,
			wantSev:  Error,
		},
		{
			name:     "empty Theme",
			mutate:   func(d *packaging.Document) { d.Theme = nil },
			wantCode: CodeMissingTheme,
			wantSev:  Error,
		},
		{
			name:     "empty WebSettings",
			mutate:   func(d *packaging.Document) { d.WebSettings = nil },
			wantCode: CodeMissingWebSettings,
			wantSev:  Warn,
		},
	}

	for _, tc := range tests {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			t.Parallel()
			doc := minimalDoc()
			tc.mutate(doc)
			issues := Validate(doc)

			found := false
			for _, iss := range issues {
				if iss.Code == tc.wantCode {
					found = true
					if iss.Severity != tc.wantSev {
						t.Errorf("code %s: expected severity %v, got %v", tc.wantCode, tc.wantSev, iss.Severity)
					}
				}
			}
			if !found {
				t.Errorf("expected issue with code %s not found in %d issues", tc.wantCode, len(issues))
				for _, iss := range issues {
					t.Logf("  %s", iss)
				}
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Validate — EMPTY_TC detection
// ---------------------------------------------------------------------------

func TestValidate_EmptyTableCell(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	// Table with one row, one cell, NO paragraphs.
	tbl := makeTable([][]shared.BlockLevelElement{{nil}})
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)

	found := false
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			found = true
			if iss.Severity != Error {
				t.Errorf("EMPTY_TC should be Error, got %v", iss.Severity)
			}
			if iss.Path == "" {
				t.Error("EMPTY_TC should include a path")
			}
			break
		}
	}
	if !found {
		t.Error("expected EMPTY_TC issue for table cell with no paragraphs")
	}
}

// ---------------------------------------------------------------------------
// Validate — valid table cell (has paragraph) → no EMPTY_TC
// ---------------------------------------------------------------------------

func TestValidate_ValidTableCell(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	tbl := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}})
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			t.Errorf("unexpected EMPTY_TC for cell that contains a paragraph: %s", iss)
		}
	}
}

// ---------------------------------------------------------------------------
// Validate — empty body warning
// ---------------------------------------------------------------------------

func TestValidate_EmptyBody(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.Document.Body.Content = nil

	issues := Validate(doc)

	found := false
	for _, iss := range issues {
		if iss.Code == CodeEmptyBody {
			found = true
			if iss.Severity != Warn {
				t.Errorf("EMPTY_BODY should be Warn, got %v", iss.Severity)
			}
		}
	}
	if !found {
		t.Error("expected EMPTY_BODY warning for body with no content")
	}
}

// ---------------------------------------------------------------------------
// Validate — nested table with empty cell
// ---------------------------------------------------------------------------

func TestValidate_NestedTableEmptyCell(t *testing.T) {
	t.Parallel()

	innerTable := makeTable([][]shared.BlockLevelElement{{nil}}) // empty inner cell

	outerTable := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					table.CT_Tc{Content: []shared.BlockLevelElement{
						&para.CT_P{}, // outer cell has a paragraph (valid)
						innerTable,   // but also a nested table with empty cell
					}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, outerTable)

	issues := Validate(doc)

	count := 0
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			count++
		}
	}
	if count != 1 {
		t.Errorf("expected 1 EMPTY_TC (inner table), got %d", count)
	}
}

// ---------------------------------------------------------------------------
// Validate — severity ordering (Fatal first)
// ---------------------------------------------------------------------------

func TestValidate_SeverityOrdering(t *testing.T) {
	t.Parallel()

	doc := &packaging.Document{
		Document:    nil, // Fatal: MISSING_DOCUMENT_BODY
		Styles:      nil, // Error: MISSING_STYLES
		Settings:    nil, // Error: MISSING_SETTINGS
		Fonts:       nil, // Error: MISSING_FONTS
		Theme:       nil, // Error: MISSING_THEME
		WebSettings: nil, // Warn:  MISSING_WEBSETTINGS
	}

	issues := Validate(doc)
	if len(issues) == 0 {
		t.Fatal("expected issues for fully empty document")
	}

	// Check that Fatal comes before Error, which comes before Warn.
	prevSev := Fatal
	for _, iss := range issues {
		if iss.Severity > prevSev {
			t.Errorf("severity ordering violated: %v after %v", iss.Severity, prevSev)
		}
		prevSev = iss.Severity
	}
}

// ---------------------------------------------------------------------------
// AutoFix — webSettings
// ---------------------------------------------------------------------------

func TestAutoFix_WebSettings(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.WebSettings = nil

	fixes := AutoFix(doc)

	if len(doc.WebSettings) == 0 {
		t.Error("AutoFix should have populated WebSettings")
	}

	found := false
	for _, f := range fixes {
		if f == "added minimal webSettings.xml part" {
			found = true
		}
	}
	if !found {
		t.Error("expected fix description for webSettings")
	}
}

// ---------------------------------------------------------------------------
// AutoFix — empty table cells
// ---------------------------------------------------------------------------

func TestAutoFix_EmptyTableCell(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					table.CT_Tc{Content: nil},                                      // empty
					table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}}, // already valid
				},
			},
		},
	}
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	fixes := AutoFix(doc)

	// The first cell should now have a paragraph.
	cell0 := firstCell(tbl)
	if !cellHasParagraph(cell0) {
		t.Error("AutoFix should have added a paragraph to the empty cell")
	}

	// The second cell should still have exactly one paragraph.
	cell1 := secondCell(tbl)
	pCount := 0
	for _, el := range cell1.Content {
		if _, ok := el.(*para.CT_P); ok {
			pCount++
		}
	}
	if pCount != 1 {
		t.Errorf("AutoFix should not have modified valid cell, paragraph count: %d", pCount)
	}

	if len(fixes) == 0 {
		t.Error("expected at least one fix description")
	}
}

// ---------------------------------------------------------------------------
// AutoFix — idempotent (running twice produces no new fixes)
// ---------------------------------------------------------------------------

func TestAutoFix_Idempotent(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.WebSettings = nil
	tbl := makeTable([][]shared.BlockLevelElement{{nil}})
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	// First pass — should produce fixes.
	fixes1 := AutoFix(doc)
	if len(fixes1) == 0 {
		t.Fatal("first AutoFix should produce fixes")
	}

	// Second pass — should produce no fixes.
	fixes2 := AutoFix(doc)
	if len(fixes2) != 0 {
		t.Errorf("second AutoFix should be idempotent, got %d fixes: %v", len(fixes2), fixes2)
	}
}

// ---------------------------------------------------------------------------
// AutoFix — nil document is safe
// ---------------------------------------------------------------------------

func TestAutoFix_NilDocument(t *testing.T) {
	t.Parallel()

	fixes := AutoFix(nil)
	if fixes != nil {
		t.Errorf("AutoFix(nil) should return nil, got %v", fixes)
	}
}

// ---------------------------------------------------------------------------
// Validate after AutoFix — issues should be reduced
// ---------------------------------------------------------------------------

func TestValidateThenAutoFix_ReducesIssues(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.WebSettings = nil
	tbl := makeTable([][]shared.BlockLevelElement{{nil}})
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	before := Validate(doc)
	beforeCount := len(before)

	AutoFix(doc)
	after := Validate(doc)

	if len(after) >= beforeCount {
		t.Errorf("AutoFix should reduce issue count: before=%d, after=%d", beforeCount, len(after))
	}

	// Specifically, no more EMPTY_TC or MISSING_WEBSETTINGS.
	for _, iss := range after {
		if iss.Code == CodeEmptyTC || iss.Code == CodeMissingWebSettings {
			t.Errorf("AutoFix should have resolved %s", iss.Code)
		}
	}
}

// ---------------------------------------------------------------------------
// Helpers — HasFatal, HasErrors, Filter, Summary
// ---------------------------------------------------------------------------

func TestHelpers(t *testing.T) {
	t.Parallel()

	issues := []Issue{
		{Severity: Fatal, Code: "A"},
		{Severity: Error, Code: "B"},
		{Severity: Warn, Code: "C"},
		{Severity: Warn, Code: "D"},
	}

	if !HasFatal(issues) {
		t.Error("HasFatal should return true")
	}
	if !HasErrors(issues) {
		t.Error("HasErrors should return true")
	}

	warns := Filter(issues, Warn)
	if len(warns) != 2 {
		t.Errorf("Filter(Warn) expected 2, got %d", len(warns))
	}

	sum := Summary(issues)
	if sum == "" || sum == "document is valid" {
		t.Errorf("Summary should describe issues, got %q", sum)
	}

	// Empty issues
	if HasFatal(nil) {
		t.Error("HasFatal(nil) should be false")
	}
	if HasErrors(nil) {
		t.Error("HasErrors(nil) should be false")
	}
	if Summary(nil) != "document is valid" {
		t.Errorf("Summary(nil) should be 'document is valid', got %q", Summary(nil))
	}
}

// ---------------------------------------------------------------------------
// Severity.String()
// ---------------------------------------------------------------------------

func TestSeverityString(t *testing.T) {
	t.Parallel()

	tests := []struct {
		sev  Severity
		want string
	}{
		{Warn, "WARN"},
		{Error, "ERROR"},
		{Fatal, "FATAL"},
		{Severity(99), "UNKNOWN"},
	}
	for _, tc := range tests {
		if got := tc.sev.String(); got != tc.want {
			t.Errorf("Severity(%d).String() = %q, want %q", tc.sev, got, tc.want)
		}
	}
}

// ---------------------------------------------------------------------------
// Issue.String()
// ---------------------------------------------------------------------------

func TestIssueString(t *testing.T) {
	t.Parallel()

	iss := Issue{
		Severity: Error,
		Part:     "word/document.xml",
		Path:     "body/tbl[0]/tr[0]/tc[0]",
		Code:     CodeEmptyTC,
		Message:  "table cell has no paragraph",
	}
	s := iss.String()
	if s == "" {
		t.Error("Issue.String() should not be empty")
	}
	// Should contain the code and severity.
	if !containsSub(s, "ERROR") || !containsSub(s, CodeEmptyTC) {
		t.Errorf("Issue.String() missing expected parts: %s", s)
	}
}

func containsSub(s, sub string) bool {
	for i := 0; i <= len(s)-len(sub); i++ {
		if s[i:i+len(sub)] == sub {
			return true
		}
	}
	return false
}
