package validator

import (
	"encoding/xml"
	"fmt"
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
// Integration: Realistic document — mixed content
// ===========================================================================

// Build a document resembling a real Word file: paragraphs, multiple tables
// with varying validity, nested tables, RawXML.
func TestIntegration_RealisticDocument(t *testing.T) {
	t.Parallel()

	// Table 1: 2×3, all cells valid.
	validTable := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
			}},
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
			}},
		},
	}

	// Table 2: 2×2, one cell has nested table with empty cell.
	innerBroken := makeTable([][]shared.BlockLevelElement{{nil}})
	tableWithNested := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}, innerBroken}},
			}},
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: nil}, // empty cell
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
			}},
		},
	}

	doc := &packaging.Document{
		Document: &body.CT_Document{
			Body: &body.CT_Body{
				Content: []shared.BlockLevelElement{
					&para.CT_P{},    // title paragraph
					validTable,       // tbl[0] — all valid
					&para.CT_P{},    // spacer
					tableWithNested,  // tbl[1] — 2 issues: nested empty + direct empty
					&para.CT_P{},    // footer paragraph
				},
			},
		},
		Styles:      &styles.CT_Styles{},
		Settings:    &settings.CT_Settings{},
		Fonts:       &fonts.CT_FontsList{},
		Theme:       []byte("<a:theme/>"),
		WebSettings: []byte("<w:webSettings/>"),
	}

	issues := Validate(doc)

	emptyCount := countCode(issues, CodeEmptyTC)
	if emptyCount != 2 {
		t.Errorf("expected 2 EMPTY_TC (1 nested + 1 direct), got %d", emptyCount)
		for _, iss := range issues {
			t.Logf("  %s", iss)
		}
	}

	// No structural issues besides EMPTY_TC.
	for _, iss := range issues {
		if iss.Code != CodeEmptyTC {
			t.Errorf("unexpected issue: %s", iss)
		}
	}

	// AutoFix should resolve everything.
	fixes := AutoFix(doc)
	if len(fixes) == 0 {
		t.Fatal("expected fixes")
	}

	afterIssues := Validate(doc)
	if len(afterIssues) != 0 {
		t.Errorf("after AutoFix, expected 0 issues, got %d:", len(afterIssues))
		for _, iss := range afterIssues {
			t.Logf("  %s", iss)
		}
	}
}

// ===========================================================================
// Integration: Deep nesting (3+ levels)
// ===========================================================================

func TestIntegration_DeeplyNestedTables(t *testing.T) {
	t.Parallel()

	// Level 3: innermost table with empty cell.
	level3 := makeTable([][]shared.BlockLevelElement{{nil}})

	// Level 2: table containing level3 (cell has paragraph + nested table).
	level2 := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}, level3}},
			}},
		},
	}

	// Level 1: table containing level2 (cell has paragraph + nested table).
	level1 := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}, level2}},
			}},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, level1)

	issues := Validate(doc)
	emptyCount := countCode(issues, CodeEmptyTC)
	if emptyCount != 1 {
		t.Errorf("expected 1 EMPTY_TC at level 3, got %d", emptyCount)
	}

	// Path should reflect 3 levels of nesting.
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			// Expect something like body/tbl[0]/tr[0]/tc[0]/tbl[0]/tr[0]/tc[0]/tbl[0]/tr[0]/tc[0]
			depth := 0
			for i := 0; i < len(iss.Path); i++ {
				if i+3 < len(iss.Path) && iss.Path[i:i+3] == "tbl" {
					depth++
				}
			}
			if depth < 3 {
				t.Errorf("expected 3 levels of tbl in path, got %d in %q", depth, iss.Path)
			}
		}
	}

	// AutoFix should fix it.
	AutoFix(doc)
	afterIssues := Validate(doc)
	if countCode(afterIssues, CodeEmptyTC) != 0 {
		t.Error("AutoFix should fix deeply nested empty cell")
	}
}

// ===========================================================================
// Integration: Full repair cycle — maximally broken document
// ===========================================================================

func TestIntegration_MaximallyBrokenDocument(t *testing.T) {
	t.Parallel()

	// Table with 3 rows × 3 cols, ALL cells empty.
	var rows []table.TblContent
	for r := 0; r < 3; r++ {
		var cells []table.RowContent
		for c := 0; c < 3; c++ {
			cells = append(cells, table.CT_Tc{Content: nil})
		}
		rows = append(rows, table.CT_Row{Content: cells})
	}

	doc := &packaging.Document{
		Document: &body.CT_Document{
			Body: &body.CT_Body{
				Content: []shared.BlockLevelElement{
					&table.CT_Tbl{Content: rows},
				},
			},
		},
		Styles:   &styles.CT_Styles{},
		Settings: &settings.CT_Settings{},
		Fonts:    &fonts.CT_FontsList{},
		Theme:    []byte("<theme/>"),
		// WebSettings intentionally nil.
	}

	// Step 1: Validate.
	issues := Validate(doc)
	t.Logf("Before AutoFix: %d issues, summary: %s", len(issues), Summary(issues))

	emptyTC := countCode(issues, CodeEmptyTC)
	if emptyTC != 9 {
		t.Errorf("3×3 empty table: expected 9 EMPTY_TC, got %d", emptyTC)
	}
	if countCode(issues, CodeMissingWebSettings) != 1 {
		t.Error("expected MISSING_WEBSETTINGS")
	}

	// Step 2: AutoFix.
	fixes := AutoFix(doc)
	t.Logf("Fixes applied: %v", fixes)

	// Step 3: Re-validate — all fixable issues should be gone.
	afterIssues := Validate(doc)
	t.Logf("After AutoFix: %d issues, summary: %s", len(afterIssues), Summary(afterIssues))

	if countCode(afterIssues, CodeEmptyTC) != 0 {
		t.Error("EMPTY_TC should be resolved after AutoFix")
	}
	if countCode(afterIssues, CodeMissingWebSettings) != 0 {
		t.Error("MISSING_WEBSETTINGS should be resolved after AutoFix")
	}
	if len(afterIssues) != 0 {
		t.Errorf("all issues should be resolved, got %d remaining", len(afterIssues))
	}

	// Step 4: Idempotent — second AutoFix does nothing.
	fixes2 := AutoFix(doc)
	if len(fixes2) != 0 {
		t.Errorf("second AutoFix should be no-op, got %d fixes", len(fixes2))
	}
}

// ===========================================================================
// Integration: Document with mixed block-level element types
// ===========================================================================

// Ensure the walker correctly counts table indices when paragraphs,
// RawXML, and tables are interleaved.
func TestIntegration_MixedBlockLevelIndexing(t *testing.T) {
	t.Parallel()

	rawSdt := shared.RawXML{
		XMLName: xml.Name{Local: "sdt", Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
		Inner:   []byte("<sdtContent><p/></sdtContent>"),
	}

	tbl0Valid := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}})
	tbl1Empty := makeTable([][]shared.BlockLevelElement{{nil}})
	tbl2Valid := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}})
	tbl3Empty := makeTable([][]shared.BlockLevelElement{{nil}})

	doc := minimalDoc()
	doc.Document.Body.Content = []shared.BlockLevelElement{
		&para.CT_P{},
		tbl0Valid,  // tbl[0]
		rawSdt,     // not a table — skipped
		&para.CT_P{},
		tbl1Empty,  // tbl[1]
		&para.CT_P{},
		tbl2Valid,  // tbl[2]
		tbl3Empty,  // tbl[3]
	}

	issues := Validate(doc)

	emptyIssues := filterCode(issues, CodeEmptyTC)
	if len(emptyIssues) != 2 {
		t.Fatalf("expected 2 EMPTY_TC, got %d", len(emptyIssues))
	}

	// First empty cell should reference tbl[1].
	if !containsSub(emptyIssues[0].Path, "tbl[1]") {
		t.Errorf("first EMPTY_TC should be tbl[1], got path %q", emptyIssues[0].Path)
	}
	// Second empty cell should reference tbl[3].
	if !containsSub(emptyIssues[1].Path, "tbl[3]") {
		t.Errorf("second EMPTY_TC should be tbl[3], got path %q", emptyIssues[1].Path)
	}
}

// ===========================================================================
// Integration: Validate output stability — same input → same output
// ===========================================================================

func TestIntegration_DeterministicOutput(t *testing.T) {
	t.Parallel()

	buildDoc := func() *packaging.Document {
		return &packaging.Document{
			Document: &body.CT_Document{
				Body: &body.CT_Body{
					Content: []shared.BlockLevelElement{
						makeTable([][]shared.BlockLevelElement{{nil, &para.CT_P{}}}),
						makeTable([][]shared.BlockLevelElement{{nil}}),
					},
				},
			},
			Styles:   &styles.CT_Styles{},
			Settings: nil, // missing
			Fonts:    &fonts.CT_FontsList{},
			Theme:    []byte("<t/>"),
		}
	}

	issues1 := Validate(buildDoc())
	issues2 := Validate(buildDoc())

	if len(issues1) != len(issues2) {
		t.Fatalf("non-deterministic: %d vs %d issues", len(issues1), len(issues2))
	}

	for i := range issues1 {
		if issues1[i].Code != issues2[i].Code ||
			issues1[i].Severity != issues2[i].Severity ||
			issues1[i].Path != issues2[i].Path ||
			issues1[i].Part != issues2[i].Part {
			t.Errorf("issue %d differs:\n  run1: %s\n  run2: %s", i, issues1[i], issues2[i])
		}
	}
}

// ===========================================================================
// Integration: AutoFix preserves table structure (TblPr, TblGrid, etc.)
// ===========================================================================

func TestIntegration_AutoFixPreservesTableProperties(t *testing.T) {
	t.Parallel()

	tbl := &table.CT_Tbl{
		TblPr: &table.CT_TblPr{
			TblW: &table.CT_TblWidth{W: 5000, Type: "pct"},
		},
		TblGrid: &table.CT_TblGrid{
			GridCol: []table.CT_TblGridCol{{W: 2500}, {W: 2500}},
		},
		Content: []table.TblContent{
			table.CT_Row{
				TrPr: &table.CT_TrPr{},
				Content: []table.RowContent{
					table.CT_Tc{
						TcPr:    &table.CT_TcPr{},
						Content: nil, // empty cell
					},
					table.CT_Tc{
						Content: []shared.BlockLevelElement{&para.CT_P{}},
					},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	AutoFix(doc)

	// TblPr should be untouched.
	if tbl.TblPr == nil || tbl.TblPr.TblW == nil || tbl.TblPr.TblW.W != 5000 {
		t.Error("AutoFix should not modify TblPr")
	}

	// TblGrid should be untouched.
	if tbl.TblGrid == nil || len(tbl.TblGrid.GridCol) != 2 {
		t.Error("AutoFix should not modify TblGrid")
	}

	// Row properties should be untouched.
	row := tbl.Content[0].(table.CT_Row)
	if row.TrPr == nil {
		t.Error("AutoFix should not remove TrPr")
	}

	// Cell properties should be untouched.
	cell := row.Content[0].(table.CT_Tc)
	if cell.TcPr == nil {
		t.Error("AutoFix should not remove TcPr from fixed cell")
	}

	// Fixed cell should now have a paragraph.
	if !cellHasParagraph(cell) {
		t.Error("fixed cell should have a paragraph")
	}
}

// ===========================================================================
// Stress: Large table
// ===========================================================================

func TestStress_LargeTable(t *testing.T) {
	t.Parallel()

	const rows, cols = 50, 20

	var tblRows []table.TblContent
	for r := 0; r < rows; r++ {
		var cells []table.RowContent
		for c := 0; c < cols; c++ {
			// Make every 7th cell empty.
			if (r*cols+c)%7 == 0 {
				cells = append(cells, table.CT_Tc{Content: nil})
			} else {
				cells = append(cells, table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}})
			}
		}
		tblRows = append(tblRows, table.CT_Row{Content: cells})
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content,
		&table.CT_Tbl{Content: tblRows})

	// Count expected empty cells.
	expectedEmpty := 0
	for i := 0; i < rows*cols; i++ {
		if i%7 == 0 {
			expectedEmpty++
		}
	}

	issues := Validate(doc)
	emptyCount := countCode(issues, CodeEmptyTC)
	if emptyCount != expectedEmpty {
		t.Errorf("expected %d EMPTY_TC in %dx%d table, got %d", expectedEmpty, rows, cols, emptyCount)
	}

	// AutoFix and verify.
	AutoFix(doc)
	afterIssues := Validate(doc)
	if countCode(afterIssues, CodeEmptyTC) != 0 {
		t.Error("AutoFix should fix all empty cells in large table")
	}
}

// ===========================================================================
// Stress: Many tables
// ===========================================================================

func TestStress_ManyTables(t *testing.T) {
	t.Parallel()

	const numTables = 100

	doc := minimalDoc()
	expectedEmpty := 0

	for i := 0; i < numTables; i++ {
		if i%3 == 0 {
			// Every 3rd table has an empty cell.
			doc.Document.Body.Content = append(doc.Document.Body.Content,
				makeTable([][]shared.BlockLevelElement{{nil}}))
			expectedEmpty++
		} else {
			doc.Document.Body.Content = append(doc.Document.Body.Content,
				makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}}))
		}
	}

	issues := Validate(doc)
	emptyCount := countCode(issues, CodeEmptyTC)
	if emptyCount != expectedEmpty {
		t.Errorf("expected %d EMPTY_TC across %d tables, got %d", expectedEmpty, numTables, emptyCount)
	}

	AutoFix(doc)
	afterIssues := Validate(doc)
	if countCode(afterIssues, CodeEmptyTC) != 0 {
		t.Error("AutoFix should fix all empty cells across many tables")
	}
}

// ===========================================================================
// Integration: All fixable issues resolved, non-fixable remain
// ===========================================================================

func TestIntegration_AutoFixLeavesNonFixableIssues(t *testing.T) {
	t.Parallel()

	doc := &packaging.Document{
		Document: &body.CT_Document{
			Body: &body.CT_Body{
				Content: []shared.BlockLevelElement{
					makeTable([][]shared.BlockLevelElement{{nil}}),
				},
			},
		},
		Styles:      nil, // Error — NOT auto-fixable
		Settings:    nil, // Error — NOT auto-fixable
		Fonts:       nil, // Error — NOT auto-fixable
		Theme:       nil, // Error — NOT auto-fixable
		WebSettings: nil, // Warn  — auto-fixable
	}

	AutoFix(doc)
	issues := Validate(doc)

	// Fixable: EMPTY_TC and MISSING_WEBSETTINGS should be gone.
	if countCode(issues, CodeEmptyTC) != 0 {
		t.Error("EMPTY_TC should be fixed")
	}
	if countCode(issues, CodeMissingWebSettings) != 0 {
		t.Error("MISSING_WEBSETTINGS should be fixed")
	}

	// Non-fixable should remain.
	nonFixable := []string{
		CodeMissingStyles,
		CodeMissingSettings,
		CodeMissingFonts,
		CodeMissingTheme,
	}
	for _, code := range nonFixable {
		if countCode(issues, code) == 0 {
			t.Errorf("non-fixable issue %s should still be present after AutoFix", code)
		}
	}
}

// ===========================================================================
// Integration: Validate result contract — every issue has non-empty Code, Message
// ===========================================================================

func TestIntegration_IssueFieldsNonEmpty(t *testing.T) {
	t.Parallel()

	// Build a doc that triggers every known issue code.
	doc := &packaging.Document{
		Document: &body.CT_Document{
			Body: &body.CT_Body{
				Content: []shared.BlockLevelElement{
					makeTable([][]shared.BlockLevelElement{{nil}}),
				},
			},
		},
		// All parts nil.
	}

	issues := Validate(doc)
	for i, iss := range issues {
		if iss.Code == "" {
			t.Errorf("issue %d has empty Code: %+v", i, iss)
		}
		if iss.Message == "" {
			t.Errorf("issue %d has empty Message: %+v", i, iss)
		}
		// Severity should be a known value.
		switch iss.Severity {
		case Warn, Error, Fatal:
			// ok
		default:
			t.Errorf("issue %d has unknown Severity %d", i, iss.Severity)
		}
	}
}

// ===========================================================================
// Integration: Validate nil doc + AutoFix nil doc — both safe
// ===========================================================================

func TestIntegration_NilDocSafety(t *testing.T) {
	t.Parallel()

	issues := Validate(nil)
	if len(issues) != 1 || issues[0].Code != CodeNilDocument {
		t.Error("Validate(nil) should return exactly 1 NIL_DOCUMENT issue")
	}

	fixes := AutoFix(nil)
	if fixes != nil {
		t.Error("AutoFix(nil) should return nil")
	}
}

// ===========================================================================
// Integration: Summary agrees with actual issue counts
// ===========================================================================

func TestIntegration_SummaryMatchesIssues(t *testing.T) {
	t.Parallel()

	doc := &packaging.Document{
		Document: &body.CT_Document{
			Body: &body.CT_Body{
				Content: []shared.BlockLevelElement{
					makeTable([][]shared.BlockLevelElement{{nil, nil}}),
				},
			},
		},
		Styles:      nil,
		Settings:    &settings.CT_Settings{},
		Fonts:       &fonts.CT_FontsList{},
		Theme:       []byte("<t/>"),
		WebSettings: nil,
	}

	issues := Validate(doc)
	summary := Summary(issues)

	nFatal := len(Filter(issues, Fatal))
	nError := len(Filter(issues, Error))
	nWarn := len(Filter(issues, Warn))

	if nFatal > 0 && !containsSub(summary, fmt.Sprintf("%d fatal", nFatal)) {
		t.Errorf("summary %q should mention %d fatal", summary, nFatal)
	}
	if nError > 0 && !containsSub(summary, fmt.Sprintf("%d error", nError)) {
		t.Errorf("summary %q should mention %d error", summary, nError)
	}
	if nWarn > 0 && !containsSub(summary, fmt.Sprintf("%d warning", nWarn)) {
		t.Errorf("summary %q should mention %d warning", summary, nWarn)
	}
}

// ===========================================================================
// helpers (local to this file)
// ===========================================================================

func filterCode(issues []Issue, code string) []Issue {
	var out []Issue
	for _, iss := range issues {
		if iss.Code == code {
			out = append(out, iss)
		}
	}
	return out
}
