package validator

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/packaging"
	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
)

// ===========================================================================
// walkBlockContent — type assertion edge cases
// ===========================================================================

// walkBlockContent matches *table.CT_Tbl (pointer). Body content should always
// store tables as pointers; if a value were stored the validator would miss it.
// This test documents that assumption.
func TestWalk_OnlyPointerTablesDetected(t *testing.T) {
	t.Parallel()

	// A *table.CT_Tbl with an empty cell — validator SHOULD detect it.
	ptrTbl := makeTable([][]shared.BlockLevelElement{{nil}})

	doc := minimalDoc()
	doc.Document.Body.Content = []shared.BlockLevelElement{&para.CT_P{}, ptrTbl}

	issues := Validate(doc)
	count := countCode(issues, CodeEmptyTC)
	if count != 1 {
		t.Errorf("pointer table: expected 1 EMPTY_TC, got %d", count)
	}
}

// Body content that is entirely RawXML (no tables, no paragraphs) should not
// trigger any table-related issues.
func TestWalk_BodyWithOnlyRawXML(t *testing.T) {
	t.Parallel()

	raw := shared.RawXML{
		XMLName: xml.Name{Local: "customXml"},
		Inner:   []byte("<data/>"),
	}

	doc := minimalDoc()
	doc.Document.Body.Content = []shared.BlockLevelElement{raw}

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			t.Errorf("no tables present — should not see EMPTY_TC: %s", iss)
		}
	}
}

// ===========================================================================
// checkTable — structural edge cases
// ===========================================================================

// Table with zero rows (empty Content slice) should produce no issues.
func TestCheckTable_EmptyTableNoRows(t *testing.T) {
	t.Parallel()

	tbl := &table.CT_Tbl{Content: nil}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			t.Errorf("table with no rows should not produce EMPTY_TC: %s", iss)
		}
	}
}

// Row with zero cells (empty Content slice) should produce no issues.
func TestCheckTable_RowWithNoCells(t *testing.T) {
	t.Parallel()

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: nil}, // row exists but has no cells
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			t.Errorf("row with no cells should not produce EMPTY_TC: %s", iss)
		}
	}
}

// Row content that includes RawRowContent should be skipped without panicking.
func TestCheckTable_RowWithRawRowContent(t *testing.T) {
	t.Parallel()

	raw := table.RawRowContent{
		RawXML: shared.RawXML{
			XMLName: xml.Name{Local: "bookmarkStart"},
			Inner:   []byte{},
		},
	}

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					raw, // not a cell — should be skipped
					table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			t.Errorf("row with RawRowContent + valid cell should not produce EMPTY_TC: %s", iss)
		}
	}
}

// Table content that includes RawTblContent should be skipped without panicking.
func TestCheckTable_TableWithRawTblContent(t *testing.T) {
	t.Parallel()

	raw := table.RawTblContent{
		RawXML: shared.RawXML{
			XMLName: xml.Name{Local: "customXml"},
			Inner:   []byte{},
		},
	}

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			raw, // not a row
			table.CT_Row{
				Content: []table.RowContent{
					table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			t.Errorf("should not see EMPTY_TC — only valid cells present: %s", iss)
		}
	}
}

// ===========================================================================
// cellHasParagraph — what counts as a paragraph
// ===========================================================================

// Cell containing only a nested table (no paragraph) is still EMPTY_TC.
func TestCellHasParagraph_OnlyNestedTable(t *testing.T) {
	t.Parallel()

	inner := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}})

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					// Cell has a table but NO paragraph — invalid per spec.
					table.CT_Tc{Content: []shared.BlockLevelElement{inner}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	if countCode(issues, CodeEmptyTC) != 1 {
		t.Error("cell with only a nested table (no <w:p>) should produce EMPTY_TC")
	}
}

// Cell containing only RawXML (no paragraph) is still EMPTY_TC.
func TestCellHasParagraph_OnlyRawXML(t *testing.T) {
	t.Parallel()

	raw := shared.RawXML{
		XMLName: xml.Name{Local: "sdt"},
		Inner:   []byte("<sdtContent/>"),
	}

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					table.CT_Tc{Content: []shared.BlockLevelElement{raw}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	if countCode(issues, CodeEmptyTC) != 1 {
		t.Error("cell with only RawXML (no <w:p>) should produce EMPTY_TC")
	}
}

// Cell with mixed content: RawXML + paragraph → valid.
func TestCellHasParagraph_RawXMLPlusParagraph(t *testing.T) {
	t.Parallel()

	raw := shared.RawXML{
		XMLName: xml.Name{Local: "sdt"},
		Inner:   []byte("<sdtContent/>"),
	}

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					table.CT_Tc{Content: []shared.BlockLevelElement{raw, &para.CT_P{}}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	if countCode(issues, CodeEmptyTC) != 0 {
		t.Error("cell with RawXML + paragraph should NOT produce EMPTY_TC")
	}
}

// ===========================================================================
// checkBody — edge cases
// ===========================================================================

// Body with only paragraphs (no tables) → no EMPTY_TC.
func TestCheckBody_OnlyParagraphs(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.Document.Body.Content = []shared.BlockLevelElement{
		&para.CT_P{},
		&para.CT_P{},
		&para.CT_P{},
	}

	issues := Validate(doc)
	if countCode(issues, CodeEmptyTC) != 0 {
		t.Error("body with only paragraphs should have no EMPTY_TC")
	}
	if countCode(issues, CodeEmptyBody) != 0 {
		t.Error("body with paragraphs should not be EMPTY_BODY")
	}
}

// Body.Content = []shared.BlockLevelElement{} (empty slice, not nil) → EMPTY_BODY.
func TestCheckBody_EmptySliceNotNil(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.Document.Body.Content = []shared.BlockLevelElement{}

	issues := Validate(doc)
	if countCode(issues, CodeEmptyBody) != 1 {
		t.Error("body with empty slice should produce EMPTY_BODY")
	}
}

// ===========================================================================
// checkRequiredParts — independence of checks
// ===========================================================================

// All parts missing simultaneously → all codes present.
func TestCheckRequiredParts_AllMissing(t *testing.T) {
	t.Parallel()

	doc := &packaging.Document{} // everything nil/zero

	issues := Validate(doc)

	wantCodes := map[string]bool{
		CodeMissingDocumentBody: false,
		CodeMissingStyles:       false,
		CodeMissingSettings:     false,
		CodeMissingFonts:        false,
		CodeMissingTheme:        false,
		CodeMissingWebSettings:  false,
	}
	for _, iss := range issues {
		wantCodes[iss.Code] = true
	}
	for code, found := range wantCodes {
		if !found {
			t.Errorf("expected code %s when all parts missing", code)
		}
	}
}

// Theme set to empty []byte{} → still MISSING_THEME (len == 0).
func TestCheckRequiredParts_EmptyByteSliceTheme(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.Theme = []byte{} // allocated but empty

	issues := Validate(doc)
	if countCode(issues, CodeMissingTheme) != 1 {
		t.Error("empty byte slice theme should produce MISSING_THEME")
	}
}

// ===========================================================================
// EMPTY_TC path correctness
// ===========================================================================

// Multiple tables in body — each gets its own tbl[N] index.
func TestEmptyTC_PathIndexing_MultipleTables(t *testing.T) {
	t.Parallel()

	tbl0 := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}}) // valid
	tbl1 := makeTable([][]shared.BlockLevelElement{{nil}})          // empty cell

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl0, tbl1)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			if !strings.Contains(iss.Path, "tbl[1]") {
				t.Errorf("empty cell is in second table, expected tbl[1] in path, got %q", iss.Path)
			}
		}
	}
}

// Multi-row table — row indices are correct.
func TestEmptyTC_PathIndexing_MultiRow(t *testing.T) {
	t.Parallel()

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}},
			}},
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: nil}, // empty cell in second row
			}},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			if !strings.Contains(iss.Path, "tr[1]") {
				t.Errorf("empty cell is in second row, expected tr[1] in path, got %q", iss.Path)
			}
		}
	}
}

// Multi-cell row — cell indices are correct.
func TestEmptyTC_PathIndexing_MultiCell(t *testing.T) {
	t.Parallel()

	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{Content: []table.RowContent{
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}}, // valid
				table.CT_Tc{Content: []shared.BlockLevelElement{&para.CT_P{}}}, // valid
				table.CT_Tc{Content: nil},                                      // empty
			}},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			if !strings.Contains(iss.Path, "tc[2]") {
				t.Errorf("empty cell is third in row, expected tc[2] in path, got %q", iss.Path)
			}
		}
	}
}

// Paragraph between two tables does not affect table indexing.
func TestEmptyTC_PathIndexing_ParagraphBetweenTables(t *testing.T) {
	t.Parallel()

	tbl0 := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}})
	tbl1 := makeTable([][]shared.BlockLevelElement{{nil}})

	doc := minimalDoc()
	doc.Document.Body.Content = []shared.BlockLevelElement{
		&para.CT_P{},
		tbl0,         // tbl[0]
		&para.CT_P{}, // paragraph in between
		tbl1,         // tbl[1]
	}

	issues := Validate(doc)
	for _, iss := range issues {
		if iss.Code == CodeEmptyTC {
			if !strings.Contains(iss.Path, "tbl[1]") {
				t.Errorf("expected tbl[1] in path, got %q", iss.Path)
			}
		}
	}
}

// ===========================================================================
// Issue.String() formatting edge cases
// ===========================================================================

// Issue with empty Path → no double slash.
func TestIssueString_EmptyPath(t *testing.T) {
	t.Parallel()

	iss := Issue{
		Severity: Warn,
		Part:     "word/document.xml",
		Path:     "",
		Code:     CodeEmptyBody,
		Message:  "test",
	}
	s := iss.String()
	if strings.Contains(s, "//") {
		t.Errorf("empty Path should not produce double slash: %s", s)
	}
	if !strings.Contains(s, "word/document.xml") {
		t.Errorf("Part should appear in output: %s", s)
	}
}

// Issue with empty Part and empty Path.
func TestIssueString_AllEmpty(t *testing.T) {
	t.Parallel()

	iss := Issue{Severity: Fatal, Code: CodeNilDocument, Message: "nil"}
	s := iss.String()
	if s == "" {
		t.Error("String() should never be empty")
	}
}

// ===========================================================================
// AutoFix — edge cases
// ===========================================================================

// AutoFix with nil Document but non-nil other fields.
func TestAutoFix_NilDocumentField(t *testing.T) {
	t.Parallel()

	doc := &packaging.Document{
		Document:    nil,
		WebSettings: nil,
	}

	// Should not panic, and should still fix webSettings.
	fixes := AutoFix(doc)

	if len(doc.WebSettings) == 0 {
		t.Error("AutoFix should fix webSettings even when Document is nil")
	}
	if len(fixes) == 0 {
		t.Error("expected webSettings fix")
	}
}

// AutoFix on document where Document exists but Body is nil.
func TestAutoFix_NilBody(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()
	doc.Document.Body = nil
	doc.WebSettings = nil

	fixes := AutoFix(doc)

	// Should fix webSettings, not panic on nil body.
	if len(doc.WebSettings) == 0 {
		t.Error("AutoFix should fix webSettings even when Body is nil")
	}
	if len(fixes) != 1 {
		t.Errorf("expected 1 fix (webSettings only), got %d: %v", len(fixes), fixes)
	}
}

// AutoFix preserves existing content in cells when adding paragraph.
func TestAutoFix_PreservesExistingCellContent(t *testing.T) {
	t.Parallel()

	inner := makeTable([][]shared.BlockLevelElement{{&para.CT_P{}}})
	raw := shared.RawXML{
		XMLName: xml.Name{Local: "sdt"},
		Inner:   []byte("<data/>"),
	}

	// Cell has a table + RawXML but no paragraph.
	tbl := &table.CT_Tbl{
		Content: []table.TblContent{
			table.CT_Row{
				Content: []table.RowContent{
					table.CT_Tc{Content: []shared.BlockLevelElement{inner, raw}},
				},
			},
		},
	}

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl)

	AutoFix(doc)

	cell := firstCell(tbl)

	// Should have 3 elements: prepended paragraph + original table + original raw.
	if len(cell.Content) != 3 {
		t.Fatalf("expected 3 elements after fix, got %d", len(cell.Content))
	}

	// First element should be the added paragraph.
	if _, ok := cell.Content[0].(*para.CT_P); !ok {
		t.Error("first element should be the prepended *para.CT_P")
	}

	// Original table should still be at index 1.
	if _, ok := cell.Content[1].(*table.CT_Tbl); !ok {
		t.Error("original nested table should be preserved at index 1")
	}
}

// AutoFix on already-valid document produces no fixes and no side effects.
func TestAutoFix_ValidDocumentNoChanges(t *testing.T) {
	t.Parallel()

	doc := minimalDoc()

	fixes := AutoFix(doc)
	if len(fixes) != 0 {
		t.Errorf("valid document should produce no fixes, got %d: %v", len(fixes), fixes)
	}

	// Validate should still return 0 issues.
	issues := Validate(doc)
	if len(issues) != 0 {
		t.Errorf("valid document should have 0 issues after no-op AutoFix, got %d", len(issues))
	}
}

// AutoFix does not modify WebSettings if already populated.
func TestAutoFix_DoesNotOverwriteWebSettings(t *testing.T) {
	t.Parallel()

	original := []byte(`<w:webSettings xmlns:w="..."><w:customSetting/></w:webSettings>`)
	doc := minimalDoc()
	doc.WebSettings = original

	AutoFix(doc)

	if string(doc.WebSettings) != string(original) {
		t.Error("AutoFix should not overwrite existing WebSettings")
	}
}

// ===========================================================================
// Multiple issues from multiple tables in one body
// ===========================================================================

func TestValidate_MultipleTablesMultipleEmptyCells(t *testing.T) {
	t.Parallel()

	tbl1 := makeTable([][]shared.BlockLevelElement{
		{nil, &para.CT_P{}}, // row 0: 1 empty, 1 valid
		{nil, nil},          // row 1: 2 empty
	})
	tbl2 := makeTable([][]shared.BlockLevelElement{
		{nil}, // row 0: 1 empty
	})

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl1, tbl2)

	issues := Validate(doc)
	emptyCount := countCode(issues, CodeEmptyTC)

	// tbl1: 1 + 2 = 3 empty cells; tbl2: 1 empty cell → total 4.
	if emptyCount != 4 {
		t.Errorf("expected 4 EMPTY_TC issues, got %d", emptyCount)
		for _, iss := range issues {
			if iss.Code == CodeEmptyTC {
				t.Logf("  %s", iss)
			}
		}
	}
}

// ===========================================================================
// AutoFix on multiple tables
// ===========================================================================

func TestAutoFix_MultipleTablesAllFixed(t *testing.T) {
	t.Parallel()

	tbl1 := makeTable([][]shared.BlockLevelElement{{nil}})
	tbl2 := makeTable([][]shared.BlockLevelElement{{nil}, {nil}})

	doc := minimalDoc()
	doc.Document.Body.Content = append(doc.Document.Body.Content, tbl1, tbl2)

	AutoFix(doc)

	// All cells should now have paragraphs.
	issues := Validate(doc)
	if countCode(issues, CodeEmptyTC) != 0 {
		t.Error("AutoFix should have fixed all empty cells across both tables")
	}
}

// ===========================================================================
// Helpers
// ===========================================================================

func countCode(issues []Issue, code string) int {
	n := 0
	for _, iss := range issues {
		if iss.Code == code {
			n++
		}
	}
	return n
}
