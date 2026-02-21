package oxml

import (
	"testing"

	"github.com/user/go-docx/pkg/docx/enum"
)

// ===========================================================================
// Table tests
// ===========================================================================

func TestNewTbl_Structure(t *testing.T) {
	tbl := NewTbl(3, 4, 9360)
	// Check tblPr present
	tblPr, err := tbl.TblPr()
	if err != nil {
		t.Fatalf("TblPr error: %v", err)
	}
	if tblPr == nil {
		t.Fatal("expected tblPr, got nil")
	}
	// Check tblGrid
	grid, err := tbl.TblGrid()
	if err != nil {
		t.Fatalf("TblGrid error: %v", err)
	}
	cols := grid.GridColList()
	if len(cols) != 4 {
		t.Errorf("expected 4 gridCol, got %d", len(cols))
	}
	// Check column widths
	for _, col := range cols {
		w := col.W()
		if w != 2340 { // 9360/4
			t.Errorf("expected col width 2340, got %d", w)
		}
	}
	// Check rows
	trs := tbl.TrList()
	if len(trs) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(trs))
	}
	// Check cells per row
	for i, tr := range trs {
		tcs := tr.TcList()
		if len(tcs) != 4 {
			t.Errorf("row %d: expected 4 cells, got %d", i, len(tcs))
		}
	}
}

func TestCT_Tbl_ColCount(t *testing.T) {
	tbl := NewTbl(2, 5, 10000)
	got, err := tbl.ColCount()
	if err != nil {
		t.Fatal(err)
	}
	if got != 5 {
		t.Errorf("expected ColCount=5, got %d", got)
	}
}

func TestCT_Tbl_ColWidths(t *testing.T) {
	tbl := NewTbl(1, 3, 9000)
	widths, err := tbl.ColWidths()
	if err != nil {
		t.Fatal(err)
	}
	if len(widths) != 3 {
		t.Fatalf("expected 3 widths, got %d", len(widths))
	}
	for _, w := range widths {
		if w != 3000 {
			t.Errorf("expected 3000, got %d", w)
		}
	}
}

func TestCT_Tbl_IterTcs(t *testing.T) {
	tbl := NewTbl(2, 3, 6000)
	tcs := tbl.IterTcs()
	if len(tcs) != 6 {
		t.Errorf("expected 6 cells, got %d", len(tcs))
	}
}

func TestCT_Tbl_TblStyleVal_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	v, err := tbl.TblStyleVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "" {
		t.Errorf("expected empty, got %q", v)
	}
	if err := tbl.SetTblStyleVal("TableGrid"); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.TblStyleVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "TableGrid" {
		t.Errorf("expected TableGrid, got %q", v)
	}
	if err := tbl.SetTblStyleVal(""); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.TblStyleVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != "" {
		t.Errorf("expected empty after clear, got %q", v)
	}
}

func TestCT_Tbl_Alignment_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	v, err := tbl.AlignmentVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	center := enum.WdTableAlignmentCenter
	if err := tbl.SetAlignmentVal(&center); err != nil {
		t.Fatal(err)
	}
	got, err := tbl.AlignmentVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != enum.WdTableAlignmentCenter {
		t.Errorf("expected Center, got %v", got)
	}
	if err := tbl.SetAlignmentVal(nil); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.AlignmentVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_Tbl_BidiVisualVal_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	v, err := tbl.BidiVisualVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	tr := true
	if err := tbl.SetBidiVisualVal(&tr); err != nil {
		t.Fatal(err)
	}
	got, err := tbl.BidiVisualVal()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != true {
		t.Errorf("expected true, got %v", got)
	}
	if err := tbl.SetBidiVisualVal(nil); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.BidiVisualVal()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
}

func TestCT_Tbl_Autofit_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	// Default should be true (no tblLayout means autofit)
	v, err := tbl.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if !v {
		t.Error("expected autofit=true by default")
	}
	if err := tbl.SetAutofit(false); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if v {
		t.Error("expected autofit=false after set")
	}
	if err := tbl.SetAutofit(true); err != nil {
		t.Fatal(err)
	}
	v, err = tbl.Autofit()
	if err != nil {
		t.Fatal(err)
	}
	if !v {
		t.Error("expected autofit=true after reset")
	}
}

func TestCT_Row_TrIdx(t *testing.T) {
	tbl := NewTbl(3, 1, 1000)
	trs := tbl.TrList()
	for i, tr := range trs {
		if got := tr.TrIdx(); got != i {
			t.Errorf("row %d: expected TrIdx=%d, got %d", i, i, got)
		}
	}
}

func TestCT_Row_TcAtGridOffset(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tr := tbl.TrList()[0]
	tc, err := tr.TcAtGridOffset(0)
	if err != nil {
		t.Fatal(err)
	}
	if tc == nil {
		t.Fatal("expected non-nil tc at offset 0")
	}
	tc2, err := tr.TcAtGridOffset(2)
	if err != nil {
		t.Fatal(err)
	}
	if tc2 == nil {
		t.Fatal("expected non-nil tc at offset 2")
	}
	_, err = tr.TcAtGridOffset(5)
	if err == nil {
		t.Error("expected error for offset 5")
	}
}

func TestCT_Row_TrHeight_RoundTrip(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tr := tbl.TrList()[0]
	if v := tr.TrHeightVal(); v != nil {
		t.Errorf("expected nil, got %d", *v)
	}
	h := 720
	if err := tr.SetTrHeightVal(&h); err != nil {
		t.Fatalf("SetTrHeightVal: %v", err)
	}
	got := tr.TrHeightVal()
	if got == nil || *got != 720 {
		t.Errorf("expected 720, got %v", got)
	}
	rule := enum.WdRowHeightRuleExactly
	if err := tr.SetTrHeightHRule(&rule); err != nil {
		t.Fatalf("SetTrHeightHRule: %v", err)
	}
	gotRule := tr.TrHeightHRule()
	if gotRule == nil || *gotRule != enum.WdRowHeightRuleExactly {
		t.Errorf("expected Exactly, got %v", gotRule)
	}
}

func TestNewTc(t *testing.T) {
	tc := NewTc()
	ps := tc.PList()
	if len(ps) != 1 {
		t.Errorf("expected 1 paragraph, got %d", len(ps))
	}
}

func TestCT_Tc_GridSpan_RoundTrip(t *testing.T) {
	tc := NewTc()
	if v := tc.GridSpanVal(); v != 1 {
		t.Errorf("expected 1, got %d", v)
	}
	if err := tc.SetGridSpanVal(3); err != nil {
		t.Fatalf("SetGridSpanVal: %v", err)
	}
	if v := tc.GridSpanVal(); v != 3 {
		t.Errorf("expected 3, got %d", v)
	}
	tc.SetGridSpanVal(1) // should remove
	if v := tc.GridSpanVal(); v != 1 {
		t.Errorf("expected 1 after reset, got %d", v)
	}
}

func TestCT_Tc_VMerge_RoundTrip(t *testing.T) {
	tc := NewTc()
	if v := tc.VMergeVal(); v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	restart := "restart"
	if err := tc.SetVMergeVal(&restart); err != nil {
		t.Fatalf("SetVMergeVal: %v", err)
	}
	got := tc.VMergeVal()
	if got == nil || *got != "restart" {
		t.Errorf("expected restart, got %v", got)
	}
	if err := tc.SetVMergeVal(nil); err != nil {
		t.Fatalf("SetVMergeVal: %v", err)
	}
	if v := tc.VMergeVal(); v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_Tc_Width_RoundTrip(t *testing.T) {
	tc := NewTc()
	if v := tc.WidthTwips(); v != nil {
		t.Errorf("expected nil, got %d", *v)
	}
	if err := tc.SetWidthTwips(2880); err != nil {
		t.Fatalf("SetWidthTwips: %v", err)
	}
	got := tc.WidthTwips()
	if got == nil || *got != 2880 {
		t.Errorf("expected 2880, got %v", got)
	}
}

func TestCT_Tc_VAlign_RoundTrip(t *testing.T) {
	tc := NewTc()
	if v := tc.VAlignVal(); v != nil {
		t.Errorf("expected nil, got %v", *v)
	}
	center := enum.WdCellVerticalAlignmentCenter
	if err := tc.SetVAlignVal(&center); err != nil {
		t.Fatalf("SetVAlignVal: %v", err)
	}
	got := tc.VAlignVal()
	if got == nil || *got != enum.WdCellVerticalAlignmentCenter {
		t.Errorf("expected center, got %v", got)
	}
	if err := tc.SetVAlignVal(nil); err != nil {
		t.Fatalf("SetVAlignVal(nil): %v", err)
	}
	if v := tc.VAlignVal(); v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_Tc_InnerContentElements(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tc := tbl.TrList()[0].TcList()[0]
	elems := tc.InnerContentElements()
	if len(elems) != 1 {
		t.Errorf("expected 1 inner element, got %d", len(elems))
	}
	if _, ok := elems[0].(*CT_P); !ok {
		t.Error("expected first element to be *CT_P")
	}
}

func TestCT_Tc_ClearContent(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tc := tbl.TrList()[0].TcList()[0]
	tc.ClearContent()
	// Should have no p or tbl children, only tcPr
	if elems := tc.InnerContentElements(); len(elems) != 0 {
		t.Errorf("expected 0 inner elements after clear, got %d", len(elems))
	}
	if tcPr := tc.TcPr(); tcPr == nil {
		t.Error("expected tcPr to be preserved")
	}
}

func TestCT_Tc_IsEmpty(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tc := tbl.TrList()[0].TcList()[0]
	if !tc.IsEmpty() {
		t.Error("expected new cell to be empty")
	}
}

func TestCT_Tc_GridOffset(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	offsets := []int{0, 1, 2}
	for i, tc := range tcs {
		if got := tc.GridOffset(); got != offsets[i] {
			t.Errorf("cell %d: expected offset %d, got %d", i, offsets[i], got)
		}
	}
}

func TestCT_Tc_LeftRight(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	if got := tcs[0].Left(); got != 0 {
		t.Errorf("expected left=0, got %d", got)
	}
	if got := tcs[0].Right(); got != 1 {
		t.Errorf("expected right=1, got %d", got)
	}
	if got := tcs[2].Right(); got != 3 {
		t.Errorf("expected right=3, got %d", got)
	}
}

func TestCT_Tc_TopBottom(t *testing.T) {
	tbl := NewTbl(2, 1, 1000)
	tcs := tbl.IterTcs()
	if got := tcs[0].Top(); got != 0 {
		t.Errorf("expected top=0, got %d", got)
	}
	if got := tcs[0].Bottom(); got != 1 {
		t.Errorf("expected bottom=1, got %d", got)
	}
	if got := tcs[1].Top(); got != 1 {
		t.Errorf("expected top=1, got %d", got)
	}
}

func TestCT_Tc_NextTc(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	next := tcs[0].NextTc()
	if next == nil {
		t.Fatal("expected next tc")
	}
	if next.E != tcs[1].E {
		t.Error("next tc should be second cell")
	}
	if last := tcs[2].NextTc(); last != nil {
		t.Error("expected nil for last cell")
	}
}

func TestCT_TblGridCol_GridColIdx(t *testing.T) {
	tbl := NewTbl(1, 4, 4000)
	grid, err := tbl.TblGrid()
	if err != nil {
		t.Fatal(err)
	}
	cols := grid.GridColList()
	for i, col := range cols {
		if got := col.GridColIdx(); got != i {
			t.Errorf("col %d: expected idx %d, got %d", i, i, got)
		}
	}
}

func TestCT_TblWidth_WidthTwips(t *testing.T) {
	tbl := NewTbl(1, 1, 2000)
	tc := tbl.TrList()[0].TcList()[0]
	tcPr := tc.TcPr()
	tcW := tcPr.TcW()
	if tcW == nil {
		t.Fatal("expected tcW")
	}
	w := tcW.WidthTwips()
	if w == nil || *w != 2000 {
		t.Errorf("expected 2000, got %v", w)
	}
}

func TestCT_Tc_Merge_Horizontal(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	topTc, err := tcs[0].Merge(tcs[2])
	if err != nil {
		t.Fatal(err)
	}
	if topTc.GridSpanVal() != 3 {
		t.Errorf("expected gridSpan=3, got %d", topTc.GridSpanVal())
	}
	// After merge, row should have only 1 tc
	remaining := tbl.TrList()[0].TcList()
	if len(remaining) != 1 {
		t.Errorf("expected 1 remaining tc, got %d", len(remaining))
	}
}

// ===========================================================================
// Section tests
// ===========================================================================

func TestCT_SectPr_PageWidth_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	if v := sp.PageWidth(); v != nil {
		t.Errorf("expected nil, got %d", *v)
	}
	w := 12240
	if err := sp.SetPageWidth(&w); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	got := sp.PageWidth()
	if got == nil || *got != 12240 {
		t.Errorf("expected 12240, got %v", got)
	}
	if err := sp.SetPageWidth(nil); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	if v := sp.PageWidth(); v != nil {
		t.Errorf("expected nil after clear, got %v", *v)
	}
}

func TestCT_SectPr_PageHeight_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	h := 15840
	if err := sp.SetPageHeight(&h); err != nil {
		t.Fatalf("SetPageHeight: %v", err)
	}
	got := sp.PageHeight()
	if got == nil || *got != 15840 {
		t.Errorf("expected 15840, got %v", got)
	}
}

func TestCT_SectPr_Orientation_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	// Default portrait
	if sp.Orientation() != enum.WdOrientationPortrait {
		t.Error("expected portrait by default")
	}
	if err := sp.SetOrientation(enum.WdOrientationLandscape); err != nil {
		t.Fatalf("SetOrientation(landscape): %v", err)
	}
	if sp.Orientation() != enum.WdOrientationLandscape {
		t.Error("expected landscape")
	}
	if err := sp.SetOrientation(enum.WdOrientationPortrait); err != nil {
		t.Fatalf("SetOrientation(portrait): %v", err)
	}
	// After setting portrait, orient attr should be removed (default)
	pgSz := sp.PgSz()
	if pgSz != nil {
		_, ok := pgSz.GetAttr("w:orient")
		if ok {
			t.Error("expected orient attr to be removed for portrait")
		}
	}
}

func TestCT_SectPr_StartType_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	if sp.StartType() != enum.WdSectionStartNewPage {
		t.Error("expected NEW_PAGE by default")
	}
	if err := sp.SetStartType(enum.WdSectionStartContinuous); err != nil {
		t.Fatal(err)
	}
	if sp.StartType() != enum.WdSectionStartContinuous {
		t.Error("expected Continuous")
	}
	if err := sp.SetStartType(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}
	if sp.Type() != nil {
		t.Error("expected type element removed for NEW_PAGE")
	}
}

func TestCT_SectPr_TitlePg_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	if sp.TitlePgVal() {
		t.Error("expected false by default")
	}
	if err := sp.SetTitlePgVal(true); err != nil {
		t.Fatalf("SetTitlePgVal: %v", err)
	}
	if !sp.TitlePgVal() {
		t.Error("expected true after set")
	}
	if err := sp.SetTitlePgVal(false); err != nil {
		t.Fatalf("SetTitlePgVal: %v", err)
	}
	if sp.TitlePg() != nil {
		t.Error("expected titlePg element removed")
	}
}

func TestCT_SectPr_Margins_RoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}

	top := 1440
	if err := sp.SetTopMargin(&top); err != nil {
		t.Fatalf("SetTopMargin: %v", err)
	}
	got := sp.TopMargin()
	if got == nil || *got != 1440 {
		t.Errorf("top: expected 1440, got %v", got)
	}

	bottom := 1440
	if err := sp.SetBottomMargin(&bottom); err != nil {
		t.Fatalf("SetBottomMargin: %v", err)
	}
	if got := sp.BottomMargin(); got == nil || *got != 1440 {
		t.Errorf("bottom: expected 1440, got %v", got)
	}

	left := 1800
	if err := sp.SetLeftMargin(&left); err != nil {
		t.Fatalf("SetLeftMargin: %v", err)
	}
	if got := sp.LeftMargin(); got == nil || *got != 1800 {
		t.Errorf("left: expected 1800, got %v", got)
	}

	right := 1800
	if err := sp.SetRightMargin(&right); err != nil {
		t.Fatalf("SetRightMargin: %v", err)
	}
	if got := sp.RightMargin(); got == nil || *got != 1800 {
		t.Errorf("right: expected 1800, got %v", got)
	}

	hdr := 720
	if err := sp.SetHeaderMargin(&hdr); err != nil {
		t.Fatalf("SetHeaderMargin: %v", err)
	}
	if got := sp.HeaderMargin(); got == nil || *got != 720 {
		t.Errorf("header: expected 720, got %v", got)
	}

	ftr := 720
	if err := sp.SetFooterMargin(&ftr); err != nil {
		t.Fatalf("SetFooterMargin: %v", err)
	}
	if got := sp.FooterMargin(); got == nil || *got != 720 {
		t.Errorf("footer: expected 720, got %v", got)
	}

	gut := 0
	if err := sp.SetGutterMargin(&gut); err != nil {
		t.Fatalf("SetGutterMargin: %v", err)
	}
}

func TestCT_SectPr_Clone(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	w := 12240
	if err := sp.SetPageWidth(&w); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	sp.E.CreateAttr("w:rsidR", "00A12345")

	cloned := sp.Clone()
	// Width should be preserved
	if cw := cloned.PageWidth(); cw == nil || *cw != 12240 {
		t.Errorf("expected cloned width 12240, got %v", cw)
	}
	// rsid should be removed
	if _, ok := cloned.GetAttr("w:rsidR"); ok {
		t.Error("expected rsid attribute to be removed in clone")
	}
	// Modifying clone shouldn't affect original
	w2 := 9999
	if err := cloned.SetPageWidth(&w2); err != nil {
		t.Fatalf("SetPageWidth: %v", err)
	}
	if orig := sp.PageWidth(); orig == nil || *orig != 12240 {
		t.Error("original should be unchanged")
	}
}

func TestCT_SectPr_HeaderFooterRef(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}

	// Add header ref
	_, err := sp.AddHeaderRef(enum.WdHeaderFooterIndexPrimary, "rId1")
	if err != nil {
		t.Fatalf("AddHeaderRef: %v", err)
	}
	ref, err := sp.GetHeaderRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if ref == nil {
		t.Fatal("expected header ref")
	}
	rId, _ := ref.RId()
	if rId != "rId1" {
		t.Errorf("expected rId1, got %s", rId)
	}

	// Add footer ref
	if _, err = sp.AddFooterRef(enum.WdHeaderFooterIndexPrimary, "rId2"); err != nil {
		t.Fatalf("AddFooterRef: %v", err)
	}
	fRef, err := sp.GetFooterRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if fRef == nil {
		t.Fatal("expected footer ref")
	}

	// Remove header ref
	removed := sp.RemoveHeaderRef(enum.WdHeaderFooterIndexPrimary)
	if removed != "rId1" {
		t.Errorf("expected removed rId1, got %s", removed)
	}
	refAfter, err := sp.GetHeaderRef(enum.WdHeaderFooterIndexPrimary)
	if err != nil {
		t.Fatal(err)
	}
	if refAfter != nil {
		t.Error("expected header ref to be removed")
	}

	// Remove footer ref
	removedF := sp.RemoveFooterRef(enum.WdHeaderFooterIndexPrimary)
	if removedF != "rId2" {
		t.Errorf("expected removed rId2, got %s", removedF)
	}
}

func TestCT_HdrFtr_InnerContentElements(t *testing.T) {
	hf := &CT_HdrFtr{Element{E: OxmlElement("w:hdr")}}
	hf.AddP()
	hf.AddTbl()
	elems := hf.InnerContentElements()
	if len(elems) != 2 {
		t.Errorf("expected 2, got %d", len(elems))
	}
}

// ===========================================================================
// Styles tests
// ===========================================================================

func TestStyleIdFromName(t *testing.T) {
	cases := []struct {
		name, expected string
	}{
		{"Heading 1", "Heading1"},
		{"heading 1", "Heading1"},
		{"caption", "Caption"},
		{"Normal", "Normal"},
		{"Table of Contents", "TableofContents"},
		{"Body Text", "BodyText"},
	}
	for _, c := range cases {
		if got := StyleIdFromName(c.name); got != c.expected {
			t.Errorf("StyleIdFromName(%q) = %q, want %q", c.name, got, c.expected)
		}
	}
}

func TestCT_Styles_GetByID(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("Heading1"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := s.SetNameVal("heading 1"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}

	found := styles.GetByID("Heading1")
	if found == nil {
		t.Fatal("expected to find style by ID")
	}
	if found.NameVal() != "heading 1" {
		t.Errorf("expected name 'heading 1', got %q", found.NameVal())
	}

	if styles.GetByID("NoSuchStyle") != nil {
		t.Error("expected nil for unknown style")
	}
}

func TestCT_Styles_GetByName(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := s.SetNameVal("Normal"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}

	found := styles.GetByName("Normal")
	if found == nil {
		t.Fatal("expected to find style by name")
	}
	if found.StyleId() != "Normal" {
		t.Errorf("expected styleId 'Normal', got %q", found.StyleId())
	}
}

func TestCT_Styles_DefaultFor(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	xmlType, err := enum.WdStyleTypeParagraph.ToXml()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if err := s.SetType(xmlType); err != nil {
		t.Fatalf("SetType: %v", err)
	}
	if err := s.SetDefault(true); err != nil {
		t.Fatalf("SetDefault: %v", err)
	}

	def := styles.DefaultFor(enum.WdStyleTypeParagraph)
	if def == nil {
		t.Fatal("expected default style")
	}
	if def.StyleId() != "Normal" {
		t.Errorf("expected Normal, got %q", def.StyleId())
	}

	// No default for character
	if styles.DefaultFor(enum.WdStyleTypeCharacter) != nil {
		t.Error("expected nil for character type")
	}
}

func TestCT_Styles_AddStyleOfType(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s, err := styles.AddStyleOfType("My Custom Style", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}

	if s.StyleId() != "MyCustomStyle" {
		t.Errorf("expected styleId MyCustomStyle, got %q", s.StyleId())
	}
	if s.NameVal() != "My Custom Style" {
		t.Errorf("expected name 'My Custom Style', got %q", s.NameVal())
	}
	if s.Type() != "paragraph" {
		t.Errorf("expected type paragraph, got %q", s.Type())
	}
	if !s.CustomStyle() {
		t.Error("expected customStyle=true for non-builtin")
	}

	// Builtin
	b, err := styles.AddStyleOfType("Heading 1", enum.WdStyleTypeParagraph, true)
	if err != nil {
		t.Fatal(err)
	}
	if b.CustomStyle() {
		t.Error("expected customStyle=false for builtin")
	}
	if b.StyleId() != "Heading1" {
		t.Errorf("expected Heading1, got %q", b.StyleId())
	}
}

func TestCT_Style_NameVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.NameVal() != "" {
		t.Errorf("expected empty, got %q", s.NameVal())
	}
	if err := s.SetNameVal("Normal"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}
	if s.NameVal() != "Normal" {
		t.Errorf("expected Normal, got %q", s.NameVal())
	}
	if err := s.SetNameVal(""); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}
	if s.NameVal() != "" {
		t.Errorf("expected empty after clear, got %q", s.NameVal())
	}
}

func TestCT_Style_BasedOnVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.BasedOnVal() != "" {
		t.Errorf("expected empty, got %q", s.BasedOnVal())
	}
	if err := s.SetBasedOnVal("Normal"); err != nil {
		t.Fatalf("SetBasedOnVal: %v", err)
	}
	if s.BasedOnVal() != "Normal" {
		t.Errorf("expected Normal, got %q", s.BasedOnVal())
	}
}

func TestCT_Style_NextVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetNextVal("Normal"); err != nil {
		t.Fatalf("SetNextVal: %v", err)
	}
	if s.NextVal() != "Normal" {
		t.Errorf("expected Normal, got %q", s.NextVal())
	}
	if err := s.SetNextVal(""); err != nil {
		t.Fatalf("SetNextVal: %v", err)
	}
	if s.NextVal() != "" {
		t.Errorf("expected empty, got %q", s.NextVal())
	}
}

func TestCT_Style_LockedVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.LockedVal() {
		t.Error("expected false by default")
	}
	if err := s.SetLockedVal(true); err != nil {
		t.Fatalf("SetLockedVal: %v", err)
	}
	if !s.LockedVal() {
		t.Error("expected true")
	}
	if err := s.SetLockedVal(false); err != nil {
		t.Fatalf("SetLockedVal: %v", err)
	}
	if s.LockedVal() {
		t.Error("expected false after clear")
	}
}

func TestCT_Style_SemiHiddenVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.SemiHiddenVal() {
		t.Error("expected false")
	}
	if err := s.SetSemiHiddenVal(true); err != nil {
		t.Fatalf("SetSemiHiddenVal: %v", err)
	}
	if !s.SemiHiddenVal() {
		t.Error("expected true")
	}
}

func TestCT_Style_QFormatVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	s.SetQFormatVal(true)
	if !s.QFormatVal() {
		t.Error("expected true")
	}
	s.SetQFormatVal(false)
	if s.QFormatVal() {
		t.Error("expected false")
	}
}

func TestCT_Style_UiPriorityVal_RoundTrip(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if s.UiPriorityVal() != nil {
		t.Error("expected nil")
	}
	v := 99
	if err := s.SetUiPriorityVal(&v); err != nil {
		t.Fatalf("SetUiPriorityVal: %v", err)
	}
	got := s.UiPriorityVal()
	if got == nil || *got != 99 {
		t.Errorf("expected 99, got %v", got)
	}
	if err := s.SetUiPriorityVal(nil); err != nil {
		t.Fatalf("SetUiPriorityVal: %v", err)
	}
	if s.UiPriorityVal() != nil {
		t.Error("expected nil after clear")
	}
}

func TestCT_Style_BaseStyle(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	normal := styles.AddStyle()
	if err := normal.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := normal.SetNameVal("Normal"); err != nil {
		t.Fatalf("SetNameVal: %v", err)
	}

	heading := styles.AddStyle()
	if err := heading.SetStyleId("Heading1"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := heading.SetBasedOnVal("Normal"); err != nil {
		t.Fatalf("SetBasedOnVal: %v", err)
	}

	base := heading.BaseStyle()
	if base == nil {
		t.Fatal("expected base style")
	}
	if base.StyleId() != "Normal" {
		t.Errorf("expected Normal, got %q", base.StyleId())
	}
}

func TestCT_Style_NextStyle(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	normal := styles.AddStyle()
	if err := normal.SetStyleId("Normal"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}

	heading := styles.AddStyle()
	if err := heading.SetStyleId("Heading1"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if err := heading.SetNextVal("Normal"); err != nil {
		t.Fatalf("SetNextVal: %v", err)
	}

	next := heading.NextStyle()
	if next == nil {
		t.Fatal("expected next style")
	}
	if next.StyleId() != "Normal" {
		t.Errorf("expected Normal, got %q", next.StyleId())
	}
}

func TestCT_Style_Delete(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s := styles.AddStyle()
	if err := s.SetStyleId("ToDelete"); err != nil {
		t.Fatalf("SetStyleId: %v", err)
	}
	if styles.GetByID("ToDelete") == nil {
		t.Fatal("style should exist before delete")
	}
	s.Delete()
	if styles.GetByID("ToDelete") != nil {
		t.Error("style should be removed after delete")
	}
}

func TestCT_Style_IsBuiltin(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s, err := styles.AddStyleOfType("Normal", enum.WdStyleTypeParagraph, true)
	if err != nil {
		t.Fatal(err)
	}
	if !s.IsBuiltin() {
		t.Error("expected builtin")
	}
	custom, err := styles.AddStyleOfType("My Style", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}
	if custom.IsBuiltin() {
		t.Error("expected not builtin")
	}
}

func TestCT_LatentStyles_GetByName(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()
	exc := ls.AddLsdException()
	if err := exc.SetName("Heading 1"); err != nil {
		t.Fatalf("SetName: %v", err)
	}
	if err := exc.SetUiPriority(9); err != nil {
		t.Fatalf("SetUiPriority: %v", err)
	}

	found := ls.GetByName("Heading 1")
	if found == nil {
		t.Fatal("expected to find exception")
	}
	if found.UiPriority() != 9 {
		t.Errorf("expected priority 9, got %d", found.UiPriority())
	}
	if ls.GetByName("NoSuch") != nil {
		t.Error("expected nil for unknown")
	}
}

func TestCT_LsdException_Delete(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()
	exc := ls.AddLsdException()
	if err := exc.SetName("ToRemove"); err != nil {
		t.Fatalf("SetName: %v", err)
	}
	if ls.GetByName("ToRemove") == nil {
		t.Fatal("should exist")
	}
	exc.Delete()
	if ls.GetByName("ToRemove") != nil {
		t.Error("should be removed")
	}
}

func TestCT_LsdException_OnOffProp(t *testing.T) {
	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()
	exc := ls.AddLsdException()
	if err := exc.SetName("Test"); err != nil {
		t.Fatalf("SetName: %v", err)
	}

	// nil by default (attr not set)
	if exc.OnOffProp("w:locked") != nil {
		// Note: SetLocked in generated code defaults to false removal
		// OnOffProp reads raw attr
	}
	tr := true
	if err := exc.SetOnOffProp("w:locked", &tr); err != nil {
		t.Fatalf("SetOnOffProp: %v", err)
	}
	got := exc.OnOffProp("w:locked")
	if got == nil || !*got {
		t.Error("expected locked=true")
	}
	if err := exc.SetOnOffProp("w:locked", nil); err != nil {
		t.Fatalf("SetOnOffProp: %v", err)
	}
	if exc.OnOffProp("w:locked") != nil {
		t.Error("expected nil after removal")
	}
}
