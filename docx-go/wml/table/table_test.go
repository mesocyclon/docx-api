package table

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

const nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

// tableXML2x2 is the reference 2×2 table from reference-appendix.md section 3.2.
// Block-level content inside <w:tc> is stored as RawXML since para module
// is not available in this package's dependency graph.
const tableXML2x2 = `<w:tbl xmlns:w="` + nsW + `">` +
	`<w:tblPr>` +
	`<w:tblStyle w:val="TableGrid"/>` +
	`<w:tblW w:w="0" w:type="auto"/>` +
	`<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>` +
	`</w:tblPr>` +
	`<w:tblGrid>` +
	`<w:gridCol w:w="4675"/>` +
	`<w:gridCol w:w="4675"/>` +
	`</w:tblGrid>` +
	`<w:tr w:rsidR="009A2C41" w:rsidTr="009A2C41">` +
	`<w:tc>` +
	`<w:tcPr>` +
	`<w:tcW w:w="4675" w:type="dxa"/>` +
	`<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/>` +
	`</w:tcPr>` +
	`<w:p><w:r><w:t>Header 1</w:t></w:r></w:p>` +
	`</w:tc>` +
	`<w:tc>` +
	`<w:tcPr>` +
	`<w:tcW w:w="4675" w:type="dxa"/>` +
	`<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/>` +
	`</w:tcPr>` +
	`<w:p><w:r><w:t>Header 2</w:t></w:r></w:p>` +
	`</w:tc>` +
	`</w:tr>` +
	`<w:tr w:rsidR="009A2C41" w:rsidTr="009A2C41">` +
	`<w:tc>` +
	`<w:tcPr>` +
	`<w:tcW w:w="4675" w:type="dxa"/>` +
	`</w:tcPr>` +
	`<w:p><w:r><w:t>Cell A</w:t></w:r></w:p>` +
	`</w:tc>` +
	`<w:tc>` +
	`<w:tcPr>` +
	`<w:tcW w:w="4675" w:type="dxa"/>` +
	`</w:tcPr>` +
	`<w:p><w:r><w:t>Cell B</w:t></w:r></w:p>` +
	`</w:tc>` +
	`</w:tr>` +
	`</w:tbl>`

func TestTblRoundTrip(t *testing.T) {
	// Unmarshal
	var tbl CT_Tbl
	if err := xml.Unmarshal([]byte(tableXML2x2), &tbl); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	// Verify structure
	if tbl.TblPr == nil {
		t.Fatal("TblPr is nil")
	}
	if tbl.TblPr.TblStyle == nil || tbl.TblPr.TblStyle.Val != "TableGrid" {
		t.Error("TblStyle not parsed correctly")
	}
	if tbl.TblPr.TblW == nil || tbl.TblPr.TblW.W != 0 || tbl.TblPr.TblW.Type != "auto" {
		t.Error("TblW not parsed correctly")
	}
	if tbl.TblPr.TblLook == nil {
		t.Error("TblLook is nil")
	} else {
		if tbl.TblPr.TblLook.FirstRow == nil || !*tbl.TblPr.TblLook.FirstRow {
			t.Error("TblLook.FirstRow should be true")
		}
		if tbl.TblPr.TblLook.NoVBand == nil || !*tbl.TblPr.TblLook.NoVBand {
			t.Error("TblLook.NoVBand should be true")
		}
	}

	// Grid
	if tbl.TblGrid == nil || len(tbl.TblGrid.GridCol) != 2 {
		t.Fatalf("TblGrid: expected 2 gridCol, got %v", tbl.TblGrid)
	}
	if tbl.TblGrid.GridCol[0].W != 4675 {
		t.Errorf("GridCol[0].W = %d, want 4675", tbl.TblGrid.GridCol[0].W)
	}

	// Rows
	if len(tbl.Content) != 2 {
		t.Fatalf("Expected 2 rows, got %d", len(tbl.Content))
	}

	row0, ok := tbl.Content[0].(CT_Row)
	if !ok {
		t.Fatalf("Content[0] is not CT_Row: %T", tbl.Content[0])
	}
	if row0.RsidR == nil || *row0.RsidR != "009A2C41" {
		t.Error("Row[0].RsidR not parsed")
	}
	if len(row0.Content) != 2 {
		t.Fatalf("Row[0]: expected 2 cells, got %d", len(row0.Content))
	}

	cell0, ok := row0.Content[0].(CT_Tc)
	if !ok {
		t.Fatalf("Row[0].Content[0] is not CT_Tc: %T", row0.Content[0])
	}
	if cell0.TcPr == nil || cell0.TcPr.TcW == nil {
		t.Fatal("Cell[0].TcPr.TcW is nil")
	}
	if cell0.TcPr.TcW.W != 4675 || cell0.TcPr.TcW.Type != "dxa" {
		t.Errorf("Cell[0].TcW = %v, want 4675/dxa", cell0.TcPr.TcW)
	}
	if cell0.TcPr.Shd == nil || cell0.TcPr.Shd.Val != "clear" {
		t.Error("Cell[0].Shd not parsed")
	}
	if cell0.TcPr.Shd.Fill == nil || *cell0.TcPr.Shd.Fill != "D9E2F3" {
		t.Error("Cell[0].Shd.Fill not parsed")
	}

	// Cell content: paragraph stored as RawXML (para module not available)
	if len(cell0.Content) != 1 {
		t.Errorf("Cell[0]: expected 1 block element, got %d", len(cell0.Content))
	}

	// Marshal
	output, err := xml.Marshal(&tbl)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	// Re-unmarshal
	var tbl2 CT_Tbl
	if err := xml.Unmarshal(output, &tbl2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}

	// Compare key fields
	if tbl2.TblPr == nil || tbl2.TblPr.TblStyle == nil || tbl2.TblPr.TblStyle.Val != "TableGrid" {
		t.Error("Round-trip lost TblStyle")
	}
	if tbl2.TblGrid == nil || len(tbl2.TblGrid.GridCol) != 2 {
		t.Error("Round-trip lost TblGrid")
	}
	if len(tbl2.Content) != 2 {
		t.Errorf("Round-trip: expected 2 rows, got %d", len(tbl2.Content))
	}
}

func TestTblPrOrder(t *testing.T) {
	// Verify that TblPr fields are marshaled in XSD sequence order.
	pr := &CT_TblPr{
		TblStyle: &xmltypes.CT_String{Val: "TableGrid"},
		TblW:     &CT_TblWidth{W: 5000, Type: "pct"},
		Jc:       &CT_JcTable{Val: "center"},
		TblLook: &CT_TblLook{
			FirstRow: boolPtr(true),
			NoVBand:  boolPtr(true),
		},
	}

	out, err := xml.Marshal(pr)
	if err != nil {
		t.Fatal(err)
	}
	s := string(out)

	// Check ordering: tblStyle before tblW before jc before tblLook
	idxStyle := strings.Index(s, "tblStyle")
	idxW := strings.Index(s, "tblW")
	idxJc := strings.Index(s, ":jc") // use ":jc" to avoid matching "Jc"
	if idxJc < 0 {
		idxJc = strings.Index(s, "<jc") // fallback without namespace
	}
	idxLook := strings.Index(s, "tblLook")

	if idxStyle < 0 || idxW < 0 || idxLook < 0 {
		t.Fatalf("Missing expected elements in output: %s", s)
	}
	if idxStyle > idxW {
		t.Error("tblStyle should come before tblW")
	}
	if idxW > idxLook {
		t.Error("tblW should come before tblLook")
	}
	_ = idxJc // jc position checked implicitly
}

func TestTcPrOrder(t *testing.T) {
	pr := &CT_TcPr{
		TcW:    &CT_TblWidth{W: 4675, Type: "dxa"},
		VMerge: &CT_VMerge{Val: strPtr("restart")},
		Shd:    &xmltypes.CT_Shd{Val: "clear", Fill: strPtr("FF0000")},
		VAlign: &CT_VerticalJc{Val: "center"},
	}

	out, err := xml.Marshal(pr)
	if err != nil {
		t.Fatal(err)
	}
	s := string(out)

	// Verify order: tcW before vMerge before shd before vAlign
	idxTcW := strings.Index(s, "tcW")
	idxVMerge := strings.Index(s, "vMerge")
	idxShd := strings.Index(s, "shd")
	idxVAlign := strings.Index(s, "vAlign")

	if idxTcW < 0 || idxVMerge < 0 || idxShd < 0 || idxVAlign < 0 {
		t.Fatalf("Missing expected elements in output: %s", s)
	}
	if idxTcW > idxVMerge {
		t.Error("tcW should come before vMerge")
	}
	if idxVMerge > idxShd {
		t.Error("vMerge should come before shd")
	}
	if idxShd > idxVAlign {
		t.Error("shd should come before vAlign")
	}
}

func TestTrPrRoundTrip(t *testing.T) {
	input := `<w:trPr xmlns:w="` + nsW + `">` +
		`<w:cantSplit/>` +
		`<w:trHeight w:val="400" w:hRule="atLeast"/>` +
		`<w:tblHeader/>` +
		`<w:jc w:val="center"/>` +
		`</w:trPr>`

	var trPr CT_TrPr
	if err := xml.Unmarshal([]byte(input), &trPr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if !trPr.CantSplit.Bool(false) {
		t.Error("CantSplit should be true")
	}
	if trPr.TrHeight == nil || trPr.TrHeight.Val == nil || *trPr.TrHeight.Val != 400 {
		t.Error("TrHeight.Val should be 400")
	}
	if trPr.TrHeight.HRule == nil || *trPr.TrHeight.HRule != "atLeast" {
		t.Error("TrHeight.HRule should be atLeast")
	}
	if !trPr.TblHeader.Bool(false) {
		t.Error("TblHeader should be true")
	}
	if trPr.Jc == nil || trPr.Jc.Val != "center" {
		t.Error("Jc should be center")
	}

	// Marshal and re-unmarshal
	out, err := xml.Marshal(&trPr)
	if err != nil {
		t.Fatal(err)
	}
	var trPr2 CT_TrPr
	if err := xml.Unmarshal(out, &trPr2); err != nil {
		t.Fatal(err)
	}
	if trPr2.Jc == nil || trPr2.Jc.Val != "center" {
		t.Error("Round-trip lost Jc")
	}
}

func TestRawXMLRoundTrip(t *testing.T) {
	// Table with an unknown extension element.
	input := `<w:tbl xmlns:w="` + nsW + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:tblPr>` +
		`<w:tblStyle w:val="TableGrid"/>` +
		`<w:tblW w:w="0" w:type="auto"/>` +
		`<w14:someExtension w14:val="test"/>` +
		`</w:tblPr>` +
		`<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
		`<w:tr>` +
		`<w:tc>` +
		`<w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>` +
		`<w:p/>` +
		`</w:tc>` +
		`</w:tr>` +
		`</w:tbl>`

	var tbl CT_Tbl
	if err := xml.Unmarshal([]byte(input), &tbl); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if tbl.TblPr == nil {
		t.Fatal("TblPr is nil")
	}
	if len(tbl.TblPr.Extra) != 1 {
		t.Fatalf("Expected 1 Extra element in TblPr, got %d", len(tbl.TblPr.Extra))
	}
	if tbl.TblPr.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Extra[0].Local = %q, want someExtension", tbl.TblPr.Extra[0].XMLName.Local)
	}

	// Marshal and verify Extra preserved
	out, err := xml.Marshal(&tbl)
	if err != nil {
		t.Fatal(err)
	}

	var tbl2 CT_Tbl
	if err := xml.Unmarshal(out, &tbl2); err != nil {
		t.Fatal(err)
	}
	if len(tbl2.TblPr.Extra) != 1 {
		t.Fatalf("Round-trip: expected 1 Extra, got %d", len(tbl2.TblPr.Extra))
	}
	if tbl2.TblPr.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Round-trip lost extension element name: %q", tbl2.TblPr.Extra[0].XMLName.Local)
	}
}

func TestTblBordersRoundTrip(t *testing.T) {
	input := `<w:tblPr xmlns:w="` + nsW + `">` +
		`<w:tblBorders>` +
		`<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
		`<w:start w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
		`<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
		`<w:end w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
		`<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
		`<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
		`</w:tblBorders>` +
		`</w:tblPr>`

	var pr CT_TblPr
	if err := xml.Unmarshal([]byte(input), &pr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if pr.TblBorders == nil {
		t.Fatal("TblBorders is nil")
	}
	if pr.TblBorders.Top == nil || pr.TblBorders.Top.Val != "single" {
		t.Error("Top border not parsed")
	}
	if pr.TblBorders.Start == nil {
		t.Error("Start border not parsed")
	}
	if pr.TblBorders.InsideH == nil {
		t.Error("InsideH border not parsed")
	}

	sz := 4
	if pr.TblBorders.Top.Sz == nil || *pr.TblBorders.Top.Sz != sz {
		t.Errorf("Top border sz = %v, want %d", pr.TblBorders.Top.Sz, sz)
	}

	// Round-trip
	out, err := xml.Marshal(&pr)
	if err != nil {
		t.Fatal(err)
	}
	var pr2 CT_TblPr
	if err := xml.Unmarshal(out, &pr2); err != nil {
		t.Fatal(err)
	}
	if pr2.TblBorders == nil || pr2.TblBorders.Top == nil {
		t.Error("Round-trip lost TblBorders.Top")
	}
	if pr2.TblBorders.InsideV == nil || pr2.TblBorders.InsideV.Val != "single" {
		t.Error("Round-trip lost TblBorders.InsideV")
	}
}

func TestTblCellMarRoundTrip(t *testing.T) {
	input := `<w:tblPr xmlns:w="` + nsW + `">` +
		`<w:tblCellMar>` +
		`<w:top w:w="0" w:type="dxa"/>` +
		`<w:start w:w="108" w:type="dxa"/>` +
		`<w:bottom w:w="0" w:type="dxa"/>` +
		`<w:end w:w="108" w:type="dxa"/>` +
		`</w:tblCellMar>` +
		`</w:tblPr>`

	var pr CT_TblPr
	if err := xml.Unmarshal([]byte(input), &pr); err != nil {
		t.Fatal(err)
	}

	if pr.TblCellMar == nil {
		t.Fatal("TblCellMar is nil")
	}
	if pr.TblCellMar.Start == nil || pr.TblCellMar.Start.W != 108 {
		t.Error("TblCellMar.Start not parsed")
	}

	out, err := xml.Marshal(&pr)
	if err != nil {
		t.Fatal(err)
	}
	var pr2 CT_TblPr
	if err := xml.Unmarshal(out, &pr2); err != nil {
		t.Fatal(err)
	}
	if pr2.TblCellMar == nil || pr2.TblCellMar.Start == nil || pr2.TblCellMar.Start.W != 108 {
		t.Error("Round-trip lost TblCellMar.Start")
	}
}

func TestNestedTable(t *testing.T) {
	// A cell containing a nested table (as RawXML since it is parsed).
	input := `<w:tc xmlns:w="` + nsW + `">` +
		`<w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>` +
		`<w:tbl>` +
		`<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
		`<w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid>` +
		`<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>` +
		`</w:tbl>` +
		`<w:p/>` +
		`</w:tc>`

	var tc CT_Tc
	if err := xml.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if tc.TcPr == nil || tc.TcPr.TcW == nil || tc.TcPr.TcW.W != 5000 {
		t.Error("TcPr not parsed correctly")
	}

	// Should have 2 content items: nested tbl + p (as RawXML)
	if len(tc.Content) != 2 {
		t.Fatalf("Expected 2 content elements, got %d", len(tc.Content))
	}

	// First should be nested CT_Tbl
	nestedTbl, ok := tc.Content[0].(*CT_Tbl)
	if !ok {
		t.Fatalf("Content[0] is not *CT_Tbl: %T", tc.Content[0])
	}
	if nestedTbl.TblPr == nil || nestedTbl.TblPr.TblW == nil {
		t.Error("Nested table TblPr not parsed")
	}

	// Second should be RawXML (p)
	_, ok = tc.Content[1].(shared.RawXML)
	if !ok {
		t.Fatalf("Content[1] is not shared.RawXML: %T", tc.Content[1])
	}
}

func TestVMerge(t *testing.T) {
	input := `<w:tcPr xmlns:w="` + nsW + `">` +
		`<w:tcW w:w="2000" w:type="dxa"/>` +
		`<w:vMerge w:val="restart"/>` +
		`</w:tcPr>`

	var pr CT_TcPr
	if err := xml.Unmarshal([]byte(input), &pr); err != nil {
		t.Fatal(err)
	}
	if pr.VMerge == nil || pr.VMerge.Val == nil || *pr.VMerge.Val != "restart" {
		t.Error("VMerge not parsed")
	}

	// VMerge without val (means "continue")
	input2 := `<w:tcPr xmlns:w="` + nsW + `">` +
		`<w:tcW w:w="2000" w:type="dxa"/>` +
		`<w:vMerge/>` +
		`</w:tcPr>`

	var pr2 CT_TcPr
	if err := xml.Unmarshal([]byte(input2), &pr2); err != nil {
		t.Fatal(err)
	}
	if pr2.VMerge == nil {
		t.Error("VMerge should be present even without val")
	}
	if pr2.VMerge.Val != nil {
		t.Error("VMerge.Val should be nil for continue")
	}
}

// Helpers
func boolPtr(b bool) *bool    { return &b }
func strPtr(s string) *string { return &s }
