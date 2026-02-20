package sectpr

import (
	"encoding/xml"
	"strings"
	"testing"
)

const nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// TestSectPrBasicRoundTrip tests unmarshal → marshal → unmarshal of a minimal sectPr.
func TestSectPrBasicRoundTrip(t *testing.T) {
	input := `<w:sectPr xmlns:w="` + nsW + `" w:rsidR="00000001">` +
		`<w:pgSz w:w="12240" w:h="15840"/>` +
		`<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
		`<w:cols w:space="720"/>` +
		`<w:docGrid w:linePitch="360"/>` +
		`</w:sectPr>`

	var sp CT_SectPr
	if err := xml.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	// Verify parsed values
	if sp.RsidR == nil || *sp.RsidR != "00000001" {
		t.Error("RsidR not parsed correctly")
	}
	if sp.PgSz == nil {
		t.Fatal("PgSz is nil")
	}
	if sp.PgSz.W != 12240 || sp.PgSz.H != 15840 {
		t.Errorf("PgSz: got W=%d H=%d, want W=12240 H=15840", sp.PgSz.W, sp.PgSz.H)
	}
	if sp.PgMar == nil {
		t.Fatal("PgMar is nil")
	}
	if sp.PgMar.Top != 1440 || sp.PgMar.Left != 1440 {
		t.Errorf("PgMar: got top=%d left=%d", sp.PgMar.Top, sp.PgMar.Left)
	}
	if sp.Cols == nil {
		t.Fatal("Cols is nil")
	}
	if sp.Cols.Space == nil || *sp.Cols.Space != 720 {
		t.Error("Cols.Space not parsed")
	}
	if sp.DocGrid == nil || sp.DocGrid.LinePitch == nil || *sp.DocGrid.LinePitch != 360 {
		t.Error("DocGrid.LinePitch not parsed")
	}

	// Marshal
	output, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	// Re-unmarshal
	var sp2 CT_SectPr
	if err := xml.Unmarshal(output, &sp2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}

	// Compare key fields
	if sp2.PgSz.W != sp.PgSz.W || sp2.PgSz.H != sp.PgSz.H {
		t.Error("round-trip lost PgSz")
	}
	if sp2.PgMar.Top != sp.PgMar.Top {
		t.Error("round-trip lost PgMar")
	}
	if *sp2.Cols.Space != *sp.Cols.Space {
		t.Error("round-trip lost Cols")
	}
	if *sp2.DocGrid.LinePitch != *sp.DocGrid.LinePitch {
		t.Error("round-trip lost DocGrid")
	}
}

// TestSectPrWithHeaders tests sectPr with header/footer references.
func TestSectPrWithHeaders(t *testing.T) {
	input := `<w:sectPr xmlns:w="` + nsW + `" xmlns:r="` + nsR + `">` +
		`<w:headerReference w:type="default" r:id="rId8"/>` +
		`<w:headerReference w:type="first" r:id="rId9"/>` +
		`<w:footerReference w:type="default" r:id="rId10"/>` +
		`<w:pgSz w:w="12240" w:h="15840"/>` +
		`<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
		`<w:titlePg/>` +
		`</w:sectPr>`

	var sp CT_SectPr
	if err := xml.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if len(sp.HeaderRefs) != 2 {
		t.Fatalf("Expected 2 header refs, got %d", len(sp.HeaderRefs))
	}
	if sp.HeaderRefs[0].Type != "default" || sp.HeaderRefs[0].RID != "rId8" {
		t.Errorf("HeaderRef[0]: got type=%q rid=%q", sp.HeaderRefs[0].Type, sp.HeaderRefs[0].RID)
	}
	if sp.HeaderRefs[1].Type != "first" || sp.HeaderRefs[1].RID != "rId9" {
		t.Errorf("HeaderRef[1]: got type=%q rid=%q", sp.HeaderRefs[1].Type, sp.HeaderRefs[1].RID)
	}
	if len(sp.FooterRefs) != 1 {
		t.Fatalf("Expected 1 footer ref, got %d", len(sp.FooterRefs))
	}
	if sp.FooterRefs[0].Type != "default" || sp.FooterRefs[0].RID != "rId10" {
		t.Errorf("FooterRef[0]: got type=%q rid=%q", sp.FooterRefs[0].Type, sp.FooterRefs[0].RID)
	}
	if sp.TitlePg == nil {
		t.Error("TitlePg is nil")
	}
	if !sp.TitlePg.Bool(false) {
		t.Error("TitlePg should be true")
	}

	// Marshal and verify round-trip
	output, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	var sp2 CT_SectPr
	if err := xml.Unmarshal(output, &sp2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if len(sp2.HeaderRefs) != 2 {
		t.Errorf("round-trip lost header refs: got %d", len(sp2.HeaderRefs))
	}
	if len(sp2.FooterRefs) != 1 {
		t.Errorf("round-trip lost footer refs: got %d", len(sp2.FooterRefs))
	}
}

// TestSectPrLandscapeColumns tests landscape page with columns.
func TestSectPrLandscapeColumns(t *testing.T) {
	input := `<w:sectPr xmlns:w="` + nsW + `">` +
		`<w:type w:val="continuous"/>` +
		`<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>` +
		`<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
		`<w:cols w:num="2" w:space="720"/>` +
		`</w:sectPr>`

	var sp CT_SectPr
	if err := xml.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if sp.Type == nil || sp.Type.Val != "continuous" {
		t.Error("Type not parsed")
	}
	if sp.PgSz.Orient == nil || *sp.PgSz.Orient != "landscape" {
		t.Error("Orient not parsed")
	}
	if sp.Cols.Num == nil || *sp.Cols.Num != 2 {
		t.Error("Cols.Num not parsed")
	}

	output, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	var sp2 CT_SectPr
	if err := xml.Unmarshal(output, &sp2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if sp2.Type == nil || sp2.Type.Val != "continuous" {
		t.Error("round-trip lost Type")
	}
	if sp2.PgSz.Orient == nil || *sp2.PgSz.Orient != "landscape" {
		t.Error("round-trip lost Orient")
	}
	if sp2.Cols.Num == nil || *sp2.Cols.Num != 2 {
		t.Error("round-trip lost Cols.Num")
	}
}

// TestSectPrRawXMLRoundTrip tests that unknown/extension elements are preserved.
func TestSectPrRawXMLRoundTrip(t *testing.T) {
	input := `<w:sectPr xmlns:w="` + nsW + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:pgSz w:w="12240" w:h="15840"/>` +
		`<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
		`<w:docGrid w:linePitch="360"/>` +
		`<w14:someExtension w14:val="test">` +
		`<w14:inner>data</w14:inner>` +
		`</w14:someExtension>` +
		`</w:sectPr>`

	var sp CT_SectPr
	if err := xml.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if len(sp.Extra) != 1 {
		t.Fatalf("Expected 1 Extra element, got %d", len(sp.Extra))
	}
	if sp.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Extra[0] name: got %q, want %q", sp.Extra[0].XMLName.Local, "someExtension")
	}

	// Marshal
	output, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	// Re-unmarshal
	var sp2 CT_SectPr
	if err := xml.Unmarshal(output, &sp2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}

	if len(sp2.Extra) != 1 {
		t.Fatalf("round-trip: Expected 1 Extra element, got %d", len(sp2.Extra))
	}
	if sp2.Extra[0].XMLName.Local != "someExtension" {
		t.Error("round-trip lost extension element name")
	}
}

// TestSectPrMarshalOrder verifies elements are written in strict xsd:sequence order.
func TestSectPrMarshalOrder(t *testing.T) {
	sp := CT_SectPr{
		DocGrid: &CT_DocGrid{LinePitch: intPtr(360)},
		PgSz:    &CT_PageSz{W: 12240, H: 15840},
		PgMar:   &CT_PageMar{Top: 1440, Right: 1440, Bottom: 1440, Left: 1440, Header: 720, Footer: 720},
		Cols:    &CT_Columns{Space: intPtr(720)},
		Type:    &CT_SectType{Val: "nextPage"},
	}

	output, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	out := string(output)

	// Verify order: type < pgSz < pgMar < cols < docGrid
	typeIdx := strings.Index(out, "type")
	pgSzIdx := strings.Index(out, "pgSz")
	pgMarIdx := strings.Index(out, "pgMar")
	colsIdx := strings.Index(out, "cols")
	docGridIdx := strings.Index(out, "docGrid")

	if typeIdx > pgSzIdx {
		t.Error("type should come before pgSz")
	}
	if pgSzIdx > pgMarIdx {
		t.Error("pgSz should come before pgMar")
	}
	if pgMarIdx > colsIdx {
		t.Error("pgMar should come before cols")
	}
	if colsIdx > docGridIdx {
		t.Error("cols should come before docGrid")
	}
}

// TestSectPrColumnsWithCol tests columns with explicit col definitions.
func TestSectPrColumnsWithCol(t *testing.T) {
	input := `<w:sectPr xmlns:w="` + nsW + `">` +
		`<w:pgSz w:w="12240" w:h="15840"/>` +
		`<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
		`<w:cols w:num="3" w:equalWidth="0">` +
		`<w:col w:w="3000" w:space="720"/>` +
		`<w:col w:w="4800" w:space="720"/>` +
		`<w:col w:w="3000"/>` +
		`</w:cols>` +
		`</w:sectPr>`

	var sp CT_SectPr
	if err := xml.Unmarshal([]byte(input), &sp); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if sp.Cols == nil {
		t.Fatal("Cols is nil")
	}
	if len(sp.Cols.Col) != 3 {
		t.Fatalf("Expected 3 columns, got %d", len(sp.Cols.Col))
	}
	if sp.Cols.Col[0].W == nil || *sp.Cols.Col[0].W != 3000 {
		t.Error("Col[0].W not parsed")
	}
	if sp.Cols.Col[1].Space == nil || *sp.Cols.Col[1].Space != 720 {
		t.Error("Col[1].Space not parsed")
	}
	if sp.Cols.Col[2].Space != nil {
		t.Error("Col[2].Space should be nil")
	}

	// Round-trip
	output, err := xml.Marshal(&sp)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	var sp2 CT_SectPr
	if err := xml.Unmarshal(output, &sp2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if len(sp2.Cols.Col) != 3 {
		t.Errorf("round-trip lost columns: got %d", len(sp2.Cols.Col))
	}
}

func intPtr(v int) *int {
	return &v
}
