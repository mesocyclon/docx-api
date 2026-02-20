package styles

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/xmltypes"
)

// minimalStylesXML is the minimal valid styles.xml from reference-appendix 2.6.
const minimalStylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="w14 w15">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi"
                  w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99"
    w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376"/>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList">
    <w:name w:val="No List"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
</w:styles>`

func TestParseMinimalStyles(t *testing.T) {
	s, err := Parse([]byte(minimalStylesXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// DocDefaults
	if s.DocDefaults == nil {
		t.Fatal("DocDefaults is nil")
	}
	if s.DocDefaults.RPrDefault == nil {
		t.Fatal("RPrDefault is nil")
	}
	if s.DocDefaults.RPrDefault.RPr == nil {
		t.Fatal("RPrDefault.RPr is nil")
	}
	if s.DocDefaults.RPrDefault.RPr.Base.Sz == nil || s.DocDefaults.RPrDefault.RPr.Base.Sz.Val != 24 {
		t.Errorf("RPrDefault sz: got %v, want 24", s.DocDefaults.RPrDefault.RPr.Base.Sz)
	}

	if s.DocDefaults.PPrDefault == nil {
		t.Fatal("PPrDefault is nil")
	}
	if s.DocDefaults.PPrDefault.PPr == nil {
		t.Fatal("PPrDefault.PPr is nil")
	}
	if s.DocDefaults.PPrDefault.PPr.Spacing == nil {
		t.Fatal("PPrDefault spacing is nil")
	}
	if s.DocDefaults.PPrDefault.PPr.Spacing.After == nil || *s.DocDefaults.PPrDefault.PPr.Spacing.After != 160 {
		t.Error("PPrDefault spacing.After != 160")
	}

	// LatentStyles
	if s.LatentStyles == nil {
		t.Fatal("LatentStyles is nil")
	}
	if s.LatentStyles.DefUIPriority == nil || *s.LatentStyles.DefUIPriority != 99 {
		t.Error("LatentStyles defUIPriority != 99")
	}
	if s.LatentStyles.Count == nil || *s.LatentStyles.Count != 376 {
		t.Error("LatentStyles count != 376")
	}

	// Styles
	if len(s.Style) != 4 {
		t.Fatalf("Expected 4 styles, got %d", len(s.Style))
	}

	// Normal
	normal := s.Style[0]
	if normal.StyleID != "Normal" {
		t.Errorf("Style 0 id: got %q, want Normal", normal.StyleID)
	}
	if normal.Type != "paragraph" {
		t.Errorf("Style 0 type: got %q, want paragraph", normal.Type)
	}
	if normal.Default == nil || !*normal.Default {
		t.Error("Style 0 default should be true")
	}
	if normal.Name == nil || normal.Name.Val != "Normal" {
		t.Error("Style 0 name != Normal")
	}
	if normal.QFormat == nil {
		t.Error("Style 0 qFormat should be set")
	}

	// TableNormal — should have tblPr
	tableNormal := s.Style[2]
	if tableNormal.StyleID != "TableNormal" {
		t.Errorf("Style 2 id: got %q, want TableNormal", tableNormal.StyleID)
	}
	if tableNormal.TblPr == nil {
		t.Error("TableNormal should have TblPr")
	}
}

func TestRoundTripMinimalStyles(t *testing.T) {
	// Parse
	s, err := Parse([]byte(minimalStylesXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// Serialize
	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// Re-parse
	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\nOutput:\n%s", err, string(out))
	}

	// Compare key fields
	if len(s2.Style) != len(s.Style) {
		t.Fatalf("Round-trip lost styles: %d → %d", len(s.Style), len(s2.Style))
	}

	for i, st := range s.Style {
		st2 := s2.Style[i]
		if st.StyleID != st2.StyleID {
			t.Errorf("Style %d: styleId %q → %q", i, st.StyleID, st2.StyleID)
		}
		if st.Type != st2.Type {
			t.Errorf("Style %d: type %q → %q", i, st.Type, st2.Type)
		}
		if (st.Name == nil) != (st2.Name == nil) {
			t.Errorf("Style %d: name presence changed", i)
		} else if st.Name != nil && st.Name.Val != st2.Name.Val {
			t.Errorf("Style %d: name %q → %q", i, st.Name.Val, st2.Name.Val)
		}
	}

	// DocDefaults round-trip
	if s2.DocDefaults == nil {
		t.Fatal("Round-trip lost DocDefaults")
	}
	if s2.DocDefaults.RPrDefault == nil || s2.DocDefaults.RPrDefault.RPr == nil {
		t.Fatal("Round-trip lost RPrDefault")
	}
	if s2.DocDefaults.RPrDefault.RPr.Base.Sz == nil || s2.DocDefaults.RPrDefault.RPr.Base.Sz.Val != 24 {
		t.Error("Round-trip lost RPrDefault sz")
	}

	// LatentStyles round-trip
	if s2.LatentStyles == nil {
		t.Fatal("Round-trip lost LatentStyles")
	}
	if s2.LatentStyles.Count == nil || *s2.LatentStyles.Count != 376 {
		t.Error("Round-trip lost LatentStyles count")
	}
}

// stylesWithExtensionXML includes an unknown w14 extension element to test
// RawXML round-trip preservation.
const stylesWithExtensionXML = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="240" w:after="0"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
    <w14:someExtension w14:val="test"/>
  </w:style>
</w:styles>`

func TestRoundTripWithExtension(t *testing.T) {
	s, err := Parse([]byte(stylesWithExtensionXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(s.Style) != 1 {
		t.Fatalf("Expected 1 style, got %d", len(s.Style))
	}

	st := s.Style[0]
	if st.StyleID != "Heading1" {
		t.Errorf("StyleID: %q, want Heading1", st.StyleID)
	}
	if st.Name == nil || st.Name.Val != "heading 1" {
		t.Error("Name != 'heading 1'")
	}
	if st.BasedOn == nil || st.BasedOn.Val != "Normal" {
		t.Error("BasedOn != Normal")
	}
	if st.Next == nil || st.Next.Val != "Normal" {
		t.Error("Next != Normal")
	}
	if st.Link == nil || st.Link.Val != "Heading1Char" {
		t.Error("Link != Heading1Char")
	}
	if st.UIpriority == nil || st.UIpriority.Val != 9 {
		t.Error("UIpriority != 9")
	}

	// PPr check
	if st.PPr == nil {
		t.Fatal("PPr is nil")
	}
	if !st.PPr.KeepNext.Bool(false) {
		t.Error("KeepNext should be true")
	}
	if st.PPr.OutlineLvl == nil || st.PPr.OutlineLvl.Val != 0 {
		t.Error("OutlineLvl != 0")
	}

	// RPr check
	if st.RPr == nil {
		t.Fatal("RPr is nil")
	}
	if st.RPr.Sz == nil || st.RPr.Sz.Val != 32 {
		t.Errorf("RPr Sz: %v, want 32", st.RPr.Sz)
	}
	if st.RPr.Color == nil || st.RPr.Color.Val != "2F5496" {
		t.Error("RPr Color != 2F5496")
	}

	// RawXML extension preserved
	if len(st.Extra) != 1 {
		t.Fatalf("Expected 1 Extra element, got %d", len(st.Extra))
	}
	if st.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Extra[0] local = %q, want someExtension", st.Extra[0].XMLName.Local)
	}

	// Serialize and re-parse
	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\nOutput:\n%s", err, string(out))
	}

	// Verify extension survives round-trip
	if len(s2.Style) != 1 {
		t.Fatalf("Round-trip: expected 1 style, got %d", len(s2.Style))
	}
	st2 := s2.Style[0]
	if len(st2.Extra) != 1 {
		t.Fatalf("Round-trip lost Extra: got %d", len(st2.Extra))
	}
	if st2.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Round-trip: Extra[0] local = %q", st2.Extra[0].XMLName.Local)
	}
	if st2.PPr == nil || !st2.PPr.KeepNext.Bool(false) {
		t.Error("Round-trip lost PPr.KeepNext")
	}
	if st2.RPr == nil || st2.RPr.Sz == nil || st2.RPr.Sz.Val != 32 {
		t.Error("Round-trip lost RPr.Sz")
	}
}

// tableStyleXML tests a table style with tblStylePr conditional formatting.
const tableStyleXML = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="39"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
      <w:rPr>
        <w:b/>
      </w:rPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastRow">
      <w:rPr>
        <w:b/>
      </w:rPr>
    </w:tblStylePr>
  </w:style>
</w:styles>`

func TestParseTableStyle(t *testing.T) {
	s, err := Parse([]byte(tableStyleXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(s.Style) != 1 {
		t.Fatalf("Expected 1 style, got %d", len(s.Style))
	}

	st := s.Style[0]
	if st.Type != "table" {
		t.Errorf("type: %q, want table", st.Type)
	}
	if st.TblPr == nil {
		t.Fatal("TblPr is nil")
	}

	// TblStylePr
	if len(st.TblStylePr) != 2 {
		t.Fatalf("Expected 2 tblStylePr, got %d", len(st.TblStylePr))
	}
	if st.TblStylePr[0].Type != "firstRow" {
		t.Errorf("tblStylePr[0] type: %q, want firstRow", st.TblStylePr[0].Type)
	}
	if st.TblStylePr[0].RPr == nil || st.TblStylePr[0].RPr.B == nil {
		t.Error("firstRow RPr.B is nil")
	}
	if st.TblStylePr[1].Type != "lastRow" {
		t.Errorf("tblStylePr[1] type: %q, want lastRow", st.TblStylePr[1].Type)
	}
}

func TestRoundTripTableStyle(t *testing.T) {
	s, err := Parse([]byte(tableStyleXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\nOutput:\n%s", err, string(out))
	}

	if len(s2.Style) != 1 {
		t.Fatalf("Round-trip: expected 1 style, got %d", len(s2.Style))
	}
	st := s2.Style[0]
	if st.StyleID != "TableGrid" {
		t.Error("StyleID != TableGrid")
	}
	if len(st.TblStylePr) != 2 {
		t.Fatalf("Round-trip: expected 2 tblStylePr, got %d", len(st.TblStylePr))
	}
	if st.TblStylePr[0].Type != "firstRow" || st.TblStylePr[0].RPr == nil {
		t.Error("Round-trip lost firstRow tblStylePr")
	}
}

// TestExtraOnStylesLevel verifies unknown elements at the root <w:styles> level
// survive round-trip.
func TestExtraOnStylesLevel(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	                 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w14:defaultImageDpi w14:val="32767"/>
</w:styles>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(s.Extra) != 1 {
		t.Fatalf("Expected 1 Extra at styles level, got %d", len(s.Extra))
	}
	if s.Extra[0].XMLName.Local != "defaultImageDpi" {
		t.Errorf("Extra local: %q", s.Extra[0].XMLName.Local)
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v", err)
	}
	if len(s2.Extra) != 1 {
		t.Fatalf("Round-trip lost Extra: got %d", len(s2.Extra))
	}
	if s2.Extra[0].XMLName.Local != "defaultImageDpi" {
		t.Errorf("Round-trip Extra local: %q", s2.Extra[0].XMLName.Local)
	}
}

// TestSerializeNewStyles verifies that creating a new CT_Styles from scratch
// produces valid XML.
func TestSerializeNewStyles(t *testing.T) {
	tr := true
	s := &CT_Styles{
		Style: []CT_Style{
			{
				Type:    "paragraph",
				Default: &tr,
				StyleID: "Normal",
				Name:    &xmltypes.CT_String{Val: "Normal"},
				QFormat: xmltypes.NewOnOff(true),
			},
		},
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	xmlStr := string(out)
	if !strings.Contains(xmlStr, `w:styles`) {
		t.Error("Output missing w:styles element")
	}
	if !strings.Contains(xmlStr, `styleId`) {
		t.Error("Output missing styleId attribute")
	}

	// Verify it can be re-parsed
	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse of new styles failed: %v\n%s", err, xmlStr)
	}
	if len(s2.Style) != 1 || s2.Style[0].StyleID != "Normal" {
		t.Errorf("Re-parsed: %d styles", len(s2.Style))
	}
}

// TestNamespacePreservation verifies that namespace declarations on the root
// element survive round-trip.
func TestNamespacePreservation(t *testing.T) {
	s, err := Parse([]byte(minimalStylesXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(s.Namespaces) == 0 {
		t.Fatal("Expected preserved namespaces, got none")
	}

	// Check we preserved xmlns:w
	found := false
	for _, attr := range s.Namespaces {
		if attr.Value == xmltypes.NSw {
			found = true
			break
		}
	}
	if !found {
		t.Error("xmlns:w not preserved in Namespaces")
	}

	// Serialize and check output contains the namespace
	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	if !strings.Contains(string(out), xmltypes.NSw) {
		t.Error("Output missing w namespace URI")
	}
}

// TestEmptyStylesRoundTrip tests a minimal empty styles document.
func TestEmptyStylesRoundTrip(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(s.Style) != 0 {
		t.Errorf("Expected 0 styles, got %d", len(s.Style))
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\n%s", err, string(out))
	}

	_ = s2
}

// TestLatentStylesExceptions verifies that lsdException children parse correctly.
func TestLatentStylesExceptions(t *testing.T) {
	input := `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:count="376">
    <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
    <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
    <w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
  </w:latentStyles>
</w:styles>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	ls := s.LatentStyles
	if ls == nil {
		t.Fatal("LatentStyles is nil")
	}
	if len(ls.LsdException) != 3 {
		t.Fatalf("Expected 3 exceptions, got %d", len(ls.LsdException))
	}
	if ls.LsdException[0].Name != "Normal" {
		t.Errorf("exc[0] name: %q", ls.LsdException[0].Name)
	}
	if ls.LsdException[2].SemiHidden == nil || !*ls.LsdException[2].SemiHidden {
		t.Error("heading 2 semiHidden should be true")
	}

	// Round-trip
	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse: %v", err)
	}
	if len(s2.LatentStyles.LsdException) != 3 {
		t.Errorf("Round-trip: expected 3 exceptions, got %d", len(s2.LatentStyles.LsdException))
	}
}

// TestXMLMarshalDirect tests xml.Marshal on CT_Styles directly.
func TestXMLMarshalDirect(t *testing.T) {
	tr := true
	s := CT_Styles{
		LatentStyles: &CT_LatentStyles{
			DefUIPriority: intPtr(99),
			Count:         intPtr(376),
		},
		Style: []CT_Style{
			{
				Type:    "paragraph",
				Default: &tr,
				StyleID: "Normal",
				Name:    &xmltypes.CT_String{Val: "Normal"},
			},
		},
	}

	out, err := xml.MarshalIndent(&s, "", "  ")
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}

	var s2 CT_Styles
	if err := xml.Unmarshal(out, &s2); err != nil {
		t.Fatalf("Unmarshal: %v\n%s", err, string(out))
	}

	if len(s2.Style) != 1 || s2.Style[0].StyleID != "Normal" {
		t.Errorf("Unmarshal mismatch: %d styles", len(s2.Style))
	}
	if s2.LatentStyles == nil || s2.LatentStyles.Count == nil || *s2.LatentStyles.Count != 376 {
		t.Error("LatentStyles count not preserved")
	}
}

func intPtr(v int) *int { return &v }
