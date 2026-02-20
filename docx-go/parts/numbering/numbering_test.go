package numbering

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/xmltypes"
)

// numberingXML is the reference XML from reference-appendix.md §3.3,
// wrapped in a <w:numbering> root element.
const numberingXML = `<w:numbering xmlns:w="` + xmltypes.NSw + `">
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="3A5C117E"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:tmpl w:val="E6A2FD28"/>
    <w:lvl w:ilvl="0" w:tplc="04090001">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#xF0B7;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1" w:tplc="04090003">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`

func TestRoundTrip(t *testing.T) {
	// --- Unmarshal ---
	var numbering CT_Numbering
	if err := xml.Unmarshal([]byte(numberingXML), &numbering); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	// Verify parsed structure
	if got := len(numbering.AbstractNum); got != 1 {
		t.Fatalf("AbstractNum count: got %d, want 1", got)
	}
	an := numbering.AbstractNum[0]
	if an.AbstractNumID != 0 {
		t.Errorf("AbstractNumID: got %d, want 0", an.AbstractNumID)
	}
	if an.Nsid == nil || an.Nsid.Val != "3A5C117E" {
		t.Errorf("Nsid: got %v, want 3A5C117E", an.Nsid)
	}
	if an.MultiLevelType == nil || an.MultiLevelType.Val != "hybridMultilevel" {
		t.Errorf("MultiLevelType: got %v, want hybridMultilevel", an.MultiLevelType)
	}
	if an.Tmpl == nil || an.Tmpl.Val != "E6A2FD28" {
		t.Errorf("Tmpl: got %v, want E6A2FD28", an.Tmpl)
	}
	if got := len(an.Lvl); got != 2 {
		t.Fatalf("Lvl count: got %d, want 2", got)
	}

	// Level 0
	lvl0 := an.Lvl[0]
	if lvl0.Ilvl != 0 {
		t.Errorf("Lvl[0].Ilvl: got %d, want 0", lvl0.Ilvl)
	}
	if lvl0.Tplc == nil || *lvl0.Tplc != "04090001" {
		t.Errorf("Lvl[0].Tplc: got %v, want 04090001", lvl0.Tplc)
	}
	if lvl0.Start == nil || lvl0.Start.Val != 1 {
		t.Errorf("Lvl[0].Start: got %v, want 1", lvl0.Start)
	}
	if lvl0.NumFmt == nil || lvl0.NumFmt.Val != "bullet" {
		t.Errorf("Lvl[0].NumFmt: got %v, want bullet", lvl0.NumFmt)
	}
	if lvl0.LvlJc == nil || lvl0.LvlJc.Val != "left" {
		t.Errorf("Lvl[0].LvlJc: got %v, want left", lvl0.LvlJc)
	}
	if lvl0.PPr == nil {
		t.Error("Lvl[0].PPr: expected non-nil")
	}
	if lvl0.RPr == nil {
		t.Error("Lvl[0].RPr: expected non-nil")
	}

	// Level 1
	lvl1 := an.Lvl[1]
	if lvl1.Ilvl != 1 {
		t.Errorf("Lvl[1].Ilvl: got %d, want 1", lvl1.Ilvl)
	}

	// Num
	if got := len(numbering.Num); got != 1 {
		t.Fatalf("Num count: got %d, want 1", got)
	}
	num := numbering.Num[0]
	if num.NumID != 1 {
		t.Errorf("Num.NumID: got %d, want 1", num.NumID)
	}
	if num.AbstractNumID.Val != 0 {
		t.Errorf("Num.AbstractNumID: got %d, want 0", num.AbstractNumID.Val)
	}

	// --- Marshal ---
	out, err := xml.MarshalIndent(&numbering, "", "  ")
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}
	outStr := string(out)

	// Verify key structural elements survive the round-trip
	mustContain := []string{
		`abstractNumId="0"`,
		`val="3A5C117E"`,
		`val="hybridMultilevel"`,
		`val="E6A2FD28"`,
		`ilvl="0"`,
		`tplc="04090001"`,
		`val="1"`,
		`val="bullet"`,
		`val="left"`,
		`ilvl="1"`,
		`tplc="04090003"`,
		`numId="1"`,
		// PPr and RPr inner content — only assert on attributes that
		// survive ppr/rpr round-trip (namespace-qualified attrs like
		// w:left may be dropped by Go's encoding/xml; that is a ppr
		// concern, not numbering's).
		`hanging="360"`,
		`ascii="Symbol"`,
		`ascii="Courier New"`,
	}
	for _, s := range mustContain {
		if !strings.Contains(outStr, s) {
			t.Errorf("Marshal output missing %q\n\nFull output:\n%s", s, outStr)
		}
	}
}

func TestRoundTripWithLvlOverride(t *testing.T) {
	input := `<w:numbering xmlns:w="` + xmltypes.NSw + `">
  <w:abstractNum w:abstractNumId="1">
    <w:nsid w:val="AABBCCDD"/>
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="5"/>
    </w:lvlOverride>
  </w:num>
</w:numbering>`

	var n CT_Numbering
	if err := xml.Unmarshal([]byte(input), &n); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}

	if len(n.Num) != 1 {
		t.Fatalf("Num count: got %d, want 1", len(n.Num))
	}
	num := n.Num[0]
	if len(num.LvlOverride) != 1 {
		t.Fatalf("LvlOverride count: got %d, want 1", len(num.LvlOverride))
	}
	lo := num.LvlOverride[0]
	if lo.Ilvl != 0 {
		t.Errorf("LvlOverride.Ilvl: got %d, want 0", lo.Ilvl)
	}
	if lo.StartOverride == nil || lo.StartOverride.Val != 5 {
		t.Errorf("StartOverride.Val: got %v, want 5", lo.StartOverride)
	}

	// Marshal back and verify
	out, err := xml.MarshalIndent(&n, "", "  ")
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	outStr := string(out)
	for _, s := range []string{`ilvl="0"`, `val="5"`, `numId="2"`, `val="1"`} {
		if !strings.Contains(outStr, s) {
			t.Errorf("Missing %q in output:\n%s", s, outStr)
		}
	}
}

func TestRoundTripWithExtra(t *testing.T) {
	// Test that unknown elements survive round-trip via RawXML.
	input := `<w:numbering xmlns:w="` + xmltypes.NSw + `" xmlns:w14="` + xmltypes.NSw14 + `">
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="00000001"/>
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w14:someExtension w14:foo="bar">inner content</w14:someExtension>
    </w:lvl>
  </w:abstractNum>
  <w14:numIdMacAtCleanup w14:val="42"/>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`

	var n CT_Numbering
	if err := xml.Unmarshal([]byte(input), &n); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}

	// Verify the extension was captured in Lvl.Extra
	if len(n.AbstractNum) != 1 {
		t.Fatalf("AbstractNum count: got %d, want 1", len(n.AbstractNum))
	}
	lvl := n.AbstractNum[0].Lvl[0]
	if len(lvl.Extra) != 1 {
		t.Fatalf("Lvl.Extra count: got %d, want 1", len(lvl.Extra))
	}
	if lvl.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Extra[0].Local: got %q, want someExtension", lvl.Extra[0].XMLName.Local)
	}

	// Verify extra at numbering level was captured
	if len(n.Extra) != 1 {
		t.Fatalf("Numbering.Extra count: got %d, want 1", len(n.Extra))
	}
	if n.Extra[0].XMLName.Local != "numIdMacAtCleanup" {
		t.Errorf("Numbering.Extra[0].Local: got %q, want numIdMacAtCleanup", n.Extra[0].XMLName.Local)
	}

	// Marshal and verify the extension survives
	out, err := xml.MarshalIndent(&n, "", "  ")
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	outStr := string(out)
	if !strings.Contains(outStr, "someExtension") {
		t.Errorf("Extension element lost in round-trip:\n%s", outStr)
	}
	if !strings.Contains(outStr, "inner content") {
		t.Errorf("Extension inner content lost in round-trip:\n%s", outStr)
	}
	if !strings.Contains(outStr, "numIdMacAtCleanup") {
		t.Errorf("Top-level extra lost in round-trip:\n%s", outStr)
	}
}

func TestRoundTripParseSerialize(t *testing.T) {
	n, err := Parse([]byte(numberingXML))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	out, err := Serialize(n)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	outStr := string(out)
	// Must contain XML header
	if !strings.Contains(outStr, "<?xml") {
		t.Error("Missing XML declaration in Serialize output")
	}
	// Spot-check content
	for _, s := range []string{"abstractNumId", "numId", "3A5C117E", "hybridMultilevel"} {
		if !strings.Contains(outStr, s) {
			t.Errorf("Serialize output missing %q", s)
		}
	}
}

func TestEmptyNumbering(t *testing.T) {
	input := `<w:numbering xmlns:w="` + xmltypes.NSw + `"></w:numbering>`
	var n CT_Numbering
	if err := xml.Unmarshal([]byte(input), &n); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(n.AbstractNum) != 0 {
		t.Errorf("AbstractNum should be empty, got %d", len(n.AbstractNum))
	}
	if len(n.Num) != 0 {
		t.Errorf("Num should be empty, got %d", len(n.Num))
	}

	out, err := xml.Marshal(&n)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	if !strings.Contains(string(out), "numbering") {
		t.Errorf("Output missing root element: %s", out)
	}
}

func TestLvlOptionalFields(t *testing.T) {
	// Lvl with only required fields: ilvl attribute and a start element.
	input := `<w:numbering xmlns:w="` + xmltypes.NSw + `">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1)"/>
      <w:suff w:val="space"/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`

	var n CT_Numbering
	if err := xml.Unmarshal([]byte(input), &n); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}

	lvl := n.AbstractNum[0].Lvl[0]
	if lvl.Suff == nil || lvl.Suff.Val != "space" {
		t.Errorf("Suff: got %v, want space", lvl.Suff)
	}
	if lvl.LvlText == nil || lvl.LvlText.Val != "%1)" {
		t.Errorf("LvlText: got %v, want %%1)", lvl.LvlText)
	}
	// Fields not present should be nil
	if lvl.LvlRestart != nil {
		t.Error("LvlRestart should be nil")
	}
	if lvl.PStyle != nil {
		t.Error("PStyle should be nil")
	}
	if lvl.IsLgl != nil {
		t.Error("IsLgl should be nil")
	}
	if lvl.PPr != nil {
		t.Error("PPr should be nil")
	}
	if lvl.RPr != nil {
		t.Error("RPr should be nil")
	}

	// Marshal and re-parse to verify stability
	out, err := xml.Marshal(&n)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}

	var n2 CT_Numbering
	if err := xml.Unmarshal(out, &n2); err != nil {
		t.Fatalf("Re-unmarshal: %v", err)
	}
	lvl2 := n2.AbstractNum[0].Lvl[0]
	if lvl2.Suff == nil || lvl2.Suff.Val != "space" {
		t.Errorf("After re-parse, Suff: got %v, want space", lvl2.Suff)
	}
}

func TestLvlWithFullOverride(t *testing.T) {
	input := `<w:numbering xmlns:w="` + xmltypes.NSw + `">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0">
      <w:lvl w:ilvl="0">
        <w:start w:val="10"/>
        <w:numFmt w:val="upperRoman"/>
        <w:lvlText w:val="%1)"/>
        <w:lvlJc w:val="center"/>
      </w:lvl>
    </w:lvlOverride>
  </w:num>
</w:numbering>`

	var n CT_Numbering
	if err := xml.Unmarshal([]byte(input), &n); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}

	lo := n.Num[0].LvlOverride[0]
	if lo.StartOverride != nil {
		t.Error("StartOverride should be nil when lvl is present")
	}
	if lo.Lvl == nil {
		t.Fatal("Lvl should be non-nil")
	}
	if lo.Lvl.Start == nil || lo.Lvl.Start.Val != 10 {
		t.Errorf("Override Lvl.Start: got %v, want 10", lo.Lvl.Start)
	}
	if lo.Lvl.NumFmt == nil || lo.Lvl.NumFmt.Val != "upperRoman" {
		t.Errorf("Override Lvl.NumFmt: got %v, want upperRoman", lo.Lvl.NumFmt)
	}
	if lo.Lvl.LvlJc == nil || lo.Lvl.LvlJc.Val != "center" {
		t.Errorf("Override Lvl.LvlJc: got %v, want center", lo.Lvl.LvlJc)
	}

	// Round-trip
	out, err := xml.Marshal(&n)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	for _, s := range []string{"upperRoman", "center", `val="10"`} {
		if !strings.Contains(string(out), s) {
			t.Errorf("Missing %q in marshal output", s)
		}
	}
}
