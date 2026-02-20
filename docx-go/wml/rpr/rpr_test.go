package rpr

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/xmltypes"
)

// nsHdr is the namespace prefix header reused by test inputs.
const nsHdr = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
	` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`

// ────────────────────────────────────────────────────────
// CT_RPrBase round-trip
// ────────────────────────────────────────────────────────

func TestRPrBaseRoundTrip(t *testing.T) {
	input := `<w:rPr ` + nsHdr + `>` +
		`<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>` +
		`<w:b/>` +
		`<w:bCs/>` +
		`<w:i/>` +
		`<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>` +
		`<w:sz w:val="32"/>` +
		`<w:szCs w:val="32"/>` +
		`<w:u w:val="single"/>` +
		`<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>` +
		`<w14:someExtension w14:val="test"/>` +
		`</w:rPr>`

	var base CT_RPrBase
	if err := xml.Unmarshal([]byte(input), &base); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	assertNotNil(t, base.RFonts, "RFonts")
	assertEqual(t, *base.RFonts.Ascii, "Arial", "RFonts.Ascii")
	assertTrue(t, base.B.Bool(false), "B")
	assertTrue(t, base.BCs.Bool(false), "BCs")
	assertTrue(t, base.I.Bool(false), "I")
	assertNotNil(t, base.Color, "Color")
	assertEqual(t, base.Color.Val, "2F5496", "Color.Val")
	assertIntEq(t, base.Sz.Val, 32, "Sz.Val")
	assertIntEq(t, base.SzCs.Val, 32, "SzCs.Val")
	assertNotNil(t, base.U, "U")
	assertEqual(t, *base.U.Val, "single", "U.Val")
	assertNotNil(t, base.Lang, "Lang")
	assertEqual(t, *base.Lang.Val, "en-US", "Lang.Val")
	assertIntEq(t, len(base.Extra), 1, "len(Extra)")
	assertEqual(t, base.Extra[0].XMLName.Local, "someExtension", "Extra[0].Local")

	// Marshal → re-unmarshal.
	output, err := xml.Marshal(&base)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	var base2 CT_RPrBase
	if err := xml.Unmarshal(output, &base2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	assertEqual(t, *base2.RFonts.Ascii, "Arial", "rt RFonts.Ascii")
	assertTrue(t, base2.B.Bool(false), "rt B")
	assertIntEq(t, base2.Sz.Val, 32, "rt Sz")
	assertIntEq(t, len(base2.Extra), 1, "rt len(Extra)")
	assertEqual(t, base2.Extra[0].XMLName.Local, "someExtension", "rt Extra name")
}

// ────────────────────────────────────────────────────────
// CT_RPr round-trip (with rPrChange)
// ────────────────────────────────────────────────────────

func TestRPrRoundTrip(t *testing.T) {
	input := `<w:rPr ` + nsHdr + `>` +
		`<w:rStyle w:val="Strong"/>` +
		`<w:b/>` +
		`<w:sz w:val="24"/>` +
		`<w:rPrChange w:id="1" w:author="Alice" w:date="2025-01-15T10:00:00Z">` +
		`<w:rPr>` +
		`<w:sz w:val="20"/>` +
		`</w:rPr>` +
		`</w:rPrChange>` +
		`<w14:ligatures w14:val="standard"/>` +
		`</w:rPr>`

	var rpr CT_RPr
	if err := xml.Unmarshal([]byte(input), &rpr); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	assertEqual(t, rpr.Base.RStyle.Val, "Strong", "RStyle")
	assertTrue(t, rpr.Base.B.Bool(false), "B")
	assertIntEq(t, rpr.Base.Sz.Val, 24, "Sz")

	assertNotNil(t, rpr.RPrChange, "RPrChange")
	assertIntEq(t, rpr.RPrChange.ID, 1, "RPrChange.ID")
	assertEqual(t, rpr.RPrChange.Author, "Alice", "RPrChange.Author")
	assertNotNil(t, rpr.RPrChange.RPr, "RPrChange.RPr")
	assertIntEq(t, rpr.RPrChange.RPr.Sz.Val, 20, "RPrChange.RPr.Sz")

	assertIntEq(t, len(rpr.Extra), 1, "len(Extra)")
	assertEqual(t, rpr.Extra[0].XMLName.Local, "ligatures", "Extra[0]")

	// Round-trip.
	output, err := xml.Marshal(&rpr)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	var rpr2 CT_RPr
	if err := xml.Unmarshal(output, &rpr2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	assertEqual(t, rpr2.Base.RStyle.Val, "Strong", "rt RStyle")
	assertNotNil(t, rpr2.RPrChange, "rt RPrChange")
	assertIntEq(t, rpr2.RPrChange.ID, 1, "rt RPrChange.ID")
	assertIntEq(t, len(rpr2.Extra), 1, "rt Extra")
}

// ────────────────────────────────────────────────────────
// CT_ParaRPr round-trip (with ins/del)
// ────────────────────────────────────────────────────────

func TestParaRPrRoundTrip(t *testing.T) {
	input := `<w:rPr ` + nsHdr + `>` +
		`<w:b/>` +
		`<w:sz w:val="28"/>` +
		`<w:ins w:id="5" w:author="Bob" w:date="2025-02-01T09:00:00Z"/>` +
		`</w:rPr>`

	var prpr CT_ParaRPr
	if err := xml.Unmarshal([]byte(input), &prpr); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	assertTrue(t, prpr.Base.B.Bool(false), "B")
	assertIntEq(t, prpr.Base.Sz.Val, 28, "Sz")
	assertNotNil(t, prpr.Ins, "Ins")
	assertIntEq(t, prpr.Ins.ID, 5, "Ins.ID")
	assertEqual(t, prpr.Ins.Author, "Bob", "Ins.Author")

	output, err := xml.Marshal(&prpr)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	var prpr2 CT_ParaRPr
	if err := xml.Unmarshal(output, &prpr2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	assertTrue(t, prpr2.Base.B.Bool(false), "rt B")
	assertNotNil(t, prpr2.Ins, "rt Ins")
	assertIntEq(t, prpr2.Ins.ID, 5, "rt Ins.ID")
}

// ────────────────────────────────────────────────────────
// Real-world XML from reference-appendix (styles.xml rPrDefault)
// ────────────────────────────────────────────────────────

func TestRealWorldStylesRPr(t *testing.T) {
	input := `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi"` +
		` w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>` +
		`<w:sz w:val="24"/>` +
		`<w:szCs w:val="24"/>` +
		`<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>` +
		`</w:rPr>`

	var base CT_RPrBase
	if err := xml.Unmarshal([]byte(input), &base); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	assertNotNil(t, base.RFonts, "RFonts")
	assertEqual(t, *base.RFonts.AsciiTheme, "minorHAnsi", "AsciiTheme")
	assertEqual(t, *base.RFonts.CSTheme, "minorBidi", "CSTheme")
	assertIntEq(t, base.Sz.Val, 24, "Sz")
	assertNotNil(t, base.Lang, "Lang")
	assertEqual(t, *base.Lang.Bidi, "ar-SA", "Lang.Bidi")
}

// ────────────────────────────────────────────────────────
// CT_OnOff semantics
// ────────────────────────────────────────────────────────

func TestOnOffNilVal(t *testing.T) {
	input := `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:b/></w:rPr>`
	var base CT_RPrBase
	if err := xml.Unmarshal([]byte(input), &base); err != nil {
		t.Fatal(err)
	}
	assertNotNil(t, base.B, "B")
	assertTrue(t, base.B.Bool(false), "B no-val → true")
}

func TestOnOffExplicitFalse(t *testing.T) {
	input := `<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:b w:val="false"/></w:rPr>`
	var base CT_RPrBase
	if err := xml.Unmarshal([]byte(input), &base); err != nil {
		t.Fatal(err)
	}
	assertNotNil(t, base.B, "B")
	if base.B.Bool(true) {
		t.Error("B val=false should be false")
	}
}

// ────────────────────────────────────────────────────────
// Marshal order verification
// ────────────────────────────────────────────────────────

func TestMarshalOrder(t *testing.T) {
	var base CT_RPrBase
	base.Sz = &xmltypes.CT_HpsMeasure{Val: 28}
	base.B = &xmltypes.CT_OnOff{} // no Val → true
	calibri := "Calibri"
	base.RFonts = &xmltypes.CT_Fonts{Ascii: &calibri}

	out, err := xml.Marshal(&base)
	if err != nil {
		t.Fatal(err)
	}
	s := string(out)

	// encoding/xml writes <rFonts xmlns="...">, <b xmlns="...">, <sz xmlns="...">
	rFontsIdx := strings.Index(s, "<rFonts")
	bIdx := strings.Index(s, "<b ")
	szIdx := strings.Index(s, "<sz ")
	if rFontsIdx < 0 || bIdx < 0 || szIdx < 0 {
		t.Fatalf("missing elements in: %s", s)
	}
	if !(rFontsIdx < bIdx && bIdx < szIdx) {
		t.Errorf("wrong order: rFonts@%d b@%d sz@%d in %s", rFontsIdx, bIdx, szIdx, s)
	}
}

// ────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────

func assertNotNil(t *testing.T, v interface{}, name string) {
	t.Helper()
	if v == nil {
		t.Fatalf("%s is nil", name)
	}
}

func assertEqual(t *testing.T, got, want, name string) {
	t.Helper()
	if got != want {
		t.Errorf("%s = %q, want %q", name, got, want)
	}
}

func assertIntEq(t *testing.T, got, want int, name string) {
	t.Helper()
	if got != want {
		t.Errorf("%s = %d, want %d", name, got, want)
	}
}

func assertTrue(t *testing.T, v bool, name string) {
	t.Helper()
	if !v {
		t.Errorf("%s = false, want true", name)
	}
}
