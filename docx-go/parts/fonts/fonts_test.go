package fonts

import (
	"encoding/xml"
	"strings"
	"testing"
)

// referenceXML is the minimal fontTable.xml from reference-appendix §2.10.
const referenceXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
	`<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
	`<w:font w:name="Calibri">` +
	`<w:panose1 w:val="020F0502020204030204"/>` +
	`<w:charset w:val="00"/>` +
	`<w:family w:val="swiss"/>` +
	`<w:pitch w:val="variable"/>` +
	`</w:font>` +
	`<w:font w:name="Times New Roman">` +
	`<w:panose1 w:val="02020603050405020304"/>` +
	`<w:charset w:val="00"/>` +
	`<w:family w:val="roman"/>` +
	`<w:pitch w:val="variable"/>` +
	`</w:font>` +
	`<w:font w:name="Calibri Light">` +
	`<w:panose1 w:val="020F0302020204030204"/>` +
	`<w:charset w:val="00"/>` +
	`<w:family w:val="swiss"/>` +
	`<w:pitch w:val="variable"/>` +
	`</w:font>` +
	`</w:fonts>`

func TestFontsListRoundTrip(t *testing.T) {
	// ---- 1. Parse ----
	fl, err := Parse([]byte(referenceXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// ---- 2. Verify parsed content ----
	if len(fl.Font) != 3 {
		t.Fatalf("expected 3 fonts, got %d", len(fl.Font))
	}

	// Calibri
	calibri := fl.Font[0]
	if calibri.Name != "Calibri" {
		t.Errorf("font[0].Name = %q, want %q", calibri.Name, "Calibri")
	}
	if calibri.Panose1 == nil || calibri.Panose1.Val != "020F0502020204030204" {
		t.Error("font[0].Panose1 not parsed correctly")
	}
	if calibri.Charset == nil || calibri.Charset.Val != "00" {
		t.Error("font[0].Charset not parsed correctly")
	}
	if calibri.Family == nil || calibri.Family.Val != "swiss" {
		t.Error("font[0].Family not parsed correctly")
	}
	if calibri.Pitch == nil || calibri.Pitch.Val != "variable" {
		t.Error("font[0].Pitch not parsed correctly")
	}

	// Times New Roman
	tnr := fl.Font[1]
	if tnr.Name != "Times New Roman" {
		t.Errorf("font[1].Name = %q, want %q", tnr.Name, "Times New Roman")
	}
	if tnr.Family == nil || tnr.Family.Val != "roman" {
		t.Error("font[1].Family not parsed correctly")
	}

	// ---- 3. Serialize ----
	output, err := Serialize(fl)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// ---- 4. Re-parse and compare ----
	fl2, err := Parse(output)
	if err != nil {
		t.Fatalf("Re-parse failed: %v", err)
	}

	if len(fl2.Font) != len(fl.Font) {
		t.Fatalf("round-trip: expected %d fonts, got %d", len(fl.Font), len(fl2.Font))
	}

	for i, orig := range fl.Font {
		rt := fl2.Font[i]
		if orig.Name != rt.Name {
			t.Errorf("round-trip: font[%d].Name = %q, want %q", i, rt.Name, orig.Name)
		}
		comparePanose(t, i, orig.Panose1, rt.Panose1)
		compareCharset(t, i, orig.Charset, rt.Charset)
		compareFamily(t, i, orig.Family, rt.Family)
		comparePitch(t, i, orig.Pitch, rt.Pitch)
	}
}

func TestFontSigRoundTrip(t *testing.T) {
	input := `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:font w:name="Calibri">` +
		`<w:panose1 w:val="020F0502020204030204"/>` +
		`<w:charset w:val="00"/>` +
		`<w:family w:val="swiss"/>` +
		`<w:pitch w:val="variable"/>` +
		`<w:sig w:usb0="E4002EFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>` +
		`</w:font>` +
		`</w:fonts>`

	fl, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(fl.Font) != 1 {
		t.Fatalf("expected 1 font, got %d", len(fl.Font))
	}

	font := fl.Font[0]
	if font.Sig == nil {
		t.Fatal("Sig not parsed")
	}
	if font.Sig.Usb0 != "E4002EFF" {
		t.Errorf("Sig.Usb0 = %q, want %q", font.Sig.Usb0, "E4002EFF")
	}
	if font.Sig.Csb0 != "000001FF" {
		t.Errorf("Sig.Csb0 = %q, want %q", font.Sig.Csb0, "000001FF")
	}

	// Round-trip
	out, err := Serialize(fl)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	fl2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v", err)
	}

	if fl2.Font[0].Sig == nil {
		t.Fatal("round-trip lost Sig")
	}
	if fl2.Font[0].Sig.Usb0 != font.Sig.Usb0 {
		t.Error("round-trip lost Sig.Usb0")
	}
	if fl2.Font[0].Sig.Csb1 != font.Sig.Csb1 {
		t.Error("round-trip lost Sig.Csb1")
	}
}

func TestFontExtraRoundTrip(t *testing.T) {
	// Font with an unknown extension element — must survive round-trip.
	input := `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:font w:name="TestFont">` +
		`<w:panose1 w:val="00000000000000000000"/>` +
		`<w:charset w:val="00"/>` +
		`<w:family w:val="auto"/>` +
		`<w:pitch w:val="default"/>` +
		`<w14:someExtension w14:val="hello"/>` +
		`</w:font>` +
		`</w:fonts>`

	fl, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	font := fl.Font[0]
	if font.Name != "TestFont" {
		t.Errorf("Name = %q, want %q", font.Name, "TestFont")
	}
	if len(font.Extra) != 1 {
		t.Fatalf("expected 1 Extra element, got %d", len(font.Extra))
	}
	if font.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Extra[0].XMLName.Local = %q, want %q", font.Extra[0].XMLName.Local, "someExtension")
	}

	// Round-trip
	out, err := Serialize(fl)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	fl2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v", err)
	}

	if len(fl2.Font[0].Extra) != 1 {
		t.Fatalf("round-trip: expected 1 Extra, got %d", len(fl2.Font[0].Extra))
	}
	if fl2.Font[0].Extra[0].XMLName.Local != "someExtension" {
		t.Error("round-trip lost Extra element name")
	}
}

func TestFontEmbedRoundTrip(t *testing.T) {
	input := `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<w:font w:name="EmbeddedFont">` +
		`<w:charset w:val="00"/>` +
		`<w:family w:val="swiss"/>` +
		`<w:pitch w:val="variable"/>` +
		`<w:embedRegular w:fontKey="{12345678-1234-1234-1234-123456789012}" r:id="rId1"/>` +
		`<w:embedBold r:id="rId2"/>` +
		`</w:font>` +
		`</w:fonts>`

	fl, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	font := fl.Font[0]
	if font.EmbedRegular == nil {
		t.Fatal("EmbedRegular not parsed")
	}
	if font.EmbedRegular.ID != "rId1" {
		t.Errorf("EmbedRegular.ID = %q, want %q", font.EmbedRegular.ID, "rId1")
	}
	if font.EmbedRegular.FontKey == nil || *font.EmbedRegular.FontKey != "{12345678-1234-1234-1234-123456789012}" {
		t.Error("EmbedRegular.FontKey not parsed")
	}
	if font.EmbedBold == nil || font.EmbedBold.ID != "rId2" {
		t.Error("EmbedBold not parsed correctly")
	}
	if font.EmbedItalic != nil {
		t.Error("EmbedItalic should be nil")
	}

	// Round-trip
	out, err := Serialize(fl)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	fl2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v", err)
	}

	rt := fl2.Font[0]
	if rt.EmbedRegular == nil || rt.EmbedRegular.ID != "rId1" {
		t.Error("round-trip lost EmbedRegular.ID")
	}
	if rt.EmbedRegular.FontKey == nil || *rt.EmbedRegular.FontKey != *font.EmbedRegular.FontKey {
		t.Error("round-trip lost EmbedRegular.FontKey")
	}
	if rt.EmbedBold == nil || rt.EmbedBold.ID != "rId2" {
		t.Error("round-trip lost EmbedBold")
	}
}

func TestEmptyFontsList(t *testing.T) {
	input := `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:fonts>`

	fl, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(fl.Font) != 0 {
		t.Errorf("expected 0 fonts, got %d", len(fl.Font))
	}

	out, err := Serialize(fl)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	fl2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v", err)
	}
	if len(fl2.Font) != 0 {
		t.Errorf("round-trip: expected 0 fonts, got %d", len(fl2.Font))
	}
}

func TestMarshalXMLElementOrder(t *testing.T) {
	// Verify that marshal preserves strict XSD element order.
	fl := &CT_FontsList{
		Font: []CT_Font{
			{
				Name:    "Test",
				Pitch:   &CT_Pitch{Val: "variable"},
				Panose1: &CT_Panose{Val: "00000000000000000000"},
				Family:  &CT_FontFamily{Val: "swiss"},
				Charset: &CT_Charset{Val: "00"},
			},
		},
	}

	out, err := Serialize(fl)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s := string(out)
	panoseIdx := strings.Index(s, "panose1")
	charsetIdx := strings.Index(s, "charset")
	familyIdx := strings.Index(s, "family")
	pitchIdx := strings.Index(s, "pitch")

	if panoseIdx < 0 || charsetIdx < 0 || familyIdx < 0 || pitchIdx < 0 {
		t.Fatalf("missing elements in output: %s", s)
	}

	// XSD order: panose1 < charset < family < pitch
	if panoseIdx >= charsetIdx {
		t.Errorf("panose1 (pos %d) should come before charset (pos %d)", panoseIdx, charsetIdx)
	}
	if charsetIdx >= familyIdx {
		t.Errorf("charset (pos %d) should come before family (pos %d)", charsetIdx, familyIdx)
	}
	if familyIdx >= pitchIdx {
		t.Errorf("family (pos %d) should come before pitch (pos %d)", familyIdx, pitchIdx)
	}
}

func TestCT_FontDirectMarshalUnmarshal(t *testing.T) {
	// Test direct xml.Marshal/Unmarshal of CT_Font.
	orig := CT_Font{
		Name:    "Arial",
		Charset: &CT_Charset{Val: "00"},
		Family:  &CT_FontFamily{Val: "swiss"},
		Pitch:   &CT_Pitch{Val: "variable"},
	}

	data, err := xml.Marshal(&orig)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	var decoded CT_Font
	if err := xml.Unmarshal(data, &decoded); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if decoded.Name != orig.Name {
		t.Errorf("Name = %q, want %q", decoded.Name, orig.Name)
	}
	if decoded.Charset == nil || decoded.Charset.Val != "00" {
		t.Error("Charset not round-tripped")
	}
	if decoded.Family == nil || decoded.Family.Val != "swiss" {
		t.Error("Family not round-tripped")
	}
	if decoded.Pitch == nil || decoded.Pitch.Val != "variable" {
		t.Error("Pitch not round-tripped")
	}
}

// ============================================================
// Helpers
// ============================================================

func comparePanose(t *testing.T, idx int, a, b *CT_Panose) {
	t.Helper()
	if (a == nil) != (b == nil) {
		t.Errorf("font[%d].Panose1: nil mismatch", idx)
		return
	}
	if a != nil && a.Val != b.Val {
		t.Errorf("font[%d].Panose1.Val = %q, want %q", idx, b.Val, a.Val)
	}
}

func compareCharset(t *testing.T, idx int, a, b *CT_Charset) {
	t.Helper()
	if (a == nil) != (b == nil) {
		t.Errorf("font[%d].Charset: nil mismatch", idx)
		return
	}
	if a != nil && a.Val != b.Val {
		t.Errorf("font[%d].Charset.Val = %q, want %q", idx, b.Val, a.Val)
	}
}

func compareFamily(t *testing.T, idx int, a, b *CT_FontFamily) {
	t.Helper()
	if (a == nil) != (b == nil) {
		t.Errorf("font[%d].Family: nil mismatch", idx)
		return
	}
	if a != nil && a.Val != b.Val {
		t.Errorf("font[%d].Family.Val = %q, want %q", idx, b.Val, a.Val)
	}
}

func comparePitch(t *testing.T, idx int, a, b *CT_Pitch) {
	t.Helper()
	if (a == nil) != (b == nil) {
		t.Errorf("font[%d].Pitch: nil mismatch", idx)
		return
	}
	if a != nil && a.Val != b.Val {
		t.Errorf("font[%d].Pitch.Val = %q, want %q", idx, b.Val, a.Val)
	}
}
