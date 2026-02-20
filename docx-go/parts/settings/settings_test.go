package settings

import (
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
)

// referenceSettingsXML is the minimal-valid settings.xml from
// reference-appendix.md §2.7.  It exercises every typed field and includes
// extension-namespace elements (w14, w15, m) that must survive round-trip.
const referenceSettingsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
            mc:Ignorable="w14 w15">
  <w:zoom w:percent="100"/>
  <w:proofState w:spelling="clean" w:grammar="clean"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"
      w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="enableOpenTypeFeatures"
      w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="doNotFlipMirrorIndents"
      w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="differentiateMultirowTableHeaders"
      w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="useWord2013TrackBottomHyphenation"
      w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>
  </w:compat>
  <w:rsids>
    <w:rsidRoot w:val="00000001"/>
    <w:rsid w:val="00000001"/>
  </w:rsids>
  <m:mathPr>
    <m:mathFont m:val="Cambria Math"/>
    <m:brkBin m:val="before"/>
    <m:brkBinSub m:val="--"/>
    <m:smallFrac m:val="0"/>
    <m:dispDef/>
    <m:lMargin m:val="0"/>
    <m:rMargin m:val="0"/>
    <m:defJc m:val="centerGroup"/>
    <m:wrapIndent m:val="1440"/>
    <m:intLim m:val="subSup"/>
    <m:naryLim m:val="undOvr"/>
  </m:mathPr>
  <w:themeFontLang w:val="en-US"/>
  <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2"
    w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4"
    w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink"
    w:followedHyperlink="followedHyperlink"/>
  <w:shapeDefaults>
    <o:shapedefaults v:ext="edit" spidmax="1026"/>
    <o:shapelayout v:ext="edit">
      <o:idmap v:ext="edit" data="1"/>
    </o:shapelayout>
  </w:shapeDefaults>
  <w:decimalSymbol w:val="."/>
  <w:listSeparator w:val=","/>
  <w14:docId w14:val="00000001"/>
  <w15:chartTrackingRefBased/>
  <w15:docId w15:val="{00000000-0000-0000-0000-000000000001}"/>
</w:settings>`

// ---------------------------------------------------------------------------
// Round-trip test
// ---------------------------------------------------------------------------

func TestSettingsRoundTrip(t *testing.T) {
	// --- 1. Unmarshal ---
	settings, err := Parse([]byte(referenceSettingsXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// --- 2. Structural checks ---
	assertZoom(t, settings)
	assertProofState(t, settings)
	assertDefaultTabStop(t, settings)
	assertCharacterSpacingControl(t, settings)
	assertCompat(t, settings)
	assertRsids(t, settings)
	assertMathPr(t, settings)
	assertThemeFontLang(t, settings)
	assertClrSchemeMapping(t, settings)
	assertShapeDefaults(t, settings)
	assertDecimalSymbol(t, settings)
	assertListSeparator(t, settings)
	assertDocId14(t, settings)
	assertDocId15(t, settings)
	assertExtras(t, settings)

	// --- 3. Marshal ---
	output, err := Serialize(settings)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// --- 4. Re-unmarshal and compare ---
	settings2, err := Parse(output)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\nOutput:\n%s", err, string(output))
	}

	compareFields(t, settings, settings2)
}

// ---------------------------------------------------------------------------
// Assertion helpers
// ---------------------------------------------------------------------------

func assertZoom(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.Zoom == nil {
		t.Fatal("Zoom is nil")
	}
	if s.Zoom.Percent != 100 {
		t.Errorf("Zoom.Percent = %d, want 100", s.Zoom.Percent)
	}
}

func assertProofState(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.ProofState == nil {
		t.Fatal("ProofState is nil")
	}
	if s.ProofState.Spelling == nil || *s.ProofState.Spelling != "clean" {
		t.Errorf("ProofState.Spelling = %v, want 'clean'", s.ProofState.Spelling)
	}
	if s.ProofState.Grammar == nil || *s.ProofState.Grammar != "clean" {
		t.Errorf("ProofState.Grammar = %v, want 'clean'", s.ProofState.Grammar)
	}
}

func assertDefaultTabStop(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.DefaultTabStop == nil {
		t.Fatal("DefaultTabStop is nil")
	}
	if s.DefaultTabStop.Val != 720 {
		t.Errorf("DefaultTabStop.Val = %d, want 720", s.DefaultTabStop.Val)
	}
}

func assertCharacterSpacingControl(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.CharacterSpacingControl == nil {
		t.Fatal("CharacterSpacingControl is nil")
	}
	if s.CharacterSpacingControl.Val != "doNotCompress" {
		t.Errorf("CharacterSpacingControl.Val = %q, want 'doNotCompress'", s.CharacterSpacingControl.Val)
	}
}

func assertCompat(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.Compat == nil {
		t.Fatal("Compat is nil")
	}
	if len(s.Compat.CompatSetting) != 6 {
		t.Errorf("Compat.CompatSetting count = %d, want 6", len(s.Compat.CompatSetting))
		return
	}
	cs := s.Compat.CompatSetting[0]
	if cs.Name != "compatibilityMode" {
		t.Errorf("Compat first setting name = %q, want 'compatibilityMode'", cs.Name)
	}
	if cs.Val != "15" {
		t.Errorf("Compat first setting val = %q, want '15'", cs.Val)
	}
}

func assertRsids(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.Rsids == nil {
		t.Fatal("Rsids is nil")
	}
	if s.Rsids.RsidRoot == nil {
		t.Fatal("Rsids.RsidRoot is nil")
	}
	if s.Rsids.RsidRoot.Val != "00000001" {
		t.Errorf("RsidRoot.Val = %q, want '00000001'", s.Rsids.RsidRoot.Val)
	}
	if len(s.Rsids.Rsid) != 1 {
		t.Errorf("Rsid count = %d, want 1", len(s.Rsids.Rsid))
	}
}

func assertMathPr(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.MathPr == nil {
		t.Fatal("MathPr is nil")
	}
	if s.MathPr.XMLName.Local != "mathPr" {
		t.Errorf("MathPr.XMLName.Local = %q, want 'mathPr'", s.MathPr.XMLName.Local)
	}
	// MathPr inner content should contain mathFont etc.
	if !strings.Contains(string(s.MathPr.Inner), "Cambria Math") {
		t.Error("MathPr inner XML should contain 'Cambria Math'")
	}
}

func assertThemeFontLang(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.ThemeFontLang == nil {
		t.Fatal("ThemeFontLang is nil")
	}
	if s.ThemeFontLang.Val == nil || *s.ThemeFontLang.Val != "en-US" {
		t.Errorf("ThemeFontLang.Val = %v, want 'en-US'", s.ThemeFontLang.Val)
	}
}

func assertClrSchemeMapping(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.ClrSchemeMapping == nil {
		t.Fatal("ClrSchemeMapping is nil")
	}
	if s.ClrSchemeMapping.Bg1 != "light1" {
		t.Errorf("ClrSchemeMapping.Bg1 = %q, want 'light1'", s.ClrSchemeMapping.Bg1)
	}
	if s.ClrSchemeMapping.T1 != "dark1" {
		t.Errorf("ClrSchemeMapping.T1 = %q, want 'dark1'", s.ClrSchemeMapping.T1)
	}
	if s.ClrSchemeMapping.Hyperlink != "hyperlink" {
		t.Errorf("ClrSchemeMapping.Hyperlink = %q, want 'hyperlink'", s.ClrSchemeMapping.Hyperlink)
	}
	if s.ClrSchemeMapping.FollowedHyperlink != "followedHyperlink" {
		t.Errorf("ClrSchemeMapping.FollowedHyperlink = %q, want 'followedHyperlink'",
			s.ClrSchemeMapping.FollowedHyperlink)
	}
}

func assertShapeDefaults(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.ShapeDefaults == nil {
		t.Fatal("ShapeDefaults is nil")
	}
	if s.ShapeDefaults.XMLName.Local != "shapeDefaults" {
		t.Errorf("ShapeDefaults.XMLName.Local = %q, want 'shapeDefaults'",
			s.ShapeDefaults.XMLName.Local)
	}
	if !strings.Contains(string(s.ShapeDefaults.Inner), "shapedefaults") {
		t.Error("ShapeDefaults inner should contain VML shape defaults")
	}
}

func assertDecimalSymbol(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.DecimalSymbol == nil {
		t.Fatal("DecimalSymbol is nil")
	}
	if s.DecimalSymbol.Val != "." {
		t.Errorf("DecimalSymbol.Val = %q, want '.'", s.DecimalSymbol.Val)
	}
}

func assertListSeparator(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.ListSeparator == nil {
		t.Fatal("ListSeparator is nil")
	}
	if s.ListSeparator.Val != "," {
		t.Errorf("ListSeparator.Val = %q, want ','", s.ListSeparator.Val)
	}
}

func assertDocId14(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.DocId14 == nil {
		t.Fatal("DocId14 is nil")
	}
	if s.DocId14.Val != "00000001" {
		t.Errorf("DocId14.Val = %q, want '00000001'", s.DocId14.Val)
	}
}

func assertDocId15(t *testing.T, s *CT_Settings) {
	t.Helper()
	if s.DocId15 == nil {
		t.Fatal("DocId15 is nil")
	}
	if s.DocId15.Val != "{00000000-0000-0000-0000-000000000001}" {
		t.Errorf("DocId15.Val = %q, want '{00000000-0000-0000-0000-000000000001}'",
			s.DocId15.Val)
	}
}

func assertExtras(t *testing.T, s *CT_Settings) {
	t.Helper()
	// <w15:chartTrackingRefBased/> should have been captured as Extra.
	found := false
	for _, raw := range s.Extra {
		if raw.XMLName.Local == "chartTrackingRefBased" {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("Expected 'chartTrackingRefBased' in Extra, got %d extras: %v",
			len(s.Extra), extraNames(s.Extra))
	}
}

// ---------------------------------------------------------------------------
// Round-trip comparison
// ---------------------------------------------------------------------------

func compareFields(t *testing.T, a, b *CT_Settings) {
	t.Helper()

	// Zoom
	if a.Zoom != nil && b.Zoom != nil {
		if a.Zoom.Percent != b.Zoom.Percent {
			t.Errorf("round-trip lost Zoom.Percent: %d → %d", a.Zoom.Percent, b.Zoom.Percent)
		}
	} else if (a.Zoom == nil) != (b.Zoom == nil) {
		t.Errorf("round-trip Zoom nil mismatch")
	}

	// DefaultTabStop
	if a.DefaultTabStop != nil && b.DefaultTabStop != nil {
		if a.DefaultTabStop.Val != b.DefaultTabStop.Val {
			t.Errorf("round-trip lost DefaultTabStop: %d → %d",
				a.DefaultTabStop.Val, b.DefaultTabStop.Val)
		}
	}

	// CharacterSpacingControl
	if a.CharacterSpacingControl != nil && b.CharacterSpacingControl != nil {
		if a.CharacterSpacingControl.Val != b.CharacterSpacingControl.Val {
			t.Errorf("round-trip lost CharacterSpacingControl: %q → %q",
				a.CharacterSpacingControl.Val, b.CharacterSpacingControl.Val)
		}
	}

	// Compat
	if a.Compat != nil && b.Compat != nil {
		if len(a.Compat.CompatSetting) != len(b.Compat.CompatSetting) {
			t.Errorf("round-trip lost compat settings: %d → %d",
				len(a.Compat.CompatSetting), len(b.Compat.CompatSetting))
		}
	}

	// Rsids
	if a.Rsids != nil && b.Rsids != nil {
		if len(a.Rsids.Rsid) != len(b.Rsids.Rsid) {
			t.Errorf("round-trip lost rsids: %d → %d",
				len(a.Rsids.Rsid), len(b.Rsids.Rsid))
		}
		if a.Rsids.RsidRoot != nil && b.Rsids.RsidRoot != nil {
			if a.Rsids.RsidRoot.Val != b.Rsids.RsidRoot.Val {
				t.Errorf("round-trip lost rsidRoot: %q → %q",
					a.Rsids.RsidRoot.Val, b.Rsids.RsidRoot.Val)
			}
		}
	}

	// ClrSchemeMapping
	if a.ClrSchemeMapping != nil && b.ClrSchemeMapping != nil {
		if a.ClrSchemeMapping.Bg1 != b.ClrSchemeMapping.Bg1 {
			t.Errorf("round-trip lost ClrSchemeMapping.Bg1: %q → %q",
				a.ClrSchemeMapping.Bg1, b.ClrSchemeMapping.Bg1)
		}
		if a.ClrSchemeMapping.FollowedHyperlink != b.ClrSchemeMapping.FollowedHyperlink {
			t.Errorf("round-trip lost ClrSchemeMapping.FollowedHyperlink")
		}
	}

	// DocId14
	if a.DocId14 != nil && b.DocId14 != nil {
		if a.DocId14.Val != b.DocId14.Val {
			t.Errorf("round-trip lost DocId14: %q → %q", a.DocId14.Val, b.DocId14.Val)
		}
	}

	// DocId15
	if a.DocId15 != nil && b.DocId15 != nil {
		if a.DocId15.Val != b.DocId15.Val {
			t.Errorf("round-trip lost DocId15: %q → %q", a.DocId15.Val, b.DocId15.Val)
		}
	}

	// Extra count
	if len(a.Extra) != len(b.Extra) {
		t.Errorf("round-trip lost Extra: %d → %d",
			len(a.Extra), len(b.Extra))
	}
	for i := 0; i < len(a.Extra) && i < len(b.Extra); i++ {
		if a.Extra[i].XMLName.Local != b.Extra[i].XMLName.Local {
			t.Errorf("round-trip Extra[%d] name: %q → %q",
				i, a.Extra[i].XMLName.Local, b.Extra[i].XMLName.Local)
		}
	}

	// Element order preserved
	if len(a.elementOrder) != len(b.elementOrder) {
		t.Errorf("round-trip changed elementOrder length: %d → %d",
			len(a.elementOrder), len(b.elementOrder))
	}
}

// ---------------------------------------------------------------------------
// Additional tests
// ---------------------------------------------------------------------------

func TestSettingsEmptyExtraRoundTrip(t *testing.T) {
	// Minimal settings with only required elements.
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:zoom w:percent="100"/>` +
		`<w:defaultTabStop w:val="720"/>` +
		`</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if s.Zoom == nil || s.Zoom.Percent != 100 {
		t.Error("Zoom not parsed")
	}
	if s.DefaultTabStop == nil || s.DefaultTabStop.Val != 720 {
		t.Error("DefaultTabStop not parsed")
	}
	if len(s.Extra) != 0 {
		t.Errorf("Expected 0 Extra, got %d", len(s.Extra))
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\nOutput: %s", err, string(out))
	}
	if s2.Zoom == nil || s2.Zoom.Percent != 100 {
		t.Error("round-trip lost Zoom")
	}
	if s2.DefaultTabStop == nil || s2.DefaultTabStop.Val != 720 {
		t.Error("round-trip lost DefaultTabStop")
	}
}

func TestSettingsOnOffFields(t *testing.T) {
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:evenAndOddHeaders/>` +
		`<w:trackRevisions w:val="true"/>` +
		`</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if !s.EvenAndOddHeaders.Bool(false) {
		t.Error("EvenAndOddHeaders should be true (empty element)")
	}
	if !s.TrackRevisions.Bool(false) {
		t.Error("TrackRevisions should be true (val='true')")
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse failed: %v\nOutput: %s", err, string(out))
	}
	if !s2.EvenAndOddHeaders.Bool(false) {
		t.Error("round-trip lost EvenAndOddHeaders")
	}
	if !s2.TrackRevisions.Bool(false) {
		t.Error("round-trip lost TrackRevisions")
	}
}

func TestSettingsElementOrderPreserved(t *testing.T) {
	// Elements intentionally in an atypical order.
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:defaultTabStop w:val="720"/>` +
		`<w:zoom w:percent="150"/>` +
		`<w14:docId w14:val="ABCD1234"/>` +
		`</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// Verify the recorded order.
	wantOrder := []string{"defaultTabStop", "zoom", "w14:docId"}
	if len(s.elementOrder) != len(wantOrder) {
		t.Fatalf("elementOrder = %v, want %v", s.elementOrder, wantOrder)
	}
	for i, key := range wantOrder {
		if s.elementOrder[i] != key {
			t.Errorf("elementOrder[%d] = %q, want %q", i, s.elementOrder[i], key)
		}
	}

	// Marshal and verify order is preserved in output.
	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	outStr := string(out)

	idxTab := strings.Index(outStr, "defaultTabStop")
	idxZoom := strings.Index(outStr, "zoom")
	idxDocId := strings.Index(outStr, "docId")

	if idxTab >= idxZoom || idxZoom >= idxDocId {
		t.Errorf("Element order not preserved in output:\n%s", outStr)
	}
}

func TestSettingsWriteProtection(t *testing.T) {
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:writeProtection w:recommended="1"/>` +
		`</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if s.WriteProtection == nil {
		t.Fatal("WriteProtection is nil")
	}
	if s.WriteProtection.Recommended == nil || !*s.WriteProtection.Recommended {
		t.Error("WriteProtection.Recommended should be true")
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse: %v", err)
	}
	if s2.WriteProtection == nil || s2.WriteProtection.Recommended == nil || !*s2.WriteProtection.Recommended {
		t.Error("round-trip lost WriteProtection.Recommended")
	}
}

func TestSettingsDocumentProtection(t *testing.T) {
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:documentProtection w:edit="readOnly" w:enforcement="1"/>` +
		`</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if s.DocumentProtection == nil {
		t.Fatal("DocumentProtection is nil")
	}
	if s.DocumentProtection.Edit == nil || *s.DocumentProtection.Edit != "readOnly" {
		t.Error("DocumentProtection.Edit should be 'readOnly'")
	}
	if s.DocumentProtection.Enforcement == nil || !*s.DocumentProtection.Enforcement {
		t.Error("DocumentProtection.Enforcement should be true")
	}
}

func TestSettingsUnknownElementsPreserved(t *testing.T) {
	// Two unknown elements that must survive round-trip.
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml">` +
		`<w:zoom w:percent="100"/>` +
		`<w:gutterAtTop/>` +
		`<w16:someFutureElement w16:attr="value"><w16:child/></w16:someFutureElement>` +
		`</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	if len(s.Extra) != 2 {
		t.Fatalf("Expected 2 Extra elements, got %d: %v", len(s.Extra), extraNames(s.Extra))
	}

	if s.Extra[0].XMLName.Local != "gutterAtTop" {
		t.Errorf("Extra[0].Local = %q, want 'gutterAtTop'", s.Extra[0].XMLName.Local)
	}
	if s.Extra[1].XMLName.Local != "someFutureElement" {
		t.Errorf("Extra[1].Local = %q, want 'someFutureElement'", s.Extra[1].XMLName.Local)
	}

	out, err := Serialize(s)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	s2, err := Parse(out)
	if err != nil {
		t.Fatalf("Re-parse: %v\n%s", err, string(out))
	}

	if len(s2.Extra) != 2 {
		t.Fatalf("round-trip Extra count: %d → %d", len(s.Extra), len(s2.Extra))
	}
	if s2.Extra[1].XMLName.Local != "someFutureElement" {
		t.Error("round-trip lost 'someFutureElement'")
	}
}

func TestSettingsCompatExtra(t *testing.T) {
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
	<w:compat>
	  <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
	  <w15:futureFlag/>
	</w:compat>
	</w:settings>`

	s, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if s.Compat == nil {
		t.Fatal("Compat is nil")
	}
	if len(s.Compat.CompatSetting) != 1 {
		t.Errorf("CompatSetting count = %d, want 1", len(s.Compat.CompatSetting))
	}
	if len(s.Compat.Extra) != 1 {
		t.Errorf("Compat.Extra count = %d, want 1", len(s.Compat.Extra))
	}
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

func extraNames(extras []shared.RawXML) []string {
	names := make([]string, len(extras))
	for i, r := range extras {
		names[i] = r.XMLName.Local
	}
	return names
}
