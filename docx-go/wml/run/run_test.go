package run

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
)

// ---------------------------------------------------------------------------
// Helper: round-trip a CT_R through marshal → unmarshal and return both.
// ---------------------------------------------------------------------------

func roundTrip(t *testing.T, input string) (*CT_R, *CT_R, string) {
	t.Helper()

	var r1 CT_R
	if err := xml.Unmarshal([]byte(input), &r1); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	out, err := xml.Marshal(&r1)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var r2 CT_R
	if err := xml.Unmarshal(out, &r2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	return &r1, &r2, string(out)
}

// ---------------------------------------------------------------------------
// Test: Simple text run
// ---------------------------------------------------------------------------

func TestRoundTripSimpleText(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:t>Hello World</w:t>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	if len(r1.Content) != 1 {
		t.Fatalf("expected 1 content element, got %d", len(r1.Content))
	}
	txt, ok := r1.Content[0].(*CT_Text)
	if !ok {
		t.Fatalf("expected *CT_Text, got %T", r1.Content[0])
	}
	if txt.Value != "Hello World" {
		t.Errorf("text = %q, want %q", txt.Value, "Hello World")
	}

	// Verify round-trip preserves text.
	if len(r2.Content) != 1 {
		t.Fatalf("round-trip: expected 1 content, got %d", len(r2.Content))
	}
	txt2, ok := r2.Content[0].(*CT_Text)
	if !ok {
		t.Fatalf("round-trip: expected *CT_Text, got %T", r2.Content[0])
	}
	if txt2.Value != txt.Value {
		t.Errorf("round-trip lost text: got %q, want %q", txt2.Value, txt.Value)
	}
}

// ---------------------------------------------------------------------------
// Test: Text with xml:space="preserve"
// ---------------------------------------------------------------------------

func TestRoundTripTextWithSpace(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:t xml:space="preserve">This is </w:t>` +
		`</w:r>`

	r1, r2, out := roundTrip(t, input)

	txt := r1.Content[0].(*CT_Text)
	if txt.Space == nil || *txt.Space != "preserve" {
		t.Error("expected xml:space='preserve'")
	}
	if txt.Value != "This is " {
		t.Errorf("text = %q, want %q", txt.Value, "This is ")
	}

	// Output should contain xml:space="preserve".
	if !strings.Contains(out, "space") {
		t.Error("marshalled output missing xml:space attribute")
	}

	txt2 := r2.Content[0].(*CT_Text)
	if txt2.Space == nil || *txt2.Space != "preserve" {
		t.Error("round-trip lost xml:space attribute")
	}
}

// ---------------------------------------------------------------------------
// Test: Run with properties (bold, font size)
// ---------------------------------------------------------------------------

func TestRoundTripRunWithRPr(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidRPr="00C83215">` +
		`<w:rPr>` +
		`<w:b/>` +
		`<w:sz w:val="32"/>` +
		`<w:szCs w:val="32"/>` +
		`</w:rPr>` +
		`<w:t>Bold Text</w:t>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	// Check attributes.
	if r1.RsidRPr == nil || *r1.RsidRPr != "00C83215" {
		t.Error("rsidRPr not parsed")
	}

	// Check rPr.
	if r1.RPr == nil {
		t.Fatal("rPr not parsed")
	}
	if !r1.RPr.Base.B.Bool(false) {
		t.Error("bold not set")
	}
	if r1.RPr.Base.Sz == nil || r1.RPr.Base.Sz.Val != 32 {
		t.Error("sz not parsed")
	}

	// Check text content.
	if len(r1.Content) != 1 {
		t.Fatalf("expected 1 content, got %d", len(r1.Content))
	}
	if r1.Content[0].(*CT_Text).Value != "Bold Text" {
		t.Error("text mismatch")
	}

	// Round-trip checks.
	if r2.RsidRPr == nil || *r2.RsidRPr != *r1.RsidRPr {
		t.Error("round-trip lost rsidRPr")
	}
	if r2.RPr == nil || !r2.RPr.Base.B.Bool(false) {
		t.Error("round-trip lost bold property")
	}
}

// ---------------------------------------------------------------------------
// Test: Break element
// ---------------------------------------------------------------------------

func TestRoundTripBreak(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:br w:type="page"/>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	if len(r1.Content) != 1 {
		t.Fatalf("expected 1 content, got %d", len(r1.Content))
	}
	br, ok := r1.Content[0].(*CT_Br)
	if !ok {
		t.Fatalf("expected *CT_Br, got %T", r1.Content[0])
	}
	if br.Type == nil || *br.Type != "page" {
		t.Error("break type not 'page'")
	}

	br2 := r2.Content[0].(*CT_Br)
	if br2.Type == nil || *br2.Type != *br.Type {
		t.Error("round-trip lost break type")
	}
}

// ---------------------------------------------------------------------------
// Test: Field char sequence (begin/separate/end)
// ---------------------------------------------------------------------------

func TestRoundTripFieldChars(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:fldChar w:fldCharType="begin"/>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	fc := r1.Content[0].(*CT_FldChar)
	if fc.FldCharType != "begin" {
		t.Errorf("fldCharType = %q, want 'begin'", fc.FldCharType)
	}

	fc2 := r2.Content[0].(*CT_FldChar)
	if fc2.FldCharType != fc.FldCharType {
		t.Error("round-trip lost fldCharType")
	}
}

// ---------------------------------------------------------------------------
// Test: InstrText
// ---------------------------------------------------------------------------

func TestRoundTripInstrText(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:instrText xml:space="preserve"> PAGE </w:instrText>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	it := r1.Content[0].(*CT_InstrText)
	if it.Value != " PAGE " {
		t.Errorf("instrText = %q, want ' PAGE '", it.Value)
	}
	if it.Space == nil || *it.Space != "preserve" {
		t.Error("xml:space not preserved")
	}

	it2 := r2.Content[0].(*CT_InstrText)
	if it2.Value != it.Value {
		t.Error("round-trip lost instrText value")
	}
}

// ---------------------------------------------------------------------------
// Test: Sym (symbol)
// ---------------------------------------------------------------------------

func TestRoundTripSym(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:sym w:font="Wingdings" w:char="F0E0"/>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	sym := r1.Content[0].(*CT_Sym)
	if sym.Font != "Wingdings" {
		t.Errorf("font = %q, want 'Wingdings'", sym.Font)
	}
	if sym.Char != "F0E0" {
		t.Errorf("char = %q, want 'F0E0'", sym.Char)
	}

	sym2 := r2.Content[0].(*CT_Sym)
	if sym2.Font != sym.Font || sym2.Char != sym.Char {
		t.Error("round-trip lost sym attributes")
	}
}

// ---------------------------------------------------------------------------
// Test: Empty elements (tab, cr, footnoteRef, etc.)
// ---------------------------------------------------------------------------

func TestRoundTripEmptyElements(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:tab/>` +
		`<w:cr/>` +
		`<w:noBreakHyphen/>` +
		`<w:footnoteRef/>` +
		`<w:continuationSeparator/>` +
		`<w:lastRenderedPageBreak/>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	if len(r1.Content) != 6 {
		t.Fatalf("expected 6 content elements, got %d", len(r1.Content))
	}

	expected := []string{"tab", "cr", "noBreakHyphen", "footnoteRef", "continuationSeparator", "lastRenderedPageBreak"}
	for i, want := range expected {
		empty, ok := r1.Content[i].(*CT_EmptyRunContent)
		if !ok {
			t.Errorf("[%d] expected *CT_EmptyRunContent, got %T", i, r1.Content[i])
			continue
		}
		if empty.XMLName.Local != want {
			t.Errorf("[%d] XMLName.Local = %q, want %q", i, empty.XMLName.Local, want)
		}
	}

	if len(r2.Content) != len(r1.Content) {
		t.Errorf("round-trip: content count %d, want %d", len(r2.Content), len(r1.Content))
	}
	for i, want := range expected {
		if e2, ok := r2.Content[i].(*CT_EmptyRunContent); ok {
			if e2.XMLName.Local != want {
				t.Errorf("round-trip [%d]: name = %q, want %q", i, e2.XMLName.Local, want)
			}
		}
	}
}

// ---------------------------------------------------------------------------
// Test: Footnote reference
// ---------------------------------------------------------------------------

func TestRoundTripFootnoteReference(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
		`<w:footnoteReference w:id="1"/>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	if len(r1.Content) != 1 {
		t.Fatalf("expected 1 content, got %d", len(r1.Content))
	}
	ref, ok := r1.Content[0].(*CT_FtnEdnRef)
	if !ok {
		t.Fatalf("expected *CT_FtnEdnRef, got %T", r1.Content[0])
	}
	if ref.ID != 1 {
		t.Errorf("id = %d, want 1", ref.ID)
	}
	if ref.XMLName.Local != "footnoteReference" {
		t.Errorf("XMLName.Local = %q, want 'footnoteReference'", ref.XMLName.Local)
	}

	ref2 := r2.Content[0].(*CT_FtnEdnRef)
	if ref2.ID != ref.ID {
		t.Error("round-trip lost footnote id")
	}
}

// ---------------------------------------------------------------------------
// Test: Unknown element preserved via RawXML
// ---------------------------------------------------------------------------

func TestRoundTripUnknownElement(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:t>text</w:t>` +
		`<w14:someExtension w14:val="custom">inner</w14:someExtension>` +
		`</w:r>`

	r1, r2, out := roundTrip(t, input)

	if len(r1.Content) != 2 {
		t.Fatalf("expected 2 content elements, got %d", len(r1.Content))
	}

	raw, ok := r1.Content[1].(*CT_RawRunContent)
	if !ok {
		t.Fatalf("expected *CT_RawRunContent, got %T", r1.Content[1])
	}
	if raw.Raw.XMLName.Local != "someExtension" {
		t.Errorf("raw local = %q, want 'someExtension'", raw.Raw.XMLName.Local)
	}

	// Check that marshalled output contains the extension element.
	if !strings.Contains(out, "someExtension") {
		t.Error("marshalled output missing unknown extension element")
	}
	if !strings.Contains(out, "inner") {
		t.Error("marshalled output missing inner content of unknown element")
	}

	// Round-trip preserves unknown elements.
	if len(r2.Content) != 2 {
		t.Fatalf("round-trip: expected 2 content, got %d", len(r2.Content))
	}
	raw2, ok := r2.Content[1].(*CT_RawRunContent)
	if !ok {
		t.Fatalf("round-trip: expected *CT_RawRunContent, got %T", r2.Content[1])
	}
	if raw2.Raw.XMLName.Local != "someExtension" {
		t.Error("round-trip lost unknown element name")
	}
}

// ---------------------------------------------------------------------------
// Test: Complex run from reference appendix 3.1
// (formatted run with rFonts, bold, color, sz, szCs)
// ---------------------------------------------------------------------------

func TestRoundTripFormattedRun(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:rsidRPr="00C83215">` +
		`<w:rPr>` +
		`<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>` +
		`<w:b/>` +
		`<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>` +
		`<w:sz w:val="32"/>` +
		`<w:szCs w:val="32"/>` +
		`</w:rPr>` +
		`<w:t>Document Title</w:t>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	// Attributes.
	if r1.RsidRPr == nil || *r1.RsidRPr != "00C83215" {
		t.Error("rsidRPr not parsed")
	}

	// rPr fields.
	if r1.RPr == nil {
		t.Fatal("rPr is nil")
	}
	b := r1.RPr.Base
	if b.RFonts == nil || (b.RFonts.Ascii == nil || *b.RFonts.Ascii != "Arial") {
		t.Error("rFonts.Ascii not 'Arial'")
	}
	if !b.B.Bool(false) {
		t.Error("bold not set")
	}
	if b.Color == nil || b.Color.Val != "2F5496" {
		t.Error("color val mismatch")
	}
	if b.Color.ThemeColor == nil || *b.Color.ThemeColor != "accent1" {
		t.Error("themeColor mismatch")
	}
	if b.Sz == nil || b.Sz.Val != 32 {
		t.Error("sz != 32")
	}

	// Content.
	txt := r1.Content[0].(*CT_Text)
	if txt.Value != "Document Title" {
		t.Errorf("text = %q", txt.Value)
	}

	// Round-trip.
	if r2.RPr == nil || !r2.RPr.Base.B.Bool(false) {
		t.Error("round-trip lost bold")
	}
	if r2.Content[0].(*CT_Text).Value != "Document Title" {
		t.Error("round-trip lost text")
	}
}

// ---------------------------------------------------------------------------
// Test: Field char PAGE sequence (from reference appendix 3.7)
// ---------------------------------------------------------------------------

func TestRoundTripPageFieldSequence(t *testing.T) {
	t.Parallel()

	// Each run in the PAGE field is separate; here we test one run at a time.
	runs := []string{
		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:t xml:space="preserve">Page </w:t></w:r>`,

		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:fldChar w:fldCharType="begin"/></w:r>`,

		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>`,

		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:fldChar w:fldCharType="separate"/></w:r>`,

		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:t>1</w:t></w:r>`,

		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
			`<w:fldChar w:fldCharType="end"/></w:r>`,
	}

	for i, input := range runs {
		var r CT_R
		if err := xml.Unmarshal([]byte(input), &r); err != nil {
			t.Fatalf("run[%d] unmarshal: %v", i, err)
		}
		out, err := xml.Marshal(&r)
		if err != nil {
			t.Fatalf("run[%d] marshal: %v", i, err)
		}
		var r2 CT_R
		if err := xml.Unmarshal(out, &r2); err != nil {
			t.Fatalf("run[%d] re-unmarshal: %v", i, err)
		}
		if len(r2.Content) != len(r.Content) {
			t.Errorf("run[%d] content count: got %d, want %d", i, len(r2.Content), len(r.Content))
		}
	}
}

// ---------------------------------------------------------------------------
// Test: Drawing element (inline image) round-trip
// ---------------------------------------------------------------------------

func TestRoundTripDrawing(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
		`<w:rPr><w:noProof/></w:rPr>` +
		`<w:drawing>` +
		`<wp:inline distT="0" distB="0" distL="0" distR="0">` +
		`<wp:extent cx="1828800" cy="1371600"/>` +
		`</wp:inline>` +
		`</w:drawing>` +
		`</w:r>`

	r1, r2, out := roundTrip(t, input)

	if len(r1.Content) != 1 {
		t.Fatalf("expected 1 content, got %d", len(r1.Content))
	}
	dr, ok := r1.Content[0].(*CT_Drawing)
	if !ok {
		t.Fatalf("expected *CT_Drawing, got %T", r1.Content[0])
	}
	if len(dr.Children) != 1 {
		t.Fatalf("expected 1 drawing child, got %d", len(dr.Children))
	}
	if dr.Children[0].XMLName.Local != "inline" {
		t.Errorf("child local = %q, want 'inline'", dr.Children[0].XMLName.Local)
	}

	// Output should contain extent.
	if !strings.Contains(out, "1828800") {
		t.Error("marshalled output missing extent cx")
	}

	// Round-trip.
	dr2 := r2.Content[0].(*CT_Drawing)
	if len(dr2.Children) != 1 || dr2.Children[0].XMLName.Local != "inline" {
		t.Error("round-trip lost drawing child")
	}
}

// ---------------------------------------------------------------------------
// Test: Footnote body run (footnoteRef + text)
// ---------------------------------------------------------------------------

func TestRoundTripFootnoteBodyRun(t *testing.T) {
	t.Parallel()

	// Run with footnoteRef (empty element).
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
		`<w:footnoteRef/>` +
		`</w:r>`

	r1, _, _ := roundTrip(t, input)

	if len(r1.Content) != 1 {
		t.Fatalf("expected 1 content, got %d", len(r1.Content))
	}
	_, ok := r1.Content[0].(*CT_EmptyRunContent)
	if !ok {
		t.Fatalf("expected *CT_EmptyRunContent, got %T", r1.Content[0])
	}
}

// ---------------------------------------------------------------------------
// Test: Multiple rsid attributes
// ---------------------------------------------------------------------------

func TestRoundTripRsidAttributes(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`w:rsidR="00F12AB3" w:rsidRPr="00D714A3" w:rsidDel="00A77B3E">` +
		`<w:t>test</w:t>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	check := func(name string, got *string, want string) {
		t.Helper()
		if got == nil || *got != want {
			actual := "<nil>"
			if got != nil {
				actual = *got
			}
			t.Errorf("%s = %s, want %s", name, actual, want)
		}
	}

	check("RsidR", r1.RsidR, "00F12AB3")
	check("RsidRPr", r1.RsidRPr, "00D714A3")
	check("RsidDel", r1.RsidDel, "00A77B3E")

	// Round-trip.
	check("RT RsidR", r2.RsidR, "00F12AB3")
	check("RT RsidRPr", r2.RsidRPr, "00D714A3")
	check("RT RsidDel", r2.RsidDel, "00A77B3E")
}

// ---------------------------------------------------------------------------
// Test: NewText auto-sets xml:space="preserve"
// ---------------------------------------------------------------------------

func TestNewTextPreserve(t *testing.T) {
	t.Parallel()

	tests := []struct {
		input     string
		wantSpace bool
	}{
		{"hello", false},
		{" leading", true},
		{"trailing ", true},
		{"\ttab", true},
		{"", false},
		{"no spaces", false},
	}

	for _, tt := range tests {
		ct := NewText(tt.input)
		hasSpace := ct.Space != nil
		if hasSpace != tt.wantSpace {
			t.Errorf("NewText(%q): space=%v, want %v", tt.input, hasSpace, tt.wantSpace)
		}
	}
}

// ---------------------------------------------------------------------------
// Test: Factory registration works
// ---------------------------------------------------------------------------

func TestRunContentFactory(t *testing.T) {
	t.Parallel()

	tests := []struct {
		local    string
		wantType string
		wantNil  bool
	}{
		{"t", "*run.CT_Text", false},
		{"br", "*run.CT_Br", false},
		{"drawing", "*run.CT_Drawing", false},
		{"fldChar", "*run.CT_FldChar", false},
		{"instrText", "*run.CT_InstrText", false},
		{"sym", "*run.CT_Sym", false},
		{"footnoteReference", "*run.CT_FtnEdnRef", false},
		{"endnoteReference", "*run.CT_FtnEdnRef", false},
		{"tab", "*run.CT_EmptyRunContent", false},
		{"cr", "*run.CT_EmptyRunContent", false},
		{"pgNum", "*run.CT_EmptyRunContent", false},
		{"separator", "*run.CT_EmptyRunContent", false},
		{"unknownElement", "", true},
	}

	for _, tt := range tests {
		name := xml.Name{Local: tt.local}
		got := runContentFactory(name)
		if tt.wantNil {
			if got != nil {
				t.Errorf("runContentFactory(%q) = %T, want nil", tt.local, got)
			}
		} else {
			if got == nil {
				t.Errorf("runContentFactory(%q) = nil, want non-nil", tt.local)
			}
		}
	}
}

// ---------------------------------------------------------------------------
// Test: CT_R paragraph content factory
// ---------------------------------------------------------------------------

func TestParagraphContentFactory(t *testing.T) {
	t.Parallel()

	name := xml.Name{
		Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		Local: "r",
	}
	el := shared.CreateParagraphContent(name)
	if el == nil {
		t.Fatal("expected non-nil for <w:r>")
	}
	if _, ok := el.(*CT_R); !ok {
		t.Errorf("expected *CT_R, got %T", el)
	}
}

// ---------------------------------------------------------------------------
// Test: Empty run (no content, no rPr)
// ---------------------------------------------------------------------------

func TestRoundTripEmptyRun(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:r>`

	r1, r2, _ := roundTrip(t, input)

	if r1.RPr != nil {
		t.Error("expected nil rPr")
	}
	if len(r1.Content) != 0 {
		t.Errorf("expected 0 content, got %d", len(r1.Content))
	}
	if r2.RPr != nil || len(r2.Content) != 0 {
		t.Error("round-trip changed empty run")
	}
}

// ---------------------------------------------------------------------------
// Test: Mixed content (text + break + fldChar + empty elements)
// ---------------------------------------------------------------------------

func TestRoundTripMixedContent(t *testing.T) {
	t.Parallel()

	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:t xml:space="preserve">Hello </w:t>` +
		`<w:tab/>` +
		`<w:t>World</w:t>` +
		`<w:br w:type="column"/>` +
		`<w:fldChar w:fldCharType="begin"/>` +
		`<w:cr/>` +
		`</w:r>`

	r1, r2, _ := roundTrip(t, input)

	if len(r1.Content) != 6 {
		t.Fatalf("expected 6 content elements, got %d", len(r1.Content))
	}

	// Verify types in order.
	types := []string{"*run.CT_Text", "*run.CT_EmptyRunContent", "*run.CT_Text",
		"*run.CT_Br", "*run.CT_FldChar", "*run.CT_EmptyRunContent"}
	for i, want := range types {
		got := typeName(r1.Content[i])
		if got != want {
			t.Errorf("content[%d]: type = %s, want %s", i, got, want)
		}
	}

	if len(r2.Content) != len(r1.Content) {
		t.Errorf("round-trip: content count %d, want %d", len(r2.Content), len(r1.Content))
	}
}

func typeName(v interface{}) string {
	switch v.(type) {
	case *CT_Text:
		return "*run.CT_Text"
	case *CT_Br:
		return "*run.CT_Br"
	case *CT_Drawing:
		return "*run.CT_Drawing"
	case *CT_FldChar:
		return "*run.CT_FldChar"
	case *CT_InstrText:
		return "*run.CT_InstrText"
	case *CT_Sym:
		return "*run.CT_Sym"
	case *CT_FtnEdnRef:
		return "*run.CT_FtnEdnRef"
	case *CT_EmptyRunContent:
		return "*run.CT_EmptyRunContent"
	case *CT_RawRunContent:
		return "*run.CT_RawRunContent"
	case shared.RawXML:
		return "shared.RawXML"
	default:
		return "unknown"
	}
}
