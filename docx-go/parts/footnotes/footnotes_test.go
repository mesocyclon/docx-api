package footnotes

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
)

// referenceXML is the footnotes.xml example from reference-appendix.md.
// It covers all three footnote types: separator, continuationSeparator, and
// a user footnote with content paragraphs.
const referenceXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
	"\n" + `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
	`<w:footnote w:type="separator" w:id="-1">` +
	`<w:p>` +
	`<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
	`<w:r><w:separator/></w:r>` +
	`</w:p>` +
	`</w:footnote>` +
	`<w:footnote w:type="continuationSeparator" w:id="0">` +
	`<w:p>` +
	`<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
	`<w:r><w:continuationSeparator/></w:r>` +
	`</w:p>` +
	`</w:footnote>` +
	`<w:footnote w:id="1">` +
	`<w:p>` +
	`<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
	`<w:r>` +
	`<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
	`<w:footnoteRef/>` +
	`</w:r>` +
	`<w:r>` +
	`<w:t xml:space="preserve"> See the original source for details.</w:t>` +
	`</w:r>` +
	`</w:p>` +
	`</w:footnote>` +
	`</w:footnotes>`

// TestRoundTrip verifies that Parse → Serialize produces output that can be
// parsed again with identical structure. Because no block-level factories are
// registered in this test, all <w:p> children are stored as RawXML — which
// is the correct behaviour for a part module that does not depend on wml/para.
func TestRoundTrip(t *testing.T) {
	t.Parallel()

	// 1. Parse reference XML.
	fn, err := Parse([]byte(referenceXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// 2. Verify structure.
	if got := len(fn.Footnote); got != 3 {
		t.Fatalf("expected 3 footnotes, got %d", got)
	}

	// Footnote 0: separator, id=-1
	assertFtnEdn(t, fn.Footnote[0], strPtr("separator"), -1)
	// Footnote 1: continuationSeparator, id=0
	assertFtnEdn(t, fn.Footnote[1], strPtr("continuationSeparator"), 0)
	// Footnote 2: no type, id=1
	assertFtnEdn(t, fn.Footnote[2], nil, 1)

	// Each footnote has 1 block-level child (<w:p>).
	for i, fe := range fn.Footnote {
		if got := len(fe.Content); got != 1 {
			t.Errorf("footnote[%d]: expected 1 content element, got %d", i, got)
		}
	}

	// 3. Serialize.
	out, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// 4. Re-parse the serialized output.
	fn2, err := Parse(out)
	if err != nil {
		t.Fatalf("re-Parse failed: %v\nXML:\n%s", err, string(out))
	}

	// 5. Compare structure.
	if len(fn2.Footnote) != len(fn.Footnote) {
		t.Fatalf("round-trip: footnote count %d != %d", len(fn2.Footnote), len(fn.Footnote))
	}
	for i := range fn.Footnote {
		a, b := fn.Footnote[i], fn2.Footnote[i]
		if ptrStr(a.Type) != ptrStr(b.Type) {
			t.Errorf("footnote[%d] type: %q != %q", i, ptrStr(a.Type), ptrStr(b.Type))
		}
		if a.ID != b.ID {
			t.Errorf("footnote[%d] id: %d != %d", i, a.ID, b.ID)
		}
		if len(a.Content) != len(b.Content) {
			t.Errorf("footnote[%d] content count: %d != %d", i, len(a.Content), len(b.Content))
		}
	}

	// 6. Verify key content survives round-trip by checking the serialized
	//    output contains expected fragments.
	outStr := string(out)
	mustContain := []string{
		`w:type="separator"`,
		`w:type="continuationSeparator"`,
		`w:id="-1"`,
		`w:id="0"`,
		`w:id="1"`,
		`FootnoteText`,
		`FootnoteReference`,
		`footnoteRef`,
		`See the original source for details.`,
	}
	for _, want := range mustContain {
		if !strings.Contains(outStr, want) {
			t.Errorf("serialized output missing %q\nGot:\n%s", want, outStr)
		}
	}
}

// TestEndnotesRoundTrip verifies that the same types handle <w:endnotes>.
func TestEndnotesRoundTrip(t *testing.T) {
	t.Parallel()

	input := `<?xml version="1.0" encoding="UTF-8"?>` +
		`<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:endnote w:type="separator" w:id="-1">` +
		`<w:p><w:r><w:separator/></w:r></w:p>` +
		`</w:endnote>` +
		`<w:endnote w:id="1">` +
		`<w:p><w:r><w:t>Endnote text</w:t></w:r></w:p>` +
		`</w:endnote>` +
		`</w:endnotes>`

	fn, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse endnotes failed: %v", err)
	}

	if fn.rootLocal != "endnotes" {
		t.Errorf("rootLocal = %q, want %q", fn.rootLocal, "endnotes")
	}
	if fn.childLocal != "endnote" {
		t.Errorf("childLocal = %q, want %q", fn.childLocal, "endnote")
	}
	if len(fn.Footnote) != 2 {
		t.Fatalf("expected 2 endnotes, got %d", len(fn.Footnote))
	}

	out, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize endnotes failed: %v", err)
	}

	outStr := string(out)
	if !strings.Contains(outStr, "w:endnotes") {
		t.Error("serialized output missing root element 'w:endnotes'")
	}
	if !strings.Contains(outStr, "w:endnote") {
		t.Error("serialized output missing child element 'w:endnote'")
	}
	if !strings.Contains(outStr, "Endnote text") {
		t.Error("serialized output missing endnote text content")
	}
}

// TestParseEmpty verifies that an empty <w:footnotes/> element parses
// without error.
func TestParseEmpty(t *testing.T) {
	t.Parallel()

	input := `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	fn, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse empty: %v", err)
	}
	if len(fn.Footnote) != 0 {
		t.Errorf("expected 0 footnotes, got %d", len(fn.Footnote))
	}
}

// TestExtraNamespacesPreserved verifies that additional namespace declarations
// on the root element survive a round-trip.
func TestExtraNamespacesPreserved(t *testing.T) {
	t.Parallel()

	input := `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
		`xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:footnote w:type="separator" w:id="-1">` +
		`<w:p><w:r><w:separator/></w:r></w:p>` +
		`</w:footnote>` +
		`</w:footnotes>`

	fn, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	out, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	outStr := string(out)
	if !strings.Contains(outStr, "xmlns:w14") {
		t.Errorf("w14 namespace lost in round-trip.\nGot:\n%s", outStr)
	}
}

// TestMarshalNewDocument verifies that a programmatically-constructed
// CT_Footnotes can be serialized.
func TestMarshalNewDocument(t *testing.T) {
	t.Parallel()

	sep := "separator"
	fn := &CT_Footnotes{
		Footnote: []CT_FtnEdn{
			{Type: &sep, ID: -1},
			{ID: 1},
		},
	}

	out, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	outStr := string(out)
	if !strings.Contains(outStr, "w:footnotes") {
		t.Error("missing root element")
	}
	if !strings.Contains(outStr, `w:type="separator"`) {
		t.Error("missing type attribute")
	}
	if !strings.Contains(outStr, `w:id="1"`) {
		t.Error("missing id=1")
	}
}

// TestUnknownBlockElementPreserved verifies that an unrecognised block-level
// element inside a footnote is preserved as RawXML through the round-trip.
func TestUnknownBlockElementPreserved(t *testing.T) {
	t.Parallel()

	input := `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:footnote w:id="1">` +
		`<w:customBlock w:val="test"><w:inner/></w:customBlock>` +
		`</w:footnote>` +
		`</w:footnotes>`

	fn, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	if len(fn.Footnote) != 1 {
		t.Fatalf("expected 1 footnote, got %d", len(fn.Footnote))
	}
	if len(fn.Footnote[0].Content) != 1 {
		t.Fatalf("expected 1 content element, got %d", len(fn.Footnote[0].Content))
	}

	// Should be stored as RawXML.
	raw, ok := fn.Footnote[0].Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("expected RawXML, got %T", fn.Footnote[0].Content[0])
	}
	if raw.XMLName.Local != "customBlock" {
		t.Errorf("RawXML local = %q, want %q", raw.XMLName.Local, "customBlock")
	}

	// Round-trip.
	out, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	outStr := string(out)
	if !strings.Contains(outStr, "customBlock") {
		t.Errorf("unknown element lost in round-trip.\nGot:\n%s", outStr)
	}
	if !strings.Contains(outStr, "inner") {
		t.Errorf("inner content of unknown element lost.\nGot:\n%s", outStr)
	}
}

// TestXMLContentRoundTrip performs a double round-trip (parse → serialize →
// parse → serialize) and verifies byte-level stability: the second
// serialization must be identical to the first.
func TestXMLContentRoundTrip(t *testing.T) {
	t.Parallel()

	fn, err := Parse([]byte(referenceXML))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	out, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	// Re-parse and re-serialize for stability check.
	fn2, err := Parse(out)
	if err != nil {
		t.Fatalf("re-Parse: %v", err)
	}
	out2, err := Serialize(fn2)
	if err != nil {
		t.Fatalf("re-Serialize: %v", err)
	}

	// The second round-trip must be byte-identical to the first.
	if string(out) != string(out2) {
		t.Errorf("second round-trip differs from first:\n--- first ---\n%s\n--- second ---\n%s",
			string(out), string(out2))
	}
}

// TestEndnotesDoubleRoundTrip verifies byte-level stability for endnotes.
func TestEndnotesDoubleRoundTrip(t *testing.T) {
	t.Parallel()

	input := `<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:endnote w:type="separator" w:id="-1">` +
		`<w:p><w:r><w:separator/></w:r></w:p>` +
		`</w:endnote>` +
		`<w:endnote w:id="1">` +
		`<w:p><w:r><w:t>Endnote.</w:t></w:r></w:p>` +
		`</w:endnote>` +
		`</w:endnotes>`

	fn, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	out1, err := Serialize(fn)
	if err != nil {
		t.Fatalf("Serialize 1: %v", err)
	}

	fn2, err := Parse(out1)
	if err != nil {
		t.Fatalf("re-Parse: %v", err)
	}
	out2, err := Serialize(fn2)
	if err != nil {
		t.Fatalf("Serialize 2: %v", err)
	}

	if string(out1) != string(out2) {
		t.Errorf("endnotes double round-trip unstable:\n--- first ---\n%s\n--- second ---\n%s",
			string(out1), string(out2))
	}
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

func strPtr(s string) *string { return &s }

func ptrStr(p *string) string {
	if p == nil {
		return "<nil>"
	}
	return *p
}

func assertFtnEdn(t *testing.T, fe CT_FtnEdn, wantType *string, wantID int) {
	t.Helper()
	if ptrStr(fe.Type) != ptrStr(wantType) {
		t.Errorf("type = %q, want %q", ptrStr(fe.Type), ptrStr(wantType))
	}
	if fe.ID != wantID {
		t.Errorf("id = %d, want %d", fe.ID, wantID)
	}
}

// Compile-time check: CT_Footnotes implements xml.Marshaler & xml.Unmarshaler.
var (
	_ xml.Marshaler   = (*CT_Footnotes)(nil)
	_ xml.Unmarshaler = (*CT_Footnotes)(nil)
)
