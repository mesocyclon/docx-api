package hdft

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// stubParagraph is a tiny test-only stand-in for CT_P.  It round-trips the
// inner XML verbatim so we can verify that hdft preserves content.
type stubParagraph struct {
	shared.BlockLevelMarker
	InnerXML []byte `xml:",innerxml"`
}

func (p *stubParagraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	st := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "p"}}
	if err := e.EncodeToken(st); err != nil {
		return err
	}
	// Replay inner tokens.
	if len(p.InnerXML) > 0 {
		dec := xml.NewDecoder(bytes.NewReader(p.InnerXML))
		for {
			tok, err := dec.Token()
			if err != nil {
				break
			}
			if err := e.EncodeToken(xml.CopyToken(tok)); err != nil {
				return err
			}
		}
	}
	return e.EncodeToken(st.End())
}

// registerStubFactory registers a block-level factory that recognises <w:p>
// elements and returns a *stubParagraph.
func registerStubFactory() {
	shared.RegisterBlockFactory(func(name xml.Name) shared.BlockLevelElement {
		if name.Local == "p" {
			return &stubParagraph{}
		}
		return nil
	})
}

func init() {
	registerStubFactory()
}

// =========================================================================
// Tests
// =========================================================================

// TestHeaderRoundTrip parses a realistic <w:hdr> from the reference appendix,
// marshals it back, and verifies the result parses identically.
func TestHeaderRoundTrip(t *testing.T) {
	input := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<w:p>` +
		`<w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr>` +
		`<w:r><w:t xml:space="preserve">Page </w:t></w:r>` +
		`<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
		`<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
		`<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
		`<w:r><w:t>1</w:t></w:r>` +
		`<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
		`</w:p>` +
		`</w:hdr>`

	hdr, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(hdr.Content) != 1 {
		t.Fatalf("expected 1 block element, got %d", len(hdr.Content))
	}

	// Marshal back.
	out, err := Serialize(hdr, "w:hdr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// Re-parse the output.
	hdr2, err := Parse(out)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}

	if len(hdr2.Content) != len(hdr.Content) {
		t.Errorf("round-trip lost block elements: %d → %d",
			len(hdr.Content), len(hdr2.Content))
	}
}

// TestFooterRoundTrip verifies the footer example from the reference appendix.
func TestFooterRoundTrip(t *testing.T) {
	input := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p>` +
		`<w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>` +
		`<w:r><w:t>Confidential</w:t></w:r>` +
		`</w:p>` +
		`</w:ftr>`

	ftr, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(ftr.Content) != 1 {
		t.Fatalf("expected 1 block element, got %d", len(ftr.Content))
	}

	out, err := Serialize(ftr, "w:ftr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// The output should contain our footer content.
	if !strings.Contains(string(out), "Confidential") {
		t.Error("Serialize output does not contain 'Confidential'")
	}
	if !strings.Contains(string(out), "w:ftr") {
		t.Error("Serialize output does not use w:ftr root element")
	}

	// Re-parse.
	ftr2, err := Parse(out)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}
	if len(ftr2.Content) != 1 {
		t.Errorf("round-trip lost block elements: want 1, got %d", len(ftr2.Content))
	}
}

// TestUnknownElementsPreserved verifies that unrecognised XML elements
// are captured as shared.RawXML and survive the round-trip.
func TestUnknownElementsPreserved(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>` +
		`<w14:someExtension w14:val="test"><w14:child>data</w14:child></w14:someExtension>` +
		`</w:hdr>`

	hdr, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(hdr.Content) != 2 {
		t.Fatalf("expected 2 block elements (1 para + 1 raw), got %d", len(hdr.Content))
	}

	// First element should be a paragraph.
	if _, ok := hdr.Content[0].(*stubParagraph); !ok {
		t.Errorf("Content[0]: expected *stubParagraph, got %T", hdr.Content[0])
	}

	// Second element should be RawXML.
	raw, ok := hdr.Content[1].(shared.RawXML)
	if !ok {
		t.Fatalf("Content[1]: expected shared.RawXML, got %T", hdr.Content[1])
	}
	if raw.XMLName.Local != "someExtension" {
		t.Errorf("RawXML local name: want 'someExtension', got %q", raw.XMLName.Local)
	}

	// Round-trip.
	out, err := Serialize(hdr, "w:hdr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	if !strings.Contains(string(out), "someExtension") {
		t.Error("round-trip lost the extension element name")
	}
	if !strings.Contains(string(out), "data") {
		t.Error("round-trip lost the extension element inner content")
	}

	// Re-parse and verify RawXML survived.
	hdr2, err := Parse(out)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}
	if len(hdr2.Content) != 2 {
		t.Fatalf("re-Parse: expected 2 block elements, got %d", len(hdr2.Content))
	}
	raw2, ok := hdr2.Content[1].(shared.RawXML)
	if !ok {
		t.Fatalf("re-Parse Content[1]: expected shared.RawXML, got %T", hdr2.Content[1])
	}
	if raw2.XMLName.Local != "someExtension" {
		t.Errorf("re-Parse RawXML local name: want 'someExtension', got %q", raw2.XMLName.Local)
	}
}

// TestMultipleParagraphs verifies headers with several block elements.
func TestMultipleParagraphs(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p><w:r><w:t>First</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Second</w:t></w:r></w:p>` +
		`<w:p><w:r><w:t>Third</w:t></w:r></w:p>` +
		`</w:hdr>`

	hdr, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(hdr.Content) != 3 {
		t.Fatalf("expected 3 paragraphs, got %d", len(hdr.Content))
	}

	out, err := Serialize(hdr, "w:hdr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	hdr2, err := Parse(out)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}
	if len(hdr2.Content) != 3 {
		t.Errorf("round-trip element count: want 3, got %d", len(hdr2.Content))
	}
}

// TestEmptyContent verifies an empty header (edge case — spec says ≥1 <w:p>
// but we should not crash).
func TestEmptyContent(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:hdr>`

	hdr, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(hdr.Content) != 0 {
		t.Fatalf("expected 0 block elements, got %d", len(hdr.Content))
	}

	out, err := Serialize(hdr, "w:hdr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	if !strings.Contains(string(out), "w:hdr") {
		t.Error("serialized output missing root element")
	}
}

// TestNamespacesPreserved verifies that xmlns attributes from the original
// root element survive the round-trip.
func TestNamespacesPreserved(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
		` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:p/></w:hdr>`

	hdr, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	if len(hdr.Namespaces) == 0 {
		t.Fatal("expected namespace attributes to be captured")
	}

	// Serialize and verify namespace declarations are present.
	out, err := Serialize(hdr, "w:hdr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	outStr := string(out)
	for _, ns := range []string{
		"schemas.openxmlformats.org/wordprocessingml/2006/main",
		"schemas.openxmlformats.org/officeDocument/2006/relationships",
		"schemas.microsoft.com/office/word/2010/wordml",
	} {
		if !strings.Contains(outStr, ns) {
			t.Errorf("serialized output missing namespace %q", ns)
		}
	}
}

// TestDefaultNamespacesForNewDocument verifies that a freshly-created
// CT_HdrFtr (no prior unmarshal) gets sensible default xmlns attributes.
func TestDefaultNamespacesForNewDocument(t *testing.T) {
	hf := &CT_HdrFtr{}
	out, err := Serialize(hf, "w:hdr")
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	outStr := string(out)
	if !strings.Contains(outStr, xmltypes.NSw) {
		t.Errorf("default namespace missing w: %q", outStr)
	}
}

// TestSerializeRootName checks that Serialize respects the rootName argument.
func TestSerializeRootName(t *testing.T) {
	hf := &CT_HdrFtr{}

	for _, root := range []string{"w:hdr", "w:ftr"} {
		out, err := Serialize(hf, root)
		if err != nil {
			t.Fatalf("Serialize(%q) failed: %v", root, err)
		}
		if !strings.Contains(string(out), "<"+root) {
			t.Errorf("Serialize(%q): output does not start with <%s", root, root)
		}
	}
}
