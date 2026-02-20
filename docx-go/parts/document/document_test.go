package document

import (
	"bytes"
	"encoding/xml"
	"io"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/body"
)

// sampleDocumentXML is taken from reference-appendix.md § 2.5.
// It represents a minimal valid word/document.xml produced by MS Word.
const sampleDocumentXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14 w15 wp14">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="77777777"
         w:rsidR="00000001" w:rsidRDefault="00000001">
      <w:r>
        <w:t>Hello World</w:t>
      </w:r>
    </w:p>
    <w:sectPr w:rsidR="00000001">
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`

// TestParseReturnsNonNil verifies that Parse produces a non-nil CT_Document
// with a non-nil Body from well-formed input.
func TestParseReturnsNonNil(t *testing.T) {
	doc, err := Parse([]byte(sampleDocumentXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if doc == nil {
		t.Fatal("Parse returned nil document")
	}
	if doc.Body == nil {
		t.Fatal("Parse returned nil Body")
	}
}

// TestParseBody verifies that the body contains expected content:
// at least one block-level element (the paragraph) and a sectPr.
func TestParseBody(t *testing.T) {
	doc, err := Parse([]byte(sampleDocumentXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(doc.Body.Content) == 0 {
		t.Error("expected at least one block-level element in Body.Content")
	}
	if doc.Body.SectPr == nil {
		t.Error("expected body-level SectPr to be parsed")
	}
}

// TestParseNamespacePreservation verifies that namespace declarations from
// the root <w:document> element are preserved during unmarshal (see patterns.md § 6).
func TestParseNamespacePreservation(t *testing.T) {
	doc, err := Parse([]byte(sampleDocumentXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(doc.Namespaces) == 0 {
		t.Fatal("expected Namespaces to be preserved from root element attributes")
	}

	// Check that at least the main WML namespace is among the preserved attrs
	found := false
	for _, attr := range doc.Namespaces {
		if attr.Value == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" {
			found = true
			break
		}
	}
	if !found {
		t.Error("expected w: namespace to be preserved in Namespaces")
	}
}

// TestSerializeIncludesXMLHeader verifies that Serialize prepends the
// standard XML declaration.
func TestSerializeIncludesXMLHeader(t *testing.T) {
	doc, err := Parse([]byte(sampleDocumentXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	out, err := Serialize(doc)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	if !bytes.HasPrefix(out, []byte(`<?xml version="1.0"`)) {
		t.Error("serialized output must start with XML declaration")
	}
}

// TestRoundTrip is the canonical round-trip test (patterns.md § 12):
//
//	unmarshal(XML) → marshal → unmarshal again → compare
//
// It verifies that Parse → Serialize → Parse produces structurally
// equivalent output, ensuring no data is lost.
func TestRoundTrip(t *testing.T) {
	// --- Pass 1: parse original XML ---
	doc1, err := Parse([]byte(sampleDocumentXML))
	if err != nil {
		t.Fatalf("Pass 1 Parse failed: %v", err)
	}

	// --- Marshal back to bytes ---
	out, err := Serialize(doc1)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// --- Pass 2: parse the re-serialized output ---
	doc2, err := Parse(out)
	if err != nil {
		t.Fatalf("Pass 2 Parse failed: %v", err)
	}

	// --- Compare structural properties ---

	// Both must have a body
	if doc2.Body == nil {
		t.Fatal("round-trip lost Body")
	}

	// Same number of block-level elements
	if len(doc2.Body.Content) != len(doc1.Body.Content) {
		t.Errorf("round-trip block content count: got %d, want %d",
			len(doc2.Body.Content), len(doc1.Body.Content))
	}

	// SectPr must survive
	if (doc1.Body.SectPr == nil) != (doc2.Body.SectPr == nil) {
		t.Error("round-trip lost or gained SectPr")
	}

	// Namespace declarations must survive
	if len(doc2.Namespaces) == 0 {
		t.Error("round-trip lost Namespaces")
	}
}

// TestRoundTripWithUnknownElements verifies that unknown (extension) elements
// at the <w:document> level survive a round-trip via the Extra field and
// shared.RawXML (patterns.md § 4).
func TestRoundTripWithUnknownElements(t *testing.T) {
	// Inject a custom extension element inside <w:body> that is unknown
	// to the parser — it should be captured as RawXML and round-tripped.
	input := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t>Test</w:t></w:r>
    </w:p>
    <w14:customExtension w14:val="preserved"/>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>`

	doc, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	out, err := Serialize(doc)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// The serialized output should contain the extension element
	if !strings.Contains(string(out), "customExtension") {
		t.Error("round-trip lost unknown extension element <w14:customExtension>")
	}
}

// TestParseInvalidXML verifies that Parse returns a meaningful error on
// malformed input.
func TestParseInvalidXML(t *testing.T) {
	_, err := Parse([]byte(`<not valid xml`))
	if err == nil {
		t.Error("expected error for invalid XML, got nil")
	}
}

// TestSerializeEmpty verifies that Serialize handles a zero-value
// CT_Document (new document scenario) without panicking.
func TestSerializeEmpty(t *testing.T) {
	doc := &body.CT_Document{}
	out, err := Serialize(doc)
	if err != nil {
		t.Fatalf("Serialize empty doc failed: %v", err)
	}
	if len(out) == 0 {
		t.Error("expected non-empty output for empty document")
	}

	// Must still have XML declaration
	if !bytes.HasPrefix(out, []byte("<?xml")) {
		t.Error("empty doc output missing XML declaration")
	}
}

// TestSerializeRoundTripXMLValidity checks that the output of Serialize is
// valid XML by attempting to tokenize the entire result.
func TestSerializeRoundTripXMLValidity(t *testing.T) {
	doc, err := Parse([]byte(sampleDocumentXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	out, err := Serialize(doc)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// Tokenize the entire output — any XML error will surface here
	dec := xml.NewDecoder(bytes.NewReader(out))
	for {
		_, err := dec.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			t.Fatalf("serialized output is not valid XML: %v", err)
		}
	}
}
