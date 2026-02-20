package body

import (
	"encoding/xml"
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// Minimal document.xml from reference-appendix §2.5 (simplified).
// ---------------------------------------------------------------------------

const minimalDocumentXML = `<w:document ` +
	`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
	`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
	`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
	`xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ` +
	`mc:Ignorable="w14">` +
	`<w:body>` +
	`<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>` +
	`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
	`</w:body>` +
	`</w:document>`

func TestDocumentUnmarshal(t *testing.T) {
	var doc CT_Document
	if err := xml.Unmarshal([]byte(minimalDocumentXML), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}

	// Namespaces preserved
	if len(doc.Namespaces) == 0 {
		t.Fatal("expected namespace declarations, got none")
	}
	foundW := false
	for _, a := range doc.Namespaces {
		if a.Value == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" {
			foundW = true
		}
	}
	if !foundW {
		t.Error("expected xmlns:w in Namespaces")
	}

	// Body present
	if doc.Body == nil {
		t.Fatal("Body is nil")
	}

	// One paragraph
	if len(doc.Body.Content) != 1 {
		t.Fatalf("expected 1 content element, got %d", len(doc.Body.Content))
	}
	pe, ok := doc.Body.Content[0].(ParagraphElement)
	if !ok {
		t.Fatalf("expected ParagraphElement, got %T", doc.Body.Content[0])
	}
	if pe.P == nil {
		t.Fatal("ParagraphElement.P is nil")
	}

	// SectPr separated
	if doc.Body.SectPr == nil {
		t.Fatal("SectPr is nil")
	}
}

// ---------------------------------------------------------------------------
// Document with multiple block types + unknown extension element.
// ---------------------------------------------------------------------------

const richBodyXML = `<w:document ` +
	`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
	`xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
	`<w:body>` +
	`<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>` +
	`<w:tbl><w:tblPr/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
	`<w14:customBlock w14:val="test"><w14:inner>data</w14:inner></w14:customBlock>` +
	`<w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>` +
	`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>` +
	`</w:body>` +
	`</w:document>`

func TestBodyContent(t *testing.T) {
	var doc CT_Document
	if err := xml.Unmarshal([]byte(richBodyXML), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if doc.Body == nil {
		t.Fatal("Body is nil")
	}

	content := doc.Body.Content
	if len(content) != 4 {
		t.Fatalf("expected 4 content elements, got %d", len(content))
	}

	// 1. paragraph
	if _, ok := content[0].(ParagraphElement); !ok {
		t.Errorf("content[0]: expected ParagraphElement, got %T", content[0])
	}
	// 2. table
	if _, ok := content[1].(TableElement); !ok {
		t.Errorf("content[1]: expected TableElement, got %T", content[1])
	}
	// 3. unknown extension element
	raw, ok := content[2].(RawBlockElement)
	if !ok {
		t.Errorf("content[2]: expected RawBlockElement, got %T", content[2])
	} else {
		if raw.Raw.XMLName.Local != "customBlock" {
			t.Errorf("content[2]: expected customBlock, got %s", raw.Raw.XMLName.Local)
		}
	}
	// 4. paragraph
	if _, ok := content[3].(ParagraphElement); !ok {
		t.Errorf("content[3]: expected ParagraphElement, got %T", content[3])
	}

	// sectPr
	if doc.Body.SectPr == nil {
		t.Error("SectPr is nil")
	}
}

// ---------------------------------------------------------------------------
// Round-trip: unmarshal → marshal → re-unmarshal → compare.
// ---------------------------------------------------------------------------

func TestRoundTrip(t *testing.T) {
	// Step 1: Unmarshal
	var doc1 CT_Document
	if err := xml.Unmarshal([]byte(richBodyXML), &doc1); err != nil {
		t.Fatalf("Unmarshal (pass 1): %v", err)
	}

	// Step 2: Marshal
	output, err := xml.Marshal(&doc1)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	t.Logf("marshal output:\n%s", output)

	// Sanity: output must be parseable XML that contains the key local
	// names.  Go's encoder may use prefix ("w:p") or default-namespace
	// form, so we check for the local name rather than a specific prefix.
	outputStr := string(output)
	for _, local := range []string{"body", "sectPr", "customBlock"} {
		if !strings.Contains(outputStr, local) {
			t.Errorf("marshal output missing element %q", local)
		}
	}

	// Step 3: Re-unmarshal
	var doc2 CT_Document
	if err := xml.Unmarshal(output, &doc2); err != nil {
		t.Fatalf("Unmarshal (pass 2): %v", err)
	}

	// Step 4: Compare structure
	if doc2.Body == nil {
		t.Fatal("round-trip: Body is nil")
	}
	if len(doc2.Body.Content) != len(doc1.Body.Content) {
		t.Fatalf("round-trip: content length %d → %d",
			len(doc1.Body.Content), len(doc2.Body.Content))
	}

	// Same types in same order
	for i := range doc1.Body.Content {
		t1 := typeName(doc1.Body.Content[i])
		t2 := typeName(doc2.Body.Content[i])
		if t1 != t2 {
			t.Errorf("round-trip content[%d]: type %s → %s", i, t1, t2)
		}
	}

	// SectPr preserved
	if (doc1.Body.SectPr == nil) != (doc2.Body.SectPr == nil) {
		t.Error("round-trip: SectPr presence changed")
	}

	// Namespace count preserved
	if len(doc2.Namespaces) == 0 {
		t.Error("round-trip: lost namespace declarations")
	}

	// Extension element preserved
	raw1, ok1 := doc1.Body.Content[2].(RawBlockElement)
	raw2, ok2 := doc2.Body.Content[2].(RawBlockElement)
	if !ok1 || !ok2 {
		t.Error("round-trip: extension element type changed")
	} else {
		if raw1.Raw.XMLName.Local != raw2.Raw.XMLName.Local {
			t.Errorf("round-trip: extension element name %s → %s",
				raw1.Raw.XMLName.Local, raw2.Raw.XMLName.Local)
		}
	}
}

// ---------------------------------------------------------------------------
// SDT round-trip
// ---------------------------------------------------------------------------

const sdtDocXML = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
	`<w:body>` +
	`<w:sdt>` +
	`<w:sdtPr><w:alias w:val="Title"/></w:sdtPr>` +
	`<w:sdtContent>` +
	`<w:p><w:r><w:t>SDT content</w:t></w:r></w:p>` +
	`</w:sdtContent>` +
	`</w:sdt>` +
	`<w:sectPr/>` +
	`</w:body>` +
	`</w:document>`

func TestSdtRoundTrip(t *testing.T) {
	var doc CT_Document
	if err := xml.Unmarshal([]byte(sdtDocXML), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if doc.Body == nil || len(doc.Body.Content) != 1 {
		t.Fatalf("expected 1 content element, got %d",
			func() int {
				if doc.Body != nil {
					return len(doc.Body.Content)
				}
				return -1
			}())
	}

	sdtEl, ok := doc.Body.Content[0].(SdtBlockElement)
	if !ok {
		t.Fatalf("expected SdtBlockElement, got %T", doc.Body.Content[0])
	}
	sdt := sdtEl.Sdt
	if sdt.SdtPr == nil {
		t.Error("SdtPr is nil")
	}
	if len(sdt.SdtContent) != 1 {
		t.Fatalf("SdtContent: expected 1 element, got %d", len(sdt.SdtContent))
	}
	if _, ok := sdt.SdtContent[0].(ParagraphElement); !ok {
		t.Errorf("SdtContent[0]: expected ParagraphElement, got %T", sdt.SdtContent[0])
	}

	// Round-trip
	out, err := xml.Marshal(&doc)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	t.Logf("sdt marshal output:\n%s", out)
	var doc2 CT_Document
	if err := xml.Unmarshal(out, &doc2); err != nil {
		t.Fatalf("Unmarshal pass 2: %v", err)
	}
	if doc2.Body == nil || len(doc2.Body.Content) != 1 {
		t.Fatal("round-trip lost SDT")
	}
	sdtEl2, ok := doc2.Body.Content[0].(SdtBlockElement)
	if !ok {
		t.Fatalf("round-trip: expected SdtBlockElement, got %T", doc2.Body.Content[0])
	}
	if sdtEl2.Sdt.SdtPr == nil {
		t.Error("round-trip lost SdtPr")
	}
	if len(sdtEl2.Sdt.SdtContent) != 1 {
		t.Errorf("round-trip: SdtContent length %d → %d",
			len(sdt.SdtContent), len(sdtEl2.Sdt.SdtContent))
	}
}

// ---------------------------------------------------------------------------
// Parse / Serialize functions
// ---------------------------------------------------------------------------

func TestParseAndSerialize(t *testing.T) {
	doc, err := Parse([]byte(minimalDocumentXML))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if doc.Body == nil {
		t.Fatal("Parse returned nil Body")
	}

	data, err := Serialize(doc)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}
	if len(data) == 0 {
		t.Fatal("Serialize returned empty data")
	}

	// Should contain the XML header
	if !strings.HasPrefix(string(data), "<?xml") {
		t.Error("Serialize output missing XML declaration")
	}

	// Re-parse
	doc2, err := Parse(data)
	if err != nil {
		t.Fatalf("Parse round-trip: %v", err)
	}
	if doc2.Body == nil {
		t.Fatal("round-trip Parse returned nil Body")
	}
	if len(doc2.Body.Content) != len(doc.Body.Content) {
		t.Errorf("round-trip content count %d → %d",
			len(doc.Body.Content), len(doc2.Body.Content))
	}
}

// ---------------------------------------------------------------------------
// Empty body
// ---------------------------------------------------------------------------

func TestEmptyBody(t *testing.T) {
	input := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body><w:sectPr/></w:body></w:document>`
	var doc CT_Document
	if err := xml.Unmarshal([]byte(input), &doc); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if doc.Body == nil {
		t.Fatal("Body is nil")
	}
	if len(doc.Body.Content) != 0 {
		t.Errorf("expected 0 content, got %d", len(doc.Body.Content))
	}
	if doc.Body.SectPr == nil {
		t.Error("SectPr is nil for empty body")
	}
}

// ---------------------------------------------------------------------------
// New document (no namespaces stored → uses defaults)
// ---------------------------------------------------------------------------

func TestNewDocumentDefaults(t *testing.T) {
	doc := &CT_Document{
		Body: &CT_Body{},
	}

	out, err := xml.Marshal(doc)
	if err != nil {
		t.Fatalf("Marshal new doc: %v", err)
	}
	outStr := string(out)

	// Should contain default namespace declaration for w:
	if !strings.Contains(outStr, "http://schemas.openxmlformats.org/wordprocessingml/2006/main") {
		t.Error("new document missing w: namespace")
	}
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

func typeName(el interface{}) string {
	switch el.(type) {
	case ParagraphElement:
		return "ParagraphElement"
	case TableElement:
		return "TableElement"
	case SdtBlockElement:
		return "SdtBlockElement"
	case RawBlockElement:
		return "RawBlockElement"
	default:
		return "unknown"
	}
}
