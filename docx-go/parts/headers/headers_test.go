package headers

import (
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
)

// Reference XML from reference-appendix.md §3.7
const headerXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
	`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
	` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
	`<w:p>` +
	`<w:pPr>` +
	`<w:pStyle w:val="Header"/>` +
	`<w:jc w:val="right"/>` +
	`</w:pPr>` +
	`<w:r>` +
	`<w:t xml:space="preserve">Page </w:t>` +
	`</w:r>` +
	`<w:r>` +
	`<w:fldChar w:fldCharType="begin"/>` +
	`</w:r>` +
	`<w:r>` +
	`<w:instrText xml:space="preserve"> PAGE </w:instrText>` +
	`</w:r>` +
	`<w:r>` +
	`<w:fldChar w:fldCharType="separate"/>` +
	`</w:r>` +
	`<w:r>` +
	`<w:t>1</w:t>` +
	`</w:r>` +
	`<w:r>` +
	`<w:fldChar w:fldCharType="end"/>` +
	`</w:r>` +
	`</w:p>` +
	`</w:hdr>`

const footerXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
	`<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
	`<w:p>` +
	`<w:pPr>` +
	`<w:pStyle w:val="Footer"/>` +
	`<w:jc w:val="center"/>` +
	`</w:pPr>` +
	`<w:r>` +
	`<w:t>Confidential</w:t>` +
	`</w:r>` +
	`</w:p>` +
	`</w:ftr>`

// ---------------------------------------------------------------------------
// Parse
// ---------------------------------------------------------------------------

func TestParseHeader(t *testing.T) {
	hf, err := Parse([]byte(headerXML))
	if err != nil {
		t.Fatalf("Parse header: %v", err)
	}
	if len(hf.Content) == 0 {
		t.Fatal("expected at least one content element")
	}
	// All children are RawXML because no block factories are registered.
	raw, ok := hf.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("expected RawXML, got %T", hf.Content[0])
	}
	if raw.XMLName.Local != "p" {
		t.Errorf("expected 'p', got %q", raw.XMLName.Local)
	}
}

func TestParseFooter(t *testing.T) {
	hf, err := Parse([]byte(footerXML))
	if err != nil {
		t.Fatalf("Parse footer: %v", err)
	}
	if len(hf.Content) != 1 {
		t.Errorf("expected 1 element, got %d", len(hf.Content))
	}
}

// ---------------------------------------------------------------------------
// Serialize — header
// ---------------------------------------------------------------------------

func TestSerializeHeader(t *testing.T) {
	hf, err := Parse([]byte(headerXML))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	data, err := Serialize(hf)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}
	output := string(data)
	if !strings.Contains(output, "<w:hdr") {
		t.Error("output must start with <w:hdr")
	}
	if !strings.Contains(output, "Page ") {
		t.Error("output must contain 'Page ' text")
	}
	if !strings.Contains(output, "fldCharType") {
		t.Error("output must contain field char elements")
	}
}

// ---------------------------------------------------------------------------
// Serialize — footer
// ---------------------------------------------------------------------------

func TestSerializeFooter(t *testing.T) {
	hf, err := Parse([]byte(footerXML))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	data, err := SerializeFooter(hf)
	if err != nil {
		t.Fatalf("SerializeFooter: %v", err)
	}
	output := string(data)
	if !strings.Contains(output, "<w:ftr") {
		t.Error("output must start with <w:ftr")
	}
	if !strings.Contains(output, "Confidential") {
		t.Error("output must contain 'Confidential'")
	}
}

// ---------------------------------------------------------------------------
// Round-trip: Parse → Serialize → Parse → compare
// ---------------------------------------------------------------------------

func TestHeaderRoundTrip(t *testing.T) {
	// Pass 1
	hf1, err := Parse([]byte(headerXML))
	if err != nil {
		t.Fatalf("Parse pass 1: %v", err)
	}

	// Serialize
	data1, err := Serialize(hf1)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	// Pass 2
	hf2, err := Parse(data1)
	if err != nil {
		t.Fatalf("Parse pass 2: %v", err)
	}

	// Compare element count
	if len(hf1.Content) != len(hf2.Content) {
		t.Fatalf("content count mismatch: %d vs %d",
			len(hf1.Content), len(hf2.Content))
	}

	// Compare element names
	for i := range hf1.Content {
		r1, ok1 := hf1.Content[i].(shared.RawXML)
		r2, ok2 := hf2.Content[i].(shared.RawXML)
		if ok1 != ok2 {
			t.Errorf("content[%d]: type mismatch", i)
			continue
		}
		if ok1 && r1.XMLName.Local != r2.XMLName.Local {
			t.Errorf("content[%d]: name mismatch: %q vs %q",
				i, r1.XMLName.Local, r2.XMLName.Local)
		}
	}

	// Serialize pass 2 → should be identical
	data2, err := Serialize(hf2)
	if err != nil {
		t.Fatalf("Serialize pass 2: %v", err)
	}

	// Structural comparison: both contain same key content.
	// Use namespace-agnostic needles because encoding/xml may choose
	// a different prefix than "w:" when marshalling typed elements
	// (CT_P, CT_R, etc.) that were created by registered block factories.
	s1, s2 := string(data1), string(data2)
	for _, needle := range []string{"w:hdr", "pStyle", "Header", "Page ", "fldCharType"} {
		if !strings.Contains(s1, needle) {
			t.Errorf("pass 1 output missing %q", needle)
		}
		if !strings.Contains(s2, needle) {
			t.Errorf("pass 2 output missing %q", needle)
		}
	}
}

func TestFooterRoundTrip(t *testing.T) {
	hf1, err := Parse([]byte(footerXML))
	if err != nil {
		t.Fatalf("Parse pass 1: %v", err)
	}

	data1, err := SerializeFooter(hf1)
	if err != nil {
		t.Fatalf("SerializeFooter: %v", err)
	}

	hf2, err := Parse(data1)
	if err != nil {
		t.Fatalf("Parse pass 2: %v", err)
	}

	if len(hf1.Content) != len(hf2.Content) {
		t.Fatalf("content count mismatch: %d vs %d",
			len(hf1.Content), len(hf2.Content))
	}

	data2, err := SerializeFooter(hf2)
	if err != nil {
		t.Fatalf("SerializeFooter pass 2: %v", err)
	}

	s1, s2 := string(data1), string(data2)
	for _, needle := range []string{"w:ftr", "Footer", "Confidential"} {
		if !strings.Contains(s1, needle) {
			t.Errorf("pass 1 missing %q", needle)
		}
		if !strings.Contains(s2, needle) {
			t.Errorf("pass 2 missing %q", needle)
		}
	}
}

// ---------------------------------------------------------------------------
// Edge: extension elements survive round-trip
// ---------------------------------------------------------------------------

func TestExtensionElementRoundTrip(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
		` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:p><w:r><w:t>Normal</w:t></w:r></w:p>` +
		`<w14:shadow w14:val="abc"><w14:nested/></w14:shadow>` +
		`</w:hdr>`

	hf, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if len(hf.Content) != 2 {
		t.Fatalf("expected 2 elements, got %d", len(hf.Content))
	}

	data, err := Serialize(hf)
	if err != nil {
		t.Fatalf("Serialize: %v", err)
	}

	output := string(data)
	if !strings.Contains(output, "Normal") {
		t.Error("lost paragraph text")
	}
	if !strings.Contains(output, "shadow") {
		t.Error("lost extension element 'shadow'")
	}
	if !strings.Contains(output, "nested") {
		t.Error("lost nested extension content")
	}
}
