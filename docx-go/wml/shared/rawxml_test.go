package shared

import (
	"encoding/xml"
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// Interface compliance (compile-time checks)
// ---------------------------------------------------------------------------

var (
	_ BlockLevelElement = RawXML{}
	_ ParagraphContent  = RawXML{}
	_ RunContent        = RawXML{}
	_ xml.Marshaler     = RawXML{}
	_ xml.Unmarshaler   = &RawXML{}
)

// ---------------------------------------------------------------------------
// RawXML round-trip: unmarshal → marshal → compare
// ---------------------------------------------------------------------------

func TestRawXMLRoundTripSelfClosing(t *testing.T) {
	t.Parallel()
	input := `<w14:shadow xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:val="test"></w14:shadow>`

	var raw RawXML
	if err := xml.Unmarshal([]byte(input), &raw); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// Verify parsed fields.
	if raw.XMLName.Local != "shadow" {
		t.Errorf("XMLName.Local = %q, want %q", raw.XMLName.Local, "shadow")
	}
	if raw.XMLName.Space != "http://schemas.microsoft.com/office/word/2010/wordml" {
		t.Errorf("XMLName.Space = %q, want w14 namespace", raw.XMLName.Space)
	}
	if len(raw.Attrs) == 0 {
		t.Fatal("expected at least 1 attribute")
	}

	// Marshal back.
	out, err := xml.Marshal(raw)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// Re-unmarshal for semantic comparison.
	var raw2 RawXML
	if err := xml.Unmarshal(out, &raw2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if raw2.XMLName.Local != raw.XMLName.Local {
		t.Errorf("round-trip lost local name: got %q, want %q", raw2.XMLName.Local, raw.XMLName.Local)
	}
	if raw2.XMLName.Space != raw.XMLName.Space {
		t.Errorf("round-trip lost namespace: got %q, want %q", raw2.XMLName.Space, raw.XMLName.Space)
	}
}

func TestRawXMLRoundTripWithChildren(t *testing.T) {
	t.Parallel()
	input := `<w14:ligatures xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:val="standard">` +
		`<w14:child>hello</w14:child>` +
		`</w14:ligatures>`

	var raw RawXML
	if err := xml.Unmarshal([]byte(input), &raw); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if raw.XMLName.Local != "ligatures" {
		t.Errorf("XMLName.Local = %q, want %q", raw.XMLName.Local, "ligatures")
	}
	if len(raw.Inner) == 0 {
		t.Error("expected non-empty Inner for element with children")
	}

	// Marshal back.
	out, err := xml.Marshal(raw)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// The output should contain the child element.
	if !strings.Contains(string(out), "hello") {
		t.Errorf("round-trip lost child text content; output = %s", string(out))
	}

	// Re-unmarshal.
	var raw2 RawXML
	if err := xml.Unmarshal(out, &raw2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if raw2.XMLName.Local != raw.XMLName.Local {
		t.Error("round-trip lost element name")
	}
	if !strings.Contains(string(raw2.Inner), "hello") {
		t.Errorf("round-trip lost inner content after second unmarshal")
	}
}

func TestRawXMLRoundTripMultipleAttrs(t *testing.T) {
	t.Parallel()
	input := `<w15:custom xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w15:foo="bar" w15:baz="qux"></w15:custom>`

	var raw RawXML
	if err := xml.Unmarshal([]byte(input), &raw); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// xmlns:w15 is a namespace declaration, not a content attribute — it is
	// filtered out by UnmarshalXML.  Only foo and baz should remain.
	if len(raw.Attrs) != 2 {
		t.Errorf("expected 2 content attrs (foo, baz), got %d: %v", len(raw.Attrs), raw.Attrs)
	}

	out, err := xml.Marshal(raw)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// Round-trip: re-unmarshal must yield the same number of content attrs.
	var raw2 RawXML
	if err := xml.Unmarshal(out, &raw2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if len(raw2.Attrs) != len(raw.Attrs) {
		t.Errorf("round-trip changed attr count: got %d, want %d", len(raw2.Attrs), len(raw.Attrs))
	}
}

func TestRawXMLEmptyElement(t *testing.T) {
	t.Parallel()
	input := `<w14:empty xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"></w14:empty>`

	var raw RawXML
	if err := xml.Unmarshal([]byte(input), &raw); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if raw.XMLName.Local != "empty" {
		t.Errorf("XMLName.Local = %q, want %q", raw.XMLName.Local, "empty")
	}

	out, err := xml.Marshal(raw)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var raw2 RawXML
	if err := xml.Unmarshal(out, &raw2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if raw2.XMLName.Local != "empty" {
		t.Error("round-trip lost element name for empty element")
	}
}

// TestRawXMLInWrapper simulates how a parent struct would decode an unknown
// child element into RawXML and then re-marshal it.
func TestRawXMLInWrapper(t *testing.T) {
	t.Parallel()

	type wrapper struct {
		XMLName xml.Name `xml:"wrapper"`
		Items   []RawXML `xml:",any"`
	}

	input := `<wrapper>` +
		`<w14:textFill xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w14:solidFill><w14:srgbClr w14:val="FF0000"></w14:srgbClr></w14:solidFill>` +
		`</w14:textFill>` +
		`<w15:otherExt xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w15:x="1"></w15:otherExt>` +
		`</wrapper>`

	var w wrapper
	if err := xml.Unmarshal([]byte(input), &w); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if len(w.Items) != 2 {
		t.Fatalf("expected 2 items, got %d", len(w.Items))
	}
	if w.Items[0].XMLName.Local != "textFill" {
		t.Errorf("first item local = %q, want textFill", w.Items[0].XMLName.Local)
	}
	if w.Items[1].XMLName.Local != "otherExt" {
		t.Errorf("second item local = %q, want otherExt", w.Items[1].XMLName.Local)
	}

	out, err := xml.Marshal(w)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var w2 wrapper
	if err := xml.Unmarshal(out, &w2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if len(w2.Items) != 2 {
		t.Errorf("round-trip lost items: got %d, want 2", len(w2.Items))
	}
	if w2.Items[0].XMLName.Local != "textFill" {
		t.Error("round-trip lost first item name")
	}
}
