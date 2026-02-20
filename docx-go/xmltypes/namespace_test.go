package xmltypes

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestNormalizeNamespace(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name  string
		input string
		want  string
	}{
		{
			"strict wordprocessingml to transitional",
			"http://purl.oclc.org/ooxml/wordprocessingml/main",
			NSw,
		},
		{
			"strict relationships to transitional",
			"http://purl.oclc.org/ooxml/officeDocument/relationships",
			NSr,
		},
		{
			"strict drawingml to transitional",
			"http://purl.oclc.org/ooxml/drawingml/main",
			NSa,
		},
		{
			"strict wordprocessingDrawing to transitional",
			"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
			NSwp,
		},
		{
			"strict math to transitional",
			"http://purl.oclc.org/ooxml/officeDocument/math",
			NSm,
		},
		{
			"strict picture to transitional",
			"http://purl.oclc.org/ooxml/drawingml/picture",
			NSpic,
		},
		{
			"transitional passes through unchanged",
			NSw,
			NSw,
		},
		{
			"unknown namespace passes through unchanged",
			"http://example.com/unknown",
			"http://example.com/unknown",
		},
		{
			"empty string passes through",
			"",
			"",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := NormalizeNamespace(tt.input); got != tt.want {
				t.Errorf("NormalizeNamespace(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}

func TestNormalizingDecoder_NormalizesStrictNS(t *testing.T) {
	t.Parallel()

	strictXML := `<document xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main">` +
		`<body></body></document>`

	dec := NewNormalizingDecoder(strings.NewReader(strictXML))

	tok, err := dec.Token()
	if err != nil {
		t.Fatal(err)
	}

	start, ok := tok.(xml.StartElement)
	if !ok {
		t.Fatal("expected StartElement")
	}

	if start.Name.Space != NSw {
		t.Errorf("namespace = %q, want %q (transitional)", start.Name.Space, NSw)
	}
	if start.Name.Local != "document" {
		t.Errorf("local = %q, want %q", start.Name.Local, "document")
	}
}

func TestNormalizingDecoder_PassesThroughTransitional(t *testing.T) {
	t.Parallel()

	transitionalXML := `<body xmlns="` + NSw + `"></body>`
	dec := NewNormalizingDecoder(strings.NewReader(transitionalXML))

	tok, err := dec.Token()
	if err != nil {
		t.Fatal(err)
	}

	start := tok.(xml.StartElement)
	if start.Name.Space != NSw {
		t.Errorf("namespace = %q, want %q", start.Name.Space, NSw)
	}
}

func TestNormalizingDecoder_NormalizesAttributes(t *testing.T) {
	t.Parallel()

	// XML with a Strict relationship namespace on an attribute
	xmlData := `<element xmlns="` + NSw +
		`" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"` +
		` r:id="rId1"></element>`

	dec := NewNormalizingDecoder(strings.NewReader(xmlData))

	tok, err := dec.Token()
	if err != nil {
		t.Fatal(err)
	}

	start := tok.(xml.StartElement)
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" && attr.Name.Space != "" {
			if attr.Name.Space != NSr {
				t.Errorf("attr namespace = %q, want %q", attr.Name.Space, NSr)
			}
		}
	}
}

func TestNormalizingDecoder_NormalizesEndElement(t *testing.T) {
	t.Parallel()

	strictXML := `<body xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"></body>`
	dec := NewNormalizingDecoder(strings.NewReader(strictXML))

	// Read start
	if _, err := dec.Token(); err != nil {
		t.Fatal(err)
	}

	// Read end
	tok, err := dec.Token()
	if err != nil {
		t.Fatal(err)
	}

	end, ok := tok.(xml.EndElement)
	if !ok {
		t.Fatal("expected EndElement")
	}
	if end.Name.Space != NSw {
		t.Errorf("end element namespace = %q, want %q", end.Name.Space, NSw)
	}
}

func TestNamespaceConstants_AreNonEmpty(t *testing.T) {
	t.Parallel()

	// Smoke-test that critical constants are populated
	constants := map[string]string{
		"NSw":   NSw,
		"NSr":   NSr,
		"NSa":   NSa,
		"NSwp":  NSwp,
		"NSpic": NSpic,
		"NSmc":  NSmc,
		"NSm":   NSm,
		"NSw14": NSw14,
		"NSw15": NSw15,
	}

	for name, val := range constants {
		if val == "" {
			t.Errorf("constant %s is empty", name)
		}
		if !strings.HasPrefix(val, "http") && !strings.HasPrefix(val, "urn:") {
			t.Errorf("constant %s = %q does not look like a URI", name, val)
		}
	}
}
