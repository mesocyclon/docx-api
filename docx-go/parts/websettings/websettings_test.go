package websettings

import (
	"bytes"
	"testing"
)

// Reference XML from reference-appendix.md §2.8 — minimal webSettings.xml
var referenceXML = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
               xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
               mc:Ignorable="w14 w15"/>`)

func TestRoundTrip(t *testing.T) {
	// Parse (unmarshal)
	parsed, err := Parse(referenceXML)
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// Serialize (marshal)
	serialized, err := Serialize(parsed)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// Compare: output must equal input byte-for-byte
	if !bytes.Equal(referenceXML, serialized) {
		t.Errorf("round-trip mismatch:\n  want: %s\n  got:  %s", referenceXML, serialized)
	}
}

func TestParseNil(t *testing.T) {
	out, err := Parse(nil)
	if err != nil {
		t.Fatalf("Parse(nil) returned error: %v", err)
	}
	if out != nil {
		t.Errorf("Parse(nil) should return nil, got %v", out)
	}
}

func TestSerializeNil(t *testing.T) {
	out, err := Serialize(nil)
	if err != nil {
		t.Fatalf("Serialize(nil) returned error: %v", err)
	}
	if out != nil {
		t.Errorf("Serialize(nil) should return nil, got %v", out)
	}
}

func TestParseEmpty(t *testing.T) {
	empty := []byte{}
	out, err := Parse(empty)
	if err != nil {
		t.Fatalf("Parse(empty) returned error: %v", err)
	}
	if !bytes.Equal(empty, out) {
		t.Errorf("Parse(empty) mismatch")
	}
}

func TestRoundTripNonTrivial(t *testing.T) {
	// A more complex webSettings with child elements (e.g., from a real Word doc)
	input := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
               xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
               mc:Ignorable="w14 w15">
  <w:divs>
    <w:div w:id="123456789">
      <w:bodyDiv w:val="1"/>
      <w:marLeft w:val="0"/>
      <w:marRight w:val="0"/>
      <w:marTop w:val="0"/>
      <w:marBottom w:val="0"/>
    </w:div>
  </w:divs>
  <w:optimizeForBrowser/>
  <w:allowPNG/>
</w:webSettings>`)

	parsed, err := Parse(input)
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	serialized, err := Serialize(parsed)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	if !bytes.Equal(input, serialized) {
		t.Errorf("round-trip mismatch for non-trivial input")
	}
}
