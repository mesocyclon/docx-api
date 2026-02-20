package theme

import (
	"bytes"
	"testing"
)

// Reference XML from reference-appendix.md §2.9 — minimal valid theme.
var referenceThemeXML = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" panose="020F0502020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`)

// TestParseRoundTrip verifies that Parse → Serialize returns identical bytes.
func TestParseRoundTrip(t *testing.T) {
	parsed, err := Parse(referenceThemeXML)
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	serialized, err := Serialize(parsed)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	if !bytes.Equal(referenceThemeXML, serialized) {
		t.Error("round-trip: output differs from input")
	}
}

// TestParseNil checks that Parse rejects nil input.
func TestParseNil(t *testing.T) {
	_, err := Parse(nil)
	if err == nil {
		t.Error("Parse(nil) should return an error")
	}
}

// TestSerializeNil checks that Serialize rejects nil input.
func TestSerializeNil(t *testing.T) {
	_, err := Serialize(nil)
	if err == nil {
		t.Error("Serialize(nil) should return an error")
	}
}

// TestParseEmpty verifies that an empty (but non-nil) slice round-trips.
func TestParseEmpty(t *testing.T) {
	data := []byte{}
	parsed, err := Parse(data)
	if err != nil {
		t.Fatalf("Parse(empty) failed: %v", err)
	}
	serialized, err := Serialize(parsed)
	if err != nil {
		t.Fatalf("Serialize(empty) failed: %v", err)
	}
	if !bytes.Equal(data, serialized) {
		t.Error("empty data round-trip failed")
	}
}

// TestDefaultXML verifies that the DefaultXML constant round-trips through
// Parse and Serialize without modification.
func TestDefaultXML(t *testing.T) {
	data := []byte(DefaultXML)

	parsed, err := Parse(data)
	if err != nil {
		t.Fatalf("Parse(DefaultXML) failed: %v", err)
	}

	serialized, err := Serialize(parsed)
	if err != nil {
		t.Fatalf("Serialize(DefaultXML) failed: %v", err)
	}

	if !bytes.Equal(data, serialized) {
		t.Error("DefaultXML round-trip: output differs from input")
	}
}

// TestDefaultXMLContainsRequiredElements does a basic sanity check that
// the DefaultXML constant contains the elements Word requires.
func TestDefaultXMLContainsRequiredElements(t *testing.T) {
	required := []string{
		"<a:theme",
		"<a:clrScheme",
		"<a:fontScheme",
		"<a:fmtScheme",
		"<a:dk1>", "<a:lt1>", "<a:dk2>", "<a:lt2>",
		"<a:accent1>", "<a:accent2>", "<a:accent3>",
		"<a:accent4>", "<a:accent5>", "<a:accent6>",
		"<a:hlink>", "<a:folHlink>",
		"<a:majorFont>", "<a:minorFont>",
		"<a:fillStyleLst>", "<a:lnStyleLst>",
		"<a:effectStyleLst>", "<a:bgFillStyleLst>",
	}
	for _, s := range required {
		if !bytes.Contains([]byte(DefaultXML), []byte(s)) {
			t.Errorf("DefaultXML missing required element: %s", s)
		}
	}
}
