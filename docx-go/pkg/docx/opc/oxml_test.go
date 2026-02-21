package opc

import (
	"testing"
)

func TestParseContentTypes(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`

	ct, err := ParseContentTypes([]byte(xml))
	if err != nil {
		t.Fatalf("ParseContentTypes: %v", err)
	}

	// Test override
	got, err := ct.ContentType("/word/document.xml")
	if err != nil {
		t.Fatalf("ContentType: %v", err)
	}
	if got != CTWmlDocumentMain {
		t.Errorf("got %q, want %q", got, CTWmlDocumentMain)
	}

	// Test default
	got, err = ct.ContentType("/word/someother.xml")
	if err != nil {
		t.Fatalf("ContentType for xml: %v", err)
	}
	if got != CTXml {
		t.Errorf("got %q, want %q", got, CTXml)
	}

	// Test missing
	_, err = ct.ContentType("/nonexistent.xyz")
	if err == nil {
		t.Error("expected error for unknown partname/extension")
	}
}

func TestParseRelationships(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com" TargetMode="External"/>
</Relationships>`

	srels, err := ParseRelationships([]byte(xml), "/")
	if err != nil {
		t.Fatalf("ParseRelationships: %v", err)
	}
	if len(srels) != 2 {
		t.Fatalf("expected 2 rels, got %d", len(srels))
	}

	// Internal rel
	if srels[0].RID != "rId1" {
		t.Errorf("expected rId1, got %q", srels[0].RID)
	}
	if srels[0].RelType != RTOfficeDocument {
		t.Errorf("wrong reltype: %q", srels[0].RelType)
	}
	if srels[0].TargetMode != TargetModeInternal {
		t.Errorf("expected Internal, got %q", srels[0].TargetMode)
	}
	if srels[0].IsExternal() {
		t.Error("expected internal relationship")
	}
	pn := srels[0].TargetPartname()
	if pn != "/word/document.xml" {
		t.Errorf("TargetPartname = %q, want /word/document.xml", pn)
	}

	// External rel
	if srels[1].RID != "rId2" {
		t.Errorf("expected rId2, got %q", srels[1].RID)
	}
	if !srels[1].IsExternal() {
		t.Error("expected external relationship")
	}
	if srels[1].TargetRef != "http://example.com" {
		t.Errorf("expected http://example.com, got %q", srels[1].TargetRef)
	}
}

func TestParseRelationships_Nil(t *testing.T) {
	srels, err := ParseRelationships(nil, "/")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(srels) != 0 {
		t.Errorf("expected empty, got %d", len(srels))
	}
}

func TestSerializeContentTypes_RoundTrip(t *testing.T) {
	parts := []PartInfo{
		{PartName: "/word/document.xml", ContentType: CTWmlDocumentMain},
		{PartName: "/word/styles.xml", ContentType: CTWmlStyles},
		{PartName: "/word/media/image1.png", ContentType: CTPng},
	}

	blob, err := SerializeContentTypes(parts)
	if err != nil {
		t.Fatalf("SerializeContentTypes: %v", err)
	}

	ct, err := ParseContentTypes(blob)
	if err != nil {
		t.Fatalf("ParseContentTypes: %v", err)
	}

	// image1.png should be resolved via default extension
	got, err := ct.ContentType("/word/media/image1.png")
	if err != nil {
		t.Fatalf("ContentType png: %v", err)
	}
	if got != CTPng {
		t.Errorf("got %q, want %q", got, CTPng)
	}

	// document.xml should be override
	got, err = ct.ContentType("/word/document.xml")
	if err != nil {
		t.Fatalf("ContentType docxml: %v", err)
	}
	if got != CTWmlDocumentMain {
		t.Errorf("got %q, want %q", got, CTWmlDocumentMain)
	}
}

func TestSerializeRelationships_RoundTrip(t *testing.T) {
	rels := NewRelationships("/word")

	// Use Load to add with known rId
	part := NewBasePart("/word/styles.xml", CTWmlStyles, nil, nil)
	rels.Load("rId1", RTStyles, "styles.xml", part, false)
	rels.Load("rId2", RTHyperlink, "http://example.com", nil, true)

	blob, err := SerializeRelationships(rels)
	if err != nil {
		t.Fatalf("SerializeRelationships: %v", err)
	}

	srels, err := ParseRelationships(blob, "/word")
	if err != nil {
		t.Fatalf("ParseRelationships: %v", err)
	}

	if len(srels) != 2 {
		t.Fatalf("expected 2 rels, got %d", len(srels))
	}

	// Check first rel
	if srels[0].RID != "rId1" {
		t.Errorf("expected rId1, got %q", srels[0].RID)
	}
	if srels[0].RelType != RTStyles {
		t.Errorf("wrong reltype: %q", srels[0].RelType)
	}

	// Check second rel (external)
	if srels[1].RID != "rId2" {
		t.Errorf("expected rId2, got %q", srels[1].RID)
	}
	if !srels[1].IsExternal() {
		t.Error("expected external")
	}
}

func TestContentTypeMap_CaseInsensitive(t *testing.T) {
	ct := NewContentTypeMap()
	ct.AddDefault("XML", CTXml)

	got, err := ct.ContentType("/test.xml")
	if err != nil {
		t.Fatalf("ContentType: %v", err)
	}
	if got != CTXml {
		t.Errorf("got %q, want %q", got, CTXml)
	}
}
