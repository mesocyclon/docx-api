package opc

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

// TestNewPackage verifies that a new empty package can be created.
func TestNewPackage(t *testing.T) {
	pkg := New()
	if pkg == nil {
		t.Fatal("New() returned nil")
	}
	if len(pkg.Parts()) != 0 {
		t.Errorf("new package should have 0 parts, got %d", len(pkg.Parts()))
	}
	if len(pkg.PackageRels()) != 0 {
		t.Errorf("new package should have 0 rels, got %d", len(pkg.PackageRels()))
	}
}

// TestAddRemovePart verifies part add/remove/lookup operations.
func TestAddRemovePart(t *testing.T) {
	pkg := New()

	// Add a part.
	pt := pkg.AddPart("/word/document.xml", "application/xml", []byte("<doc/>"))
	if pt == nil {
		t.Fatal("AddPart returned nil")
	}
	if pt.Name != "/word/document.xml" {
		t.Errorf("expected /word/document.xml, got %s", pt.Name)
	}

	// Lookup.
	found, ok := pkg.Part("/word/document.xml")
	if !ok || found != pt {
		t.Error("Part() didn't find added part")
	}

	// Lookup without leading slash (should still normalize).
	found2, ok2 := pkg.Part("word/document.xml")
	if !ok2 || found2 != pt {
		t.Error("Part() didn't normalize name")
	}

	// Parts list.
	parts := pkg.Parts()
	if len(parts) != 1 {
		t.Errorf("expected 1 part, got %d", len(parts))
	}

	// Remove.
	if !pkg.RemovePart("/word/document.xml") {
		t.Error("RemovePart returned false")
	}
	if _, ok := pkg.Part("/word/document.xml"); ok {
		t.Error("part still exists after removal")
	}

	// Remove nonexistent.
	if pkg.RemovePart("/nonexistent.xml") {
		t.Error("RemovePart returned true for nonexistent")
	}
}

// TestPackageRels verifies package-level relationship operations.
func TestPackageRels(t *testing.T) {
	pkg := New()

	id1 := pkg.AddPackageRel(RelOfficeDocument, "word/document.xml")
	if id1 != "rId1" {
		t.Errorf("expected rId1, got %s", id1)
	}

	id2 := pkg.AddPackageRel(RelCoreProperties, "docProps/core.xml")
	if id2 != "rId2" {
		t.Errorf("expected rId2, got %s", id2)
	}

	rels := pkg.PackageRels()
	if len(rels) != 2 {
		t.Fatalf("expected 2 rels, got %d", len(rels))
	}

	byType := pkg.PackageRelsByType(RelOfficeDocument)
	if len(byType) != 1 {
		t.Errorf("expected 1 rel of type OfficeDocument, got %d", len(byType))
	}
	if byType[0].Target != "word/document.xml" {
		t.Errorf("wrong target: %s", byType[0].Target)
	}
}

// TestPartRels verifies part-level relationship operations.
func TestPartRels(t *testing.T) {
	pkg := New()
	pt := pkg.AddPart("/word/document.xml", "application/xml", []byte("<doc/>"))

	id1 := pt.AddRel(RelStyles, "styles.xml")
	if id1 != "rId1" {
		t.Errorf("expected rId1, got %s", id1)
	}

	id2 := pt.AddExternalRel(RelHyperlink, "https://example.com")
	if id2 != "rId2" {
		t.Errorf("expected rId2, got %s", id2)
	}

	// RelsByType.
	stylesRels := pt.RelsByType(RelStyles)
	if len(stylesRels) != 1 {
		t.Fatalf("expected 1 styles rel, got %d", len(stylesRels))
	}

	// RelByID.
	rel, ok := pt.RelByID("rId2")
	if !ok {
		t.Fatal("RelByID didn't find rId2")
	}
	if rel.TargetMode != "External" {
		t.Error("expected External target mode")
	}
	if rel.Target != "https://example.com" {
		t.Errorf("wrong target: %s", rel.Target)
	}

	// Not found.
	_, ok = pt.RelByID("rId999")
	if ok {
		t.Error("RelByID should return false for nonexistent")
	}
}

// TestSaveAndOpenRoundTrip writes a package then reads it back.
func TestSaveAndOpenRoundTrip(t *testing.T) {
	// Build a package.
	pkg := New()
	pkg.AddPackageRel(RelOfficeDocument, "word/document.xml")
	pkg.AddPackageRel(RelCoreProperties, "docProps/core.xml")

	docContent := []byte(`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`)
	docPart := pkg.AddPart("/word/document.xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
		docContent)
	docPart.AddRel(RelStyles, "styles.xml")
	docPart.AddRel(RelSettings, "settings.xml")
	docPart.AddExternalRel(RelHyperlink, "https://example.com")

	stylesContent := []byte(`<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg.AddPart("/word/styles.xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
		stylesContent)

	settingsContent := []byte(`<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
	pkg.AddPart("/word/settings.xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
		settingsContent)

	coreContent := []byte(`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`)
	pkg.AddPart("/docProps/core.xml",
		"application/vnd.openxmlformats-package.core-properties+xml",
		coreContent)

	// Save to buffer.
	var buf bytes.Buffer
	if err := pkg.SaveWriter(&buf); err != nil {
		t.Fatalf("SaveWriter failed: %v", err)
	}

	// Re-open from buffer.
	reader := bytes.NewReader(buf.Bytes())
	pkg2, err := OpenReader(reader, int64(buf.Len()))
	if err != nil {
		t.Fatalf("OpenReader failed: %v", err)
	}

	// Verify package-level rels.
	pkgRels := pkg2.PackageRels()
	if len(pkgRels) != 2 {
		t.Errorf("expected 2 package rels, got %d", len(pkgRels))
	}

	// Verify parts exist.
	parts := pkg2.Parts()
	if len(parts) != 4 {
		t.Errorf("expected 4 parts, got %d", len(parts))
	}

	// Verify document part data.
	docPart2, ok := pkg2.Part("/word/document.xml")
	if !ok {
		t.Fatal("document.xml not found")
	}
	if !bytes.Equal(docPart2.Data, docContent) {
		t.Error("document.xml content mismatch")
	}

	// Verify document part content type.
	if docPart2.ContentType != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" {
		t.Errorf("wrong content type: %s", docPart2.ContentType)
	}

	// Verify document part rels.
	if len(docPart2.Rels) != 3 {
		t.Errorf("expected 3 doc rels, got %d", len(docPart2.Rels))
	}

	// Verify external rel survived round-trip.
	hypRels := docPart2.RelsByType(RelHyperlink)
	if len(hypRels) != 1 {
		t.Fatalf("expected 1 hyperlink rel, got %d", len(hypRels))
	}
	if hypRels[0].TargetMode != "External" {
		t.Error("hyperlink TargetMode should be External")
	}
	if hypRels[0].Target != "https://example.com" {
		t.Errorf("hyperlink target mismatch: %s", hypRels[0].Target)
	}
}

// TestContentTypesParsing verifies parsing of [Content_Types].xml.
func TestContentTypesParsing(t *testing.T) {
	input := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`

	defaults, overrides, err := parseContentTypes([]byte(input))
	if err != nil {
		t.Fatalf("parseContentTypes failed: %v", err)
	}

	if len(defaults) != 3 {
		t.Errorf("expected 3 defaults, got %d", len(defaults))
	}
	if defaults["png"] != "image/png" {
		t.Errorf("png default mismatch: %s", defaults["png"])
	}

	if len(overrides) != 2 {
		t.Errorf("expected 2 overrides, got %d", len(overrides))
	}
	if overrides["/word/document.xml"] != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" {
		t.Error("document.xml override mismatch")
	}
}

// TestRelsParsing verifies parsing of .rels XML.
func TestRelsParsing(t *testing.T) {
	input := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`

	rels, err := parseRels([]byte(input))
	if err != nil {
		t.Fatalf("parseRels failed: %v", err)
	}
	if len(rels) != 3 {
		t.Fatalf("expected 3 rels, got %d", len(rels))
	}

	if rels[0].ID != "rId1" {
		t.Errorf("first rel ID: %s", rels[0].ID)
	}
	if rels[0].Type != RelOfficeDocument {
		t.Errorf("first rel type: %s", rels[0].Type)
	}
	if rels[0].Target != "word/document.xml" {
		t.Errorf("first rel target: %s", rels[0].Target)
	}
}

// TestRelsRoundTrip verifies that rels survive marshal → unmarshal.
func TestRelsRoundTrip(t *testing.T) {
	original := []Relationship{
		{ID: "rId1", Type: RelStyles, Target: "styles.xml"},
		{ID: "rId2", Type: RelHyperlink, Target: "https://example.com", TargetMode: "External"},
		{ID: "rId3", Type: RelImage, Target: "media/image1.png"},
	}

	data, err := buildRels(original)
	if err != nil {
		t.Fatalf("buildRels failed: %v", err)
	}

	parsed, err := parseRels(data)
	if err != nil {
		t.Fatalf("parseRels failed: %v", err)
	}

	if len(parsed) != len(original) {
		t.Fatalf("expected %d rels, got %d", len(original), len(parsed))
	}

	for i := range original {
		if parsed[i].ID != original[i].ID {
			t.Errorf("rel[%d] ID: got %s, want %s", i, parsed[i].ID, original[i].ID)
		}
		if parsed[i].Type != original[i].Type {
			t.Errorf("rel[%d] Type: got %s, want %s", i, parsed[i].Type, original[i].Type)
		}
		if parsed[i].Target != original[i].Target {
			t.Errorf("rel[%d] Target: got %s, want %s", i, parsed[i].Target, original[i].Target)
		}
		if parsed[i].TargetMode != original[i].TargetMode {
			t.Errorf("rel[%d] TargetMode: got %s, want %s", i, parsed[i].TargetMode, original[i].TargetMode)
		}
	}
}

// TestContentTypesRoundTrip verifies [Content_Types].xml round-trip.
func TestContentTypesRoundTrip(t *testing.T) {
	pkg := New()
	pkg.defaults["rels"] = "application/vnd.openxmlformats-package.relationships+xml"
	pkg.defaults["xml"] = "application/xml"

	pkg.AddPart("/word/document.xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
		nil)
	pkg.AddPart("/word/styles.xml",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
		nil)

	data, err := buildContentTypes(pkg)
	if err != nil {
		t.Fatalf("buildContentTypes failed: %v", err)
	}

	// Verify it's valid XML.
	var ct xmlTypes
	if err := xml.Unmarshal(data, &ct); err != nil {
		t.Fatalf("generated content types XML invalid: %v", err)
	}

	if len(ct.Overrides) != 2 {
		t.Errorf("expected 2 overrides, got %d", len(ct.Overrides))
	}

	// Check it starts with XML declaration.
	if !strings.HasPrefix(string(data), "<?xml") {
		t.Error("missing XML declaration")
	}
}

// TestNormalizeName verifies path normalization.
func TestNormalizeName(t *testing.T) {
	tests := []struct {
		input, expected string
	}{
		{"word/document.xml", "/word/document.xml"},
		{"/word/document.xml", "/word/document.xml"},
		{"word\\document.xml", "/word/document.xml"},
	}
	for _, tt := range tests {
		got := normalizeName(tt.input)
		if got != tt.expected {
			t.Errorf("normalizeName(%q) = %q, want %q", tt.input, got, tt.expected)
		}
	}
}

// TestRelsZipPath verifies .rels path generation.
func TestRelsZipPath(t *testing.T) {
	tests := []struct {
		input, expected string
	}{
		{"/word/document.xml", "word/_rels/document.xml.rels"},
		{"", "_rels/.rels"},
		{"/", "_rels/.rels"},
	}
	for _, tt := range tests {
		got := relsZipPath(tt.input)
		if got != tt.expected {
			t.Errorf("relsZipPath(%q) = %q, want %q", tt.input, got, tt.expected)
		}
	}
}

// TestDoubleClose verifies that SaveWriter doesn't fail with nil returns.
func TestEmptyPackageSave(t *testing.T) {
	pkg := New()
	var buf bytes.Buffer
	if err := pkg.SaveWriter(&buf); err != nil {
		t.Fatalf("SaveWriter for empty package failed: %v", err)
	}
	if buf.Len() == 0 {
		t.Error("expected non-empty ZIP output")
	}
}
