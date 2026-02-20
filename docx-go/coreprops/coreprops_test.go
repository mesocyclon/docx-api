package coreprops

import (
	"encoding/xml"
	"strings"
	"testing"
	"time"
)

// Reference XML from reference-appendix.md section 2.11
const coreXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title/>
  <dc:subject/>
  <dc:creator>Author</dc:creator>
  <cp:keywords/>
  <dc:description/>
  <cp:lastModifiedBy>Author</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`

// Reference XML from reference-appendix.md section 2.12
const appXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal</Template>
  <TotalTime>0</TotalTime>
  <Pages>1</Pages>
  <Words>0</Words>
  <Characters>0</Characters>
  <Application>docx-go</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>1</Lines>
  <Paragraphs>1</Paragraphs>
  <ScaleCrop>false</ScaleCrop>
  <Company/>
  <LinksUpToDate>false</LinksUpToDate>
  <CharactersWithSpaces>0</CharactersWithSpaces>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`

func TestParseCoreProperties(t *testing.T) {
	cp, err := ParseCore([]byte(coreXML))
	if err != nil {
		t.Fatalf("ParseCore failed: %v", err)
	}

	if cp.Creator != "Author" {
		t.Errorf("Creator = %q, want %q", cp.Creator, "Author")
	}
	if cp.LastModifiedBy != "Author" {
		t.Errorf("LastModifiedBy = %q, want %q", cp.LastModifiedBy, "Author")
	}
	if cp.Revision != "1" {
		t.Errorf("Revision = %q, want %q", cp.Revision, "1")
	}
	expected := time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC)
	if !cp.Created.Equal(expected) {
		t.Errorf("Created = %v, want %v", cp.Created, expected)
	}
	if !cp.Modified.Equal(expected) {
		t.Errorf("Modified = %v, want %v", cp.Modified, expected)
	}
	if cp.Title != "" {
		t.Errorf("Title = %q, want empty", cp.Title)
	}
}

func TestParseAppProperties(t *testing.T) {
	ap, err := ParseApp([]byte(appXML))
	if err != nil {
		t.Fatalf("ParseApp failed: %v", err)
	}

	if ap.Template != "Normal" {
		t.Errorf("Template = %q, want %q", ap.Template, "Normal")
	}
	if ap.Pages != 1 {
		t.Errorf("Pages = %d, want 1", ap.Pages)
	}
	if ap.Application != "docx-go" {
		t.Errorf("Application = %q, want %q", ap.Application, "docx-go")
	}
	if ap.AppVersion != "16.0000" {
		t.Errorf("AppVersion = %q, want %q", ap.AppVersion, "16.0000")
	}
	if ap.Lines != 1 {
		t.Errorf("Lines = %d, want 1", ap.Lines)
	}
	if ap.Paragraphs != 1 {
		t.Errorf("Paragraphs = %d, want 1", ap.Paragraphs)
	}
	if ap.Company != "" {
		t.Errorf("Company = %q, want empty", ap.Company)
	}
}

func TestCoreRoundTrip(t *testing.T) {
	// Parse reference XML
	cp1, err := ParseCore([]byte(coreXML))
	if err != nil {
		t.Fatalf("ParseCore failed: %v", err)
	}

	// Serialize
	data, err := SerializeCore(cp1)
	if err != nil {
		t.Fatalf("SerializeCore failed: %v", err)
	}

	// Verify output is valid XML
	if !strings.Contains(string(data), "<?xml") {
		t.Error("output missing XML declaration")
	}

	// Re-parse
	cp2, err := ParseCore(data)
	if err != nil {
		t.Fatalf("re-ParseCore failed: %v", err)
	}

	// Compare fields
	if cp2.Creator != cp1.Creator {
		t.Errorf("round-trip Creator: got %q, want %q", cp2.Creator, cp1.Creator)
	}
	if cp2.LastModifiedBy != cp1.LastModifiedBy {
		t.Errorf("round-trip LastModifiedBy: got %q, want %q", cp2.LastModifiedBy, cp1.LastModifiedBy)
	}
	if cp2.Revision != cp1.Revision {
		t.Errorf("round-trip Revision: got %q, want %q", cp2.Revision, cp1.Revision)
	}
	if !cp2.Created.Equal(cp1.Created) {
		t.Errorf("round-trip Created: got %v, want %v", cp2.Created, cp1.Created)
	}
	if !cp2.Modified.Equal(cp1.Modified) {
		t.Errorf("round-trip Modified: got %v, want %v", cp2.Modified, cp1.Modified)
	}
	if cp2.Title != cp1.Title {
		t.Errorf("round-trip Title: got %q, want %q", cp2.Title, cp1.Title)
	}
	if cp2.Subject != cp1.Subject {
		t.Errorf("round-trip Subject: got %q, want %q", cp2.Subject, cp1.Subject)
	}
	if cp2.Keywords != cp1.Keywords {
		t.Errorf("round-trip Keywords: got %q, want %q", cp2.Keywords, cp1.Keywords)
	}
	if cp2.Description != cp1.Description {
		t.Errorf("round-trip Description: got %q, want %q", cp2.Description, cp1.Description)
	}
}

func TestAppRoundTrip(t *testing.T) {
	// Parse reference XML
	ap1, err := ParseApp([]byte(appXML))
	if err != nil {
		t.Fatalf("ParseApp failed: %v", err)
	}

	// Serialize
	data, err := SerializeApp(ap1)
	if err != nil {
		t.Fatalf("SerializeApp failed: %v", err)
	}

	// Verify output is valid XML
	if !strings.Contains(string(data), "<?xml") {
		t.Error("output missing XML declaration")
	}

	// Re-parse
	ap2, err := ParseApp(data)
	if err != nil {
		t.Fatalf("re-ParseApp failed: %v", err)
	}

	// Compare fields
	if ap2.Template != ap1.Template {
		t.Errorf("round-trip Template: got %q, want %q", ap2.Template, ap1.Template)
	}
	if ap2.Pages != ap1.Pages {
		t.Errorf("round-trip Pages: got %d, want %d", ap2.Pages, ap1.Pages)
	}
	if ap2.Words != ap1.Words {
		t.Errorf("round-trip Words: got %d, want %d", ap2.Words, ap1.Words)
	}
	if ap2.Characters != ap1.Characters {
		t.Errorf("round-trip Characters: got %d, want %d", ap2.Characters, ap1.Characters)
	}
	if ap2.Application != ap1.Application {
		t.Errorf("round-trip Application: got %q, want %q", ap2.Application, ap1.Application)
	}
	if ap2.Lines != ap1.Lines {
		t.Errorf("round-trip Lines: got %d, want %d", ap2.Lines, ap1.Lines)
	}
	if ap2.Paragraphs != ap1.Paragraphs {
		t.Errorf("round-trip Paragraphs: got %d, want %d", ap2.Paragraphs, ap1.Paragraphs)
	}
	if ap2.Company != ap1.Company {
		t.Errorf("round-trip Company: got %q, want %q", ap2.Company, ap1.Company)
	}
	if ap2.AppVersion != ap1.AppVersion {
		t.Errorf("round-trip AppVersion: got %q, want %q", ap2.AppVersion, ap1.AppVersion)
	}
	if ap2.CharactersWithSpaces != ap1.CharactersWithSpaces {
		t.Errorf("round-trip CharactersWithSpaces: got %d, want %d", ap2.CharactersWithSpaces, ap1.CharactersWithSpaces)
	}
}

func TestDefaultCore(t *testing.T) {
	cp := DefaultCore("Test User")
	if cp.Creator != "Test User" {
		t.Errorf("Creator = %q, want %q", cp.Creator, "Test User")
	}
	if cp.LastModifiedBy != "Test User" {
		t.Errorf("LastModifiedBy = %q, want %q", cp.LastModifiedBy, "Test User")
	}
	if cp.Revision != "1" {
		t.Errorf("Revision = %q, want %q", cp.Revision, "1")
	}
	if cp.Created.IsZero() {
		t.Error("Created should not be zero")
	}
	if cp.Modified.IsZero() {
		t.Error("Modified should not be zero")
	}
}

func TestDefaultApp(t *testing.T) {
	ap := DefaultApp()
	if ap.Template != "Normal" {
		t.Errorf("Template = %q, want %q", ap.Template, "Normal")
	}
	if ap.Application != "docx-go" {
		t.Errorf("Application = %q, want %q", ap.Application, "docx-go")
	}
	if ap.Pages != 1 {
		t.Errorf("Pages = %d, want 1", ap.Pages)
	}
}

func TestSerializeCoreXMLStructure(t *testing.T) {
	cp := DefaultCore("Author")
	data, err := SerializeCore(cp)
	if err != nil {
		t.Fatalf("SerializeCore failed: %v", err)
	}
	s := string(data)

	// Verify namespace declarations
	if !strings.Contains(s, `xmlns:cp=`) {
		t.Error("missing cp namespace declaration")
	}
	if !strings.Contains(s, `xmlns:dc=`) {
		t.Error("missing dc namespace declaration")
	}
	if !strings.Contains(s, `xmlns:dcterms=`) {
		t.Error("missing dcterms namespace declaration")
	}
	if !strings.Contains(s, `xmlns:xsi=`) {
		t.Error("missing xsi namespace declaration")
	}

	// Verify element prefixes
	if !strings.Contains(s, "<dc:creator>") {
		t.Error("missing dc:creator element")
	}
	if !strings.Contains(s, "<cp:revision>") {
		t.Error("missing cp:revision element")
	}
	if !strings.Contains(s, "<dcterms:created") {
		t.Error("missing dcterms:created element")
	}
	if !strings.Contains(s, `xsi:type="dcterms:W3CDTF"`) {
		t.Error("missing xsi:type attribute on dcterms:created")
	}

	// Verify the output is valid XML by attempting to parse
	var probe struct {
		XMLName xml.Name
	}
	if err := xml.Unmarshal(data, &probe); err != nil {
		t.Errorf("serialized XML is not valid: %v", err)
	}
}

func TestSerializeAppXMLStructure(t *testing.T) {
	ap := DefaultApp()
	data, err := SerializeApp(ap)
	if err != nil {
		t.Fatalf("SerializeApp failed: %v", err)
	}
	s := string(data)

	// Verify namespace declarations
	if !strings.Contains(s, `xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"`) {
		t.Error("missing default namespace declaration")
	}
	if !strings.Contains(s, `xmlns:vt=`) {
		t.Error("missing vt namespace declaration")
	}

	// Verify elements
	if !strings.Contains(s, "<Template>Normal</Template>") {
		t.Error("missing Template element")
	}
	if !strings.Contains(s, "<Application>docx-go</Application>") {
		t.Error("missing Application element")
	}
}

func TestCoreWithAllFields(t *testing.T) {
	cp := &CoreProperties{
		Title:          "Test Title",
		Subject:        "Test Subject",
		Creator:        "Creator",
		Keywords:       "go, docx, test",
		Description:    "A test document",
		LastModifiedBy: "Editor",
		Revision:       "3",
		Created:        time.Date(2025, 6, 15, 10, 30, 0, 0, time.UTC),
		Modified:       time.Date(2025, 6, 16, 14, 0, 0, 0, time.UTC),
		Category:       "Testing",
		ContentStatus:  "Draft",
	}

	data, err := SerializeCore(cp)
	if err != nil {
		t.Fatalf("SerializeCore failed: %v", err)
	}

	cp2, err := ParseCore(data)
	if err != nil {
		t.Fatalf("ParseCore failed: %v", err)
	}

	if cp2.Title != cp.Title {
		t.Errorf("Title: got %q, want %q", cp2.Title, cp.Title)
	}
	if cp2.Subject != cp.Subject {
		t.Errorf("Subject: got %q, want %q", cp2.Subject, cp.Subject)
	}
	if cp2.Keywords != cp.Keywords {
		t.Errorf("Keywords: got %q, want %q", cp2.Keywords, cp.Keywords)
	}
	if cp2.Description != cp.Description {
		t.Errorf("Description: got %q, want %q", cp2.Description, cp.Description)
	}
	if cp2.Category != cp.Category {
		t.Errorf("Category: got %q, want %q", cp2.Category, cp.Category)
	}
	if cp2.ContentStatus != cp.ContentStatus {
		t.Errorf("ContentStatus: got %q, want %q", cp2.ContentStatus, cp.ContentStatus)
	}
	if !cp2.Created.Equal(cp.Created) {
		t.Errorf("Created: got %v, want %v", cp2.Created, cp.Created)
	}
	if !cp2.Modified.Equal(cp.Modified) {
		t.Errorf("Modified: got %v, want %v", cp2.Modified, cp.Modified)
	}
}

func TestAppWithNonZeroValues(t *testing.T) {
	ap := &AppProperties{
		Template:             "Custom",
		TotalTime:            120,
		Pages:                5,
		Words:                1500,
		Characters:           8000,
		Application:          "Microsoft Office Word",
		DocSecurity:          0,
		Lines:                100,
		Paragraphs:           25,
		Company:              "Acme Corp",
		AppVersion:           "16.0000",
		CharactersWithSpaces: 9500,
	}

	data, err := SerializeApp(ap)
	if err != nil {
		t.Fatalf("SerializeApp failed: %v", err)
	}

	ap2, err := ParseApp(data)
	if err != nil {
		t.Fatalf("ParseApp failed: %v", err)
	}

	if ap2.TotalTime != ap.TotalTime {
		t.Errorf("TotalTime: got %d, want %d", ap2.TotalTime, ap.TotalTime)
	}
	if ap2.Pages != ap.Pages {
		t.Errorf("Pages: got %d, want %d", ap2.Pages, ap.Pages)
	}
	if ap2.Words != ap.Words {
		t.Errorf("Words: got %d, want %d", ap2.Words, ap.Words)
	}
	if ap2.Company != ap.Company {
		t.Errorf("Company: got %q, want %q", ap2.Company, ap.Company)
	}
	if ap2.CharactersWithSpaces != ap.CharactersWithSpaces {
		t.Errorf("CharactersWithSpaces: got %d, want %d", ap2.CharactersWithSpaces, ap.CharactersWithSpaces)
	}
}
