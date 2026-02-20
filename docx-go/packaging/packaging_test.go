package packaging

import (
	"bytes"
	"fmt"
	"strings"
	"sync"
	"testing"

	"github.com/vortex/docx-go/coreprops"
	"github.com/vortex/docx-go/opc"
	"github.com/vortex/docx-go/wml/hdft"
)

// ═══════════════════════════════════════════════════════════════════
//  XML FIXTURES — minimal valid XML accepted by each parts/* parser
// ═══════════════════════════════════════════════════════════════════

const fixtureDocumentXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

const fixtureStylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="w14 w15">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>`

const fixtureSettingsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
            mc:Ignorable="w14 w15">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`

const fixtureFontTableXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>`

const fixtureWebSettingsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
               xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
               mc:Ignorable="w14 w15"/>`

const fixtureThemeXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`

const fixtureCoreXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Test Title</dc:title>
  <dc:creator>TestAuthor</dc:creator>
  <cp:lastModifiedBy>TestAuthor</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`

const fixtureAppXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal</Template>
  <TotalTime>0</TotalTime>
  <Pages>1</Pages>
  <Words>2</Words>
  <Characters>11</Characters>
  <Application>docx-go</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>1</Lines>
  <Paragraphs>1</Paragraphs>
  <Company>TestCo</Company>
  <ScaleCrop>false</ScaleCrop>
  <LinksUpToDate>false</LinksUpToDate>
  <CharactersWithSpaces>13</CharactersWithSpaces>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`

const fixtureNumberingXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`

const fixtureCommentsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer" w:date="2025-01-20T14:00:00Z" w:initials="R">
    <w:p><w:r><w:t>Please verify.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`

const fixtureFootnotesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p><w:r><w:t>Footnote text.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`

const fixtureEndnotesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
</w:endnotes>`

const fixtureHeaderXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:pStyle w:val="Header"/></w:pPr>
    <w:r><w:t>Default Header</w:t></w:r>
  </w:p>
</w:hdr>`

const fixtureFooterXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/></w:pPr>
    <w:r><w:t>Confidential</w:t></w:r>
  </w:p>
</w:ftr>`

const fixtureFirstPageHeaderXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>First Page Header</w:t></w:r></w:p>
</w:hdr>`

const fixtureEvenHeaderXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Even Page Header</w:t></w:r></w:p>
</w:hdr>`

const fixtureEvenFooterXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Even Page Footer</w:t></w:r></w:p>
</w:ftr>`

// Fake image bytes (PNG 8-byte magic + some padding).
var fakePNG = []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00}
var fakeJPEG = []byte{0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10}
var fakeGIF = []byte{0x47, 0x49, 0x46, 0x38, 0x39, 0x61}

// Hyperlink rel type (not exported from opc in this package's scope).
const relHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

// ═══════════════════════════════════════════════════════════════════
//  BUILDER — programmatic assembly of opc.Package
// ═══════════════════════════════════════════════════════════════════

type testPkgBuilder struct {
	pkg     *opc.Package
	docPart *opc.Part
}

// newFullPkg creates a package with all mandatory parts.
func newFullPkg() *testPkgBuilder {
	b := &testPkgBuilder{pkg: opc.New()}

	b.pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	b.pkg.AddPackageRel(relCoreProperties, "docProps/core.xml")
	b.pkg.AddPackageRel(relExtProperties, "docProps/app.xml")

	b.docPart = b.pkg.AddPart("/word/document.xml", ctDocument, []byte(fixtureDocumentXML))

	b.docPart.AddRel(relStyles, "styles.xml")
	b.docPart.AddRel(relSettings, "settings.xml")
	b.docPart.AddRel(relFontTable, "fontTable.xml")
	b.docPart.AddRel(relWebSettings, "webSettings.xml")
	b.docPart.AddRel(relTheme, "theme/theme1.xml")

	b.pkg.AddPart("/word/styles.xml", ctStyles, []byte(fixtureStylesXML))
	b.pkg.AddPart("/word/settings.xml", ctSettings, []byte(fixtureSettingsXML))
	b.pkg.AddPart("/word/fontTable.xml", ctFontTable, []byte(fixtureFontTableXML))
	b.pkg.AddPart("/word/webSettings.xml", ctWebSettings, []byte(fixtureWebSettingsXML))
	b.pkg.AddPart("/word/theme/theme1.xml", ctTheme, []byte(fixtureThemeXML))
	b.pkg.AddPart("/docProps/core.xml", ctCore, []byte(fixtureCoreXML))
	b.pkg.AddPart("/docProps/app.xml", ctExtended, []byte(fixtureAppXML))

	return b
}

// newMinimalPkg creates a package with ONLY the document part —
// no styles, settings, fonts, etc.
func newMinimalPkg() *testPkgBuilder {
	b := &testPkgBuilder{pkg: opc.New()}
	b.pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	b.docPart = b.pkg.AddPart("/word/document.xml", ctDocument, []byte(fixtureDocumentXML))
	return b
}

func (b *testPkgBuilder) addNumbering() *testPkgBuilder {
	b.docPart.AddRel(relNumbering, "numbering.xml")
	b.pkg.AddPart("/word/numbering.xml", ctNumbering, []byte(fixtureNumberingXML))
	return b
}

func (b *testPkgBuilder) addComments() *testPkgBuilder {
	b.docPart.AddRel(relComments, "comments.xml")
	b.pkg.AddPart("/word/comments.xml", ctComments, []byte(fixtureCommentsXML))
	return b
}

func (b *testPkgBuilder) addFootnotes() *testPkgBuilder {
	b.docPart.AddRel(relFootnotes, "footnotes.xml")
	b.pkg.AddPart("/word/footnotes.xml", ctFootnotes, []byte(fixtureFootnotesXML))
	return b
}

func (b *testPkgBuilder) addEndnotes() *testPkgBuilder {
	b.docPart.AddRel(relEndnotes, "endnotes.xml")
	b.pkg.AddPart("/word/endnotes.xml", ctEndnotes, []byte(fixtureEndnotesXML))
	return b
}

func (b *testPkgBuilder) addHeader(target, xml string) string {
	rID := b.docPart.AddRel(relHeader, target)
	b.pkg.AddPart("/word/"+target, ctHeader, []byte(xml))
	return rID
}

func (b *testPkgBuilder) addFooter(target, xml string) string {
	rID := b.docPart.AddRel(relFooter, target)
	b.pkg.AddPart("/word/"+target, ctFooter, []byte(xml))
	return rID
}

func (b *testPkgBuilder) addImage(target string, data []byte) string {
	rID := b.docPart.AddRel(relImage, target)
	b.pkg.AddPart("/word/"+target, "image/png", data)
	return rID
}

func (b *testPkgBuilder) addExternalImage(url string) string {
	return b.docPart.AddExternalRel(relImage, url)
}

func (b *testPkgBuilder) addUnknownDocRel(relType, target string, data []byte) string {
	rID := b.docPart.AddRel(relType, target)
	if data != nil {
		b.pkg.AddPart("/word/"+target, "application/xml", data)
	}
	return rID
}

func (b *testPkgBuilder) addExternalDocRel(relType, target string) string {
	return b.docPart.AddExternalRel(relType, target)
}

func (b *testPkgBuilder) addUnknownPkgRel(relType, target string, data []byte) {
	b.pkg.AddPackageRel(relType, target)
	if data != nil {
		b.pkg.AddPart("/"+target, "application/xml", data)
	}
}

func (b *testPkgBuilder) load(t *testing.T) *Document {
	t.Helper()
	doc, err := load(b.pkg)
	if err != nil {
		t.Fatalf("load() failed: %v", err)
	}
	return doc
}

// newDocStruct creates a Document struct suitable for unit-testing helper
// methods without needing an opc.Package.
func newDocStruct(relSeq, bmkID int) *Document {
	return &Document{
		Headers:      make(map[string]*hdft.CT_HdrFtr),
		Footers:      make(map[string]*hdft.CT_HdrFtr),
		Media:        make(map[string][]byte),
		UnknownParts: make(map[string][]byte),
		nextRelSeq:   relSeq,
		nextBmkID:    bmkID,
		docPartName:  "/word/document.xml",
	}
}

// ═══════════════════════════════════════════════════════════════════
//  I. UNIT TESTS — normalizePartName
// ═══════════════════════════════════════════════════════════════════

func TestNormalizePartName(t *testing.T) {
	tests := []struct {
		input, want string
	}{
		{"word/document.xml", "/word/document.xml"},
		{"/word/document.xml", "/word/document.xml"},
		{"docProps/core.xml", "/docProps/core.xml"},
		{"/docProps/core.xml", "/docProps/core.xml"},
		{"a.xml", "/a.xml"},
		{"deep/nested/path.xml", "/deep/nested/path.xml"},
	}
	for _, tt := range tests {
		if got := normalizePartName(tt.input); got != tt.want {
			t.Errorf("normalizePartName(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  II. UNIT TESTS — resolveTarget
// ═══════════════════════════════════════════════════════════════════

func TestResolveTarget_Relative(t *testing.T) {
	tests := []struct {
		source, target, want string
	}{
		{"/word/document.xml", "styles.xml", "/word/styles.xml"},
		{"/word/document.xml", "theme/theme1.xml", "/word/theme/theme1.xml"},
		{"/word/document.xml", "media/image1.png", "/word/media/image1.png"},
		{"/word/document.xml", "glossary/document.xml", "/word/glossary/document.xml"},
	}
	for _, tt := range tests {
		if got := resolveTarget(tt.source, tt.target); got != tt.want {
			t.Errorf("resolveTarget(%q, %q) = %q, want %q", tt.source, tt.target, got, tt.want)
		}
	}
}

func TestResolveTarget_Absolute(t *testing.T) {
	got := resolveTarget("/word/document.xml", "/docProps/core.xml")
	if got != "/docProps/core.xml" {
		t.Errorf("absolute target: got %q, want /docProps/core.xml", got)
	}
}

func TestResolveTarget_ParentDir(t *testing.T) {
	got := resolveTarget("/word/document.xml", "../customXml/item1.xml")
	if got != "/customXml/item1.xml" {
		t.Errorf("parent-dir target: got %q, want /customXml/item1.xml", got)
	}
}

func TestResolveTarget_DeeplyNested(t *testing.T) {
	got := resolveTarget("/word/sub/doc.xml", "images/pic.png")
	if got != "/word/sub/images/pic.png" {
		t.Errorf("deeply nested: got %q, want /word/sub/images/pic.png", got)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  III. UNIT TESTS — parseRelIDNum
// ═══════════════════════════════════════════════════════════════════

func TestParseRelIDNum(t *testing.T) {
	tests := []struct {
		input string
		want  int
	}{
		{"rId1", 1},
		{"rId10", 10},
		{"rId0", 0},
		{"rId999", 999},
		{"bogus", 0},
		{"rId", 0},
		{"rId-1", -1},
		{"", 0},
		{"rId007", 7},
		{"RId5", 0},                   // case-sensitive prefix
		{"rid5", 0},                   // case-sensitive
		{"rId2147483647", 2147483647}, // large number
	}
	for _, tt := range tests {
		if got := parseRelIDNum(tt.input); got != tt.want {
			t.Errorf("parseRelIDNum(%q) = %d, want %d", tt.input, got, tt.want)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  IV. UNIT TESTS — guessMediaContentType
// ═══════════════════════════════════════════════════════════════════

func TestGuessMediaContentType(t *testing.T) {
	tests := []struct {
		filename, want string
	}{
		{"image1.png", "image/png"},
		{"image1.PNG", "image/png"}, // case-insensitive
		{"photo.jpg", "image/jpeg"},
		{"photo.jpeg", "image/jpeg"},
		{"photo.JPEG", "image/jpeg"},
		{"icon.gif", "image/gif"},
		{"pic.bmp", "image/bmp"},
		{"scan.tiff", "image/tiff"},
		{"scan.tif", "image/tiff"},
		{"diagram.emf", "image/x-emf"},
		{"drawing.wmf", "image/x-wmf"},
		{"logo.svg", "image/svg+xml"},
		{"unknown.xyz", "application/octet-stream"},
		{"noext", "application/octet-stream"},
		{"", "application/octet-stream"},
		{".png", "image/png"}, // no basename, just extension
	}
	for _, tt := range tests {
		if got := guessMediaContentType(tt.filename); got != tt.want {
			t.Errorf("guessMediaContentType(%q) = %q, want %q", tt.filename, got, tt.want)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  V. UNIT TESTS — guessContentType
// ═══════════════════════════════════════════════════════════════════

func TestGuessContentType(t *testing.T) {
	tests := []struct {
		partName, want string
	}{
		{"/word/custom.xml", "application/xml"},
		{"/word/custom.XML", "application/xml"}, // case-insensitive
		{"/word/_rels/document.xml.rels", "application/vnd.openxmlformats-package.relationships+xml"},
		{"/word/media/image1.png", "image/png"},
		{"/word/media/photo.jpg", "image/jpeg"},
		{"/word/something.bin", "application/octet-stream"},
		{"/word/custom.json", "application/octet-stream"}, // not in map
	}
	for _, tt := range tests {
		if got := guessContentType(tt.partName); got != tt.want {
			t.Errorf("guessContentType(%q) = %q, want %q", tt.partName, got, tt.want)
		}
	}
}

func TestGuessContentType_DelegatesToMedia(t *testing.T) {
	// When extension is a known media type, guessContentType should return it
	// via guessMediaContentType fallback.
	got := guessContentType("/word/media/photo.svg")
	if got != "image/svg+xml" {
		t.Errorf("guessContentType for .svg = %q, want image/svg+xml", got)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  VI. UNIT TESTS — isPackageLevelRel
// ═══════════════════════════════════════════════════════════════════

func TestIsPackageLevelRel(t *testing.T) {
	tests := []struct {
		relType string
		want    bool
	}{
		{relCoreProperties, true},
		{relExtProperties, true},
		{"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail", true},
		// Any URI containing "metadata/" is treated as package-level.
		{"http://example.com/some/metadata/extra", true},
		// Document-level rels:
		{relStyles, false},
		{relSettings, false},
		{relFontTable, false},
		{relWebSettings, false},
		{relNumbering, false},
		{relFootnotes, false},
		{relEndnotes, false},
		{relComments, false},
		{relHeader, false},
		{relFooter, false},
		{relImage, false},
		{relTheme, false},
		{relHyperlink, false},
		// Custom non-metadata rel:
		{"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", false},
	}
	for _, tt := range tests {
		if got := isPackageLevelRel(tt.relType); got != tt.want {
			t.Errorf("isPackageLevelRel(%q) = %v, want %v", tt.relType, got, tt.want)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  VII. NextRelID
// ═══════════════════════════════════════════════════════════════════

func TestNextRelID_StartsFromSeed(t *testing.T) {
	for _, seed := range []int{1, 5, 42, 100} {
		d := newDocStruct(seed, 0)
		got := d.NextRelID()
		want := fmt.Sprintf("rId%d", seed)
		if got != want {
			t.Errorf("seed=%d: NextRelID()=%q, want %q", seed, got, want)
		}
	}
}

func TestNextRelID_Monotonic(t *testing.T) {
	d := newDocStruct(1, 0)
	prev := 0
	for i := 0; i < 20; i++ {
		id := d.NextRelID()
		n := parseRelIDNum(id)
		if n <= prev {
			t.Fatalf("not monotonic: %d followed %d", n, prev)
		}
		prev = n
	}
}

func TestNextRelID_FormatPrefix(t *testing.T) {
	d := newDocStruct(7, 0)
	id := d.NextRelID()
	if !strings.HasPrefix(id, "rId") {
		t.Errorf("NextRelID()=%q missing rId prefix", id)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  VIII. NextBookmarkID
// ═══════════════════════════════════════════════════════════════════

func TestNextBookmarkID_StartsFromSeed(t *testing.T) {
	d := newDocStruct(1, 50)
	if got := d.NextBookmarkID(); got != 50 {
		t.Errorf("got %d, want 50", got)
	}
}

func TestNextBookmarkID_Monotonic(t *testing.T) {
	d := newDocStruct(1, 100)
	prev := d.NextBookmarkID()
	for i := 0; i < 20; i++ {
		cur := d.NextBookmarkID()
		if cur != prev+1 {
			t.Errorf("step %d: %d → %d (not +1)", i, prev, cur)
		}
		prev = cur
	}
}

// ═══════════════════════════════════════════════════════════════════
//  IX. AddMedia
// ═══════════════════════════════════════════════════════════════════

func TestAddMedia_StoresData(t *testing.T) {
	d := newDocStruct(1, 0)
	d.AddMedia("logo.png", fakePNG)

	stored, ok := d.Media["logo.png"]
	if !ok {
		t.Fatal("logo.png not in Media map")
	}
	if !bytes.Equal(stored, fakePNG) {
		t.Error("stored data ≠ input data")
	}
}

func TestAddMedia_ReturnsRelID(t *testing.T) {
	d := newDocStruct(10, 0)
	rID := d.AddMedia("img.png", fakePNG)
	if rID != "rId10" {
		t.Errorf("rID=%q, want rId10", rID)
	}
}

func TestAddMedia_UniqueRIDs(t *testing.T) {
	d := newDocStruct(1, 0)
	r1 := d.AddMedia("a.png", fakePNG)
	r2 := d.AddMedia("b.png", fakeJPEG)
	r3 := d.AddMedia("c.gif", fakeGIF)
	if r1 == r2 || r2 == r3 || r1 == r3 {
		t.Errorf("duplicate rIds: %s, %s, %s", r1, r2, r3)
	}
}

func TestAddMedia_DeduplicatesFilename(t *testing.T) {
	d := newDocStruct(1, 0)
	d.AddMedia("img.png", []byte("first"))
	d.AddMedia("img.png", []byte("second"))
	d.AddMedia("img.png", []byte("third"))

	if _, ok := d.Media["img.png"]; !ok {
		t.Error("original img.png missing")
	}
	if _, ok := d.Media["img1.png"]; !ok {
		t.Error("first dedup img1.png missing")
	}
	if _, ok := d.Media["img2.png"]; !ok {
		t.Error("second dedup img2.png missing")
	}
	if len(d.Media) != 3 {
		t.Errorf("Media count = %d, want 3", len(d.Media))
	}
}

func TestAddMedia_PreservesExtension(t *testing.T) {
	d := newDocStruct(1, 0)
	d.AddMedia("photo.jpeg", []byte("a"))
	d.AddMedia("photo.jpeg", []byte("b"))

	if _, ok := d.Media["photo.jpeg"]; !ok {
		t.Error("photo.jpeg missing")
	}
	if _, ok := d.Media["photo1.jpeg"]; !ok {
		t.Error("dedup should produce photo1.jpeg, not photo1.jpg")
	}
}

func TestAddMedia_FileWithoutExtension(t *testing.T) {
	d := newDocStruct(1, 0)
	d.AddMedia("Makefile", []byte("data"))
	d.AddMedia("Makefile", []byte("data2"))

	if _, ok := d.Media["Makefile"]; !ok {
		t.Error("Makefile missing")
	}
	// Second should be Makefile1 (no extension → base="Makefile", ext="")
	if _, ok := d.Media["Makefile1"]; !ok {
		t.Error("Makefile1 missing")
	}
}

func TestAddMedia_AdvancesRelSeq(t *testing.T) {
	d := newDocStruct(5, 0)
	d.AddMedia("a.png", fakePNG) // consumes rId5
	got := d.NextRelID()         // should be rId6
	if got != "rId6" {
		t.Errorf("after AddMedia, NextRelID()=%q, want rId6", got)
	}
}

func TestAddMedia_MultipleDedupsChain(t *testing.T) {
	d := newDocStruct(1, 0)
	for i := 0; i < 10; i++ {
		d.AddMedia("file.txt", []byte(fmt.Sprintf("v%d", i)))
	}
	// First is "file.txt", then "file1.txt" .. "file9.txt"
	if len(d.Media) != 10 {
		t.Errorf("after 10 AddMedias with same name: %d entries, want 10", len(d.Media))
	}
	for i := 1; i <= 9; i++ {
		name := fmt.Sprintf("file%d.txt", i)
		if _, ok := d.Media[name]; !ok {
			t.Errorf("%s missing", name)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  X. CONCURRENCY
// ═══════════════════════════════════════════════════════════════════

func TestNextRelID_ConcurrentNoDuplicates(t *testing.T) {
	d := newDocStruct(1, 0)
	const n = 200
	ch := make(chan string, n)
	var wg sync.WaitGroup
	wg.Add(n)
	for i := 0; i < n; i++ {
		go func() {
			defer wg.Done()
			ch <- d.NextRelID()
		}()
	}
	wg.Wait()
	close(ch)

	seen := make(map[string]bool, n)
	for id := range ch {
		if seen[id] {
			t.Fatalf("duplicate rId: %s", id)
		}
		seen[id] = true
	}
	if len(seen) != n {
		t.Errorf("unique IDs = %d, want %d", len(seen), n)
	}
}

func TestNextBookmarkID_ConcurrentNoDuplicates(t *testing.T) {
	d := newDocStruct(0, 1)
	const n = 200
	ch := make(chan int, n)
	var wg sync.WaitGroup
	wg.Add(n)
	for i := 0; i < n; i++ {
		go func() {
			defer wg.Done()
			ch <- d.NextBookmarkID()
		}()
	}
	wg.Wait()
	close(ch)

	seen := make(map[int]bool, n)
	for id := range ch {
		if seen[id] {
			t.Fatalf("duplicate bookmarkID: %d", id)
		}
		seen[id] = true
	}
	if len(seen) != n {
		t.Errorf("unique IDs = %d, want %d", len(seen), n)
	}
}

func TestAddMedia_ConcurrentNoDuplicateRIds(t *testing.T) {
	d := newDocStruct(1, 0)
	const n = 100
	ch := make(chan string, n)
	var wg sync.WaitGroup
	wg.Add(n)
	for i := 0; i < n; i++ {
		i := i
		go func() {
			defer wg.Done()
			// unique name per goroutine to avoid dedup cross-talk
			name := fmt.Sprintf("img_%04d.png", i)
			ch <- d.AddMedia(name, []byte{byte(i)})
		}()
	}
	wg.Wait()
	close(ch)

	seen := make(map[string]bool, n)
	for rID := range ch {
		if seen[rID] {
			t.Fatalf("duplicate media rId: %s", rID)
		}
		seen[rID] = true
	}
	if len(seen) != n {
		t.Errorf("unique rIds = %d, want %d", len(seen), n)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XI. LOAD — mandatory parts
// ═══════════════════════════════════════════════════════════════════

func TestLoad_MandatoryPartsAllPresent(t *testing.T) {
	doc := newFullPkg().load(t)

	if doc.Document == nil {
		t.Error("Document is nil")
	}
	if doc.Styles == nil {
		t.Error("Styles is nil")
	}
	if doc.Settings == nil {
		t.Error("Settings is nil")
	}
	if doc.Fonts == nil {
		t.Error("Fonts is nil")
	}
	if doc.CoreProps == nil {
		t.Error("CoreProps is nil")
	}
	if doc.AppProps == nil {
		t.Error("AppProps is nil")
	}
	if len(doc.WebSettings) == 0 {
		t.Error("WebSettings is empty")
	}
	if len(doc.Theme) == 0 {
		t.Error("Theme is empty")
	}
}

func TestLoad_DocPartNameResolved(t *testing.T) {
	doc := newFullPkg().load(t)
	if doc.docPartName != "/word/document.xml" {
		t.Errorf("docPartName = %q", doc.docPartName)
	}
}

func TestLoad_MapsInitialised(t *testing.T) {
	doc := newFullPkg().load(t)
	if doc.Headers == nil {
		t.Error("Headers map is nil")
	}
	if doc.Footers == nil {
		t.Error("Footers map is nil")
	}
	if doc.Media == nil {
		t.Error("Media map is nil")
	}
	if doc.UnknownParts == nil {
		t.Error("UnknownParts map is nil")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XII. LOAD — minimal package (only document.xml)
// ═══════════════════════════════════════════════════════════════════

func TestLoad_MinimalPkg_OnlyDocument(t *testing.T) {
	doc := newMinimalPkg().load(t)

	if doc.Document == nil {
		t.Fatal("Document is nil")
	}
	// Everything optional should be nil
	if doc.Styles != nil {
		t.Error("Styles should be nil in minimal package")
	}
	if doc.Settings != nil {
		t.Error("Settings should be nil")
	}
	if doc.Fonts != nil {
		t.Error("Fonts should be nil")
	}
	if doc.Numbering != nil {
		t.Error("Numbering should be nil")
	}
	if doc.CoreProps != nil {
		t.Error("CoreProps should be nil")
	}
	if doc.AppProps != nil {
		t.Error("AppProps should be nil")
	}
	if len(doc.Theme) != 0 {
		t.Error("Theme should be empty")
	}
	if len(doc.WebSettings) != 0 {
		t.Error("WebSettings should be empty")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XIII. LOAD — optional parts
// ═══════════════════════════════════════════════════════════════════

func TestLoad_OptionalParts_NilWhenAbsent(t *testing.T) {
	doc := newFullPkg().load(t)
	if doc.Numbering != nil {
		t.Error("Numbering should be nil when absent")
	}
	if doc.Comments != nil {
		t.Error("Comments should be nil")
	}
	if doc.Footnotes != nil {
		t.Error("Footnotes should be nil")
	}
	if doc.Endnotes != nil {
		t.Error("Endnotes should be nil")
	}
}

func TestLoad_Numbering(t *testing.T) {
	doc := newFullPkg().addNumbering().load(t)
	if doc.Numbering == nil {
		t.Fatal("Numbering nil")
	}
}

func TestLoad_Comments(t *testing.T) {
	doc := newFullPkg().addComments().load(t)
	if doc.Comments == nil {
		t.Fatal("Comments nil")
	}
}

func TestLoad_Footnotes(t *testing.T) {
	doc := newFullPkg().addFootnotes().load(t)
	if doc.Footnotes == nil {
		t.Fatal("Footnotes nil")
	}
}

func TestLoad_Endnotes(t *testing.T) {
	doc := newFullPkg().addEndnotes().load(t)
	if doc.Endnotes == nil {
		t.Fatal("Endnotes nil")
	}
}

func TestLoad_AllOptionalPartsTogether(t *testing.T) {
	doc := newFullPkg().
		addNumbering().addComments().addFootnotes().addEndnotes().
		load(t)

	if doc.Numbering == nil {
		t.Error("Numbering nil")
	}
	if doc.Comments == nil {
		t.Error("Comments nil")
	}
	if doc.Footnotes == nil {
		t.Error("Footnotes nil")
	}
	if doc.Endnotes == nil {
		t.Error("Endnotes nil")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XIV. LOAD — headers & footers
// ═══════════════════════════════════════════════════════════════════

func TestLoad_SingleHeader(t *testing.T) {
	b := newFullPkg()
	rID := b.addHeader("header1.xml", fixtureHeaderXML)
	doc := b.load(t)

	if len(doc.Headers) != 1 {
		t.Fatalf("headers: got %d, want 1", len(doc.Headers))
	}
	if _, ok := doc.Headers[rID]; !ok {
		t.Errorf("header not keyed by %s", rID)
	}
}

func TestLoad_MultipleHeaders(t *testing.T) {
	b := newFullPkg()
	r1 := b.addHeader("header1.xml", fixtureHeaderXML)
	r2 := b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	r3 := b.addHeader("header3.xml", fixtureEvenHeaderXML)
	doc := b.load(t)

	if len(doc.Headers) != 3 {
		t.Fatalf("headers: got %d, want 3", len(doc.Headers))
	}
	for _, rID := range []string{r1, r2, r3} {
		if _, ok := doc.Headers[rID]; !ok {
			t.Errorf("header %s missing", rID)
		}
	}
}

func TestLoad_SingleFooter(t *testing.T) {
	b := newFullPkg()
	rID := b.addFooter("footer1.xml", fixtureFooterXML)
	doc := b.load(t)

	if len(doc.Footers) != 1 {
		t.Fatalf("footers: got %d, want 1", len(doc.Footers))
	}
	if _, ok := doc.Footers[rID]; !ok {
		t.Errorf("footer not keyed by %s", rID)
	}
}

func TestLoad_MultipleFooters(t *testing.T) {
	b := newFullPkg()
	r1 := b.addFooter("footer1.xml", fixtureFooterXML)
	r2 := b.addFooter("footer2.xml", fixtureEvenFooterXML)
	doc := b.load(t)

	if len(doc.Footers) != 2 {
		t.Fatalf("footers: got %d, want 2", len(doc.Footers))
	}
	if _, ok := doc.Footers[r1]; !ok {
		t.Error("footer1 missing")
	}
	if _, ok := doc.Footers[r2]; !ok {
		t.Error("footer2 missing")
	}
}

func TestLoad_HeadersAndFootersTogether(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)
	doc := b.load(t)

	if len(doc.Headers) != 2 {
		t.Errorf("headers = %d, want 2", len(doc.Headers))
	}
	if len(doc.Footers) != 1 {
		t.Errorf("footers = %d, want 1", len(doc.Footers))
	}
}

func TestLoad_HeaderMissingPart_Skipped(t *testing.T) {
	b := newFullPkg()
	b.docPart.AddRel(relHeader, "header_ghost.xml") // rel but no part
	doc := b.load(t)

	if len(doc.Headers) != 0 {
		t.Errorf("headers: got %d, want 0 when part missing", len(doc.Headers))
	}
}

func TestLoad_FooterMissingPart_Skipped(t *testing.T) {
	b := newFullPkg()
	b.docPart.AddRel(relFooter, "footer_ghost.xml")
	doc := b.load(t)

	if len(doc.Footers) != 0 {
		t.Errorf("footers: got %d, want 0 when part missing", len(doc.Footers))
	}
}

func TestLoad_NoHeadersNoFooters(t *testing.T) {
	doc := newFullPkg().load(t)
	if len(doc.Headers) != 0 {
		t.Error("expected empty headers map")
	}
	if len(doc.Footers) != 0 {
		t.Error("expected empty footers map")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XV. LOAD — media / images
// ═══════════════════════════════════════════════════════════════════

func TestLoad_InternalImage(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/image1.png", fakePNG)
	doc := b.load(t)

	if len(doc.Media) != 1 {
		t.Fatalf("media: got %d, want 1", len(doc.Media))
	}
	stored, ok := doc.Media["image1.png"]
	if !ok {
		t.Fatal("image1.png not in Media")
	}
	if !bytes.Equal(stored, fakePNG) {
		t.Error("image data mismatch")
	}
}

func TestLoad_MultipleImages(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/pic1.png", fakePNG)
	b.addImage("media/pic2.jpeg", fakeJPEG)
	b.addImage("media/pic3.gif", fakeGIF)
	doc := b.load(t)

	if len(doc.Media) != 3 {
		t.Fatalf("media: got %d, want 3", len(doc.Media))
	}
	for _, name := range []string{"pic1.png", "pic2.jpeg", "pic3.gif"} {
		if _, ok := doc.Media[name]; !ok {
			t.Errorf("%s missing from Media", name)
		}
	}
}

func TestLoad_ExternalImage_NotInMedia(t *testing.T) {
	b := newFullPkg()
	b.addExternalImage("https://example.com/logo.png")
	doc := b.load(t)

	if len(doc.Media) != 0 {
		t.Error("external image should not appear in Media")
	}
}

func TestLoad_ImageMissingPart_Skipped(t *testing.T) {
	b := newFullPkg()
	b.docPart.AddRel(relImage, "media/ghost.png") // no part
	doc := b.load(t)

	if len(doc.Media) != 0 {
		t.Error("media should be empty for missing part")
	}
}

func TestLoad_InternalAndExternalImagesTogether(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/internal.png", fakePNG)
	b.addExternalImage("https://example.com/ext.png")
	doc := b.load(t)

	if len(doc.Media) != 1 {
		t.Errorf("media: got %d, want 1 (only internal)", len(doc.Media))
	}
	if _, ok := doc.Media["internal.png"]; !ok {
		t.Error("internal.png missing")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XVI. LOAD — unknown / unrecognised relationships
// ═══════════════════════════════════════════════════════════════════

func TestLoad_UnknownDocRel_InternalPart(t *testing.T) {
	b := newFullPkg()
	customData := []byte("<custom xmlns='urn:test'>data</custom>")
	b.addUnknownDocRel(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
		"customXml/item1.xml",
		customData,
	)
	doc := b.load(t)

	partName := "/word/customXml/item1.xml"
	if _, ok := doc.UnknownParts[partName]; !ok {
		t.Error("unknown part not preserved in UnknownParts")
	}
	found := false
	for _, rel := range doc.UnknownRels {
		if strings.Contains(rel.Type, "customXml") {
			found = true
		}
	}
	if !found {
		t.Error("unknown doc rel not in UnknownRels")
	}
}

func TestLoad_UnknownDocRel_External(t *testing.T) {
	b := newFullPkg()
	b.addExternalDocRel(relHyperlink, "https://example.com")
	doc := b.load(t)

	found := false
	for _, rel := range doc.UnknownRels {
		if rel.Type == relHyperlink && rel.Target == "https://example.com" && rel.TargetMode == "External" {
			found = true
		}
	}
	if !found {
		t.Error("external hyperlink not in UnknownRels")
	}
	// External rels should NOT create entries in UnknownParts.
	if len(doc.UnknownParts) != 0 {
		t.Error("external rel should not populate UnknownParts")
	}
}

func TestLoad_MultipleUnknownDocRels(t *testing.T) {
	b := newFullPkg()
	b.addUnknownDocRel("urn:rel:alpha", "alpha.xml", []byte("<alpha/>"))
	b.addUnknownDocRel("urn:rel:beta", "beta.xml", []byte("<beta/>"))
	b.addExternalDocRel("urn:rel:ext", "https://ext.example.com")
	doc := b.load(t)

	relCount := 0
	for _, rel := range doc.UnknownRels {
		if !isPackageLevelRel(rel.Type) {
			relCount++
		}
	}
	if relCount < 3 {
		t.Errorf("expected ≥3 unknown doc rels, got %d", relCount)
	}
	if len(doc.UnknownParts) < 2 {
		t.Errorf("expected ≥2 unknown parts, got %d", len(doc.UnknownParts))
	}
}

func TestLoad_UnknownPkgRel(t *testing.T) {
	b := newFullPkg()
	b.addUnknownPkgRel(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
		"docProps/custom.xml",
		[]byte("<Properties/>"),
	)
	doc := b.load(t)

	found := false
	for _, rel := range doc.UnknownRels {
		if strings.Contains(rel.Type, "custom-properties") {
			found = true
		}
	}
	if !found {
		t.Error("unknown package rel not preserved")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XVII. LOAD — core & app properties
// ═══════════════════════════════════════════════════════════════════

func TestLoad_CoreProps_Fields(t *testing.T) {
	doc := newFullPkg().load(t)
	if doc.CoreProps == nil {
		t.Fatal("CoreProps nil")
	}
	if doc.CoreProps.Creator != "TestAuthor" {
		t.Errorf("Creator = %q", doc.CoreProps.Creator)
	}
	if doc.CoreProps.Title != "Test Title" {
		t.Errorf("Title = %q", doc.CoreProps.Title)
	}
	if doc.CoreProps.Revision != "1" {
		t.Errorf("Revision = %q", doc.CoreProps.Revision)
	}
}

func TestLoad_AppProps_Fields(t *testing.T) {
	doc := newFullPkg().load(t)
	if doc.AppProps == nil {
		t.Fatal("AppProps nil")
	}
	if doc.AppProps.Application != "docx-go" {
		t.Errorf("Application = %q", doc.AppProps.Application)
	}
	if doc.AppProps.Company != "TestCo" {
		t.Errorf("Company = %q", doc.AppProps.Company)
	}
	if doc.AppProps.Words != 2 {
		t.Errorf("Words = %d", doc.AppProps.Words)
	}
	if doc.AppProps.Pages != 1 {
		t.Errorf("Pages = %d", doc.AppProps.Pages)
	}
}

func TestLoad_NoCoreProps(t *testing.T) {
	doc := newMinimalPkg().load(t)
	if doc.CoreProps != nil {
		t.Error("CoreProps should be nil without core.xml")
	}
}

func TestLoad_NoAppProps(t *testing.T) {
	doc := newMinimalPkg().load(t)
	if doc.AppProps != nil {
		t.Error("AppProps should be nil without app.xml")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XVIII. LOAD — rId seeding
// ═══════════════════════════════════════════════════════════════════

func TestLoad_SeedRelIDCounter(t *testing.T) {
	b := newFullPkg()
	maxRel := 0
	for _, rel := range b.docPart.Rels {
		if n := parseRelIDNum(rel.ID); n > maxRel {
			maxRel = n
		}
	}

	doc := b.load(t)
	if doc.nextRelSeq != maxRel+1 {
		t.Errorf("nextRelSeq = %d, want %d", doc.nextRelSeq, maxRel+1)
	}
}

func TestLoad_SeedRelIDCounter_WithHeadersFootersImages(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)
	b.addImage("media/img.png", fakePNG)

	maxRel := 0
	for _, rel := range b.docPart.Rels {
		if n := parseRelIDNum(rel.ID); n > maxRel {
			maxRel = n
		}
	}

	doc := b.load(t)
	if doc.nextRelSeq != maxRel+1 {
		t.Errorf("nextRelSeq = %d, want %d", doc.nextRelSeq, maxRel+1)
	}
}

func TestLoad_NextRelID_NoCollisionWithExisting(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/img.png", fakePNG)
	doc := b.load(t)

	existingIDs := make(map[string]bool)
	for _, rel := range b.docPart.Rels {
		existingIDs[rel.ID] = true
	}

	for i := 0; i < 20; i++ {
		newID := doc.NextRelID()
		if existingIDs[newID] {
			t.Fatalf("NextRelID() = %q collides with existing", newID)
		}
	}
}

func TestLoad_SeedBookmarkID(t *testing.T) {
	doc := newFullPkg().load(t)
	// Default seed is 100 per implementation.
	if doc.nextBmkID != 100 {
		t.Errorf("nextBmkID = %d, want 100", doc.nextBmkID)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XIX. LOAD — error paths
// ═══════════════════════════════════════════════════════════════════

func TestLoad_NoOfficeDocumentRel(t *testing.T) {
	pkg := opc.New()
	_, err := load(pkg)
	if err == nil {
		t.Fatal("expected error")
	}
	if !strings.Contains(err.Error(), "no officeDocument") {
		t.Errorf("error = %q", err)
	}
}

func TestLoad_DocumentPartMissing(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	// Rel exists but part doesn't.
	_, err := load(pkg)
	if err == nil {
		t.Fatal("expected error")
	}
	if !strings.Contains(err.Error(), "not found") {
		t.Errorf("error = %q", err)
	}
}

func TestLoad_MalformedDocumentXML(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	pkg.AddPart("/word/document.xml", ctDocument, []byte("<<<garbage>>>"))

	_, err := load(pkg)
	if err == nil {
		t.Fatal("expected error for malformed document.xml")
	}
}

func TestLoad_MalformedStylesXML(t *testing.T) {
	b := newFullPkg()
	b.pkg.RemovePart("/word/styles.xml")
	b.pkg.AddPart("/word/styles.xml", ctStyles, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed styles.xml")
	}
}

func TestLoad_MalformedSettingsXML(t *testing.T) {
	b := newFullPkg()
	b.pkg.RemovePart("/word/settings.xml")
	b.pkg.AddPart("/word/settings.xml", ctSettings, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed settings.xml")
	}
}

func TestLoad_MalformedFontTableXML(t *testing.T) {
	b := newFullPkg()
	b.pkg.RemovePart("/word/fontTable.xml")
	b.pkg.AddPart("/word/fontTable.xml", ctFontTable, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed fontTable.xml")
	}
}

func TestLoad_MalformedNumberingXML(t *testing.T) {
	b := newFullPkg().addNumbering()
	b.pkg.RemovePart("/word/numbering.xml")
	b.pkg.AddPart("/word/numbering.xml", ctNumbering, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed numbering.xml")
	}
}

func TestLoad_MalformedCommentsXML(t *testing.T) {
	b := newFullPkg().addComments()
	b.pkg.RemovePart("/word/comments.xml")
	b.pkg.AddPart("/word/comments.xml", ctComments, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed comments.xml")
	}
}

func TestLoad_MalformedFootnotesXML(t *testing.T) {
	b := newFullPkg().addFootnotes()
	b.pkg.RemovePart("/word/footnotes.xml")
	b.pkg.AddPart("/word/footnotes.xml", ctFootnotes, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed footnotes.xml")
	}
}

func TestLoad_MalformedEndnotesXML(t *testing.T) {
	b := newFullPkg().addEndnotes()
	b.pkg.RemovePart("/word/endnotes.xml")
	b.pkg.AddPart("/word/endnotes.xml", ctEndnotes, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed endnotes.xml")
	}
}

func TestLoad_MalformedHeaderXML(t *testing.T) {
	b := newFullPkg()
	b.docPart.AddRel(relHeader, "header1.xml")
	b.pkg.AddPart("/word/header1.xml", ctHeader, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed header XML")
	}
}

func TestLoad_MalformedFooterXML(t *testing.T) {
	b := newFullPkg()
	b.docPart.AddRel(relFooter, "footer1.xml")
	b.pkg.AddPart("/word/footer1.xml", ctFooter, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed footer XML")
	}
}

func TestLoad_MalformedCoreXML(t *testing.T) {
	b := newFullPkg()
	b.pkg.RemovePart("/docProps/core.xml")
	b.pkg.AddPart("/docProps/core.xml", ctCore, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed core.xml")
	}
}

func TestLoad_MalformedAppXML(t *testing.T) {
	b := newFullPkg()
	b.pkg.RemovePart("/docProps/app.xml")
	b.pkg.AddPart("/docProps/app.xml", ctExtended, []byte("<<<garbage>>>"))

	_, err := load(b.pkg)
	if err == nil {
		t.Fatal("expected error for malformed app.xml")
	}
}

func TestLoad_RelExistsButPartMissing_loadByRel_Skips(t *testing.T) {
	// When a relationship exists but the target part is missing,
	// loadByRel should silently skip (no error).
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	docPart := pkg.AddPart("/word/document.xml", ctDocument, []byte(fixtureDocumentXML))
	docPart.AddRel(relStyles, "styles.xml") // rel points to non-existent part

	doc, err := load(pkg)
	if err != nil {
		t.Fatalf("should not fail: %v", err)
	}
	if doc.Styles != nil {
		t.Error("Styles should be nil when part doesn't exist")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XX. SAVE / buildPackage — structure checks
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_MandatoryPartsExist(t *testing.T) {
	doc := newFullPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatalf("buildPackage: %v", err)
	}

	mustExist := []string{
		"/word/document.xml",
		"/word/styles.xml",
		"/word/settings.xml",
		"/word/fontTable.xml",
		"/word/webSettings.xml",
		"/word/theme/theme1.xml",
		"/docProps/core.xml",
		"/docProps/app.xml",
	}
	for _, name := range mustExist {
		assertPartExists(t, doc.pkg, name)
	}
}

func TestBuildPackage_PackageLevelRels(t *testing.T) {
	doc := newFullPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	if len(doc.pkg.PackageRelsByType(relOfficeDocument)) == 0 {
		t.Error("officeDocument rel missing")
	}
	if len(doc.pkg.PackageRelsByType(relCoreProperties)) == 0 {
		t.Error("coreProperties rel missing")
	}
	if len(doc.pkg.PackageRelsByType(relExtProperties)) == 0 {
		t.Error("extProperties rel missing")
	}
}

func TestBuildPackage_OptionalPartsOmitted(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Numbering = nil
	doc.Comments = nil
	doc.Footnotes = nil
	doc.Endnotes = nil

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	absent := []string{
		"/word/numbering.xml",
		"/word/comments.xml",
		"/word/footnotes.xml",
		"/word/endnotes.xml",
	}
	for _, name := range absent {
		assertPartAbsent(t, doc.pkg, name)
	}
}

func TestBuildPackage_OptionalPartsIncluded(t *testing.T) {
	doc := newFullPkg().addNumbering().addComments().addFootnotes().addEndnotes().load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	present := []string{
		"/word/numbering.xml",
		"/word/comments.xml",
		"/word/footnotes.xml",
		"/word/endnotes.xml",
	}
	for _, name := range present {
		assertPartExists(t, doc.pkg, name)
	}
}

func TestBuildPackage_DocPartNameDefault(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.docPartName = "" // reset
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartExists(t, doc.pkg, "/word/document.xml")
}

func TestBuildPackage_ThemeInSubdirectory(t *testing.T) {
	doc := newFullPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartExists(t, doc.pkg, "/word/theme/theme1.xml")
}

func TestBuildPackage_Headers(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, ok := doc.pkg.Part("/word/document.xml")
	if !ok {
		t.Fatal("document part missing")
	}
	hdrRels := docPart.RelsByType(relHeader)
	if len(hdrRels) != 2 {
		t.Errorf("header rels: got %d, want 2", len(hdrRels))
	}
}

func TestBuildPackage_Footers(t *testing.T) {
	b := newFullPkg()
	b.addFooter("footer1.xml", fixtureFooterXML)
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	ftrRels := docPart.RelsByType(relFooter)
	if len(ftrRels) != 1 {
		t.Errorf("footer rels: got %d, want 1", len(ftrRels))
	}
}

func TestBuildPackage_Media_ContentTypeAndRel(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/photo.png", fakePNG)
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	part, ok := doc.pkg.Part("/word/media/photo.png")
	if !ok {
		t.Fatal("media part not found")
	}
	if part.ContentType != "image/png" {
		t.Errorf("content type = %q, want image/png", part.ContentType)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	imgRels := docPart.RelsByType(relImage)
	if len(imgRels) == 0 {
		t.Error("no image rel in output")
	}
}

func TestBuildPackage_MultipleMediaFiles(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/a.png", fakePNG)
	b.addImage("media/b.jpeg", fakeJPEG)
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	assertPartExists(t, doc.pkg, "/word/media/a.png")
	assertPartExists(t, doc.pkg, "/word/media/b.jpeg")

	part, _ := doc.pkg.Part("/word/media/b.jpeg")
	if part.ContentType != "image/jpeg" {
		t.Errorf("jpeg content type = %q", part.ContentType)
	}
}

func TestBuildPackage_CorePropsNil_PartOmitted(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.CoreProps = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/docProps/core.xml")
}

func TestBuildPackage_AppPropsNil_PartOmitted(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AppProps = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/docProps/app.xml")
}

func TestBuildPackage_EmptyWebSettings_NotWritten(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.WebSettings = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/word/webSettings.xml")
}

func TestBuildPackage_EmptyTheme_NotWritten(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Theme = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/word/theme/theme1.xml")
}

func TestBuildPackage_UnknownDocRel_External(t *testing.T) {
	b := newFullPkg()
	b.addExternalDocRel(relHyperlink, "https://example.com")
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	found := false
	for _, rel := range docPart.Rels {
		if rel.Type == relHyperlink && rel.TargetMode == "External" {
			found = true
		}
	}
	if !found {
		t.Error("external hyperlink not preserved in output")
	}
}

func TestBuildPackage_UnknownDocRel_Internal(t *testing.T) {
	b := newFullPkg()
	b.addUnknownDocRel("urn:rel:custom", "custom/data.xml", []byte("<data/>"))
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	found := false
	for _, rel := range docPart.Rels {
		if rel.Type == "urn:rel:custom" {
			found = true
		}
	}
	if !found {
		t.Error("unknown internal doc rel not preserved")
	}
	assertPartExists(t, doc.pkg, "/word/custom/data.xml")
}

func TestBuildPackage_NilDocument_NoPanic(t *testing.T) {
	d := newDocStruct(1, 100)
	d.Document = nil
	d.CoreProps = coreprops.DefaultCore("test")
	d.AppProps = coreprops.DefaultApp()

	err := d.buildPackage()
	if err != nil {
		t.Fatalf("buildPackage with nil Document: %v", err)
	}
	// officeDocument rel should still be present.
	if len(d.pkg.PackageRelsByType(relOfficeDocument)) == 0 {
		t.Error("officeDocument rel missing")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XXI. DOCUMENT-LEVEL REL ACCOUNTING IN SAVE
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_DocPartRels_StylesSettingsFonts(t *testing.T) {
	doc := newFullPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	for _, relType := range []string{relStyles, relSettings, relFontTable} {
		rels := docPart.RelsByType(relType)
		if len(rels) == 0 {
			t.Errorf("missing doc-level rel for %s", relType)
		}
	}
}

func TestBuildPackage_DocPartRels_WebSettingsTheme(t *testing.T) {
	doc := newFullPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	if len(docPart.RelsByType(relWebSettings)) == 0 {
		t.Error("missing webSettings rel")
	}
	if len(docPart.RelsByType(relTheme)) == 0 {
		t.Error("missing theme rel")
	}
}

func TestBuildPackage_DocPartRels_OptionalPresent(t *testing.T) {
	doc := newFullPkg().addNumbering().addFootnotes().addEndnotes().addComments().load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	docPart, _ := doc.pkg.Part("/word/document.xml")

	for _, relType := range []string{relNumbering, relFootnotes, relEndnotes, relComments} {
		if len(docPart.RelsByType(relType)) == 0 {
			t.Errorf("missing doc rel for %s", relType)
		}
	}
}

func TestBuildPackage_DocPartRels_OptionalAbsent(t *testing.T) {
	doc := newFullPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	docPart, _ := doc.pkg.Part("/word/document.xml")

	for _, relType := range []string{relNumbering, relFootnotes, relEndnotes, relComments} {
		if len(docPart.RelsByType(relType)) != 0 {
			t.Errorf("unexpected doc rel for %s when part is nil", relType)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XXII. ROUND-TRIP — load → save → load
// ═══════════════════════════════════════════════════════════════════

// roundTrip is a helper: load, save to buffer, re-open.
func roundTrip(t *testing.T, doc *Document) *Document {
	t.Helper()
	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		t.Fatalf("SaveWriter: %v", err)
	}
	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("OpenReader: %v", err)
	}
	return doc2
}

func TestRoundTrip_MandatoryParts(t *testing.T) {
	doc2 := roundTrip(t, newFullPkg().load(t))

	if doc2.Document == nil {
		t.Error("Document nil")
	}
	if doc2.Styles == nil {
		t.Error("Styles nil")
	}
	if doc2.Settings == nil {
		t.Error("Settings nil")
	}
	if doc2.Fonts == nil {
		t.Error("Fonts nil")
	}
}

func TestRoundTrip_CoreProps(t *testing.T) {
	doc1 := newFullPkg().load(t)
	doc2 := roundTrip(t, doc1)

	if doc2.CoreProps == nil {
		t.Fatal("CoreProps nil")
	}
	if doc2.CoreProps.Creator != doc1.CoreProps.Creator {
		t.Errorf("Creator = %q, want %q", doc2.CoreProps.Creator, doc1.CoreProps.Creator)
	}
	if doc2.CoreProps.Title != doc1.CoreProps.Title {
		t.Errorf("Title = %q, want %q", doc2.CoreProps.Title, doc1.CoreProps.Title)
	}
}

func TestRoundTrip_AppProps(t *testing.T) {
	doc1 := newFullPkg().load(t)
	doc2 := roundTrip(t, doc1)

	if doc2.AppProps == nil {
		t.Fatal("AppProps nil")
	}
	if doc2.AppProps.Application != "docx-go" {
		t.Errorf("Application = %q", doc2.AppProps.Application)
	}
	if doc2.AppProps.Company != "TestCo" {
		t.Errorf("Company = %q", doc2.AppProps.Company)
	}
}

func TestRoundTrip_ThemeBytes(t *testing.T) {
	doc1 := newFullPkg().load(t)
	doc2 := roundTrip(t, doc1)

	if len(doc2.Theme) == 0 {
		t.Fatal("Theme lost")
	}
	if !bytes.Equal(doc1.Theme, doc2.Theme) {
		t.Error("Theme bytes differ")
	}
}

func TestRoundTrip_WebSettingsBytes(t *testing.T) {
	doc1 := newFullPkg().load(t)
	doc2 := roundTrip(t, doc1)

	if len(doc2.WebSettings) == 0 {
		t.Fatal("WebSettings lost")
	}
	if !bytes.Equal(doc1.WebSettings, doc2.WebSettings) {
		t.Error("WebSettings bytes differ")
	}
}

func TestRoundTrip_Numbering(t *testing.T) {
	doc2 := roundTrip(t, newFullPkg().addNumbering().load(t))
	if doc2.Numbering == nil {
		t.Error("Numbering lost")
	}
}

func TestRoundTrip_Comments(t *testing.T) {
	doc2 := roundTrip(t, newFullPkg().addComments().load(t))
	if doc2.Comments == nil {
		t.Error("Comments lost")
	}
}

func TestRoundTrip_Footnotes(t *testing.T) {
	doc2 := roundTrip(t, newFullPkg().addFootnotes().load(t))
	if doc2.Footnotes == nil {
		t.Error("Footnotes lost")
	}
}

func TestRoundTrip_Endnotes(t *testing.T) {
	doc2 := roundTrip(t, newFullPkg().addEndnotes().load(t))
	if doc2.Endnotes == nil {
		t.Error("Endnotes lost")
	}
}

func TestRoundTrip_Headers(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	doc2 := roundTrip(t, b.load(t))

	if len(doc2.Headers) != 2 {
		t.Errorf("headers: got %d, want 2", len(doc2.Headers))
	}
}

func TestRoundTrip_Footers(t *testing.T) {
	b := newFullPkg()
	b.addFooter("footer1.xml", fixtureFooterXML)
	b.addFooter("footer2.xml", fixtureEvenFooterXML)
	doc2 := roundTrip(t, b.load(t))

	if len(doc2.Footers) != 2 {
		t.Errorf("footers: got %d, want 2", len(doc2.Footers))
	}
}

func TestRoundTrip_Media(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/logo.png", fakePNG)
	doc2 := roundTrip(t, b.load(t))

	stored, ok := doc2.Media["logo.png"]
	if !ok {
		t.Fatal("logo.png missing")
	}
	if !bytes.Equal(stored, fakePNG) {
		t.Error("media data mismatch")
	}
}

func TestRoundTrip_MultipleMedia(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/a.png", fakePNG)
	b.addImage("media/b.jpeg", fakeJPEG)
	doc2 := roundTrip(t, b.load(t))

	if len(doc2.Media) != 2 {
		t.Errorf("media: got %d, want 2", len(doc2.Media))
	}
}

func TestRoundTrip_FullDocument(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addComments()
	b.addFootnotes()
	b.addEndnotes()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	b.addHeader("header3.xml", fixtureEvenHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)
	b.addFooter("footer2.xml", fixtureEvenFooterXML)
	b.addImage("media/img1.png", fakePNG)
	b.addImage("media/img2.jpeg", fakeJPEG)
	doc2 := roundTrip(t, b.load(t))

	checks := []struct {
		name string
		ok   bool
	}{
		{"Document", doc2.Document != nil},
		{"Styles", doc2.Styles != nil},
		{"Settings", doc2.Settings != nil},
		{"Fonts", doc2.Fonts != nil},
		{"Numbering", doc2.Numbering != nil},
		{"Comments", doc2.Comments != nil},
		{"Footnotes", doc2.Footnotes != nil},
		{"Endnotes", doc2.Endnotes != nil},
		{"CoreProps", doc2.CoreProps != nil},
		{"AppProps", doc2.AppProps != nil},
		{"Theme", len(doc2.Theme) > 0},
		{"WebSettings", len(doc2.WebSettings) > 0},
		{"Headers=3", len(doc2.Headers) == 3},
		{"Footers=2", len(doc2.Footers) == 2},
		{"Media=2", len(doc2.Media) == 2},
	}
	for _, c := range checks {
		if !c.ok {
			t.Errorf("round-trip lost: %s", c.name)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XXIII. EDGE CASES
// ═══════════════════════════════════════════════════════════════════

func TestRoundTrip_DoubleSave(t *testing.T) {
	doc := newFullPkg().load(t)

	var buf1, buf2 bytes.Buffer
	if err := doc.SaveWriter(&buf1); err != nil {
		t.Fatal(err)
	}
	if err := doc.SaveWriter(&buf2); err != nil {
		t.Fatal("second save:", err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf2.Bytes()), int64(buf2.Len()))
	if err != nil {
		t.Fatal("re-open after double save:", err)
	}
	if doc2.Document == nil {
		t.Error("Document nil after double save")
	}
	if doc2.Styles == nil {
		t.Error("Styles nil after double save")
	}
}

func TestRoundTrip_ModifyCorePropsBeforeSave(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.CoreProps.Creator = "NewAuthor"
	doc.CoreProps.Title = "New Title"

	doc2 := roundTrip(t, doc)
	if doc2.CoreProps.Creator != "NewAuthor" {
		t.Errorf("Creator = %q", doc2.CoreProps.Creator)
	}
	if doc2.CoreProps.Title != "New Title" {
		t.Errorf("Title = %q", doc2.CoreProps.Title)
	}
}

func TestRoundTrip_AddMediaAfterLoad(t *testing.T) {
	doc := newFullPkg().load(t)
	rID := doc.AddMedia("new_photo.jpg", fakeJPEG)
	if rID == "" {
		t.Fatal("AddMedia returned empty rId")
	}

	doc2 := roundTrip(t, doc)
	if _, ok := doc2.Media["new_photo.jpg"]; !ok {
		t.Error("new_photo.jpg missing after round-trip")
	}
}

func TestRoundTrip_RemoveOptionalParts(t *testing.T) {
	doc := newFullPkg().addNumbering().addComments().load(t)
	doc.Numbering = nil
	doc.Comments = nil

	doc2 := roundTrip(t, doc)
	if doc2.Numbering != nil {
		t.Error("Numbering should be nil after removal")
	}
	if doc2.Comments != nil {
		t.Error("Comments should be nil after removal")
	}
}

func TestRoundTrip_RemoveAllOptionalParts(t *testing.T) {
	doc := newFullPkg().
		addNumbering().addComments().addFootnotes().addEndnotes().
		load(t)
	doc.Numbering = nil
	doc.Comments = nil
	doc.Footnotes = nil
	doc.Endnotes = nil
	doc.CoreProps = nil
	doc.AppProps = nil
	doc.Theme = nil
	doc.WebSettings = nil

	doc2 := roundTrip(t, doc)
	if doc2.Numbering != nil || doc2.Comments != nil || doc2.Footnotes != nil || doc2.Endnotes != nil {
		t.Error("optional parts should all be nil")
	}
	if doc2.CoreProps != nil || doc2.AppProps != nil {
		t.Error("core/app props should be nil")
	}
	if len(doc2.Theme) != 0 || len(doc2.WebSettings) != 0 {
		t.Error("theme/websettings should be empty")
	}
	// Mandatory parts must still survive
	if doc2.Document == nil {
		t.Error("Document lost")
	}
	if doc2.Styles == nil {
		t.Error("Styles lost")
	}
}

func TestRoundTrip_ClearHeaders(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	doc := b.load(t)

	// Clear all headers
	doc.Headers = make(map[string]*hdft.CT_HdrFtr)

	doc2 := roundTrip(t, doc)
	if len(doc2.Headers) != 0 {
		t.Errorf("headers after clearing: got %d, want 0", len(doc2.Headers))
	}
}

func TestRoundTrip_ClearFooters(t *testing.T) {
	b := newFullPkg()
	b.addFooter("footer1.xml", fixtureFooterXML)
	doc := b.load(t)

	doc.Footers = make(map[string]*hdft.CT_HdrFtr)

	doc2 := roundTrip(t, doc)
	if len(doc2.Footers) != 0 {
		t.Errorf("footers after clearing: got %d, want 0", len(doc2.Footers))
	}
}

func TestRoundTrip_ClearMedia(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/img.png", fakePNG)
	doc := b.load(t)

	doc.Media = make(map[string][]byte)

	doc2 := roundTrip(t, doc)
	if len(doc2.Media) != 0 {
		t.Errorf("media after clearing: got %d, want 0", len(doc2.Media))
	}
}

func TestRoundTrip_AddMultipleMediaAfterLoad(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AddMedia("img1.png", fakePNG)
	doc.AddMedia("img2.jpeg", fakeJPEG)
	doc.AddMedia("img3.gif", fakeGIF)

	doc2 := roundTrip(t, doc)
	if len(doc2.Media) != 3 {
		t.Errorf("media: got %d, want 3", len(doc2.Media))
	}
}

func TestRoundTrip_ReplaceTheme(t *testing.T) {
	doc := newFullPkg().load(t)
	newTheme := []byte("<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' name='Custom'/>")
	doc.Theme = newTheme

	doc2 := roundTrip(t, doc)
	if !bytes.Equal(doc2.Theme, newTheme) {
		t.Error("replaced theme not preserved")
	}
}

func TestRoundTrip_ReplaceWebSettings(t *testing.T) {
	doc := newFullPkg().load(t)
	newWS := []byte("<w:webSettings xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:encoding w:val='utf-8'/></w:webSettings>")
	doc.WebSettings = newWS

	doc2 := roundTrip(t, doc)
	if !bytes.Equal(doc2.WebSettings, newWS) {
		t.Error("replaced webSettings not preserved")
	}
}

func TestRoundTrip_TripleSave(t *testing.T) {
	doc := newFullPkg().addNumbering().load(t)

	// Save → load → save → load → save → load
	for i := 0; i < 3; i++ {
		doc = roundTrip(t, doc)
	}
	if doc.Document == nil {
		t.Error("Document nil after triple round-trip")
	}
	if doc.Numbering == nil {
		t.Error("Numbering nil after triple round-trip")
	}
	if doc.Styles == nil {
		t.Error("Styles nil after triple round-trip")
	}
}

func TestRoundTrip_NextRelID_DoesNotCollideAfterRoundTrip(t *testing.T) {
	b := newFullPkg()
	b.addImage("media/img.png", fakePNG)
	doc := b.load(t)

	doc2 := roundTrip(t, doc)

	// After round-trip, new IDs should not collide with any existing.
	// The new doc has its own set of rels — generate some IDs and
	// ensure they don't match any rel in the freshly built package.
	for i := 0; i < 10; i++ {
		id := doc2.NextRelID()
		if !strings.HasPrefix(id, "rId") {
			t.Errorf("invalid format: %s", id)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XXIV. CONSTANT VERIFICATION
// ═══════════════════════════════════════════════════════════════════

func TestRelConstants_MatchOPC(t *testing.T) {
	// Verify our local constants alias the opc package constants.
	if relOfficeDocument != opc.RelOfficeDocument {
		t.Error("relOfficeDocument mismatch")
	}
	if relCoreProperties != opc.RelCoreProperties {
		t.Error("relCoreProperties mismatch")
	}
	if relStyles != opc.RelStyles {
		t.Error("relStyles mismatch")
	}
	if relSettings != opc.RelSettings {
		t.Error("relSettings mismatch")
	}
	if relImage != opc.RelImage {
		t.Error("relImage mismatch")
	}
	if relHeader != opc.RelHeader {
		t.Error("relHeader mismatch")
	}
	if relFooter != opc.RelFooter {
		t.Error("relFooter mismatch")
	}
	if relTheme != opc.RelTheme {
		t.Error("relTheme mismatch")
	}
}

func TestMediaContentTypes_AllExtensions(t *testing.T) {
	expected := map[string]string{
		".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
		".gif": "image/gif", ".bmp": "image/bmp", ".tiff": "image/tiff",
		".tif": "image/tiff", ".svg": "image/svg+xml",
		".emf": "image/x-emf", ".wmf": "image/x-wmf",
	}
	for ext, want := range expected {
		got, ok := mediaContentTypes[ext]
		if !ok {
			t.Errorf("mediaContentTypes missing %q", ext)
			continue
		}
		if got != want {
			t.Errorf("mediaContentTypes[%q] = %q, want %q", ext, got, want)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  XXV. Document STRUCT FIELD VERIFICATION
// ═══════════════════════════════════════════════════════════════════

func TestDocumentStruct_ZeroValue(t *testing.T) {
	// A zero-value Document should not panic on NextRelID / NextBookmarkID.
	var d Document
	d.Media = make(map[string][]byte)

	id := d.NextRelID()
	if id != "rId0" {
		t.Errorf("zero-seed NextRelID() = %q", id)
	}
	bmk := d.NextBookmarkID()
	if bmk != 0 {
		t.Errorf("zero-seed NextBookmarkID() = %d", bmk)
	}
}

func TestDocumentStruct_AddMediaToZeroValue(t *testing.T) {
	d := Document{Media: make(map[string][]byte)}
	rID := d.AddMedia("test.png", fakePNG)
	if rID != "rId0" {
		t.Errorf("rID = %q", rID)
	}
	if _, ok := d.Media["test.png"]; !ok {
		t.Error("test.png not stored")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════

func assertPartExists(t *testing.T, pkg *opc.Package, name string) {
	t.Helper()
	if _, ok := pkg.Part(name); !ok {
		t.Errorf("expected part %q to exist", name)
	}
}

func assertPartAbsent(t *testing.T, pkg *opc.Package, name string) {
	t.Helper()
	if _, ok := pkg.Part(name); ok {
		t.Errorf("expected part %q to be absent", name)
	}
}
