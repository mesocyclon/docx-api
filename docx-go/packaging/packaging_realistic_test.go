package packaging

// Tests with realistic XML content from reference-appendix §2-3.
// Simulates documents produced by MS Word with:
// - rich formatting (headings, bold/italic, tables)
// - track changes (ins/del)
// - comments, footnotes, endnotes
// - inline images
// - headers with PAGE field codes
// - mc:AlternateContent (VML fallback)
// - section breaks (portrait→landscape)
// - numbered lists
// - hyperlinks
// - Microsoft extension rels (commentsExtended, people)

import (
	"bytes"
	"strings"
	"testing"

	"github.com/vortex/docx-go/opc"
	"github.com/vortex/docx-go/wml/hdft"
)

// ═══════════════════════════════════════════════════════════════════
//  REALISTIC XML FIXTURES (from reference-appendix)
// ═══════════════════════════════════════════════════════════════════

// Document with: heading, bold/italic runs, table 2×2, track changes
// (ins/del), comment ref, footnote ref, inline image ref (rId10),
// numbered list, section break, mc:AlternateContent, two sectPr
// with headerReference/footerReference. This is a realistic Word output.
const realDocumentXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <w:p w:rsidR="00A77B3E" w:rsidRDefault="00A77B3E" w:rsidP="00A77B3E">
      <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>
      <w:r w:rsidRPr="00C83215">
        <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/>
          <w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
        <w:t>Contract Agreement</w:t>
      </w:r>
    </w:p>
    <w:p w:rsidR="00B22C47" w:rsidRDefault="00B22C47">
      <w:r><w:t xml:space="preserve">This is </w:t></w:r>
      <w:r w:rsidRPr="00D714A3"><w:rPr><w:b/><w:bCs/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>italic</w:t></w:r>
      <w:r><w:t xml:space="preserve"> text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">The contract term is </w:t></w:r>
      <w:del w:id="1" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">
        <w:r w:rsidDel="00F12AB3"><w:rPr><w:b/></w:rPr><w:delText>30</w:delText></w:r>
      </w:del>
      <w:ins w:id="2" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">
        <w:r w:rsidR="00F12AB3"><w:rPr><w:b/></w:rPr><w:t>60</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> days.</w:t></w:r>
    </w:p>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>This text has a comment.</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>See footnote</w:t></w:r>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
    </w:p>
    <w:p>
      <w:r><w:rPr><w:noProof/></w:rPr>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0"
                     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <wp:extent cx="1828800" cy="1371600"/>
            <wp:docPr id="1" name="Picture 1" descr="Logo"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>
        <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>
      <w:tr w:rsidR="009A2C41">
        <w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Header 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Header 2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr w:rsidR="009A2C41">
        <w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps"><w:drawing><wp:anchor><!-- shape --></wp:anchor></w:drawing></mc:Choice>
        <mc:Fallback><w:pict><v:rect><!-- VML --></v:rect></w:pict></mc:Fallback>
      </mc:AlternateContent>
    </w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="ListParagraph"/>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>First bullet item</w:t></w:r></w:p>
    <w:p><w:pPr>
      <w:sectPr>
        <w:headerReference w:type="default" r:id="rId8"/>
        <w:footerReference w:type="default" r:id="rId9"/>
        <w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
                 w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr>
    </w:pPr></w:p>
    <w:p><w:r><w:t>Landscape section.</w:t></w:r></w:p>
    <w:sectPr w:rsidR="00000001">
      <w:headerReference w:type="default" r:id="rId8"/>
      <w:headerReference w:type="first" r:id="rId11"/>
      <w:footerReference w:type="default" r:id="rId9"/>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:num="2" w:space="720"/><w:titlePg/><w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`

// Styles with docDefaults + 4 required defaults + Heading1, ListParagraph, TableGrid
const realStylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376"/>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`

// Header with PAGE field code (ref-appendix §3.7)
const realHeaderWithPage = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="right"/></w:pPr>
    <w:r><w:t xml:space="preserve">Page </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:hdr>`

const realFirstHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>CONFIDENTIAL</w:t></w:r></w:p>
</w:hdr>`

const realFooter = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:t>© 2025 TestCo</w:t></w:r></w:p>
</w:ftr>`

// Comments (ref-appendix §3.5) — two comments from different authors
const realCommentsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer" w:date="2025-01-20T14:00:00Z" w:initials="R">
    <w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
    <w:r><w:t>Please verify.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="3" w:author="Legal" w:date="2025-01-21T09:00:00Z" w:initials="L">
    <w:p><w:r><w:t>Needs legal review.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`

// Footnotes with mandatory separators (ref-appendix §3.8, §5.5)
const realFootnotesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> See the original source for details.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`

const realEndnotesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
</w:endnotes>`

// Numbering with bullet + decimal (ref-appendix §3.3)
const realNumberingXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:nsid w:val="3A5C117E"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="E6A2FD28"/>
    <w:lvl w:ilvl="0" w:tplc="04090001"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
    <w:lvl w:ilvl="1" w:tplc="04090003"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`

// Microsoft extension parts (commentsExtended, people)
const msCommentsExtXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w15"/>`

const msPeopleXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:person w15:author="Jane Smith"/></w15:people>`

// Fake PNG with real header
var realPNG = []byte{
	0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
	0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
}

// ═══════════════════════════════════════════════════════════════════
//  BUILDER — production-grade package
// ═══════════════════════════════════════════════════════════════════

func newRealisticPkg() *testPkgBuilder {
	b := &testPkgBuilder{pkg: opc.New()}

	b.pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	b.pkg.AddPackageRel(relCoreProperties, "docProps/core.xml")
	b.pkg.AddPackageRel(relExtProperties, "docProps/app.xml")

	b.docPart = b.pkg.AddPart("/word/document.xml", ctDocument, []byte(realDocumentXML))

	// Document-level rels (rId1..rId16)
	b.docPart.AddRel(relStyles, "styles.xml")
	b.docPart.AddRel(relSettings, "settings.xml")
	b.docPart.AddRel(relWebSettings, "webSettings.xml")
	b.docPart.AddRel(relFontTable, "fontTable.xml")
	b.docPart.AddRel(relTheme, "theme/theme1.xml")
	b.docPart.AddRel(relNumbering, "numbering.xml")
	b.docPart.AddRel(relComments, "comments.xml")
	b.docPart.AddRel(relHeader, "header1.xml")   // default
	b.docPart.AddRel(relFooter, "footer1.xml")   // default
	b.docPart.AddRel(relImage, "media/logo.png") // rId10
	b.docPart.AddRel(relHeader, "header2.xml")   // first
	b.docPart.AddRel(relFootnotes, "footnotes.xml")
	b.docPart.AddRel(relEndnotes, "endnotes.xml")
	b.docPart.AddExternalRel(relHyperlink, "https://example.com/contract")
	// MS extensions (unknown to packaging module)
	b.docPart.AddRel("http://schemas.microsoft.com/office/2011/relationships/commentsExtended", "commentsExtended.xml")
	b.docPart.AddRel("http://schemas.microsoft.com/office/2011/relationships/people", "people.xml")

	// Parts
	b.pkg.AddPart("/word/styles.xml", ctStyles, []byte(realStylesXML))
	b.pkg.AddPart("/word/settings.xml", ctSettings, []byte(fixtureSettingsXML))
	b.pkg.AddPart("/word/webSettings.xml", ctWebSettings, []byte(fixtureWebSettingsXML))
	b.pkg.AddPart("/word/fontTable.xml", ctFontTable, []byte(fixtureFontTableXML))
	b.pkg.AddPart("/word/theme/theme1.xml", ctTheme, []byte(fixtureThemeXML))
	b.pkg.AddPart("/word/numbering.xml", ctNumbering, []byte(realNumberingXML))
	b.pkg.AddPart("/word/comments.xml", ctComments, []byte(realCommentsXML))
	b.pkg.AddPart("/word/header1.xml", ctHeader, []byte(realHeaderWithPage))
	b.pkg.AddPart("/word/header2.xml", ctHeader, []byte(realFirstHeader))
	b.pkg.AddPart("/word/footer1.xml", ctFooter, []byte(realFooter))
	b.pkg.AddPart("/word/footnotes.xml", ctFootnotes, []byte(realFootnotesXML))
	b.pkg.AddPart("/word/endnotes.xml", ctEndnotes, []byte(realEndnotesXML))
	b.pkg.AddPart("/word/media/logo.png", "image/png", realPNG)
	b.pkg.AddPart("/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml", []byte(msCommentsExtXML))
	b.pkg.AddPart("/word/people.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml", []byte(msPeopleXML))
	b.pkg.AddPart("/docProps/core.xml", ctCore, []byte(fixtureCoreXML))
	b.pkg.AddPart("/docProps/app.xml", ctExtended, []byte(fixtureAppXML))

	return b
}

// ═══════════════════════════════════════════════════════════════════
//  LOAD — all parts from realistic package
// ═══════════════════════════════════════════════════════════════════

func TestRealistic_LoadAllParts(t *testing.T) {
	doc := newRealisticPkg().load(t)

	checks := []struct {
		name string
		ok   bool
	}{
		{"Document", doc.Document != nil},
		{"Styles", doc.Styles != nil},
		{"Settings", doc.Settings != nil},
		{"Fonts", doc.Fonts != nil},
		{"Numbering", doc.Numbering != nil},
		{"Comments", doc.Comments != nil},
		{"Footnotes", doc.Footnotes != nil},
		{"Endnotes", doc.Endnotes != nil},
		{"CoreProps", doc.CoreProps != nil},
		{"AppProps", doc.AppProps != nil},
		{"Theme>0", len(doc.Theme) > 0},
		{"WebSettings>0", len(doc.WebSettings) > 0},
		{"Headers=2", len(doc.Headers) == 2},
		{"Footers=1", len(doc.Footers) == 1},
		{"Media=1", len(doc.Media) == 1},
	}
	for _, c := range checks {
		if !c.ok {
			t.Errorf("realistic load: %s failed", c.name)
		}
	}
}

func TestRealistic_MediaBytes(t *testing.T) {
	doc := newRealisticPkg().load(t)
	if stored, ok := doc.Media["logo.png"]; !ok {
		t.Fatal("logo.png missing")
	} else if !bytes.Equal(stored, realPNG) {
		t.Error("PNG bytes changed")
	}
}

func TestRealistic_HyperlinkPreserved(t *testing.T) {
	doc := newRealisticPkg().load(t)
	found := false
	for _, rel := range doc.UnknownRels {
		if rel.Type == relHyperlink && rel.TargetMode == "External" &&
			rel.Target == "https://example.com/contract" {
			found = true
		}
	}
	if !found {
		t.Error("hyperlink not preserved")
	}
}

func TestRealistic_MSExtensionRels(t *testing.T) {
	doc := newRealisticPkg().load(t)
	foundCE, foundPeople := false, false
	for _, rel := range doc.UnknownRels {
		if strings.Contains(rel.Type, "commentsExtended") {
			foundCE = true
		}
		if strings.Contains(rel.Type, "people") {
			foundPeople = true
		}
	}
	if !foundCE {
		t.Error("commentsExtended rel missing")
	}
	if !foundPeople {
		t.Error("people rel missing")
	}
}

func TestRealistic_MSExtensionParts(t *testing.T) {
	doc := newRealisticPkg().load(t)
	if _, ok := doc.UnknownParts["/word/commentsExtended.xml"]; !ok {
		t.Error("commentsExtended part missing")
	}
	if _, ok := doc.UnknownParts["/word/people.xml"]; !ok {
		t.Error("people part missing")
	}
}

func TestRealistic_RelSeeding(t *testing.T) {
	b := newRealisticPkg()
	maxRel := 0
	for _, rel := range b.docPart.Rels {
		if n := parseRelIDNum(rel.ID); n > maxRel {
			maxRel = n
		}
	}
	doc := b.load(t)
	if doc.nextRelSeq <= maxRel {
		t.Errorf("nextRelSeq=%d ≤ maxRel=%d", doc.nextRelSeq, maxRel)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  ROUND-TRIP — realistic document
// ═══════════════════════════════════════════════════════════════════

func TestRealistic_RoundTrip_AllParts(t *testing.T) {
	doc2 := roundTrip(t, newRealisticPkg().load(t))

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
		{"Headers=2", len(doc2.Headers) == 2},
		{"Footers=1", len(doc2.Footers) == 1},
		{"Media=1", len(doc2.Media) == 1},
	}
	for _, c := range checks {
		if !c.ok {
			t.Errorf("round-trip lost: %s", c.name)
		}
	}
}

func TestRealistic_RoundTrip_MediaIdentical(t *testing.T) {
	doc2 := roundTrip(t, newRealisticPkg().load(t))
	if stored, ok := doc2.Media["logo.png"]; !ok {
		t.Fatal("logo.png missing")
	} else if !bytes.Equal(stored, realPNG) {
		t.Error("PNG bytes differ")
	}
}

func TestRealistic_RoundTrip_ThemeIdentical(t *testing.T) {
	doc1 := newRealisticPkg().load(t)
	doc2 := roundTrip(t, doc1)
	if !bytes.Equal(doc1.Theme, doc2.Theme) {
		t.Error("Theme bytes differ")
	}
}

func TestRealistic_RoundTrip_UnknownParts(t *testing.T) {
	doc1 := newRealisticPkg().load(t)
	doc2 := roundTrip(t, doc1)

	for name, data := range doc1.UnknownParts {
		data2, ok := doc2.UnknownParts[name]
		if !ok {
			t.Errorf("unknown part %s lost", name)
			continue
		}
		if !bytes.Equal(data, data2) {
			t.Errorf("unknown part %s bytes differ", name)
		}
	}
}

func TestRealistic_RoundTrip_ExternalRels(t *testing.T) {
	doc1 := newRealisticPkg().load(t)
	doc2 := roundTrip(t, doc1)

	countExt := func(d *Document) int {
		n := 0
		for _, rel := range d.UnknownRels {
			if rel.TargetMode == "External" {
				n++
			}
		}
		return n
	}
	if countExt(doc2) != countExt(doc1) {
		t.Errorf("external rels: %d → %d", countExt(doc1), countExt(doc2))
	}
}

func TestRealistic_RoundTrip_CoreProps(t *testing.T) {
	doc1 := newRealisticPkg().load(t)
	doc2 := roundTrip(t, doc1)
	if doc2.CoreProps.Creator != doc1.CoreProps.Creator {
		t.Errorf("Creator: %q → %q", doc1.CoreProps.Creator, doc2.CoreProps.Creator)
	}
	if doc2.CoreProps.Title != doc1.CoreProps.Title {
		t.Errorf("Title: %q → %q", doc1.CoreProps.Title, doc2.CoreProps.Title)
	}
}

func TestRealistic_RoundTrip_AppProps(t *testing.T) {
	doc1 := newRealisticPkg().load(t)
	doc2 := roundTrip(t, doc1)
	if doc2.AppProps.Application != doc1.AppProps.Application {
		t.Errorf("Application: %q → %q", doc1.AppProps.Application, doc2.AppProps.Application)
	}
	if doc2.AppProps.Words != doc1.AppProps.Words {
		t.Errorf("Words: %d → %d", doc1.AppProps.Words, doc2.AppProps.Words)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MULTI-PASS ROUND-TRIP — stability under repeated saves
// ═══════════════════════════════════════════════════════════════════

func TestRealistic_DoubleRoundTrip(t *testing.T) {
	doc := newRealisticPkg().load(t)
	doc = roundTrip(t, doc)
	doc = roundTrip(t, doc)

	if doc.Document == nil || doc.Numbering == nil || doc.Comments == nil {
		t.Error("parts lost after 2× round-trip")
	}
	if len(doc.Headers) != 2 || len(doc.Media) != 1 {
		t.Error("headers/media count changed")
	}
}

func TestRealistic_QuadrupleRoundTrip(t *testing.T) {
	doc := newRealisticPkg().load(t)
	for i := 0; i < 4; i++ {
		doc = roundTrip(t, doc)
	}
	if doc.Document == nil || doc.Styles == nil || doc.Numbering == nil {
		t.Error("parts lost after 4× round-trip")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MUTATIONS BEFORE SAVE
// ═══════════════════════════════════════════════════════════════════

func TestRealistic_AddMedia_RoundTrip(t *testing.T) {
	doc := newRealisticPkg().load(t)
	rID := doc.AddMedia("diagram.svg", []byte("<svg/>"))
	if rID == "" {
		t.Fatal("empty rId")
	}

	doc2 := roundTrip(t, doc)
	if _, ok := doc2.Media["diagram.svg"]; !ok {
		t.Error("diagram.svg lost")
	}
	if _, ok := doc2.Media["logo.png"]; !ok {
		t.Error("logo.png lost")
	}
}

func TestRealistic_RemoveOptionalParts_RoundTrip(t *testing.T) {
	doc := newRealisticPkg().load(t)
	doc.Comments = nil
	doc.Footnotes = nil
	doc.Endnotes = nil

	doc2 := roundTrip(t, doc)
	if doc2.Comments != nil || doc2.Footnotes != nil || doc2.Endnotes != nil {
		t.Error("removed parts should be nil")
	}
	if doc2.Numbering == nil || len(doc2.Headers) != 2 {
		t.Error("other parts should survive")
	}
}

func TestRealistic_ModifyCoreProps_RoundTrip(t *testing.T) {
	doc := newRealisticPkg().load(t)
	doc.CoreProps.Title = "Updated Contract"
	doc.CoreProps.Creator = "Alice"

	doc2 := roundTrip(t, doc)
	if doc2.CoreProps.Title != "Updated Contract" {
		t.Errorf("Title = %q", doc2.CoreProps.Title)
	}
	if doc2.CoreProps.Creator != "Alice" {
		t.Errorf("Creator = %q", doc2.CoreProps.Creator)
	}
}

func TestRealistic_ClearHeadersFootersMedia_RoundTrip(t *testing.T) {
	doc := newRealisticPkg().load(t)
	doc.Media = make(map[string][]byte)
	doc.Headers = make(map[string]*hdft.CT_HdrFtr)
	doc.Footers = make(map[string]*hdft.CT_HdrFtr)

	doc2 := roundTrip(t, doc)
	if len(doc2.Media) != 0 || len(doc2.Headers) != 0 || len(doc2.Footers) != 0 {
		t.Error("cleared maps should be empty")
	}
	if doc2.Document == nil {
		t.Error("Document lost")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  BUILD PACKAGE — relationship graph
// ═══════════════════════════════════════════════════════════════════

func TestRealistic_BuildPackage_DocRels(t *testing.T) {
	doc := newRealisticPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, ok := doc.pkg.Part("/word/document.xml")
	if !ok {
		t.Fatal("document part missing")
	}

	expected := map[string]int{
		relStyles: 1, relSettings: 1, relWebSettings: 1,
		relFontTable: 1, relTheme: 1, relNumbering: 1,
		relComments: 1, relFootnotes: 1, relEndnotes: 1,
		relHeader: 2, relFooter: 1, relImage: 1,
	}
	for relType, want := range expected {
		got := len(docPart.RelsByType(relType))
		if got != want {
			t.Errorf("%s rels: %d, want %d", relType, got, want)
		}
	}
}

func TestRealistic_BuildPackage_ExternalRel(t *testing.T) {
	doc := newRealisticPkg().load(t)
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
		t.Error("hyperlink not in output")
	}
}

func TestRealistic_BuildPackage_MSExtensionRels(t *testing.T) {
	doc := newRealisticPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	docPart, _ := doc.pkg.Part("/word/document.xml")

	foundCE, foundPeople := false, false
	for _, rel := range docPart.Rels {
		if strings.Contains(rel.Type, "commentsExtended") {
			foundCE = true
		}
		if strings.Contains(rel.Type, "people") {
			foundPeople = true
		}
	}
	if !foundCE {
		t.Error("commentsExtended rel missing")
	}
	if !foundPeople {
		t.Error("people rel missing")
	}
}

func TestRealistic_BuildPackage_MSExtensionParts(t *testing.T) {
	doc := newRealisticPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartExists(t, doc.pkg, "/word/commentsExtended.xml")
	assertPartExists(t, doc.pkg, "/word/people.xml")
}

func TestRealistic_BuildPackage_ImageCT(t *testing.T) {
	doc := newRealisticPkg().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	part, ok := doc.pkg.Part("/word/media/logo.png")
	if !ok {
		t.Fatal("media part missing")
	}
	if part.ContentType != "image/png" {
		t.Errorf("CT=%q", part.ContentType)
	}
}
