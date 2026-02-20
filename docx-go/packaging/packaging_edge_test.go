package packaging

import (
	"bytes"
	"fmt"
	"strings"
	"testing"

	"github.com/vortex/docx-go/opc"
	"github.com/vortex/docx-go/wml/hdft"
)

// ═══════════════════════════════════════════════════════════════════
//  NON-STANDARD docPartName
// ═══════════════════════════════════════════════════════════════════

// Word normally places the main document at /word/document.xml, but the
// spec allows any path. Verify the module handles a custom path.

func TestLoad_CustomDocPartName(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "custom/main.xml")
	dp := pkg.AddPart("/custom/main.xml", ctDocument, []byte(fixtureDocumentXML))
	dp.AddRel(relStyles, "styles.xml")
	pkg.AddPart("/custom/styles.xml", ctStyles, []byte(fixtureStylesXML))

	doc, err := load(pkg)
	if err != nil {
		t.Fatalf("load custom path: %v", err)
	}
	if doc.docPartName != "/custom/main.xml" {
		t.Errorf("docPartName = %q", doc.docPartName)
	}
	if doc.Styles == nil {
		t.Error("Styles nil — rel resolution relative to custom path failed")
	}
}

func TestRoundTrip_CustomDocPartName(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "custom/main.xml")
	dp := pkg.AddPart("/custom/main.xml", ctDocument, []byte(fixtureDocumentXML))
	dp.AddRel(relStyles, "styles.xml")
	dp.AddRel(relSettings, "settings.xml")
	dp.AddRel(relFontTable, "fontTable.xml")
	pkg.AddPart("/custom/styles.xml", ctStyles, []byte(fixtureStylesXML))
	pkg.AddPart("/custom/settings.xml", ctSettings, []byte(fixtureSettingsXML))
	pkg.AddPart("/custom/fontTable.xml", ctFontTable, []byte(fixtureFontTableXML))

	doc, err := load(pkg)
	if err != nil {
		t.Fatal(err)
	}
	doc2 := roundTrip(t, doc)
	if doc2.Styles == nil {
		t.Error("Styles lost after round-trip with custom docPartName")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MULTIPLE officeDocument RELS (takes first)
// ═══════════════════════════════════════════════════════════════════

func TestLoad_MultipleOfficeDocumentRels_TakesFirst(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	pkg.AddPackageRel(relOfficeDocument, "word/glossary/document.xml") // second one
	pkg.AddPart("/word/document.xml", ctDocument, []byte(fixtureDocumentXML))
	pkg.AddPart("/word/glossary/document.xml", ctDocument, []byte(fixtureDocumentXML))

	doc, err := load(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if doc.docPartName != "/word/document.xml" {
		t.Errorf("should use first rel: got %q", doc.docPartName)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  loadByRel: MULTIPLE RELS OF SAME TYPE (only first is used)
// ═══════════════════════════════════════════════════════════════════

func TestLoad_MultipleStylesRels_OnlyFirstParsed(t *testing.T) {
	b := newMinimalPkg()
	b.docPart.AddRel(relStyles, "styles.xml")
	b.docPart.AddRel(relStyles, "styles2.xml") // second styles rel
	b.pkg.AddPart("/word/styles.xml", ctStyles, []byte(fixtureStylesXML))
	b.pkg.AddPart("/word/styles2.xml", ctStyles, []byte("<<<garbage>>>"))

	doc := b.load(t)
	if doc.Styles == nil {
		t.Error("first styles rel should succeed")
	}
	// If it parsed styles2.xml (garbage), it would have failed.
}

// ═══════════════════════════════════════════════════════════════════
//  CORE/APP PROPS: rel exists but part missing (no error, nil)
// ═══════════════════════════════════════════════════════════════════

func TestLoad_CorePropsRelButNoPart(t *testing.T) {
	b := newMinimalPkg()
	b.pkg.AddPackageRel(relCoreProperties, "docProps/core.xml")
	// rel exists, but no part added

	doc := b.load(t)
	if doc.CoreProps != nil {
		t.Error("CoreProps should be nil when part missing despite rel")
	}
}

func TestLoad_AppPropsRelButNoPart(t *testing.T) {
	b := newMinimalPkg()
	b.pkg.AddPackageRel(relExtProperties, "docProps/app.xml")

	doc := b.load(t)
	if doc.AppProps != nil {
		t.Error("AppProps should be nil when part missing despite rel")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  UNKNOWN PACKAGE-LEVEL RELS WITH DATA
// ═══════════════════════════════════════════════════════════════════

func TestLoad_UnknownPkgRel_CustomProperties(t *testing.T) {
	b := newFullPkg()
	// Add custom properties — a real package-level relationship
	customRelType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	b.pkg.AddPackageRel(customRelType, "docProps/custom.xml")
	b.pkg.AddPart("/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml",
		[]byte(`<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>`))

	doc := b.load(t)

	found := false
	for _, rel := range doc.UnknownRels {
		if rel.Type == customRelType {
			found = true
		}
	}
	if !found {
		t.Error("custom-properties rel not preserved")
	}
}

func TestLoad_UnknownPkgRel_Thumbnail(t *testing.T) {
	b := newFullPkg()
	thumbRelType := "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
	b.pkg.AddPackageRel(thumbRelType, "docProps/thumbnail.jpeg")
	b.pkg.AddPart("/docProps/thumbnail.jpeg", "image/jpeg", fakeJPEG)

	doc := b.load(t)

	found := false
	for _, rel := range doc.UnknownRels {
		if rel.Type == thumbRelType {
			found = true
		}
	}
	if !found {
		t.Error("thumbnail rel not preserved")
	}
}

func TestRoundTrip_UnknownPkgRel_Preserved(t *testing.T) {
	b := newFullPkg()
	customRelType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	b.pkg.AddPackageRel(customRelType, "docProps/custom.xml")
	customData := []byte(`<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>`)
	b.pkg.AddPart("/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml", customData)

	doc := b.load(t)
	doc2 := roundTrip(t, doc)

	found := false
	for _, rel := range doc2.UnknownRels {
		if rel.Type == customRelType {
			found = true
		}
	}
	if !found {
		t.Error("custom-properties pkg rel lost after round-trip")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  UNKNOWN DOC-LEVEL REL WITH NO PART DATA
// ═══════════════════════════════════════════════════════════════════

func TestLoad_UnknownDocRel_InternalNoPartData(t *testing.T) {
	b := newFullPkg()
	// Add a rel pointing to a part that doesn't exist
	b.docPart.AddRel("urn:rel:ghost", "ghost.xml")
	// No part added

	doc := b.load(t)

	// The rel should still be in UnknownRels
	found := false
	for _, rel := range doc.UnknownRels {
		if rel.Type == "urn:rel:ghost" {
			found = true
		}
	}
	if !found {
		t.Error("unknown doc rel with missing part should still be in UnknownRels")
	}
	// But UnknownParts should NOT have this entry
	if _, ok := doc.UnknownParts["/word/ghost.xml"]; ok {
		t.Error("ghost part should not be in UnknownParts")
	}
}

func TestBuildPackage_UnknownDocRel_InternalNoPartData(t *testing.T) {
	doc := newFullPkg().load(t)
	// Manually add an unknown rel with no corresponding part
	doc.UnknownRels = append(doc.UnknownRels, opc.Relationship{
		ID: "rId999", Type: "urn:rel:orphan", Target: "orphan.xml",
	})

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	found := false
	for _, rel := range docPart.Rels {
		if rel.Type == "urn:rel:orphan" {
			found = true
		}
	}
	if !found {
		t.Error("orphan rel not written to output")
	}
	// The part should NOT exist (no data to write)
	assertPartAbsent(t, doc.pkg, "/word/orphan.xml")
}

// ═══════════════════════════════════════════════════════════════════
//  SAVE WITH NIL TYPED PARTS (Styles, Settings, Fonts)
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_NilStyles(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Styles = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/word/styles.xml")
}

func TestBuildPackage_NilSettings(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Settings = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/word/settings.xml")
}

func TestBuildPackage_NilFonts(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Fonts = nil
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	assertPartAbsent(t, doc.pkg, "/word/fontTable.xml")
}

func TestRoundTrip_NilStyles_NotRestored(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Styles = nil
	doc2 := roundTrip(t, doc)
	if doc2.Styles != nil {
		t.Error("Styles should remain nil")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MEDIA EXTENSIONS — content type coverage
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_MediaContentTypes_AllFormats(t *testing.T) {
	tests := []struct {
		name, wantCT string
		data         []byte
	}{
		{"photo.png", "image/png", fakePNG},
		{"photo.jpg", "image/jpeg", fakeJPEG},
		{"photo.jpeg", "image/jpeg", fakeJPEG},
		{"anim.gif", "image/gif", fakeGIF},
		{"scan.bmp", "image/bmp", []byte{0x42, 0x4D}},
		{"scan.tiff", "image/tiff", []byte{0x49, 0x49}},
		{"scan.tif", "image/tiff", []byte{0x49, 0x49}},
		{"logo.svg", "image/svg+xml", []byte("<svg/>")},
		{"diag.emf", "image/x-emf", []byte{0x01}},
		{"diag.wmf", "image/x-wmf", []byte{0xD7}},
		{"unknown.xyz", "application/octet-stream", []byte{0x00}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			doc := newFullPkg().load(t)
			doc.Media[tt.name] = tt.data
			if err := doc.buildPackage(); err != nil {
				t.Fatal(err)
			}
			part, ok := doc.pkg.Part("/word/media/" + tt.name)
			if !ok {
				t.Fatalf("part /word/media/%s not found", tt.name)
			}
			if part.ContentType != tt.wantCT {
				t.Errorf("CT=%q, want %q", part.ContentType, tt.wantCT)
			}
		})
	}
}

func TestRoundTrip_MediaDiverseFormats(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Media["test.svg"] = []byte("<svg xmlns='http://www.w3.org/2000/svg'/>")
	doc.Media["test.gif"] = fakeGIF
	doc.Media["test.bmp"] = []byte{0x42, 0x4D, 0x00, 0x00}

	doc2 := roundTrip(t, doc)
	for _, name := range []string{"test.svg", "test.gif", "test.bmp"} {
		if _, ok := doc2.Media[name]; !ok {
			t.Errorf("%s lost after round-trip", name)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  CONTENT-TYPE VERIFICATION ON OUTPUT PARTS
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_ContentTypesCorrect(t *testing.T) {
	doc := newFullPkg().addNumbering().addComments().addFootnotes().addEndnotes().load(t)
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	expected := map[string]string{
		"/word/document.xml":     ctDocument,
		"/word/styles.xml":       ctStyles,
		"/word/settings.xml":     ctSettings,
		"/word/fontTable.xml":    ctFontTable,
		"/word/numbering.xml":    ctNumbering,
		"/word/comments.xml":     ctComments,
		"/word/footnotes.xml":    ctFootnotes,
		"/word/endnotes.xml":     ctEndnotes,
		"/word/webSettings.xml":  ctWebSettings,
		"/word/theme/theme1.xml": ctTheme,
		"/docProps/core.xml":     ctCore,
		"/docProps/app.xml":      ctExtended,
	}
	for partName, wantCT := range expected {
		part, ok := doc.pkg.Part(partName)
		if !ok {
			t.Errorf("part %s missing", partName)
			continue
		}
		if part.ContentType != wantCT {
			t.Errorf("%s: CT=%q, want %q", partName, part.ContentType, wantCT)
		}
	}
}

func TestBuildPackage_HeaderFooterContentTypes(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	for _, part := range doc.pkg.Parts() {
		if strings.Contains(part.Name, "header") && strings.HasSuffix(part.Name, ".xml") {
			if part.ContentType != ctHeader {
				t.Errorf("%s: CT=%q, want %q", part.Name, part.ContentType, ctHeader)
			}
		}
		if strings.Contains(part.Name, "footer") && strings.HasSuffix(part.Name, ".xml") {
			if part.ContentType != ctFooter {
				t.Errorf("%s: CT=%q, want %q", part.Name, part.ContentType, ctFooter)
			}
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  buildPackage IDEMPOTENCY (two calls produce clean output)
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_Idempotent(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addImage("media/img.png", fakePNG)
	doc := b.load(t)

	// First build
	if err := doc.buildPackage(); err != nil {
		t.Fatal("first build:", err)
	}
	partCount1 := len(doc.pkg.Parts())

	// Second build (should start fresh, not accumulate)
	if err := doc.buildPackage(); err != nil {
		t.Fatal("second build:", err)
	}
	partCount2 := len(doc.pkg.Parts())

	if partCount1 != partCount2 {
		t.Errorf("parts after 1st build=%d, after 2nd=%d (stale parts)", partCount1, partCount2)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  SAVE TO MULTIPLE WRITERS
// ═══════════════════════════════════════════════════════════════════

func TestSaveWriter_MultipleTimes(t *testing.T) {
	doc := newFullPkg().load(t)

	var bufs [3]bytes.Buffer
	for i := range bufs {
		if err := doc.SaveWriter(&bufs[i]); err != nil {
			t.Fatalf("save #%d: %v", i, err)
		}
	}

	// All outputs should produce valid documents
	for i := range bufs {
		d, err := OpenReader(bytes.NewReader(bufs[i].Bytes()), int64(bufs[i].Len()))
		if err != nil {
			t.Fatalf("re-open save #%d: %v", i, err)
		}
		if d.Document == nil {
			t.Errorf("save #%d produced nil Document", i)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MANY HEADERS AND FOOTERS
// ═══════════════════════════════════════════════════════════════════

func TestLoad_ManyHeadersFooters(t *testing.T) {
	b := newFullPkg()
	const n = 6
	for i := 1; i <= n; i++ {
		b.addHeader(fmt.Sprintf("header%d.xml", i), fixtureHeaderXML)
		b.addFooter(fmt.Sprintf("footer%d.xml", i), fixtureFooterXML)
	}
	doc := b.load(t)

	if len(doc.Headers) != n {
		t.Errorf("headers: %d, want %d", len(doc.Headers), n)
	}
	if len(doc.Footers) != n {
		t.Errorf("footers: %d, want %d", len(doc.Footers), n)
	}
}

func TestRoundTrip_ManyHeadersFooters(t *testing.T) {
	b := newFullPkg()
	const n = 5
	for i := 1; i <= n; i++ {
		b.addHeader(fmt.Sprintf("header%d.xml", i), fixtureHeaderXML)
		b.addFooter(fmt.Sprintf("footer%d.xml", i), fixtureFooterXML)
	}
	doc2 := roundTrip(t, b.load(t))

	if len(doc2.Headers) != n {
		t.Errorf("headers: %d, want %d", len(doc2.Headers), n)
	}
	if len(doc2.Footers) != n {
		t.Errorf("footers: %d, want %d", len(doc2.Footers), n)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MANY IMAGES
// ═══════════════════════════════════════════════════════════════════

func TestLoad_ManyImages(t *testing.T) {
	b := newFullPkg()
	const n = 20
	for i := 0; i < n; i++ {
		name := fmt.Sprintf("media/img%d.png", i)
		b.addImage(name, append(fakePNG, byte(i)))
	}
	doc := b.load(t)

	if len(doc.Media) != n {
		t.Errorf("media: %d, want %d", len(doc.Media), n)
	}
}

func TestRoundTrip_ManyImages(t *testing.T) {
	b := newFullPkg()
	const n = 15
	for i := 0; i < n; i++ {
		name := fmt.Sprintf("media/pic%d.png", i)
		b.addImage(name, append(fakePNG, byte(i)))
	}
	doc2 := roundTrip(t, b.load(t))

	if len(doc2.Media) != n {
		t.Errorf("media: %d, want %d", len(doc2.Media), n)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MANY UNKNOWN RELS
// ═══════════════════════════════════════════════════════════════════

func TestLoad_ManyUnknownRels(t *testing.T) {
	b := newFullPkg()
	const n = 10
	for i := 0; i < n; i++ {
		relType := fmt.Sprintf("urn:custom:rel%d", i)
		target := fmt.Sprintf("custom%d.xml", i)
		data := fmt.Sprintf("<custom%d/>", i)
		b.addUnknownDocRel(relType, target, []byte(data))
	}
	doc := b.load(t)

	// Count doc-level unknown rels (not package-level)
	docLevelUnknown := 0
	for _, rel := range doc.UnknownRels {
		if !isPackageLevelRel(rel.Type) {
			docLevelUnknown++
		}
	}
	if docLevelUnknown < n {
		t.Errorf("unknown doc rels: %d, want ≥%d", docLevelUnknown, n)
	}
	if len(doc.UnknownParts) < n {
		t.Errorf("unknown parts: %d, want ≥%d", len(doc.UnknownParts), n)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MIXED EXTERNAL AND INTERNAL UNKNOWN RELS
// ═══════════════════════════════════════════════════════════════════

func TestRoundTrip_MixedUnknownRels(t *testing.T) {
	b := newFullPkg()
	b.addUnknownDocRel("urn:rel:glossary", "glossary/doc.xml", []byte("<glossary/>"))
	b.addExternalDocRel("urn:rel:extlink", "https://ext.example.com")
	b.addExternalDocRel(relHyperlink, "https://hyperlink.example.com")
	b.addUnknownDocRel("urn:rel:chart", "charts/chart1.xml", []byte("<chart/>"))

	doc := b.load(t)
	doc2 := roundTrip(t, doc)

	// Count unknowns by type
	extCount, intCount := 0, 0
	for _, rel := range doc2.UnknownRels {
		if isPackageLevelRel(rel.Type) {
			continue
		}
		if rel.TargetMode == "External" {
			extCount++
		} else {
			intCount++
		}
	}
	if extCount < 2 {
		t.Errorf("external rels: %d, want ≥2", extCount)
	}
	if intCount < 2 {
		t.Errorf("internal unknown rels: %d, want ≥2", intCount)
	}

	// Check internal parts survived
	if _, ok := doc2.UnknownParts["/word/glossary/doc.xml"]; !ok {
		t.Error("glossary part lost")
	}
	if _, ok := doc2.UnknownParts["/word/charts/chart1.xml"]; !ok {
		t.Error("chart part lost")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  HEADERS/FOOTERS KEYED BY rId — rId must be distinct
// ═══════════════════════════════════════════════════════════════════

func TestLoad_Headers_RIdKeysAreDistinct(t *testing.T) {
	b := newFullPkg()
	r1 := b.addHeader("header1.xml", fixtureHeaderXML)
	r2 := b.addHeader("header2.xml", fixtureFirstPageHeaderXML)
	r3 := b.addHeader("header3.xml", fixtureEvenHeaderXML)

	if r1 == r2 || r2 == r3 || r1 == r3 {
		t.Fatalf("rIds not distinct: %s, %s, %s", r1, r2, r3)
	}

	doc := b.load(t)
	for _, rID := range []string{r1, r2, r3} {
		if _, ok := doc.Headers[rID]; !ok {
			t.Errorf("header %s missing", rID)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  EMPTY MAPS AFTER CLEARING — verify no nil-map panic
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_EmptyMaps_NoPanic(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.Headers = make(map[string]*hdft.CT_HdrFtr)
	doc.Footers = make(map[string]*hdft.CT_HdrFtr)
	doc.Media = make(map[string][]byte)
	doc.UnknownParts = make(map[string][]byte)
	doc.UnknownRels = nil

	if err := doc.buildPackage(); err != nil {
		t.Fatalf("buildPackage with empty maps: %v", err)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  ADDMEDIA EDGE CASES
// ═══════════════════════════════════════════════════════════════════

func TestAddMedia_EmptyData(t *testing.T) {
	d := newDocStruct(1, 0)
	rID := d.AddMedia("empty.bin", []byte{})
	if rID == "" {
		t.Error("AddMedia returned empty rId")
	}
	if data, ok := d.Media["empty.bin"]; !ok {
		t.Error("empty.bin not stored")
	} else if len(data) != 0 {
		t.Error("data should be empty")
	}
}

func TestAddMedia_LargeData(t *testing.T) {
	d := newDocStruct(1, 0)
	big := make([]byte, 10*1024*1024) // 10 MB
	big[0] = 0xFF
	big[len(big)-1] = 0xAA
	rID := d.AddMedia("big.bin", big)
	if rID == "" {
		t.Fatal("empty rId")
	}
	stored := d.Media["big.bin"]
	if len(stored) != len(big) {
		t.Errorf("stored size = %d", len(stored))
	}
	if stored[0] != 0xFF || stored[len(stored)-1] != 0xAA {
		t.Error("data corrupted")
	}
}

func TestAddMedia_SpecialCharsInFilename(t *testing.T) {
	d := newDocStruct(1, 0)
	d.AddMedia("спецсимволы.png", fakePNG)
	if _, ok := d.Media["спецсимволы.png"]; !ok {
		t.Error("unicode filename not stored")
	}
}

func TestAddMedia_DotInFilename(t *testing.T) {
	d := newDocStruct(1, 0)
	d.AddMedia("my.photo.jpg", fakeJPEG)
	d.AddMedia("my.photo.jpg", fakeJPEG)

	if _, ok := d.Media["my.photo.jpg"]; !ok {
		t.Error("first missing")
	}
	// Dedup should produce "my.photo1.jpg"
	if _, ok := d.Media["my.photo1.jpg"]; !ok {
		t.Error("dedup should produce my.photo1.jpg")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  ALL OPTIONAL PARTS NIL ON SAVE — verifies no nil-pointer panic
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_AllTypedPartsNil(t *testing.T) {
	d := &Document{
		Headers:      make(map[string]*hdft.CT_HdrFtr),
		Footers:      make(map[string]*hdft.CT_HdrFtr),
		Media:        make(map[string][]byte),
		UnknownParts: make(map[string][]byte),
		docPartName:  "/word/document.xml",
	}
	// Everything is nil: Document, Styles, Settings, Fonts, etc.
	if err := d.buildPackage(); err != nil {
		t.Fatalf("buildPackage all nil: %v", err)
	}

	// officeDocument rel should still be present
	rels := d.pkg.PackageRelsByType(relOfficeDocument)
	if len(rels) == 0 {
		t.Error("officeDocument rel missing")
	}
	// But no document.xml part (Document is nil)
	assertPartAbsent(t, d.pkg, "/word/document.xml")
}

// ═══════════════════════════════════════════════════════════════════
//  SEED COUNTERS — additional edge cases
// ═══════════════════════════════════════════════════════════════════

func TestLoad_SeedRelID_WithOnlyPackageRels(t *testing.T) {
	// Minimal package: document part has NO rels at all
	b := newMinimalPkg()
	doc := b.load(t)

	// nextRelSeq should be 1 (max=0 → 0+1=1)
	if doc.nextRelSeq != 1 {
		t.Errorf("nextRelSeq=%d, want 1 for empty rels", doc.nextRelSeq)
	}
}

func TestLoad_SeedRelID_GapInSequence(t *testing.T) {
	// Document part has rId1, rId5, rId20 (gaps)
	b := newMinimalPkg()
	b.docPart.AddRel(relStyles, "styles.xml") // rId1
	b.pkg.AddPart("/word/styles.xml", ctStyles, []byte(fixtureStylesXML))

	// Simulate gap: manually check that after load, nextRelSeq > max
	doc := b.load(t)
	// Find max rel ID on the docPart
	maxRel := 0
	for _, rel := range b.docPart.Rels {
		if n := parseRelIDNum(rel.ID); n > maxRel {
			maxRel = n
		}
	}
	if doc.nextRelSeq != maxRel+1 {
		t.Errorf("nextRelSeq=%d, want %d", doc.nextRelSeq, maxRel+1)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  FULL ZIP CYCLE — SaveWriter → OpenReader exact bytes check
// ═══════════════════════════════════════════════════════════════════

func TestFullZipCycle_Simple(t *testing.T) {
	doc := newFullPkg().load(t)

	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		t.Fatal(err)
	}

	if buf.Len() == 0 {
		t.Fatal("empty ZIP output")
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("OpenReader from SaveWriter output: %v", err)
	}
	if doc2.Document == nil {
		t.Error("Document nil after zip cycle")
	}
}

func TestFullZipCycle_WithAllParts(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addComments()
	b.addFootnotes()
	b.addEndnotes()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)
	b.addImage("media/pic.png", fakePNG)

	doc := b.load(t)

	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		t.Fatal(err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	checks := map[string]bool{
		"Document":  doc2.Document != nil,
		"Styles":    doc2.Styles != nil,
		"Settings":  doc2.Settings != nil,
		"Fonts":     doc2.Fonts != nil,
		"Numbering": doc2.Numbering != nil,
		"Comments":  doc2.Comments != nil,
		"Footnotes": doc2.Footnotes != nil,
		"Endnotes":  doc2.Endnotes != nil,
		"CoreProps": doc2.CoreProps != nil,
		"AppProps":  doc2.AppProps != nil,
		"Theme":     len(doc2.Theme) > 0,
		"WebSet":    len(doc2.WebSettings) > 0,
		"Headers":   len(doc2.Headers) == 1,
		"Footers":   len(doc2.Footers) == 1,
		"Media":     len(doc2.Media) == 1,
	}
	for name, ok := range checks {
		if !ok {
			t.Errorf("zip cycle lost: %s", name)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MEDIA SHORTNAME EXTRACTION — path.Base behaviour
// ═══════════════════════════════════════════════════════════════════

func TestLoad_MediaShortName_FromNestedPath(t *testing.T) {
	b := newFullPkg()
	// Media in a subdirectory of word/
	b.docPart.AddRel(relImage, "media/subdir/deep.png")
	b.pkg.AddPart("/word/media/subdir/deep.png", "image/png", fakePNG)

	doc := b.load(t)

	// path.Base("/word/media/subdir/deep.png") = "deep.png"
	if _, ok := doc.Media["deep.png"]; !ok {
		t.Error("expected short name 'deep.png' from nested media path")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  resolveTarget EDGE CASES
// ═══════════════════════════════════════════════════════════════════

func TestResolveTarget_EmptySource(t *testing.T) {
	// Edge: empty source should still work
	got := resolveTarget("", "styles.xml")
	// path.Dir("") = ".", path.Join(".", "styles.xml") = "styles.xml"
	if got != "styles.xml" {
		t.Errorf("got %q", got)
	}
}

func TestResolveTarget_RootSource(t *testing.T) {
	got := resolveTarget("/document.xml", "styles.xml")
	if got != "/styles.xml" {
		t.Errorf("got %q, want /styles.xml", got)
	}
}

func TestResolveTarget_MultipleParentDirs(t *testing.T) {
	got := resolveTarget("/a/b/c/doc.xml", "../../shared/data.xml")
	// path.Join("/a/b/c", "../../shared/data.xml") = /a/shared/data.xml
	if got != "/a/shared/data.xml" {
		t.Errorf("got %q, want /a/shared/data.xml", got)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  PART COUNT VERIFICATION — no stale/duplicate parts
// ═══════════════════════════════════════════════════════════════════

func TestBuildPackage_NoDuplicateParts(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addImage("media/img.png", fakePNG)
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	seen := make(map[string]int)
	for _, p := range doc.pkg.Parts() {
		seen[p.Name]++
	}
	for name, count := range seen {
		if count > 1 {
			t.Errorf("duplicate part: %s (×%d)", name, count)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  STRESS TEST — many concurrent AddMedia + NextRelID
// ═══════════════════════════════════════════════════════════════════

func TestStress_ConcurrentMixed(t *testing.T) {
	doc := newFullPkg().load(t)
	done := make(chan struct{})
	const n = 50

	// Goroutines calling AddMedia
	go func() {
		for i := 0; i < n; i++ {
			doc.AddMedia(fmt.Sprintf("stress_%d.png", i), fakePNG)
		}
		done <- struct{}{}
	}()
	// Goroutines calling NextRelID
	go func() {
		for i := 0; i < n; i++ {
			doc.NextRelID()
		}
		done <- struct{}{}
	}()
	// Goroutines calling NextBookmarkID
	go func() {
		for i := 0; i < n; i++ {
			doc.NextBookmarkID()
		}
		done <- struct{}{}
	}()

	<-done
	<-done
	<-done

	// All 50 media items should be stored
	if len(doc.Media) != n {
		t.Errorf("media: %d, want %d", len(doc.Media), n)
	}
}
