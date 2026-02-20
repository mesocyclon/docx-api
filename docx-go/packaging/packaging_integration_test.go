package packaging

// Integration tests exercising multi-step workflows, cross-part
// interactions, sequential mutations across round-trips, and
// relationship integrity after complex operations.

import (
	"bytes"
	"fmt"
	"strings"
	"testing"

	"github.com/vortex/docx-go/coreprops"
	"github.com/vortex/docx-go/opc"
	"github.com/vortex/docx-go/wml/hdft"
)

// ═══════════════════════════════════════════════════════════════════
//  MULTI-STEP MUTATION CHAINS
// ═══════════════════════════════════════════════════════════════════

// Load → add media → remove comments → save → load → add more → save → verify
func TestIntegration_ChainedMutations(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addComments()
	b.addFootnotes()
	b.addEndnotes()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)
	b.addImage("media/pic.png", fakePNG)

	// Step 1: Load
	doc := b.load(t)
	if doc.Comments == nil {
		t.Fatal("comments nil after initial load")
	}

	// Step 2: Add media, remove comments
	rId1 := doc.AddMedia("chart.svg", []byte("<svg/>"))
	doc.Comments = nil

	// Step 3: Round-trip
	doc = roundTrip(t, doc)
	if doc.Comments != nil {
		t.Error("comments should be nil after removal")
	}
	if _, ok := doc.Media["chart.svg"]; !ok {
		t.Error("chart.svg lost")
	}
	if _, ok := doc.Media["pic.png"]; !ok {
		t.Error("original pic.png lost")
	}

	// Step 4: More mutations — add header-like content, modify props
	rId2 := doc.AddMedia("logo.jpg", fakeJPEG)
	doc.CoreProps.Title = "Revised Document"
	doc.Endnotes = nil

	// Step 5: Round-trip again
	doc = roundTrip(t, doc)
	if _, ok := doc.Media["logo.jpg"]; !ok {
		t.Error("logo.jpg lost in second round-trip")
	}
	if _, ok := doc.Media["chart.svg"]; !ok {
		t.Error("chart.svg lost in second round-trip")
	}
	if doc.CoreProps.Title != "Revised Document" {
		t.Errorf("Title=%q", doc.CoreProps.Title)
	}
	if doc.Endnotes != nil {
		t.Error("endnotes should remain nil")
	}
	if doc.Numbering == nil {
		t.Error("numbering should survive")
	}
	if doc.Footnotes == nil {
		t.Error("footnotes should survive")
	}

	// rIds should be non-empty. Note: after round-trip the seed resets
	// from package rels, so rId2 may equal rId1 (which was consumed in
	// the previous session). That's correct — seedRelIDCounter only
	// knows about rels written into the package.
	if rId1 == "" {
		t.Error("rId1 empty")
	}
	if rId2 == "" {
		t.Error("rId2 empty")
	}
}

// Load → add 10 images one-by-one with round-trips between each
func TestIntegration_IncrementalMediaAddition(t *testing.T) {
	doc := newFullPkg().load(t)

	for i := 0; i < 10; i++ {
		name := fmt.Sprintf("step%d.png", i)
		doc.AddMedia(name, append(fakePNG, byte(i)))
		doc = roundTrip(t, doc)
	}

	if len(doc.Media) != 10 {
		t.Errorf("media count = %d, want 10", len(doc.Media))
	}
	for i := 0; i < 10; i++ {
		name := fmt.Sprintf("step%d.png", i)
		data, ok := doc.Media[name]
		if !ok {
			t.Errorf("%s missing", name)
			continue
		}
		if data[len(data)-1] != byte(i) {
			t.Errorf("%s data corrupted", name)
		}
	}
}

// Load → remove all optional parts one-by-one, round-trip each time
func TestIntegration_IncrementalPartRemoval(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addComments()
	b.addFootnotes()
	b.addEndnotes()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addFooter("footer1.xml", fixtureFooterXML)

	doc := b.load(t)

	doc.Numbering = nil
	doc = roundTrip(t, doc)
	if doc.Numbering != nil {
		t.Error("numbering should be nil")
	}

	doc.Comments = nil
	doc = roundTrip(t, doc)
	if doc.Comments != nil {
		t.Error("comments should be nil")
	}

	doc.Footnotes = nil
	doc = roundTrip(t, doc)
	if doc.Footnotes != nil {
		t.Error("footnotes should be nil")
	}

	doc.Endnotes = nil
	doc = roundTrip(t, doc)
	if doc.Endnotes != nil {
		t.Error("endnotes should be nil")
	}

	doc.Headers = make(map[string]*hdft.CT_HdrFtr)
	doc = roundTrip(t, doc)
	if len(doc.Headers) != 0 {
		t.Error("headers should be empty")
	}

	doc.Footers = make(map[string]*hdft.CT_HdrFtr)
	doc = roundTrip(t, doc)
	if len(doc.Footers) != 0 {
		t.Error("footers should be empty")
	}

	// Mandatory parts should still survive
	if doc.Document == nil || doc.Styles == nil || doc.Settings == nil || doc.Fonts == nil {
		t.Error("mandatory parts lost during incremental removal")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  RELATIONSHIP ID INTEGRITY AFTER COMPLEX OPERATIONS
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_NextRelID_NeverCollidesAfterMutations(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addImage("media/img.png", fakePNG)

	doc := b.load(t)

	// Generate several rIds within this session — must be unique
	ids := make(map[string]bool)
	for i := 0; i < 20; i++ {
		id := doc.NextRelID()
		if ids[id] {
			t.Fatalf("duplicate rId: %s at iteration %d", id, i)
		}
		ids[id] = true
	}

	// Round-trip — seed resets from package rels (IDs consumed above
	// via NextRelID are NOT written as rels, so they are "forgotten").
	doc = roundTrip(t, doc)

	// After round-trip: generated IDs should be unique within this
	// new session AND must not collide with actual rels in the package.
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	docPart, _ := doc.pkg.Part(doc.docPartName)
	existingRels := make(map[string]bool)
	for _, rel := range docPart.Rels {
		existingRels[rel.ID] = true
	}

	// Re-load to get fresh counters
	doc = roundTrip(t, doc)

	postIDs := make(map[string]bool)
	for i := 0; i < 20; i++ {
		id := doc.NextRelID()
		if postIDs[id] {
			t.Fatalf("duplicate rId in post-RT session: %s at iteration %d", id, i)
		}
		if existingRels[id] {
			t.Fatalf("rId collides with existing package rel: %s", id)
		}
		postIDs[id] = true
	}
}

func TestIntegration_AddMedia_RIdsAreUnique(t *testing.T) {
	doc := newFullPkg().load(t)

	ids := make(map[string]bool)
	for i := 0; i < 30; i++ {
		rId := doc.AddMedia(fmt.Sprintf("img%d.png", i), fakePNG)
		if ids[rId] {
			t.Fatalf("duplicate media rId: %s", rId)
		}
		ids[rId] = true
	}

	// Mix NextRelID calls between AddMedia
	for i := 0; i < 10; i++ {
		id := doc.NextRelID()
		if ids[id] {
			t.Fatalf("NextRelID collision with AddMedia: %s", id)
		}
		ids[id] = true
	}
}

func TestIntegration_RelIDs_UniqueAcrossRoundTrip(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AddMedia("a.png", fakePNG)
	doc.AddMedia("b.png", fakePNG)

	doc = roundTrip(t, doc)

	// After round-trip, add more media and call NextRelID
	doc.AddMedia("c.png", fakePNG)
	doc.AddMedia("d.png", fakePNG)
	id1 := doc.NextRelID()
	id2 := doc.NextRelID()

	if id1 == id2 {
		t.Errorf("duplicate: %s", id1)
	}

	// Build package and check for duplicate rIds on doc part
	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}
	docPart, _ := doc.pkg.Part("/word/document.xml")
	seen := make(map[string]bool)
	for _, rel := range docPart.Rels {
		if seen[rel.ID] {
			t.Errorf("duplicate rel ID in output: %s", rel.ID)
		}
		seen[rel.ID] = true
	}
}

// ═══════════════════════════════════════════════════════════════════
//  ALL HEADER TYPES: even + default + first
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_AllHeaderTypes(t *testing.T) {
	b := newFullPkg()
	b.addHeader("header1.xml", fixtureHeaderXML)          // default
	b.addHeader("header2.xml", fixtureFirstPageHeaderXML) // first
	b.addHeader("header3.xml", fixtureEvenHeaderXML)      // even
	b.addFooter("footer1.xml", fixtureFooterXML)          // default
	b.addFooter("footer2.xml", fixtureEvenFooterXML)      // even

	doc := b.load(t)
	if len(doc.Headers) != 3 {
		t.Errorf("headers: %d, want 3", len(doc.Headers))
	}
	if len(doc.Footers) != 2 {
		t.Errorf("footers: %d, want 2", len(doc.Footers))
	}

	doc2 := roundTrip(t, doc)
	if len(doc2.Headers) != 3 {
		t.Errorf("headers after RT: %d, want 3", len(doc2.Headers))
	}
	if len(doc2.Footers) != 2 {
		t.Errorf("footers after RT: %d, want 2", len(doc2.Footers))
	}
}

// ═══════════════════════════════════════════════════════════════════
//  MULTIPLE EXTERNAL HYPERLINKS + UNKNOWN INTERNAL RELS
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_MixedExternalAndInternalUnknownRels(t *testing.T) {
	b := newFullPkg()

	// External hyperlinks
	for i := 0; i < 5; i++ {
		b.addExternalDocRel(relHyperlink, fmt.Sprintf("https://example.com/link%d", i))
	}
	// External OLE
	b.addExternalDocRel(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
		"https://external-ole.example.com/obj",
	)
	// Internal unknown: glossary, chart, customXml
	b.addUnknownDocRel(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
		"glossary/document.xml",
		[]byte(`<glossary/>`),
	)
	b.addUnknownDocRel(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
		"charts/chart1.xml",
		[]byte(`<chart/>`),
	)
	b.addUnknownDocRel(
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
		"../customXml/item1.xml",
		[]byte(`<customXml/>`),
	)

	doc := b.load(t)

	extCount, intCount := 0, 0
	for _, rel := range doc.UnknownRels {
		if isPackageLevelRel(rel.Type) {
			continue
		}
		if rel.TargetMode == "External" {
			extCount++
		} else {
			intCount++
		}
	}
	if extCount < 6 {
		t.Errorf("external unknowns: %d, want ≥6", extCount)
	}
	if intCount < 3 {
		t.Errorf("internal unknowns: %d, want ≥3", intCount)
	}

	// Round-trip preserves all
	doc2 := roundTrip(t, doc)

	extCount2, intCount2 := 0, 0
	for _, rel := range doc2.UnknownRels {
		if isPackageLevelRel(rel.Type) {
			continue
		}
		if rel.TargetMode == "External" {
			extCount2++
		} else {
			intCount2++
		}
	}
	if extCount2 != extCount {
		t.Errorf("external rels: %d → %d", extCount, extCount2)
	}
	if intCount2 != intCount {
		t.Errorf("internal rels: %d → %d", intCount, intCount2)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  PKG-LEVEL + DOC-LEVEL UNKNOWN RELS TOGETHER
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_PkgAndDocLevelUnknowns(t *testing.T) {
	b := newFullPkg()

	// Package-level unknown
	thumbRelType := "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
	b.pkg.AddPackageRel(thumbRelType, "docProps/thumbnail.jpeg")
	b.pkg.AddPart("/docProps/thumbnail.jpeg", "image/jpeg", fakeJPEG)

	// Document-level unknown
	b.addUnknownDocRel("urn:custom:chart", "charts/chart1.xml", []byte(`<chart/>`))
	b.addExternalDocRel(relHyperlink, "https://example.com")

	doc := b.load(t)

	// Both should be in UnknownRels
	hasPkg, hasDocInternal, hasDocExternal := false, false, false
	for _, rel := range doc.UnknownRels {
		if rel.Type == thumbRelType {
			hasPkg = true
		}
		if rel.Type == "urn:custom:chart" {
			hasDocInternal = true
		}
		if rel.Type == relHyperlink && rel.TargetMode == "External" {
			hasDocExternal = true
		}
	}
	if !hasPkg {
		t.Error("pkg-level unknown missing")
	}
	if !hasDocInternal {
		t.Error("doc-level internal unknown missing")
	}
	if !hasDocExternal {
		t.Error("doc-level external unknown missing")
	}

	// Round-trip
	doc2 := roundTrip(t, doc)
	hasPkg2, hasDocInternal2, hasDocExternal2 := false, false, false
	for _, rel := range doc2.UnknownRels {
		if rel.Type == thumbRelType {
			hasPkg2 = true
		}
		if rel.Type == "urn:custom:chart" {
			hasDocInternal2 = true
		}
		if rel.Type == relHyperlink && rel.TargetMode == "External" {
			hasDocExternal2 = true
		}
	}
	if !hasPkg2 {
		t.Error("pkg-level unknown lost")
	}
	if !hasDocInternal2 {
		t.Error("doc-level internal unknown lost")
	}
	if !hasDocExternal2 {
		t.Error("doc-level external unknown lost")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  EXACT RELATIONSHIP TARGETS AFTER BUILD
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_BuildPackage_RelTargets(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	doc := b.load(t)

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")

	// Expected: for each known type there should be a rel with a
	// sensible target filename.
	expectations := []struct {
		relType    string
		targetPart string // substring of target
	}{
		{relStyles, "styles.xml"},
		{relSettings, "settings.xml"},
		{relWebSettings, "webSettings.xml"},
		{relFontTable, "fontTable.xml"},
		{relTheme, "theme/theme1.xml"},
		{relNumbering, "numbering.xml"},
	}
	for _, exp := range expectations {
		rels := docPart.RelsByType(exp.relType)
		if len(rels) == 0 {
			t.Errorf("no rel for %s", exp.relType)
			continue
		}
		if !strings.Contains(rels[0].Target, exp.targetPart) {
			t.Errorf("%s target=%q, want substring %q", exp.relType, rels[0].Target, exp.targetPart)
		}
	}
}

func TestIntegration_BuildPackage_MediaRelTargets(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AddMedia("photo.jpg", fakeJPEG)
	doc.AddMedia("diagram.svg", []byte("<svg/>"))

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")
	imageRels := docPart.RelsByType(relImage)

	targets := make(map[string]bool)
	for _, rel := range imageRels {
		targets[rel.Target] = true
	}
	if !targets["media/photo.jpg"] {
		t.Error("missing media/photo.jpg rel target")
	}
	if !targets["media/diagram.svg"] {
		t.Error("missing media/diagram.svg rel target")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  REPLACE TYPED PARTS AND VERIFY
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_ReplaceTheme_RoundTrip(t *testing.T) {
	doc := newFullPkg().load(t)
	original := doc.Theme

	replacement := []byte(`<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="NewTheme"><a:themeElements/></a:theme>`)
	doc.Theme = replacement

	doc2 := roundTrip(t, doc)
	if bytes.Equal(doc2.Theme, original) {
		t.Error("theme should have changed")
	}
	// We can't check exact bytes because the part goes through theme.Parse/Serialize,
	// but length should differ from original
	if len(doc2.Theme) == 0 {
		t.Error("theme is empty after replacement")
	}
}

func TestIntegration_ReplaceWebSettings_RoundTrip(t *testing.T) {
	doc := newFullPkg().load(t)
	replacement := []byte(`<?xml version="1.0"?><w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/></w:webSettings>`)
	doc.WebSettings = replacement

	doc2 := roundTrip(t, doc)
	if len(doc2.WebSettings) == 0 {
		t.Error("webSettings empty")
	}
}

func TestIntegration_ReplaceCoreProps_NewObject(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.CoreProps = &coreprops.CoreProperties{
		Title:    "Brand New Title",
		Creator:  "New Author",
		Revision: "999",
	}

	doc2 := roundTrip(t, doc)
	if doc2.CoreProps.Title != "Brand New Title" {
		t.Errorf("Title=%q", doc2.CoreProps.Title)
	}
	if doc2.CoreProps.Creator != "New Author" {
		t.Errorf("Creator=%q", doc2.CoreProps.Creator)
	}
}

func TestIntegration_ReplaceAppProps_NewObject(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AppProps = &coreprops.AppProperties{
		Application: "TestApp",
		Company:     "TestCo",
		Words:       42,
		Pages:       7,
	}

	doc2 := roundTrip(t, doc)
	if doc2.AppProps.Application != "TestApp" {
		t.Errorf("App=%q", doc2.AppProps.Application)
	}
	if doc2.AppProps.Words != 42 {
		t.Errorf("Words=%d", doc2.AppProps.Words)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  ADD THEN REMOVE ACROSS ROUND-TRIPS
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_AddThenRemoveComments(t *testing.T) {
	// Start without comments
	doc := newFullPkg().load(t)
	if doc.Comments != nil {
		t.Skip("base pkg already has comments")
	}

	// The packaging module doesn't create typed Comments out of thin air,
	// but we can simulate by loading a pkg that has them, then removing.
	b := newFullPkg()
	b.addComments()
	doc = b.load(t)
	if doc.Comments == nil {
		t.Fatal("should have comments")
	}

	// Remove and round-trip
	doc.Comments = nil
	doc = roundTrip(t, doc)
	if doc.Comments != nil {
		t.Error("comments should be gone")
	}
}

func TestIntegration_AddThenRemoveMedia(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AddMedia("temp.png", fakePNG)

	doc = roundTrip(t, doc)
	if _, ok := doc.Media["temp.png"]; !ok {
		t.Fatal("media should exist")
	}

	// Remove it
	delete(doc.Media, "temp.png")
	doc = roundTrip(t, doc)
	if _, ok := doc.Media["temp.png"]; ok {
		t.Error("media should be removed")
	}
}

func TestIntegration_AddRemoveAdd_Media(t *testing.T) {
	doc := newFullPkg().load(t)

	doc.AddMedia("file.png", []byte{1, 2, 3})
	doc = roundTrip(t, doc)

	delete(doc.Media, "file.png")
	doc = roundTrip(t, doc)

	doc.AddMedia("file.png", []byte{4, 5, 6})
	doc = roundTrip(t, doc)

	data, ok := doc.Media["file.png"]
	if !ok {
		t.Fatal("file.png missing after re-add")
	}
	if !bytes.Equal(data, []byte{4, 5, 6}) {
		t.Errorf("data = %v, want [4 5 6]", data)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  GRACEFUL DEGRADATION — missing optional mandatory parts
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_MinimalPkg_OnlyDocAndStyles(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	dp := pkg.AddPart("/word/document.xml", ctDocument, []byte(fixtureDocumentXML))
	dp.AddRel(relStyles, "styles.xml")
	pkg.AddPart("/word/styles.xml", ctStyles, []byte(fixtureStylesXML))
	// No settings, no fonts, no theme, no webSettings

	doc, err := load(pkg)
	if err != nil {
		t.Fatal(err)
	}
	if doc.Document == nil || doc.Styles == nil {
		t.Error("present parts should be parsed")
	}
	if doc.Settings != nil || doc.Fonts != nil {
		t.Error("absent parts should be nil")
	}
	if len(doc.Theme) != 0 || len(doc.WebSettings) != 0 {
		t.Error("absent raw parts should be empty")
	}
}

func TestIntegration_MinimalPkg_RoundTrip(t *testing.T) {
	pkg := opc.New()
	pkg.AddPackageRel(relOfficeDocument, "word/document.xml")
	dp := pkg.AddPart("/word/document.xml", ctDocument, []byte(fixtureDocumentXML))
	dp.AddRel(relStyles, "styles.xml")
	pkg.AddPart("/word/styles.xml", ctStyles, []byte(fixtureStylesXML))

	doc, err := load(pkg)
	if err != nil {
		t.Fatal(err)
	}

	doc2 := roundTrip(t, doc)
	if doc2.Document == nil {
		t.Error("Document nil")
	}
	if doc2.Styles == nil {
		t.Error("Styles nil")
	}
	if doc2.Settings != nil {
		t.Error("Settings should remain nil")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  SAME BASE NAME, DIFFERENT EXTENSIONS — media dedup
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_MediaSameBaseDiffExt(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AddMedia("logo.png", fakePNG)
	doc.AddMedia("logo.jpg", fakeJPEG)
	doc.AddMedia("logo.gif", fakeGIF)

	doc2 := roundTrip(t, doc)
	for _, name := range []string{"logo.png", "logo.jpg", "logo.gif"} {
		if _, ok := doc2.Media[name]; !ok {
			t.Errorf("%s missing after round-trip", name)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  IMAGES COLLISION IN MEDIA MAP
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_MediaCollision_DedupThenRoundTrip(t *testing.T) {
	doc := newFullPkg().load(t)
	doc.AddMedia("image.png", []byte{1})
	doc.AddMedia("image.png", []byte{2}) // dedup → image1.png
	doc.AddMedia("image.png", []byte{3}) // dedup → image2.png

	if len(doc.Media) != 3 {
		t.Fatalf("media count = %d, want 3", len(doc.Media))
	}

	doc2 := roundTrip(t, doc)
	if len(doc2.Media) != 3 {
		t.Errorf("media after RT = %d, want 3", len(doc2.Media))
	}
}

// ═══════════════════════════════════════════════════════════════════
//  FULL ZIP CYCLE WITH MUTATIONS
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_ZipCycle_WithMutations(t *testing.T) {
	b := newFullPkg()
	b.addComments()
	b.addFootnotes()
	b.addHeader("header1.xml", fixtureHeaderXML)
	b.addImage("media/img.png", fakePNG)

	doc := b.load(t)

	// Mutations
	doc.AddMedia("extra.svg", []byte("<svg/>"))
	doc.CoreProps.Title = "Mutated"
	doc.Comments = nil

	// Save to buffer
	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		t.Fatal(err)
	}

	// Re-open
	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	if doc2.CoreProps.Title != "Mutated" {
		t.Errorf("Title=%q", doc2.CoreProps.Title)
	}
	if doc2.Comments != nil {
		t.Error("comments should be nil")
	}
	if _, ok := doc2.Media["extra.svg"]; !ok {
		t.Error("extra.svg missing")
	}
	if _, ok := doc2.Media["img.png"]; !ok {
		t.Error("img.png missing")
	}
	if doc2.Footnotes == nil {
		t.Error("footnotes lost")
	}
	if len(doc2.Headers) != 1 {
		t.Errorf("headers = %d", len(doc2.Headers))
	}
}

// ═══════════════════════════════════════════════════════════════════
//  LARGE DOCUMENT — many parts simultaneously
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_LargeDocument(t *testing.T) {
	b := newFullPkg()
	b.addNumbering()
	b.addComments()
	b.addFootnotes()
	b.addEndnotes()

	// 10 headers, 10 footers
	for i := 1; i <= 10; i++ {
		b.addHeader(fmt.Sprintf("header%d.xml", i), fixtureHeaderXML)
		b.addFooter(fmt.Sprintf("footer%d.xml", i), fixtureFooterXML)
	}
	// 20 images
	for i := 0; i < 20; i++ {
		b.addImage(fmt.Sprintf("media/img%02d.png", i), append(fakePNG, byte(i)))
	}
	// 5 unknown internal rels
	for i := 0; i < 5; i++ {
		b.addUnknownDocRel(
			fmt.Sprintf("urn:custom:%d", i),
			fmt.Sprintf("custom%d.xml", i),
			[]byte(fmt.Sprintf("<custom%d/>", i)),
		)
	}
	// 3 external hyperlinks
	for i := 0; i < 3; i++ {
		b.addExternalDocRel(relHyperlink, fmt.Sprintf("https://example.com/%d", i))
	}

	doc := b.load(t)

	if len(doc.Headers) != 10 {
		t.Errorf("headers = %d", len(doc.Headers))
	}
	if len(doc.Footers) != 10 {
		t.Errorf("footers = %d", len(doc.Footers))
	}
	if len(doc.Media) != 20 {
		t.Errorf("media = %d", len(doc.Media))
	}

	// Round-trip the whole thing
	doc2 := roundTrip(t, doc)

	if len(doc2.Headers) != 10 {
		t.Errorf("headers after RT = %d", len(doc2.Headers))
	}
	if len(doc2.Footers) != 10 {
		t.Errorf("footers after RT = %d", len(doc2.Footers))
	}
	if len(doc2.Media) != 20 {
		t.Errorf("media after RT = %d", len(doc2.Media))
	}
	if doc2.Numbering == nil || doc2.Comments == nil || doc2.Footnotes == nil || doc2.Endnotes == nil {
		t.Error("optional typed parts lost")
	}

	// Verify image data integrity
	for i := 0; i < 20; i++ {
		name := fmt.Sprintf("img%02d.png", i)
		data, ok := doc2.Media[name]
		if !ok {
			t.Errorf("%s missing", name)
			continue
		}
		if data[len(data)-1] != byte(i) {
			t.Errorf("%s data[last] = %d, want %d", name, data[len(data)-1], i)
		}
	}
}

// ═══════════════════════════════════════════════════════════════════
//  REALISTIC: FULL DOCUMENT WITH ALL EXTENSIONS (from newRealisticPkg)
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_RealisticPkg_ZipCycle(t *testing.T) {
	doc := newRealisticPkg().load(t)

	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		t.Fatal(err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	parts := []struct {
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
		{"Headers", len(doc2.Headers) == 2},
		{"Footers", len(doc2.Footers) == 1},
		{"Media", len(doc2.Media) == 1},
	}
	for _, p := range parts {
		if !p.ok {
			t.Errorf("zip cycle lost: %s", p.name)
		}
	}
}

func TestIntegration_RealisticPkg_MutateAndZipCycle(t *testing.T) {
	doc := newRealisticPkg().load(t)

	// Mutate everything
	doc.AddMedia("extra.png", fakePNG)
	doc.CoreProps.Title = "Modified"
	doc.Comments = nil
	doc.Endnotes = nil

	var buf bytes.Buffer
	if err := doc.SaveWriter(&buf); err != nil {
		t.Fatal(err)
	}

	doc2, err := OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	if doc2.CoreProps.Title != "Modified" {
		t.Errorf("Title=%q", doc2.CoreProps.Title)
	}
	if doc2.Comments != nil {
		t.Error("comments should be nil")
	}
	if doc2.Endnotes != nil {
		t.Error("endnotes should be nil")
	}
	if _, ok := doc2.Media["extra.png"]; !ok {
		t.Error("extra.png missing")
	}
	if _, ok := doc2.Media["logo.png"]; !ok {
		t.Error("original logo.png missing")
	}
	if doc2.Numbering == nil || doc2.Footnotes == nil {
		t.Error("remaining optional parts lost")
	}
	if len(doc2.Headers) != 2 {
		t.Errorf("headers=%d", len(doc2.Headers))
	}
}

// ═══════════════════════════════════════════════════════════════════
//  BUILDPACKAGE CONSISTENCY — verify no stale rels after mutations
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_BuildPackage_NoStaleRels(t *testing.T) {
	b := newFullPkg()
	b.addComments()
	b.addFootnotes()
	b.addImage("media/old.png", fakePNG)
	doc := b.load(t)

	// Remove parts
	doc.Comments = nil
	doc.Footnotes = nil
	delete(doc.Media, "old.png")

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	docPart, _ := doc.pkg.Part("/word/document.xml")

	// No comments/footnotes/image rels should exist
	if rels := docPart.RelsByType(relComments); len(rels) != 0 {
		t.Errorf("stale comments rels: %d", len(rels))
	}
	if rels := docPart.RelsByType(relFootnotes); len(rels) != 0 {
		t.Errorf("stale footnotes rels: %d", len(rels))
	}
	if rels := docPart.RelsByType(relImage); len(rels) != 0 {
		t.Errorf("stale image rels: %d", len(rels))
	}

	// No parts should exist
	assertPartAbsent(t, doc.pkg, "/word/comments.xml")
	assertPartAbsent(t, doc.pkg, "/word/footnotes.xml")
	assertPartAbsent(t, doc.pkg, "/word/media/old.png")
}

func TestIntegration_BuildPackage_AddNewParts(t *testing.T) {
	doc := newFullPkg().load(t) // no numbering/comments initially

	// Simulate adding: in reality, packaging just checks nil vs non-nil.
	// We load a package that has them.
	b := newFullPkg()
	b.addNumbering()
	b.addComments()
	doc2 := b.load(t)

	if err := doc2.buildPackage(); err != nil {
		t.Fatal(err)
	}

	assertPartExists(t, doc2.pkg, "/word/numbering.xml")
	assertPartExists(t, doc2.pkg, "/word/comments.xml")

	docPart, _ := doc2.pkg.Part("/word/document.xml")
	if rels := docPart.RelsByType(relNumbering); len(rels) != 1 {
		t.Error("numbering rel missing")
	}
	if rels := docPart.RelsByType(relComments); len(rels) != 1 {
		t.Error("comments rel missing")
	}

	_ = doc // silence unused
}

// ═══════════════════════════════════════════════════════════════════
//  BOOKMARK ID INTEGRITY THROUGH ROUND-TRIPS
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_BookmarkIDs_NoCollisionAfterRoundTrip(t *testing.T) {
	doc := newFullPkg().load(t)

	// Within a single session — IDs must be unique and monotonic
	ids := make(map[int]bool)
	prev := -1
	for i := 0; i < 20; i++ {
		id := doc.NextBookmarkID()
		if ids[id] {
			t.Fatalf("duplicate bookmark ID: %d", id)
		}
		if id <= prev {
			t.Fatalf("non-monotonic: %d after %d", id, prev)
		}
		ids[id] = true
		prev = id
	}

	// After round-trip, seed resets (seedBookmarkIDCounter always
	// starts at a fixed value). Verify uniqueness in the new session.
	doc = roundTrip(t, doc)

	postIDs := make(map[int]bool)
	prev = -1
	for i := 0; i < 20; i++ {
		id := doc.NextBookmarkID()
		if postIDs[id] {
			t.Fatalf("duplicate bookmark ID in post-RT session: %d", id)
		}
		if id <= prev {
			t.Fatalf("non-monotonic post-RT: %d after %d", id, prev)
		}
		postIDs[id] = true
		prev = id
	}
}

// ═══════════════════════════════════════════════════════════════════
//  EDGE: SAVE → RE-OPEN → MODIFY → SAVE → RE-OPEN (chained zip)
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_ChainedZipCycles(t *testing.T) {
	doc := newFullPkg().load(t)

	for i := 0; i < 5; i++ {
		doc.AddMedia(fmt.Sprintf("round%d.png", i), fakePNG)
		doc.CoreProps.Title = fmt.Sprintf("Round %d", i)

		var buf bytes.Buffer
		if err := doc.SaveWriter(&buf); err != nil {
			t.Fatalf("save round %d: %v", i, err)
		}

		var err error
		doc, err = OpenReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
		if err != nil {
			t.Fatalf("open round %d: %v", i, err)
		}
	}

	// After 5 chained zip cycles, we should have 5 media files
	if len(doc.Media) != 5 {
		t.Errorf("media = %d, want 5", len(doc.Media))
	}
	if doc.CoreProps.Title != "Round 4" {
		t.Errorf("title = %q", doc.CoreProps.Title)
	}
}

// ═══════════════════════════════════════════════════════════════════
//  EDGE: ALL UNKNOWN REL DATA BYTES PRESERVED EXACTLY
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_UnknownPartBytes_ExactPreservation(t *testing.T) {
	// Binary-ish data that shouldn't be modified
	binaryData := make([]byte, 256)
	for i := range binaryData {
		binaryData[i] = byte(i)
	}

	b := newFullPkg()
	b.addUnknownDocRel("urn:custom:binary", "binary.dat", binaryData)

	doc := b.load(t)

	stored := doc.UnknownParts["/word/binary.dat"]
	if !bytes.Equal(stored, binaryData) {
		t.Error("binary data modified during load")
	}

	doc2 := roundTrip(t, doc)
	stored2 := doc2.UnknownParts["/word/binary.dat"]
	if !bytes.Equal(stored2, binaryData) {
		t.Error("binary data modified during round-trip")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  PACKAGE-LEVEL RELS: officeDocument always present
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_BuildPackage_OfficeDocumentRelAlwaysPresent(t *testing.T) {
	// Even with nil CoreProps/AppProps
	doc := newFullPkg().load(t)
	doc.CoreProps = nil
	doc.AppProps = nil

	if err := doc.buildPackage(); err != nil {
		t.Fatal(err)
	}

	rels := doc.pkg.PackageRelsByType(relOfficeDocument)
	if len(rels) != 1 {
		t.Errorf("officeDocument rels = %d", len(rels))
	}

	// Core/App rels should NOT be present
	if rels := doc.pkg.PackageRelsByType(relCoreProperties); len(rels) != 0 {
		t.Error("core rels should be absent")
	}
	if rels := doc.pkg.PackageRelsByType(relExtProperties); len(rels) != 0 {
		t.Error("ext rels should be absent")
	}
}

// ═══════════════════════════════════════════════════════════════════
//  HEADERS/FOOTERS WITH SEPARATE CONTENT VERIFY INDIVIDUALITY
// ═══════════════════════════════════════════════════════════════════

func TestIntegration_HeadersRetainIndividualContent(t *testing.T) {
	b := newFullPkg()
	r1 := b.addHeader("header1.xml", fixtureHeaderXML)
	r2 := b.addHeader("header2.xml", fixtureFirstPageHeaderXML)

	doc := b.load(t)

	h1 := doc.Headers[r1]
	h2 := doc.Headers[r2]
	if h1 == nil || h2 == nil {
		t.Fatal("nil header(s)")
	}
	// They should be distinct objects (different content)
	if h1 == h2 {
		t.Error("headers should be distinct objects")
	}
}
