package docx

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// -----------------------------------------------------------------------
// section_test.go — Section / Sections (Batch 1)
// Mirrors Python: tests/test_section.py
// -----------------------------------------------------------------------

// Mirrors Python: Sections.it_knows_how_many_sections
func TestSections_Len(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	if sections.Len() < 1 {
		t.Errorf("Sections.Len() = %d, want >= 1", sections.Len())
	}
}

// Mirrors Python: Sections.it_can_iterate
func TestSections_Iter(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	iter := sections.Iter()
	if len(iter) != sections.Len() {
		t.Errorf("len(Iter()) = %d, want %d", len(iter), sections.Len())
	}
	for i, s := range iter {
		if s == nil {
			t.Errorf("Iter()[%d] is nil", i)
		}
	}
}

// Mirrors Python: Sections.it_can_access_by_index
func TestSections_Get(t *testing.T) {
	doc := mustNewDoc(t)
	sections := doc.Sections()
	s, err := sections.Get(0)
	if err != nil {
		t.Fatal(err)
	}
	if s == nil {
		t.Error("Get(0) returned nil")
	}

	// Out of range
	_, err = sections.Get(999)
	if err == nil {
		t.Error("expected error for Get(999)")
	}
}

// Mirrors Python: it_knows_its_start_type / it_can_change
func TestSection_StartType(t *testing.T) {
	sectPr := makeSectPr(t, ``)
	sec := newSection(sectPr, nil)

	// Set to new page
	if err := sec.SetStartType(enum.WdSectionStartNewPage); err != nil {
		t.Fatal(err)
	}
	got, err := sec.StartType()
	if err != nil {
		t.Fatal(err)
	}
	if got != enum.WdSectionStartNewPage {
		t.Errorf("StartType() = %v, want %v", got, enum.WdSectionStartNewPage)
	}

	// Change to continuous
	if err := sec.SetStartType(enum.WdSectionStartContinuous); err != nil {
		t.Fatal(err)
	}
	got2, err := sec.StartType()
	if err != nil {
		t.Fatal(err)
	}
	if got2 != enum.WdSectionStartContinuous {
		t.Errorf("StartType() = %v, want %v", got2, enum.WdSectionStartContinuous)
	}
}

// Mirrors Python: it_knows_its_page_orientation / it_can_change
func TestSection_Orientation(t *testing.T) {
	sectPr := makeSectPr(t, `<w:pgSz w:w="12240" w:h="15840"/>`)
	sec := newSection(sectPr, nil)

	// Set to landscape
	if err := sec.SetOrientation(enum.WdOrientationLandscape); err != nil {
		t.Fatal(err)
	}
	got, err := sec.Orientation()
	if err != nil {
		t.Fatal(err)
	}
	if got != enum.WdOrientationLandscape {
		t.Errorf("Orientation() = %v, want LANDSCAPE", got)
	}

	// Set to portrait
	if err := sec.SetOrientation(enum.WdOrientationPortrait); err != nil {
		t.Fatal(err)
	}
	got2, err := sec.Orientation()
	if err != nil {
		t.Fatal(err)
	}
	if got2 != enum.WdOrientationPortrait {
		t.Errorf("Orientation() = %v, want PORTRAIT", got2)
	}
}

// Mirrors Python: it_knows_its_page_dimensions (complete set/get)
func TestSection_PageDimensions_SetGet(t *testing.T) {
	sectPr := makeSectPr(t, `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>`)
	sec := newSection(sectPr, nil)

	// Page Width
	w, err := sec.PageWidth()
	if err != nil {
		t.Fatal(err)
	}
	if w == nil || *w != 12240 {
		t.Errorf("PageWidth = %v, want 12240", w)
	}
	newW := 15840
	if err := sec.SetPageWidth(&newW); err != nil {
		t.Fatal(err)
	}
	w2, _ := sec.PageWidth()
	if w2 == nil || *w2 != 15840 {
		t.Errorf("PageWidth after set = %v, want 15840", w2)
	}

	// Page Height
	h, err := sec.PageHeight()
	if err != nil {
		t.Fatal(err)
	}
	if h == nil || *h != 15840 {
		t.Errorf("PageHeight = %v, want 15840", h)
	}
}

// Mirrors Python: margins set/get
func TestSection_Margins_SetGet(t *testing.T) {
	sectPr := makeSectPr(t, `<w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>`)
	sec := newSection(sectPr, nil)

	// Read
	top, err := sec.TopMargin()
	if err != nil {
		t.Fatal(err)
	}
	if top == nil || *top != 1440 {
		t.Errorf("TopMargin = %v, want 1440", top)
	}

	// Set
	v := 2000
	if err := sec.SetTopMargin(&v); err != nil {
		t.Fatal(err)
	}
	top2, _ := sec.TopMargin()
	if top2 == nil || *top2 != 2000 {
		t.Errorf("TopMargin after set = %v, want 2000", top2)
	}

	// Bottom
	bot, _ := sec.BottomMargin()
	if bot == nil || *bot != 1440 {
		t.Errorf("BottomMargin = %v, want 1440", bot)
	}

	// Left
	left, _ := sec.LeftMargin()
	if left == nil || *left != 1800 {
		t.Errorf("LeftMargin = %v, want 1800", left)
	}

	// Right
	right, _ := sec.RightMargin()
	if right == nil || *right != 1800 {
		t.Errorf("RightMargin = %v, want 1800", right)
	}
}

// Mirrors Python: it_knows_when_it_displays_a_distinct_first_page_header
func TestSection_DifferentFirstPageHeaderFooter(t *testing.T) {
	// Without titlePg
	sectPr1 := makeSectPr(t, ``)
	sec1 := newSection(sectPr1, nil)
	if sec1.DifferentFirstPageHeaderFooter() {
		t.Error("expected false when titlePg absent")
	}

	// With titlePg
	sectPr2 := makeSectPr(t, `<w:titlePg/>`)
	sec2 := newSection(sectPr2, nil)
	if !sec2.DifferentFirstPageHeaderFooter() {
		t.Error("expected true when titlePg present")
	}

	// Set
	if err := sec1.SetDifferentFirstPageHeaderFooter(true); err != nil {
		t.Fatal(err)
	}
	if !sec1.DifferentFirstPageHeaderFooter() {
		t.Error("expected true after SetDifferentFirstPageHeaderFooter(true)")
	}
}

// Helper: check Sections from a document with body-level sectPr
func makeSectionsDoc(t *testing.T, bodySectPrXml string) *oxml.CT_Document {
	t.Helper()
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body><w:p/><w:sectPr>` + bodySectPrXml + `</w:sectPr></w:body></w:document>`
	el := mustParseXml(t, xml)
	return &oxml.CT_Document{Element: *el}
}
