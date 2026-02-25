package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Section provides access to section and page setup settings.
//
// Mirrors Python Section.
type Section struct {
	sectPr  *oxml.CT_SectPr
	docPart *parts.DocumentPart
}

// NewSection creates a new Section proxy.
func NewSection(sectPr *oxml.CT_SectPr, docPart *parts.DocumentPart) *Section {
	return &Section{sectPr: sectPr, docPart: docPart}
}

// BottomMargin returns the bottom margin in twips, or nil if not set.
func (s *Section) BottomMargin() (*int, error) { return s.sectPr.BottomMargin() }

// SetBottomMargin sets the bottom margin in twips.
func (s *Section) SetBottomMargin(v *int) error { return s.sectPr.SetBottomMargin(v) }

// TopMargin returns the top margin in twips, or nil if not set.
func (s *Section) TopMargin() (*int, error) { return s.sectPr.TopMargin() }

// SetTopMargin sets the top margin in twips.
func (s *Section) SetTopMargin(v *int) error { return s.sectPr.SetTopMargin(v) }

// LeftMargin returns the left margin in twips, or nil if not set.
func (s *Section) LeftMargin() (*int, error) { return s.sectPr.LeftMargin() }

// SetLeftMargin sets the left margin in twips.
func (s *Section) SetLeftMargin(v *int) error { return s.sectPr.SetLeftMargin(v) }

// RightMargin returns the right margin in twips, or nil if not set.
func (s *Section) RightMargin() (*int, error) { return s.sectPr.RightMargin() }

// SetRightMargin sets the right margin in twips.
func (s *Section) SetRightMargin(v *int) error { return s.sectPr.SetRightMargin(v) }

// PageWidth returns the page width in twips, or nil if not set.
func (s *Section) PageWidth() (*int, error) { return s.sectPr.PageWidth() }

// SetPageWidth sets the page width in twips.
func (s *Section) SetPageWidth(v *int) error { return s.sectPr.SetPageWidth(v) }

// PageHeight returns the page height in twips, or nil if not set.
func (s *Section) PageHeight() (*int, error) { return s.sectPr.PageHeight() }

// SetPageHeight sets the page height in twips.
func (s *Section) SetPageHeight(v *int) error { return s.sectPr.SetPageHeight(v) }

// Orientation returns the page orientation.
func (s *Section) Orientation() (enum.WdOrientation, error) { return s.sectPr.Orientation() }

// SetOrientation sets the page orientation.
func (s *Section) SetOrientation(v enum.WdOrientation) error { return s.sectPr.SetOrientation(v) }

// StartType returns the section start type.
func (s *Section) StartType() (enum.WdSectionStart, error) { return s.sectPr.StartType() }

// SetStartType sets the section start type.
func (s *Section) SetStartType(v enum.WdSectionStart) error { return s.sectPr.SetStartType(v) }

// Gutter returns the gutter in twips, or nil if not set.
func (s *Section) Gutter() (*int, error) { return s.sectPr.GutterMargin() }

// SetGutter sets the gutter in twips.
func (s *Section) SetGutter(v *int) error { return s.sectPr.SetGutterMargin(v) }

// HeaderDistance returns the header distance in twips, or nil if not set.
func (s *Section) HeaderDistance() (*int, error) { return s.sectPr.HeaderMargin() }

// SetHeaderDistance sets the header distance.
func (s *Section) SetHeaderDistance(v *int) error { return s.sectPr.SetHeaderMargin(v) }

// FooterDistance returns the footer distance in twips, or nil if not set.
func (s *Section) FooterDistance() (*int, error) { return s.sectPr.FooterMargin() }

// SetFooterDistance sets the footer distance.
func (s *Section) SetFooterDistance(v *int) error { return s.sectPr.SetFooterMargin(v) }

// DifferentFirstPageHeaderFooter returns true if this section displays a distinct
// first-page header and footer.
func (s *Section) DifferentFirstPageHeaderFooter() bool { return s.sectPr.TitlePgVal() }

// SetDifferentFirstPageHeaderFooter sets the first-page header/footer flag.
func (s *Section) SetDifferentFirstPageHeaderFooter(v bool) error {
	return s.sectPr.SetTitlePgVal(v)
}

// Header returns the default (primary) page header.
func (s *Section) Header() *Header {
	return NewHeader(s.sectPr, s.docPart, enum.WdHeaderFooterIndexPrimary)
}

// Footer returns the default (primary) page footer.
func (s *Section) Footer() *Footer {
	return NewFooter(s.sectPr, s.docPart, enum.WdHeaderFooterIndexPrimary)
}

// EvenPageHeader returns the even-page header.
func (s *Section) EvenPageHeader() *Header {
	return NewHeader(s.sectPr, s.docPart, enum.WdHeaderFooterIndexEvenPage)
}

// EvenPageFooter returns the even-page footer.
func (s *Section) EvenPageFooter() *Footer {
	return NewFooter(s.sectPr, s.docPart, enum.WdHeaderFooterIndexEvenPage)
}

// FirstPageHeader returns the first-page header.
func (s *Section) FirstPageHeader() *Header {
	return NewHeader(s.sectPr, s.docPart, enum.WdHeaderFooterIndexFirstPage)
}

// FirstPageFooter returns the first-page footer.
func (s *Section) FirstPageFooter() *Footer {
	return NewFooter(s.sectPr, s.docPart, enum.WdHeaderFooterIndexFirstPage)
}

// IterInnerContent returns paragraphs and tables in this section body.
//
// Mirrors Python Section.iter_inner_content → CT_SectPr.iter_inner_content →
// _SectBlockElementIterator. Only returns block-items belonging to THIS section,
// not the entire document body.
func (s *Section) IterInnerContent() []*InnerContentItem {
	body := s.sectPr.RawElement().Parent()
	if body == nil {
		return nil
	}

	// Determine section boundaries. Sections are delimited by sectPr elements:
	// - Paragraph-based: w:p/w:pPr/w:sectPr marks end of a section (the p itself belongs to that section)
	// - Body-based: w:body/w:sectPr marks the last section
	//
	// We walk body children once, collecting (start, end) ranges for each section.

	type sectionRange struct {
		startIdx int // inclusive index into body children
		endIdx   int // exclusive index into body children
		sectPrEl *etree.Element
	}

	children := body.ChildElements()
	var ranges []sectionRange
	rangeStart := 0

	for i, child := range children {
		if child.Space == "w" && child.Tag == "p" {
			// Check if this paragraph contains w:pPr/w:sectPr
			if pSectPr := findParagraphSectPr(child); pSectPr != nil {
				// This paragraph (inclusive) ends a section
				ranges = append(ranges, sectionRange{
					startIdx: rangeStart,
					endIdx:   i + 1, // include this p
					sectPrEl: pSectPr,
				})
				rangeStart = i + 1
			}
		} else if child.Space == "w" && child.Tag == "sectPr" {
			// Body-level sectPr: last section
			ranges = append(ranges, sectionRange{
				startIdx: rangeStart,
				endIdx:   i, // exclude the sectPr itself
				sectPrEl: child,
			})
		}
	}

	// Find which range matches our sectPr
	for _, sr := range ranges {
		if sr.sectPrEl == s.sectPr.RawElement() {
			return collectBlockItems(children[sr.startIdx:sr.endIdx], s.docPart)
		}
	}

	return nil
}

// findParagraphSectPr returns the w:sectPr element inside w:p/w:pPr, or nil.
func findParagraphSectPr(p *etree.Element) *etree.Element {
	for _, child := range p.ChildElements() {
		if child.Space == "w" && child.Tag == "pPr" {
			for _, gc := range child.ChildElements() {
				if gc.Space == "w" && gc.Tag == "sectPr" {
					return gc
				}
			}
		}
	}
	return nil
}

// collectBlockItems filters elements for w:p and w:tbl, wrapping them as
// InnerContentItems with the appropriate proxy objects.
func collectBlockItems(elems []*etree.Element, docPart *parts.DocumentPart) []*InnerContentItem {
	var result []*InnerContentItem
	var sp *parts.StoryPart
	if docPart != nil {
		sp = &docPart.StoryPart
	}
	for _, child := range elems {
		switch {
		case child.Space == "w" && child.Tag == "p":
			p := &oxml.CT_P{Element: oxml.WrapElement(child)}
			result = append(result, &InnerContentItem{paragraph: NewParagraph(p, sp)})
		case child.Space == "w" && child.Tag == "tbl":
			tbl := &oxml.CT_Tbl{Element: oxml.WrapElement(child)}
			result = append(result, &InnerContentItem{table: NewTable(tbl, sp)})
		}
	}
	return result
}

// --------------------------------------------------------------------------
// Sections
// --------------------------------------------------------------------------

// Sections is a sequence of Section objects corresponding to the sections in a document.
//
// Mirrors Python Sections(Sequence).
type Sections struct {
	docElm  *oxml.CT_Document
	docPart *parts.DocumentPart
}

// NewSections creates a new Sections proxy.
func NewSections(docElm *oxml.CT_Document, docPart *parts.DocumentPart) *Sections {
	return &Sections{docElm: docElm, docPart: docPart}
}

// Len returns the number of sections.
func (ss *Sections) Len() int {
	return len(ss.docElm.SectPrList())
}

// Get returns the section at the given index.
func (ss *Sections) Get(idx int) (*Section, error) {
	lst := ss.docElm.SectPrList()
	if idx < 0 || idx >= len(lst) {
		return nil, fmt.Errorf("docx: section index [%d] out of range", idx)
	}
	return NewSection(lst[idx], ss.docPart), nil
}

// Iter returns all sections in document order.
func (ss *Sections) Iter() []*Section {
	lst := ss.docElm.SectPrList()
	result := make([]*Section, len(lst))
	for i, sp := range lst {
		result[i] = NewSection(sp, ss.docPart)
	}
	return result
}

// --------------------------------------------------------------------------
// Header / Footer — _BaseHeaderFooter pattern
// --------------------------------------------------------------------------

// Header is a proxy for a page header.
//
// Mirrors Python _Header(_BaseHeaderFooter(BlockItemContainer)).
// Provides BlockItemContainer methods (AddParagraph, AddTable, Paragraphs,
// Tables, IterInnerContent) by delegating to the underlying header part.
type Header struct {
	sectPr  *oxml.CT_SectPr
	docPart *parts.DocumentPart
	index   enum.WdHeaderFooterIndex
}

// NewHeader creates a new Header proxy.
func NewHeader(sectPr *oxml.CT_SectPr, docPart *parts.DocumentPart, index enum.WdHeaderFooterIndex) *Header {
	return &Header{sectPr: sectPr, docPart: docPart, index: index}
}

// IsLinkedToPrevious returns true if this header uses the definition from
// the prior section.
//
// Mirrors Python _BaseHeaderFooter.is_linked_to_previous.
func (h *Header) IsLinkedToPrevious() bool {
	return !h.hasDefinition()
}

// SetIsLinkedToPrevious sets the linked-to-previous state.
func (h *Header) SetIsLinkedToPrevious(v bool) error {
	if v == h.IsLinkedToPrevious() {
		return nil
	}
	if v {
		h.dropDefinition()
		return nil
	}
	_, err := h.addDefinition()
	return err
}

// AddParagraph appends a new paragraph to this header.
//
// Mirrors Python BlockItemContainer.add_paragraph (inherited by _BaseHeaderFooter).
func (h *Header) AddParagraph(text string, style interface{}) (*Paragraph, error) {
	bic, err := h.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: header add paragraph: %w", err)
	}
	return bic.AddParagraph(text, style)
}

// AddTable appends a new table to this header.
//
// Mirrors Python BlockItemContainer.add_table (inherited by _BaseHeaderFooter).
func (h *Header) AddTable(rows, cols int, widthTwips int) (*Table, error) {
	bic, err := h.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: header add table: %w", err)
	}
	return bic.AddTable(rows, cols, widthTwips)
}

// Paragraphs returns the paragraphs in this header.
//
// Mirrors Python BlockItemContainer.paragraphs (inherited by _BaseHeaderFooter).
func (h *Header) Paragraphs() []*Paragraph {
	bic, err := h.blockItemContainer()
	if err != nil {
		return nil
	}
	return bic.Paragraphs()
}

// Tables returns the tables in this header.
//
// Mirrors Python BlockItemContainer.tables (inherited by _BaseHeaderFooter).
func (h *Header) Tables() []*Table {
	bic, err := h.blockItemContainer()
	if err != nil {
		return nil
	}
	return bic.Tables()
}

// IterInnerContent returns paragraphs and tables in this header in document order.
//
// Mirrors Python BlockItemContainer.iter_inner_content (inherited by _BaseHeaderFooter).
func (h *Header) IterInnerContent() []*InnerContentItem {
	bic, err := h.blockItemContainer()
	if err != nil {
		return nil
	}
	return bic.IterInnerContent()
}

// Part returns the HeaderPart as a StoryPart. This overrides the part
// accessor to provide the correct StoryPart for style resolution and
// image insertion in header content.
//
// Mirrors Python _BaseHeaderFooter.part property.
func (h *Header) Part() *parts.StoryPart {
	hp := h.getOrAddDefinition()
	if hp == nil {
		return nil
	}
	return &hp.StoryPart
}

// blockItemContainer creates a BlockItemContainer backed by the header part's
// element and StoryPart. Created fresh each call to match Python's property
// behavior (no stale cache if definition changes).
func (h *Header) blockItemContainer() (*BlockItemContainer, error) {
	hp := h.getOrAddDefinition()
	if hp == nil {
		return nil, fmt.Errorf("docx: failed to resolve header definition")
	}
	el := hp.Element()
	if el == nil {
		return nil, fmt.Errorf("docx: header part has nil element")
	}
	bic := NewBlockItemContainer(el, &hp.StoryPart)
	return &bic, nil
}

func (h *Header) hasDefinition() bool {
	ref, _ := h.sectPr.GetHeaderRef(h.index)
	return ref != nil
}

func (h *Header) addDefinition() (*parts.HeaderPart, error) {
	hp, rId, err := h.docPart.AddHeaderPart()
	if err != nil {
		return nil, err
	}
	if _, err := h.sectPr.AddHeaderRef(h.index, rId); err != nil {
		return nil, err
	}
	return hp, nil
}

func (h *Header) dropDefinition() {
	rId := h.sectPr.RemoveHeaderRef(h.index)
	if rId != "" {
		h.docPart.DropHeaderPart(rId)
	}
}

// getOrAddDefinition mirrors Python _BaseHeaderFooter._get_or_add_definition.
// Recursive: if linked, walk to prior section's header; if first section and
// linked, add new definition.
func (h *Header) getOrAddDefinition() *parts.HeaderPart {
	if h.hasDefinition() {
		return h.definition()
	}
	prior := h.priorHeader()
	if prior != nil {
		return prior.getOrAddDefinition()
	}
	hp, _ := h.addDefinition()
	return hp
}

func (h *Header) definition() *parts.HeaderPart {
	ref, _ := h.sectPr.GetHeaderRef(h.index)
	if ref == nil {
		return nil
	}
	rId, _ := ref.RId()
	hp, err := h.docPart.HeaderPartByRID(rId)
	if err != nil {
		return nil
	}
	return hp
}

func (h *Header) priorHeader() *Header {
	prev := h.sectPr.PrecedingSectPr()
	if prev == nil {
		return nil
	}
	return NewHeader(prev, h.docPart, h.index)
}

// --------------------------------------------------------------------------
// Footer
// --------------------------------------------------------------------------

// Footer is a proxy for a page footer.
//
// Mirrors Python _Footer(_BaseHeaderFooter(BlockItemContainer)).
// Provides BlockItemContainer methods (AddParagraph, AddTable, Paragraphs,
// Tables, IterInnerContent) by delegating to the underlying footer part.
type Footer struct {
	sectPr  *oxml.CT_SectPr
	docPart *parts.DocumentPart
	index   enum.WdHeaderFooterIndex
}

// NewFooter creates a new Footer proxy.
func NewFooter(sectPr *oxml.CT_SectPr, docPart *parts.DocumentPart, index enum.WdHeaderFooterIndex) *Footer {
	return &Footer{sectPr: sectPr, docPart: docPart, index: index}
}

// IsLinkedToPrevious returns true if this footer uses the definition from
// the prior section.
func (f *Footer) IsLinkedToPrevious() bool {
	return !f.hasDefinition()
}

// SetIsLinkedToPrevious sets the linked-to-previous state.
func (f *Footer) SetIsLinkedToPrevious(v bool) error {
	if v == f.IsLinkedToPrevious() {
		return nil
	}
	if v {
		f.dropDefinition()
		return nil
	}
	_, err := f.addDefinition()
	return err
}

// AddParagraph appends a new paragraph to this footer.
//
// Mirrors Python BlockItemContainer.add_paragraph (inherited by _BaseHeaderFooter).
func (f *Footer) AddParagraph(text string, style interface{}) (*Paragraph, error) {
	bic, err := f.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: footer add paragraph: %w", err)
	}
	return bic.AddParagraph(text, style)
}

// AddTable appends a new table to this footer.
//
// Mirrors Python BlockItemContainer.add_table (inherited by _BaseHeaderFooter).
func (f *Footer) AddTable(rows, cols int, widthTwips int) (*Table, error) {
	bic, err := f.blockItemContainer()
	if err != nil {
		return nil, fmt.Errorf("docx: footer add table: %w", err)
	}
	return bic.AddTable(rows, cols, widthTwips)
}

// Paragraphs returns the paragraphs in this footer.
//
// Mirrors Python BlockItemContainer.paragraphs (inherited by _BaseHeaderFooter).
func (f *Footer) Paragraphs() []*Paragraph {
	bic, err := f.blockItemContainer()
	if err != nil {
		return nil
	}
	return bic.Paragraphs()
}

// Tables returns the tables in this footer.
//
// Mirrors Python BlockItemContainer.tables (inherited by _BaseHeaderFooter).
func (f *Footer) Tables() []*Table {
	bic, err := f.blockItemContainer()
	if err != nil {
		return nil
	}
	return bic.Tables()
}

// IterInnerContent returns paragraphs and tables in this footer in document order.
//
// Mirrors Python BlockItemContainer.iter_inner_content (inherited by _BaseHeaderFooter).
func (f *Footer) IterInnerContent() []*InnerContentItem {
	bic, err := f.blockItemContainer()
	if err != nil {
		return nil
	}
	return bic.IterInnerContent()
}

// Part returns the FooterPart as a StoryPart. This overrides the part
// accessor to provide the correct StoryPart for style resolution and
// image insertion in footer content.
//
// Mirrors Python _BaseHeaderFooter.part property.
func (f *Footer) Part() *parts.StoryPart {
	fp := f.getOrAddDefinition()
	if fp == nil {
		return nil
	}
	return &fp.StoryPart
}

// blockItemContainer creates a BlockItemContainer backed by the footer part's
// element and StoryPart. Created fresh each call to match Python's property
// behavior.
func (f *Footer) blockItemContainer() (*BlockItemContainer, error) {
	fp := f.getOrAddDefinition()
	if fp == nil {
		return nil, fmt.Errorf("docx: failed to resolve footer definition")
	}
	el := fp.Element()
	if el == nil {
		return nil, fmt.Errorf("docx: footer part has nil element")
	}
	bic := NewBlockItemContainer(el, &fp.StoryPart)
	return &bic, nil
}

func (f *Footer) hasDefinition() bool {
	ref, _ := f.sectPr.GetFooterRef(f.index)
	return ref != nil
}

func (f *Footer) addDefinition() (*parts.FooterPart, error) {
	fp, rId, err := f.docPart.AddFooterPart()
	if err != nil {
		return nil, err
	}
	if _, err := f.sectPr.AddFooterRef(f.index, rId); err != nil {
		return nil, err
	}
	return fp, nil
}

func (f *Footer) dropDefinition() {
	rId := f.sectPr.RemoveFooterRef(f.index)
	if rId != "" {
		f.docPart.DropRel(rId)
	}
}

func (f *Footer) getOrAddDefinition() *parts.FooterPart {
	if f.hasDefinition() {
		return f.definition()
	}
	prior := f.priorFooter()
	if prior != nil {
		return prior.getOrAddDefinition()
	}
	fp, _ := f.addDefinition()
	return fp
}

func (f *Footer) definition() *parts.FooterPart {
	ref, _ := f.sectPr.GetFooterRef(f.index)
	if ref == nil {
		return nil
	}
	rId, _ := ref.RId()
	fp, err := f.docPart.FooterPartByRID(rId)
	if err != nil {
		return nil
	}
	return fp
}

func (f *Footer) priorFooter() *Footer {
	prev := f.sectPr.PrecedingSectPr()
	if prev == nil {
		return nil
	}
	return NewFooter(prev, f.docPart, f.index)
}
