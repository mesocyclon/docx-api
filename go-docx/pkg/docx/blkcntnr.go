package docx

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// InnerContentItem represents either a *Paragraph or a *Table found in a
// block-item container. Callers inspect the type via Paragraph() / Table().
type InnerContentItem struct {
	paragraph *Paragraph
	table     *Table
}

// IsParagraph returns true if this item is a paragraph.
func (it *InnerContentItem) IsParagraph() bool { return it.paragraph != nil }

// IsTable returns true if this item is a table.
func (it *InnerContentItem) IsTable() bool { return it.table != nil }

// Paragraph returns the paragraph, or nil if this item is a table.
func (it *InnerContentItem) Paragraph() *Paragraph { return it.paragraph }

// Table returns the table, or nil if this item is a paragraph.
func (it *InnerContentItem) Table() *Table { return it.table }

// BlockItemContainer is the base for proxy objects that can contain block items
// (paragraphs and tables). These include Body, Cell, Header, Footer, and Comment.
//
// Mirrors Python BlockItemContainer(StoryChild).
type BlockItemContainer struct {
	element *etree.Element // CT_Body | CT_Comment | CT_HdrFtr | CT_Tc
	part    *parts.StoryPart
}

// NewBlockItemContainer creates a new BlockItemContainer.
func NewBlockItemContainer(element *etree.Element, part *parts.StoryPart) BlockItemContainer {
	return BlockItemContainer{element: element, part: part}
}

// AddParagraph appends a new paragraph to the end of this container. If text is
// non-empty, it is placed in a single run. If style is non-nil (string name),
// the paragraph style is applied.
//
// Mirrors Python BlockItemContainer.add_paragraph.
func (c *BlockItemContainer) AddParagraph(text string, style interface{}) (*Paragraph, error) {
	p := c.addP()
	para := NewParagraph(p, c.part)
	if text != "" {
		if _, err := para.AddRun(text, nil); err != nil {
			return nil, fmt.Errorf("docx: adding run to paragraph: %w", err)
		}
	}
	if style != nil {
		if err := para.SetStyle(style); err != nil {
			return nil, fmt.Errorf("docx: setting paragraph style: %w", err)
		}
	}
	return para, nil
}

// AddTable appends a new table with the given rows, columns, and width (twips).
// The table is inserted before any trailing w:sectPr to maintain schema order.
//
// Mirrors Python BlockItemContainer.add_table (_insert_tbl with successor w:sectPr).
func (c *BlockItemContainer) AddTable(rows, cols int, widthTwips int) (*Table, error) {
	tbl := oxml.NewTbl(rows, cols, widthTwips)
	c.insertBeforeSectPr(tbl.E)
	return NewTable(tbl, c.part), nil
}

// IterInnerContent returns a slice of InnerContentItems (Paragraph or Table)
// in document order.
//
// Mirrors Python BlockItemContainer.iter_inner_content.
func (c *BlockItemContainer) IterInnerContent() []*InnerContentItem {
	var result []*InnerContentItem
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			p := &oxml.CT_P{Element: oxml.Element{E: child}}
			result = append(result, &InnerContentItem{paragraph: NewParagraph(p, c.part)})
		} else if child.Space == "w" && child.Tag == "tbl" {
			tbl := &oxml.CT_Tbl{Element: oxml.Element{E: child}}
			result = append(result, &InnerContentItem{table: NewTable(tbl, c.part)})
		}
	}
	return result
}

// Paragraphs returns all paragraphs in this container, in document order.
//
// Mirrors Python BlockItemContainer.paragraphs.
func (c *BlockItemContainer) Paragraphs() []*Paragraph {
	var result []*Paragraph
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			p := &oxml.CT_P{Element: oxml.Element{E: child}}
			result = append(result, NewParagraph(p, c.part))
		}
	}
	return result
}

// Tables returns all tables in this container, in document order.
//
// Mirrors Python BlockItemContainer.tables.
func (c *BlockItemContainer) Tables() []*Table {
	var result []*Table
	for _, child := range c.element.ChildElements() {
		if child.Space == "w" && child.Tag == "tbl" {
			tbl := &oxml.CT_Tbl{Element: oxml.Element{E: child}}
			result = append(result, NewTable(tbl, c.part))
		}
	}
	return result
}

// Element returns the backing etree element.
func (c *BlockItemContainer) Element() *etree.Element { return c.element }

// Part returns the story part this container belongs to.
func (c *BlockItemContainer) Part() *parts.StoryPart { return c.part }

// addP creates and inserts a new <w:p> element before any trailing w:sectPr.
func (c *BlockItemContainer) addP() *oxml.CT_P {
	pE := etree.NewElement("p")
	pE.Space = "w"
	c.insertBeforeSectPr(pE)
	return &oxml.CT_P{Element: oxml.Element{E: pE}}
}

// insertBeforeSectPr inserts child into this container, placing it just before
// the first direct child <w:sectPr> if one exists. If no w:sectPr child is
// present, the element is appended to the end. This matches the Python
// xmlchemy successor constraint: both w:p and w:tbl have successors=('w:sectPr',).
//
// Only CT_Body has a trailing w:sectPr; for CT_Tc, CT_HdrFtr, and CT_Comment
// there is no such child, so this degrades to a simple append.
func (c *BlockItemContainer) insertBeforeSectPr(child *etree.Element) {
	children := c.element.Child
	for i, tok := range children {
		if el, ok := tok.(*etree.Element); ok {
			if el.Tag == "sectPr" && el.Space == "w" {
				// Remove child from any existing parent
				if p := child.Parent(); p != nil {
					p.RemoveChild(child)
				}
				c.element.InsertChildAt(i, child)
				return
			}
		}
	}
	// No w:sectPr — append normally
	c.element.AddChild(child)
}
