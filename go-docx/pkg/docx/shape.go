package docx

import (
	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// Namespace URIs for shape type detection.
const (
	nsPicture = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	nsChart   = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	nsDiagram = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
)

// InlineShapes is a sequence of InlineShape instances found in a document body.
//
// Mirrors Python InlineShapes(Parented).
type InlineShapes struct {
	body *etree.Element // CT_Body element
}

// newInlineShapes creates a new InlineShapes proxy.
func newInlineShapes(body *etree.Element) *InlineShapes {
	return &InlineShapes{body: body}
}

// Len returns the number of inline shapes in the document.
func (iss *InlineShapes) Len() int {
	return len(iss.inlineList())
}

// Get returns the inline shape at the given index.
func (iss *InlineShapes) Get(idx int) (*InlineShape, error) {
	list := iss.inlineList()
	if idx < 0 || idx >= len(list) {
		return nil, errIndexOutOfRange("InlineShapes", idx, len(list))
	}
	return &InlineShape{inline: list[idx]}, nil
}

// Iter returns all inline shapes in the document.
func (iss *InlineShapes) Iter() []*InlineShape {
	list := iss.inlineList()
	result := make([]*InlineShape, len(list))
	for i, il := range list {
		result[i] = &InlineShape{inline: il}
	}
	return result
}

// inlineList walks the body tree to find all wp:inline elements
// (equivalent to Python's //w:p/w:r/w:drawing/wp:inline xpath).
func (iss *InlineShapes) inlineList() []*oxml.CT_Inline {
	var result []*oxml.CT_Inline
	for _, p := range iss.body.ChildElements() {
		if !(p.Space == "w" && p.Tag == "p") {
			continue
		}
		for _, r := range p.ChildElements() {
			if !(r.Space == "w" && r.Tag == "r") {
				continue
			}
			for _, drawing := range r.ChildElements() {
				if !(drawing.Space == "w" && drawing.Tag == "drawing") {
					continue
				}
				for _, inline := range drawing.ChildElements() {
					if inline.Space == "wp" && inline.Tag == "inline" {
						result = append(result, &oxml.CT_Inline{Element: oxml.WrapElement(inline)})
					}
				}
			}
		}
	}
	return result
}

// InlineShape is a proxy for a <wp:inline> element representing an inline graphical object.
//
// Mirrors Python InlineShape.
type InlineShape struct {
	inline *oxml.CT_Inline
}

// newInlineShape creates a new InlineShape proxy.
func newInlineShape(elm *oxml.CT_Inline) *InlineShape {
	return &InlineShape{inline: elm}
}

// Height returns the display height of this inline shape as a Length (EMU).
//
// Mirrors Python InlineShape.height (getter).
func (is *InlineShape) Height() (Length, error) {
	cy, err := is.inline.ExtentCy()
	if err != nil {
		return 0, err
	}
	return Length(cy), nil
}

// SetHeight sets the display height of this inline shape.
//
// Mirrors Python InlineShape.height (setter).
func (is *InlineShape) SetHeight(v Length) error {
	cy := int64(v)
	if err := is.inline.SetExtentCy(cy); err != nil {
		return err
	}
	// Also update the spPr transform if accessible
	graphic, err := is.inline.Graphic()
	if err != nil {
		return nil // no graphic, just extent was enough
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return nil
	}
	pic := gd.Pic()
	if pic == nil {
		return nil
	}
	spPr, err := pic.SpPr()
	if err != nil {
		return nil
	}
	return spPr.SetCy(cy)
}

// Width returns the display width of this inline shape as a Length (EMU).
//
// Mirrors Python InlineShape.width (getter).
func (is *InlineShape) Width() (Length, error) {
	cx, err := is.inline.ExtentCx()
	if err != nil {
		return 0, err
	}
	return Length(cx), nil
}

// SetWidth sets the display width of this inline shape.
//
// Mirrors Python InlineShape.width (setter).
func (is *InlineShape) SetWidth(v Length) error {
	cx := int64(v)
	if err := is.inline.SetExtentCx(cx); err != nil {
		return err
	}
	// Also update the spPr transform if accessible
	graphic, err := is.inline.Graphic()
	if err != nil {
		return nil
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return nil
	}
	pic := gd.Pic()
	if pic == nil {
		return nil
	}
	spPr, err := pic.SpPr()
	if err != nil {
		return nil
	}
	return spPr.SetCx(cx)
}

// Type returns the type of this inline shape (PICTURE, LINKED_PICTURE, CHART,
// SMART_ART, or NOT_IMPLEMENTED).
//
// Mirrors Python InlineShape.type.
func (is *InlineShape) Type() enum.WdInlineShapeType {
	graphic, err := is.inline.Graphic()
	if err != nil {
		return enum.WdInlineShapeTypeNotImplemented
	}
	gd, err := graphic.GraphicData()
	if err != nil {
		return enum.WdInlineShapeTypeNotImplemented
	}
	uri, err := gd.Uri()
	if err != nil {
		return enum.WdInlineShapeTypeNotImplemented
	}

	switch uri {
	case nsPicture:
		pic := gd.Pic()
		if pic == nil {
			return enum.WdInlineShapeTypePicture
		}
		bf, err := pic.BlipFill()
		if err != nil {
			return enum.WdInlineShapeTypePicture
		}
		blip := bf.Blip()
		if blip == nil {
			return enum.WdInlineShapeTypePicture
		}
		if blip.Link() != "" {
			return enum.WdInlineShapeTypeLinkedPicture
		}
		return enum.WdInlineShapeTypePicture
	case nsChart:
		return enum.WdInlineShapeTypeChart
	case nsDiagram:
		return enum.WdInlineShapeTypeSmartArt
	default:
		return enum.WdInlineShapeTypeNotImplemented
	}
}
