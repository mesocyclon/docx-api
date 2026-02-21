package oxml

import "fmt"

// ===========================================================================
// CT_Inline — custom methods
// ===========================================================================

// NewPicInline creates a new <wp:inline> element containing a <pic:pic> element
// for an inline image. shape_id is an integer identifier, rId is the relationship
// id for the image part, filename is the original image name, cx and cy are the
// image dimensions in EMU.
func NewPicInline(shapeId int, rId, filename string, cx, cy int64) *CT_Inline {
	pic := newPicture(0, filename, rId, cx, cy)
	return newInline(cx, cy, shapeId, pic)
}

// newInline creates a <wp:inline> skeleton and fills it with the given values.
func newInline(cx, cy int64, shapeId int, pic *CT_Picture) *CT_Inline {
	xml := fmt.Sprintf(
		`<wp:inline `+
			`xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" `+
			`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `+
			`xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" `+
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+
			`<wp:extent cx="914400" cy="914400"/>`+
			`<wp:docPr id="666" name="unnamed"/>`+
			`<wp:cNvGraphicFramePr>`+
			`<a:graphicFrameLocks noChangeAspect="1"/>`+
			`</wp:cNvGraphicFramePr>`+
			`<a:graphic>`+
			`<a:graphicData uri="URI not set"/>`+
			`</a:graphic>`+
			`</wp:inline>`,
	)
	el, err := ParseXml([]byte(xml))
	if err != nil {
		panic(fmt.Sprintf("shape_custom: failed to parse inline XML: %v", err))
	}
	inline := &CT_Inline{Element{E: el}}

	// Set extent dimensions
	inline.Extent().SetCx(cx)
	inline.Extent().SetCy(cy)

	// Set docPr
	inline.DocPr().SetId(shapeId)
	inline.DocPr().SetName(fmt.Sprintf("Picture %d", shapeId))

	// Set graphic data URI and insert the picture element
	gd := inline.Graphic().GraphicData()
	gd.SetUri("http://schemas.openxmlformats.org/drawingml/2006/picture")
	// Insert pic:pic into graphicData
	gd.E.AddChild(pic.E)

	return inline
}

// newPicture creates a minimum viable <pic:pic> element.
func newPicture(picId int, filename, rId string, cx, cy int64) *CT_Picture {
	xml := fmt.Sprintf(
		`<pic:pic `+
			`xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" `+
			`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `+
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+
			`<pic:nvPicPr>`+
			`<pic:cNvPr id="666" name="unnamed"/>`+
			`<pic:cNvPicPr/>`+
			`</pic:nvPicPr>`+
			`<pic:blipFill>`+
			`<a:blip/>`+
			`<a:stretch>`+
			`<a:fillRect/>`+
			`</a:stretch>`+
			`</pic:blipFill>`+
			`<pic:spPr>`+
			`<a:xfrm>`+
			`<a:off x="0" y="0"/>`+
			`<a:ext cx="914400" cy="914400"/>`+
			`</a:xfrm>`+
			`<a:prstGeom prst="rect"/>`+
			`</pic:spPr>`+
			`</pic:pic>`,
	)
	el, err := ParseXml([]byte(xml))
	if err != nil {
		panic(fmt.Sprintf("shape_custom: failed to parse pic XML: %v", err))
	}
	pic := &CT_Picture{Element{E: el}}

	// Set picture properties
	pic.NvPicPr().CNvPr().SetId(picId)
	pic.NvPicPr().CNvPr().SetName(filename)
	pic.BlipFill().Blip().SetEmbed(rId)
	pic.SpPr().SetCx(cx)
	pic.SpPr().SetCy(cy)

	return pic
}

// ExtentCx returns the width of the inline image in EMU.
func (i *CT_Inline) ExtentCx() int64 {
	v, _ := i.Extent().Cx()
	return v
}

// ExtentCy returns the height of the inline image in EMU.
func (i *CT_Inline) ExtentCy() int64 {
	v, _ := i.Extent().Cy()
	return v
}

// SetExtentCx sets the width of the inline image in EMU.
func (i *CT_Inline) SetExtentCx(v int64) {
	i.Extent().SetCx(v)
}

// SetExtentCy sets the height of the inline image in EMU.
func (i *CT_Inline) SetExtentCy(v int64) {
	i.Extent().SetCy(v)
}

// ===========================================================================
// CT_ShapeProperties — custom methods
// ===========================================================================

// Cx returns the shape width in EMU via xfrm/ext/@cx, or nil if not present.
func (sp *CT_ShapeProperties) Cx() *int64 {
	xfrm := sp.Xfrm()
	if xfrm == nil {
		return nil
	}
	return xfrm.CxVal()
}

// SetCx sets the shape width in EMU via xfrm/ext/@cx.
func (sp *CT_ShapeProperties) SetCx(v int64) {
	xfrm := sp.GetOrAddXfrm()
	xfrm.SetCxVal(v)
}

// Cy returns the shape height in EMU via xfrm/ext/@cy, or nil if not present.
func (sp *CT_ShapeProperties) Cy() *int64 {
	xfrm := sp.Xfrm()
	if xfrm == nil {
		return nil
	}
	return xfrm.CyVal()
}

// SetCy sets the shape height in EMU via xfrm/ext/@cy.
func (sp *CT_ShapeProperties) SetCy(v int64) {
	xfrm := sp.GetOrAddXfrm()
	xfrm.SetCyVal(v)
}

// ===========================================================================
// CT_Transform2D — custom methods
// ===========================================================================

// CxVal returns the width in EMU from ext/@cx, or nil if ext is not present.
func (t *CT_Transform2D) CxVal() *int64 {
	ext := t.Ext()
	if ext == nil {
		return nil
	}
	v, err := ext.Cx()
	if err != nil {
		return nil
	}
	return &v
}

// SetCxVal sets the width in EMU on ext/@cx, creating ext if needed.
func (t *CT_Transform2D) SetCxVal(v int64) {
	ext := t.GetOrAddExt()
	ext.SetCx(v)
}

// CyVal returns the height in EMU from ext/@cy, or nil if ext is not present.
func (t *CT_Transform2D) CyVal() *int64 {
	ext := t.Ext()
	if ext == nil {
		return nil
	}
	v, err := ext.Cy()
	if err != nil {
		return nil
	}
	return &v
}

// SetCyVal sets the height in EMU on ext/@cy, creating ext if needed.
func (t *CT_Transform2D) SetCyVal(v int64) {
	ext := t.GetOrAddExt()
	ext.SetCy(v)
}
