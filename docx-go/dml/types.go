// Package dml implements the DrawingML types used for inline and floating
// images inside WordprocessingML documents.
//
// Contract: C-19 in contracts.md
// Imports: xmltypes, wml/shared
package dml

import (
	"github.com/vortex/docx-go/wml/shared"
)

// ---------------------------------------------------------------------------
// WordprocessingDrawing types  (wp: namespace)
// ---------------------------------------------------------------------------

// WP_Inline represents an inline image (<wp:inline>).
type WP_Inline struct {
	DistT        int              `xml:"distT,attr"`
	DistB        int              `xml:"distB,attr"`
	DistL        int              `xml:"distL,attr"`
	DistR        int              `xml:"distR,attr"`
	Extent       WP_Extent        // wp:extent
	EffectExtent *WP_EffectExtent // wp:effectExtent
	DocPr        WP_DocPr         // wp:docPr
	Graphic      A_Graphic        // a:graphic
	Extra        []shared.RawXML  // unknown child elements (e.g. wp:cNvGraphicFramePr)
}

// WP_Anchor represents a floating image (<wp:anchor>).
type WP_Anchor struct {
	BehindDoc      bool
	DistT, DistB   int
	DistL, DistR   int
	RelativeHeight int
	SimplePos      bool
	Locked         bool
	LayoutInCell   bool
	AllowOverlap   bool
	SimplePosXY    *WP_Point        // wp:simplePos
	PositionH      WP_PosH          // wp:positionH
	PositionV      WP_PosV          // wp:positionV
	Extent         WP_Extent        // wp:extent
	EffectExtent   *WP_EffectExtent // wp:effectExtent
	WrapType       interface{}      // WP_WrapNone | WP_WrapSquare | WP_WrapTight | WP_WrapTopAndBottom
	DocPr          WP_DocPr         // wp:docPr
	Graphic        A_Graphic        // a:graphic
	Extra          []shared.RawXML  // unknown child elements
}

// WP_Extent holds the size of a drawing in EMU.
type WP_Extent struct {
	CX int64 `xml:"cx,attr"`
	CY int64 `xml:"cy,attr"`
}

// WP_EffectExtent holds the additional space reserved for effects (EMU).
type WP_EffectExtent struct {
	L, T, R, B int64
}

// WP_DocPr holds non-visual drawing properties.
type WP_DocPr struct {
	ID    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Descr string `xml:"descr,attr,omitempty"`
}

// WP_PosH describes horizontal positioning of a floating image.
type WP_PosH struct {
	RelativeFrom string  // relativeFrom attribute
	PosOffset    *int64  // wp:posOffset child
	Align        *string // wp:align child
}

// WP_PosV describes vertical positioning of a floating image.
type WP_PosV struct {
	RelativeFrom string
	PosOffset    *int64
	Align        *string
}

// WP_Point is a simple (x,y) coordinate pair in EMU.
type WP_Point struct{ X, Y int64 }

// ---------------------------------------------------------------------------
// Wrap types
// ---------------------------------------------------------------------------

// WP_WrapNone means no text wrapping.
type WP_WrapNone struct{}

// WP_WrapSquare wraps text around a rectangular boundary.
type WP_WrapSquare struct{ WrapText string }

// WP_WrapTight wraps text tightly around the image contour.
type WP_WrapTight struct{ WrapText string }

// WP_WrapTopAndBottom places text only above and below.
type WP_WrapTopAndBottom struct{}

// ---------------------------------------------------------------------------
// DrawingML Graphic types  (a: namespace)
// ---------------------------------------------------------------------------

// A_Graphic wraps the graphicData payload (<a:graphic>).
type A_Graphic struct {
	GraphicData A_GraphicData
}

// A_GraphicData carries either a typed picture or raw fallback data.
type A_GraphicData struct {
	URI     string          // uri attribute
	Pic     *PIC_Pic        // pic:pic (when URI is the picture namespace)
	RawData *shared.RawXML  // fallback for non-picture graphic data
}

// ---------------------------------------------------------------------------
// Picture types  (pic: namespace)
// ---------------------------------------------------------------------------

// PIC_Pic represents a picture element (<pic:pic>).
type PIC_Pic struct {
	NvPicPr  PIC_NvPicPr  // pic:nvPicPr
	BlipFill PIC_BlipFill // pic:blipFill
	SpPr     A_SpPr       // pic:spPr
}

// PIC_NvPicPr contains non-visual picture properties.
type PIC_NvPicPr struct {
	CNvPr WP_DocPr // pic:cNvPr (reuses DocPr shape)
}

// PIC_BlipFill contains the image fill for a picture.
type PIC_BlipFill struct {
	Blip    A_Blip     // a:blip
	Stretch *A_Stretch // a:stretch
}

// A_Blip references an image resource via relationship ID.
type A_Blip struct {
	Embed string `xml:"embed,attr"` // r:embed
	Link  string `xml:"link,attr,omitempty"`
}

// A_Stretch marks the blip as stretched to fill.
type A_Stretch struct{}

// ---------------------------------------------------------------------------
// Shape Properties types  (a: namespace)
// ---------------------------------------------------------------------------

// A_SpPr holds shape properties for the picture.
type A_SpPr struct {
	Xfrm     *A_Xfrm     // a:xfrm
	PrstGeom *A_PrstGeom  // a:prstGeom
	Extra    []shared.RawXML
}

// A_Xfrm describes the transform (offset + extent).
type A_Xfrm struct {
	Off A_Off // a:off
	Ext A_Ext // a:ext
}

// A_Off is an offset point.
type A_Off struct{ X, Y int64 }

// A_Ext is an extent (width × height).
type A_Ext struct{ CX, CY int64 }

// A_PrstGeom describes a preset geometry shape.
type A_PrstGeom struct {
	Prst string // prst attribute ("rect", "ellipse", …)
}
