// Package sectpr implements CT_SectPr — section properties for OOXML documents.
package sectpr

import (
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_SectPr represents section properties (w:sectPr).
// Elements MUST be marshalled in strict xsd:sequence order.
type CT_SectPr struct {
	// EG_HdrFtrReferences
	HeaderRefs []CT_HdrFtrRef
	FooterRefs []CT_HdrFtrRef
	// EG_SectPrContents (STRICT ORDER)
	FootnotePr    *CT_FtnProps
	EndnotePr     *CT_EdnProps
	Type          *CT_SectType
	PgSz          *CT_PageSz
	PgMar         *CT_PageMar
	PaperSrc      *CT_PaperSource
	PgBorders     *CT_PageBorders
	LnNumType     *CT_LineNumber
	PgNumType     *CT_PageNumber
	Cols          *CT_Columns
	FormProt      *xmltypes.CT_OnOff
	VAlign        *CT_VerticalJc
	NoEndnote     *xmltypes.CT_OnOff
	TitlePg       *xmltypes.CT_OnOff
	TextDirection *CT_TextDirection
	Bidi          *xmltypes.CT_OnOff
	RtlGutter     *xmltypes.CT_OnOff
	DocGrid       *CT_DocGrid
	// Attributes
	RsidR    *string `xml:"rsidR,attr,omitempty"`
	RsidSect *string `xml:"rsidSect,attr,omitempty"`
	// Unknown/extension elements (round-trip)
	Extra []shared.RawXML
}

// CT_HdrFtrRef is a header or footer reference.
// XML: <w:headerReference w:type="default" r:id="rId8"/>
//
//	<w:footerReference w:type="default" r:id="rId10"/>
type CT_HdrFtrRef struct {
	Type string `xml:"type,attr"` // "default"|"first"|"even"
	RID  string `xml:"id,attr"`   // r:id → relationship
}

// CT_PageSz represents page size (w:pgSz).
type CT_PageSz struct {
	W      int     `xml:"w,attr"`
	H      int     `xml:"h,attr"`
	Orient *string `xml:"orient,attr,omitempty"`
	Code   *int    `xml:"code,attr,omitempty"`
}

// CT_PageMar represents page margins (w:pgMar).
type CT_PageMar struct {
	Top    int `xml:"top,attr"`
	Right  int `xml:"right,attr"`
	Bottom int `xml:"bottom,attr"`
	Left   int `xml:"left,attr"`
	Header int `xml:"header,attr"`
	Footer int `xml:"footer,attr"`
	Gutter int `xml:"gutter,attr"`
}

// CT_Columns represents column definitions (w:cols).
type CT_Columns struct {
	EqualWidth *bool       `xml:"equalWidth,attr,omitempty"`
	Space      *int        `xml:"space,attr,omitempty"`
	Num        *int        `xml:"num,attr,omitempty"`
	Sep        *bool       `xml:"sep,attr,omitempty"`
	Col        []CT_Column `xml:"col,omitempty"`
}

// CT_Column represents a single column definition.
type CT_Column struct {
	W     *int `xml:"w,attr,omitempty"`
	Space *int `xml:"space,attr,omitempty"`
}

// CT_DocGrid represents document grid settings (w:docGrid).
type CT_DocGrid struct {
	Type      *string `xml:"type,attr,omitempty"`
	LinePitch *int    `xml:"linePitch,attr,omitempty"`
	CharSpace *int    `xml:"charSpace,attr,omitempty"`
}

// CT_SectType represents section type (w:type).
type CT_SectType struct {
	Val string `xml:"val,attr"` // "nextPage"|"nextColumn"|"continuous"|"evenPage"|"oddPage"
}

// CT_VerticalJc represents vertical alignment (w:vAlign).
type CT_VerticalJc struct {
	Val string `xml:"val,attr"` // "top"|"center"|"both"|"bottom"
}

// CT_TextDirection represents text flow direction (w:textDirection).
type CT_TextDirection struct {
	Val string `xml:"val,attr"`
}

// CT_PageNumber represents page numbering settings (w:pgNumType).
type CT_PageNumber struct {
	Fmt       *string `xml:"fmt,attr,omitempty"`
	Start     *int    `xml:"start,attr,omitempty"`
	ChapStyle *string `xml:"chapStyle,attr,omitempty"`
	ChapSep   *string `xml:"chapSep,attr,omitempty"`
}

// CT_PageBorders represents page borders (w:pgBorders).
type CT_PageBorders struct {
	OffsetFrom *string        `xml:"offsetFrom,attr,omitempty"`
	ZOrder     *string        `xml:"zOrder,attr,omitempty"`
	Display    *string        `xml:"display,attr,omitempty"`
	Top        *CT_PageBorder `xml:"top,omitempty"`
	Left       *CT_PageBorder `xml:"left,omitempty"`
	Bottom     *CT_PageBorder `xml:"bottom,omitempty"`
	Right      *CT_PageBorder `xml:"right,omitempty"`
}

// CT_PageBorder represents a single page border.
type CT_PageBorder struct {
	Val        string  `xml:"val,attr"`
	Sz         *int    `xml:"sz,attr,omitempty"`
	Space      *int    `xml:"space,attr,omitempty"`
	Color      *string `xml:"color,attr,omitempty"`
	ThemeColor *string `xml:"themeColor,attr,omitempty"`
	Shadow     *bool   `xml:"shadow,attr,omitempty"`
	Frame      *bool   `xml:"frame,attr,omitempty"`
}

// CT_LineNumber represents line numbering settings (w:lnNumType).
type CT_LineNumber struct {
	CountBy  *int    `xml:"countBy,attr,omitempty"`
	Start    *int    `xml:"start,attr,omitempty"`
	Restart  *string `xml:"restart,attr,omitempty"`
	Distance *int    `xml:"distance,attr,omitempty"`
}

// CT_PaperSource represents printer tray settings (w:paperSrc).
type CT_PaperSource struct {
	First *int `xml:"first,attr,omitempty"`
	Other *int `xml:"other,attr,omitempty"`
}

// CT_FtnProps represents footnote properties (w:footnotePr).
type CT_FtnProps struct {
	Pos        *CT_FtnPos     `xml:"pos,omitempty"`
	NumFmt     *CT_NumFmt     `xml:"numFmt,omitempty"`
	NumStart   *CT_NumStart   `xml:"numStart,omitempty"`
	NumRestart *CT_NumRestart `xml:"numRestart,omitempty"`
}

// CT_EdnProps represents endnote properties (w:endnotePr).
type CT_EdnProps struct {
	Pos        *CT_EdnPos     `xml:"pos,omitempty"`
	NumFmt     *CT_NumFmt     `xml:"numFmt,omitempty"`
	NumStart   *CT_NumStart   `xml:"numStart,omitempty"`
	NumRestart *CT_NumRestart `xml:"numRestart,omitempty"`
}

// Helper types for footnote/endnote properties.
type CT_FtnPos struct {
	Val string `xml:"val,attr"`
}
type CT_EdnPos struct {
	Val string `xml:"val,attr"`
}
type CT_NumFmt struct {
	Val string `xml:"val,attr"`
}
type CT_NumStart struct {
	Val int `xml:"val,attr"`
}
type CT_NumRestart struct {
	Val string `xml:"val,attr"`
}
