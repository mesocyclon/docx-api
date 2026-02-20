// Package ppr implements CT_PPr and related types for paragraph properties
// in the WordprocessingML (wml) namespace.
//
// Contract: C-11 in contracts.md
// Imports: xmltypes, wml/rpr, wml/shared
package ppr

import (
	"github.com/vortex/docx-go/wml/rpr"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_PPr — full paragraph properties.
type CT_PPr struct {
	Base      CT_PPrBase
	RPr       *rpr.CT_ParaRPr // from package wml/rpr (NOT a cycle: rpr does not import ppr)
	SectPr    *CT_SectPrRef   // raw XML reference (avoid pulling in sectpr)
	PPrChange *CT_PPrChange
	Extra     []shared.RawXML
}

// CT_PPrBase — paragraph properties in STRICT ORDER (xsd:sequence!).
// Violation of order → Word shows "file is corrupted".
type CT_PPrBase struct {
	PStyle              *xmltypes.CT_String        //  1. pStyle
	KeepNext            *xmltypes.CT_OnOff         //  2. keepNext
	KeepLines           *xmltypes.CT_OnOff         //  3. keepLines
	PageBreakBefore     *xmltypes.CT_OnOff         //  4. pageBreakBefore
	FramePr             *CT_FramePr                //  5. framePr
	WidowControl        *xmltypes.CT_OnOff         //  6. widowControl
	NumPr               *CT_NumPr                  //  7. numPr
	SuppressLineNumbers *xmltypes.CT_OnOff         //  8
	PBdr                *CT_PBdr                   //  9. pBdr
	Shd                 *xmltypes.CT_Shd           // 10. shd
	Tabs                *CT_Tabs                   // 11. tabs
	SuppressAutoHyphens *xmltypes.CT_OnOff         // 12
	Kinsoku             *xmltypes.CT_OnOff         // 13
	WordWrap            *xmltypes.CT_OnOff         // 14
	OverflowPunct       *xmltypes.CT_OnOff         // 15
	TopLinePunct        *xmltypes.CT_OnOff         // 16
	AutoSpaceDE         *xmltypes.CT_OnOff         // 17
	AutoSpaceDN         *xmltypes.CT_OnOff         // 18
	Bidi                *xmltypes.CT_OnOff         // 19
	AdjustRightInd      *xmltypes.CT_OnOff         // 20
	SnapToGrid          *xmltypes.CT_OnOff         // 21
	Spacing             *CT_Spacing                // 22. spacing
	Ind                 *CT_Ind                    // 23. ind
	ContextualSpacing   *xmltypes.CT_OnOff         // 24
	MirrorIndents       *xmltypes.CT_OnOff         // 25
	SuppressOverlap     *xmltypes.CT_OnOff         // 26
	Jc                  *CT_Jc                     // 27. jc
	TextDirection       *CT_TextDirection          // 28
	TextAlignment       *CT_TextAlignment          // 29
	TextboxTightWrap    *CT_TextboxTightWrap       // 30
	OutlineLvl          *xmltypes.CT_DecimalNumber // 31
	DivId               *xmltypes.CT_DecimalNumber // 32
	CnfStyle            *CT_Cnf                    // 33
	Extra               []shared.RawXML
}

// CT_Spacing — paragraph spacing attributes.
type CT_Spacing struct {
	Before            *int    `xml:"before,attr,omitempty"`
	BeforeLines       *int    `xml:"beforeLines,attr,omitempty"`
	BeforeAutospacing *bool   `xml:"beforeAutospacing,attr,omitempty"`
	After             *int    `xml:"after,attr,omitempty"`
	AfterLines        *int    `xml:"afterLines,attr,omitempty"`
	AfterAutospacing  *bool   `xml:"afterAutospacing,attr,omitempty"`
	Line              *int    `xml:"line,attr,omitempty"`
	LineRule          *string `xml:"lineRule,attr,omitempty"`
}

// CT_Ind — paragraph indentation attributes.
type CT_Ind struct {
	Start          *int `xml:"start,attr,omitempty"`
	StartChars     *int `xml:"startChars,attr,omitempty"`
	End            *int `xml:"end,attr,omitempty"`
	EndChars       *int `xml:"endChars,attr,omitempty"`
	Hanging        *int `xml:"hanging,attr,omitempty"`
	HangingChars   *int `xml:"hangingChars,attr,omitempty"`
	FirstLine      *int `xml:"firstLine,attr,omitempty"`
	FirstLineChars *int `xml:"firstLineChars,attr,omitempty"`
}

// CT_Jc — paragraph justification.
type CT_Jc struct {
	Val string `xml:"val,attr"`
}

// CT_NumPr — numbering properties.
type CT_NumPr struct {
	Ilvl  *xmltypes.CT_DecimalNumber
	NumId *xmltypes.CT_DecimalNumber
}

// CT_Tabs — tab stop collection.
type CT_Tabs struct {
	Tab []CT_TabStop
}

// CT_TabStop — individual tab stop.
type CT_TabStop struct {
	Val    string  `xml:"val,attr"`
	Pos    int     `xml:"pos,attr"`
	Leader *string `xml:"leader,attr,omitempty"`
}

// CT_PBdr — paragraph borders.
type CT_PBdr struct {
	Top     *xmltypes.CT_Border
	Bottom  *xmltypes.CT_Border
	Left    *xmltypes.CT_Border
	Right   *xmltypes.CT_Border
	Between *xmltypes.CT_Border
	Bar     *xmltypes.CT_Border
}

// CT_FramePr — frame/drop cap properties (attribute-only element).
type CT_FramePr struct {
	DropCap    *string `xml:"dropCap,attr,omitempty"`
	Lines      *int    `xml:"lines,attr,omitempty"`
	W          *int    `xml:"w,attr,omitempty"`
	H          *int    `xml:"h,attr,omitempty"`
	HSpace     *int    `xml:"hSpace,attr,omitempty"`
	VSpace     *int    `xml:"vSpace,attr,omitempty"`
	Wrap       *string `xml:"wrap,attr,omitempty"`
	HAnchor    *string `xml:"hAnchor,attr,omitempty"`
	VAnchor    *string `xml:"vAnchor,attr,omitempty"`
	X          *int    `xml:"x,attr,omitempty"`
	XAlign     *string `xml:"xAlign,attr,omitempty"`
	Y          *int    `xml:"y,attr,omitempty"`
	YAlign     *string `xml:"yAlign,attr,omitempty"`
	HRule      *string `xml:"hRule,attr,omitempty"`
	AnchorLock *bool   `xml:"anchorLock,attr,omitempty"`
}

// CT_TextDirection — text flow direction.
type CT_TextDirection struct {
	Val string `xml:"val,attr"`
}

// CT_TextAlignment — vertical text alignment.
type CT_TextAlignment struct {
	Val string `xml:"val,attr"`
}

// CT_TextboxTightWrap — textbox tight wrap.
type CT_TextboxTightWrap struct {
	Val string `xml:"val,attr"`
}

// CT_Cnf — conditional formatting (12 boolean attributes).
type CT_Cnf struct {
	Val              *string `xml:"val,attr,omitempty"`
	FirstRow         *bool   `xml:"firstRow,attr,omitempty"`
	LastRow          *bool   `xml:"lastRow,attr,omitempty"`
	FirstColumn      *bool   `xml:"firstColumn,attr,omitempty"`
	LastColumn       *bool   `xml:"lastColumn,attr,omitempty"`
	OddVBand         *bool   `xml:"oddVBand,attr,omitempty"`
	EvenVBand        *bool   `xml:"evenVBand,attr,omitempty"`
	OddHBand         *bool   `xml:"oddHBand,attr,omitempty"`
	EvenHBand        *bool   `xml:"evenHBand,attr,omitempty"`
	FirstRowFirstCol *bool   `xml:"firstRowFirstColumn,attr,omitempty"`
	FirstRowLastCol  *bool   `xml:"firstRowLastColumn,attr,omitempty"`
	LastRowFirstCol  *bool   `xml:"lastRowFirstColumn,attr,omitempty"`
	LastRowLastCol   *bool   `xml:"lastRowLastColumn,attr,omitempty"`
}

// CT_SectPrRef — raw XML reference to section properties (avoid importing sectpr).
type CT_SectPrRef struct {
	InnerXML []byte `xml:",innerxml"`
}

// CT_PPrChange — track changes for paragraph properties.
type CT_PPrChange struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Date   string `xml:"date,attr,omitempty"`
	PPr    *CT_PPrBase
}
