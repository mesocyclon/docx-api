// Package table implements WML table types (CT_Tbl, CT_Row, CT_Tc, etc.)
// Contract: C-13 in contracts.md.
// Imports ONLY: xmltypes, wml/shared.
package table

import (
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_Tbl represents a table element <w:tbl>.
type CT_Tbl struct {
	shared.BlockLevelMarker // implements shared.BlockLevelElement
	TblPr                   *CT_TblPr
	TblGrid                 *CT_TblGrid
	// Content: rows + bookmarks + track changes.
	Content []TblContent
	Extra   []shared.RawXML
}

// TblContent — interface for table-level content (rows or unknown elements).
type TblContent interface {
	isTblContent()
}

// CT_TblPr — table properties in STRICT XSD sequence order.
type CT_TblPr struct {
	TblStyle            *xmltypes.CT_String
	TblpPr              *CT_TblPPr
	TblOverlap          *CT_TblOverlap
	BidiVisual          *xmltypes.CT_OnOff
	TblStyleRowBandSize *xmltypes.CT_DecimalNumber
	TblStyleColBandSize *xmltypes.CT_DecimalNumber
	TblW                *CT_TblWidth
	Jc                  *CT_JcTable
	TblCellSpacing      *CT_TblWidth
	TblInd              *CT_TblWidth
	TblBorders          *CT_TblBorders
	Shd                 *xmltypes.CT_Shd
	TblLayout           *CT_TblLayoutType
	TblCellMar          *CT_TblCellMar
	TblLook             *CT_TblLook
	TblCaption          *xmltypes.CT_String
	TblDescription      *xmltypes.CT_String
	TblPrChange         *CT_TblPrChange
	Extra               []shared.RawXML
}

// CT_TblGrid — table grid definition.
type CT_TblGrid struct {
	GridCol []CT_TblGridCol
}

// CT_TblGridCol — single grid column width.
type CT_TblGridCol struct {
	W int `xml:"w,attr"`
}

// CT_Row represents a table row <w:tr>.
type CT_Row struct {
	TblPrEx *CT_TblPrEx
	TrPr    *CT_TrPr
	Content []RowContent
	// Attributes
	RsidR  *string `xml:"rsidR,attr,omitempty"`
	RsidTr *string `xml:"rsidTr,attr,omitempty"`
	Extra  []shared.RawXML
}

// CT_Row implements TblContent.
func (CT_Row) isTblContent() {}

// RowContent — interface for row-level content (cells or unknown elements).
type RowContent interface {
	isRowContent()
}

// CT_Tc represents a table cell <w:tc>.
type CT_Tc struct {
	TcPr    *CT_TcPr
	Content []shared.BlockLevelElement
	// IMPORTANT: always contains ≥1 element (minimum empty <w:p/>).
}

// CT_Tc implements RowContent.
func (CT_Tc) isRowContent() {}

// RawTblContent wraps shared.RawXML for unknown table-level elements.
type RawTblContent struct {
	shared.RawXML
}

func (RawTblContent) isTblContent() {}

// RawRowContent wraps shared.RawXML for unknown row-level elements.
type RawRowContent struct {
	shared.RawXML
}

func (RawRowContent) isRowContent() {}

// CT_TcPr — cell properties in STRICT XSD sequence order.
type CT_TcPr struct {
	CnfStyle      *CT_Cnf
	TcW           *CT_TblWidth
	GridSpan      *xmltypes.CT_DecimalNumber
	HMerge        *CT_HMerge
	VMerge        *CT_VMerge
	TcBorders     *CT_TcBorders
	Shd           *xmltypes.CT_Shd
	NoWrap        *xmltypes.CT_OnOff
	TcMar         *CT_TblCellMar
	TextDirection *CT_TextDirection
	TcFitText     *xmltypes.CT_OnOff
	VAlign        *CT_VerticalJc
	HideMark      *xmltypes.CT_OnOff
	TcPrChange    *CT_TcPrChange
	Extra         []shared.RawXML
}

// CT_TrPr — row properties.
type CT_TrPr struct {
	CnfStyle       *CT_Cnf
	GridBefore     *xmltypes.CT_DecimalNumber
	GridAfter      *xmltypes.CT_DecimalNumber
	WBefore        *CT_TblWidth
	WAfter         *CT_TblWidth
	CantSplit      *xmltypes.CT_OnOff
	TrHeight       *CT_Height
	TblHeader      *xmltypes.CT_OnOff
	TblCellSpacing *CT_TblWidth
	Jc             *CT_JcTable
	Hidden         *xmltypes.CT_OnOff
	TrPrChange     *CT_TrPrChange
	Extra          []shared.RawXML
}

// CT_TblWidth — width specification (value + type).
type CT_TblWidth struct {
	W    int    `xml:"w,attr"`
	Type string `xml:"type,attr"`
}

// CT_TblBorders — table-level borders.
type CT_TblBorders struct {
	Top     *xmltypes.CT_Border
	Start   *xmltypes.CT_Border // or Left in transitional
	Bottom  *xmltypes.CT_Border
	End     *xmltypes.CT_Border // or Right in transitional
	InsideH *xmltypes.CT_Border
	InsideV *xmltypes.CT_Border
}

// CT_TcBorders — cell-level borders.
type CT_TcBorders struct {
	Top     *xmltypes.CT_Border
	Start   *xmltypes.CT_Border
	Bottom  *xmltypes.CT_Border
	End     *xmltypes.CT_Border
	InsideH *xmltypes.CT_Border
	InsideV *xmltypes.CT_Border
	Tl2br   *xmltypes.CT_Border
	Tr2bl   *xmltypes.CT_Border
}

// CT_TblCellMar — cell margins.
type CT_TblCellMar struct {
	Top    *CT_TblWidth
	Start  *CT_TblWidth
	Bottom *CT_TblWidth
	End    *CT_TblWidth
}

// CT_TblLook — conditional formatting flags.
type CT_TblLook struct {
	FirstRow    *bool `xml:"firstRow,attr,omitempty"`
	LastRow     *bool `xml:"lastRow,attr,omitempty"`
	FirstColumn *bool `xml:"firstColumn,attr,omitempty"`
	LastColumn  *bool `xml:"lastColumn,attr,omitempty"`
	NoHBand     *bool `xml:"noHBand,attr,omitempty"`
	NoVBand     *bool `xml:"noVBand,attr,omitempty"`
}

// CT_Height — row height.
type CT_Height struct {
	Val   *int    `xml:"val,attr,omitempty"`
	HRule *string `xml:"hRule,attr,omitempty"`
}

// CT_HMerge — horizontal merge attribute.
type CT_HMerge struct {
	Val *string `xml:"val,attr,omitempty"`
}

// CT_VMerge — vertical merge attribute.
type CT_VMerge struct {
	Val *string `xml:"val,attr,omitempty"`
}

// CT_TblLayoutType — table layout mode.
type CT_TblLayoutType struct {
	Type string `xml:"type,attr"`
}

// CT_JcTable — table justification.
type CT_JcTable struct {
	Val string `xml:"val,attr"`
}

// CT_TblPPr — floating table positioning (placeholder struct).
type CT_TblPPr struct {
	LeftFromText   *int    `xml:"leftFromText,attr,omitempty"`
	RightFromText  *int    `xml:"rightFromText,attr,omitempty"`
	TopFromText    *int    `xml:"topFromText,attr,omitempty"`
	BottomFromText *int    `xml:"bottomFromText,attr,omitempty"`
	VertAnchor     *string `xml:"vertAnchor,attr,omitempty"`
	HorzAnchor     *string `xml:"horzAnchor,attr,omitempty"`
	TblpXSpec      *string `xml:"tblpXSpec,attr,omitempty"`
	TblpYSpec      *string `xml:"tblpYSpec,attr,omitempty"`
	TblpX          *int    `xml:"tblpX,attr,omitempty"`
	TblpY          *int    `xml:"tblpY,attr,omitempty"`
}

// CT_TblOverlap — table overlap.
type CT_TblOverlap struct {
	Val string `xml:"val,attr"`
}

// CT_VerticalJc — vertical alignment.
type CT_VerticalJc struct {
	Val string `xml:"val,attr"`
}

// CT_TextDirection — text direction.
type CT_TextDirection struct {
	Val string `xml:"val,attr"`
}

// CT_Cnf — conditional formatting (12 boolean attrs).
type CT_Cnf struct {
	Val                 *string `xml:"val,attr,omitempty"`
	FirstRow            *bool   `xml:"firstRow,attr,omitempty"`
	LastRow             *bool   `xml:"lastRow,attr,omitempty"`
	FirstColumn         *bool   `xml:"firstColumn,attr,omitempty"`
	LastColumn          *bool   `xml:"lastColumn,attr,omitempty"`
	OddVBand            *bool   `xml:"oddVBand,attr,omitempty"`
	EvenVBand           *bool   `xml:"evenVBand,attr,omitempty"`
	OddHBand            *bool   `xml:"oddHBand,attr,omitempty"`
	EvenHBand           *bool   `xml:"evenHBand,attr,omitempty"`
	FirstRowFirstColumn *bool   `xml:"firstRowFirstColumn,attr,omitempty"`
	FirstRowLastColumn  *bool   `xml:"firstRowLastColumn,attr,omitempty"`
	LastRowFirstColumn  *bool   `xml:"lastRowFirstColumn,attr,omitempty"`
	LastRowLastColumn   *bool   `xml:"lastRowLastColumn,attr,omitempty"`
}

// CT_TblPrChange — track change for table properties.
type CT_TblPrChange struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Date   string `xml:"date,attr,omitempty"`
}

// CT_TrPrChange — track change for row properties.
type CT_TrPrChange struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Date   string `xml:"date,attr,omitempty"`
}

// CT_TcPrChange — track change for cell properties.
type CT_TcPrChange struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Date   string `xml:"date,attr,omitempty"`
}

// CT_TblPrEx — row-level table property overrides (placeholder).
type CT_TblPrEx struct {
	TblW       *CT_TblWidth
	Jc         *CT_JcTable
	TblBorders *CT_TblBorders
	Shd        *xmltypes.CT_Shd
	TblLayout  *CT_TblLayoutType
	TblCellMar *CT_TblCellMar
	TblLook    *CT_TblLook
	Extra      []shared.RawXML
}
