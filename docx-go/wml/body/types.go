// Package body implements CT_Document and CT_Body — the top-level containers
// of a WordprocessingML document.xml part.
//
// Contract: C-17 in contracts.md.
// Imports:  xmltypes, wml/para, wml/table, wml/sectpr, wml/shared.
package body

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/sectpr"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
)

// ---------------------------------------------------------------------------
// CT_Document — root element <w:document>
// ---------------------------------------------------------------------------

// CT_Document represents the root <w:document> element of document.xml.
type CT_Document struct {
	Body  *CT_Body
	Extra []shared.RawXML

	// Namespaces stores all namespace declarations (xmlns:*) from the
	// original <w:document> element so they survive a round-trip.
	// See patterns.md §6.
	Namespaces []xml.Attr
}

// ---------------------------------------------------------------------------
// CT_Body — <w:body>
// ---------------------------------------------------------------------------

// CT_Body represents the document body.  It contains an ordered sequence of
// block-level elements (paragraphs, tables, SDTs, …) and an optional
// trailing section properties element.
type CT_Body struct {
	Content []shared.BlockLevelElement // p, tbl, sdt, or unknown
	SectPr  *sectpr.CT_SectPr         // trailing body-level <w:sectPr>
}

// ---------------------------------------------------------------------------
// Wrapper types implementing shared.BlockLevelElement
// ---------------------------------------------------------------------------

// ParagraphElement wraps a paragraph and satisfies BlockLevelElement.
type ParagraphElement struct {
	shared.BlockLevelMarker
	P *para.CT_P
}

// TableElement wraps a table and satisfies BlockLevelElement.
type TableElement struct {
	shared.BlockLevelMarker
	T *table.CT_Tbl
}

// SdtBlockElement wraps a block-level structured document tag.
type SdtBlockElement struct {
	shared.BlockLevelMarker
	Sdt *CT_SdtBlock
}

// RawBlockElement wraps an unrecognised block-level element for round-trip.
type RawBlockElement struct {
	shared.BlockLevelMarker
	Raw shared.RawXML
}

// ---------------------------------------------------------------------------
// CT_SdtBlock — <w:sdt> at block level
// ---------------------------------------------------------------------------

// CT_SdtBlock represents a block-level structured document tag (content
// control).  SdtPr and SdtEndPr are stored as raw XML because their full
// schema is very complex and not needed for an MVP.
type CT_SdtBlock struct {
	SdtPr      *shared.RawXML
	SdtEndPr   *shared.RawXML
	SdtContent []shared.BlockLevelElement
}
