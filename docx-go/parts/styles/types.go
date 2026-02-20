// Package styles implements parsing and serialisation of the styles part
// (word/styles.xml) of an OOXML WordprocessingML document.
//
// Contract: C-21 in contracts.md
// Dependencies: xmltypes, wml/rpr, wml/ppr, wml/table, wml/shared
package styles

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/ppr"
	"github.com/vortex/docx-go/wml/rpr"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_Styles is the root element of word/styles.xml (<w:styles>).
type CT_Styles struct {
	DocDefaults  *CT_DocDefaults
	LatentStyles *CT_LatentStyles
	Style        []CT_Style
	Extra        []shared.RawXML

	// Namespaces preserves the original xmlns declarations for round-trip.
	Namespaces []xml.Attr
}

// CT_DocDefaults contains default run and paragraph properties.
type CT_DocDefaults struct {
	RPrDefault *CT_RPrDefault
	PPrDefault *CT_PPrDefault
}

// CT_RPrDefault wraps default run properties.
type CT_RPrDefault struct {
	RPr *rpr.CT_RPr
}

// CT_PPrDefault wraps default paragraph properties.
type CT_PPrDefault struct {
	PPr *ppr.CT_PPrBase
}

// CT_LatentStyles contains latent style defaults and per-style exceptions.
type CT_LatentStyles struct {
	DefLockedState    *bool `xml:"defLockedState,attr,omitempty"`
	DefUIPriority     *int  `xml:"defUIPriority,attr,omitempty"`
	DefSemiHidden     *bool `xml:"defSemiHidden,attr,omitempty"`
	DefUnhideWhenUsed *bool `xml:"defUnhideWhenUsed,attr,omitempty"`
	DefQFormat        *bool `xml:"defQFormat,attr,omitempty"`
	Count             *int  `xml:"count,attr,omitempty"`
	LsdException      []CT_LsdException
}

// CT_LsdException overrides latent style defaults for a named style.
type CT_LsdException struct {
	Name           string `xml:"name,attr"`
	Locked         *bool  `xml:"locked,attr,omitempty"`
	UIPriority     *int   `xml:"uiPriority,attr,omitempty"`
	SemiHidden     *bool  `xml:"semiHidden,attr,omitempty"`
	UnhideWhenUsed *bool  `xml:"unhideWhenUsed,attr,omitempty"`
	QFormat        *bool  `xml:"qFormat,attr,omitempty"`
}

// CT_Style represents a single style definition.
// Children are in STRICT ORDER per XSD sequence (see patterns.md 2.8).
type CT_Style struct {
	// Attributes
	Type        string `xml:"type,attr"`
	Default     *bool  `xml:"default,attr,omitempty"`
	CustomStyle *bool  `xml:"customStyle,attr,omitempty"`
	StyleID     string `xml:"styleId,attr"`

	// Elements (STRICT ORDER — see styleFieldMap in marshal.go)
	Name            *xmltypes.CT_String
	Aliases         *xmltypes.CT_String
	BasedOn         *xmltypes.CT_String
	Next            *xmltypes.CT_String
	Link            *xmltypes.CT_String
	AutoRedefine    *xmltypes.CT_OnOff
	Hidden          *xmltypes.CT_OnOff
	UIpriority      *xmltypes.CT_DecimalNumber
	SemiHidden      *xmltypes.CT_OnOff
	UnhideWhenUsed  *xmltypes.CT_OnOff
	QFormat         *xmltypes.CT_OnOff
	Locked          *xmltypes.CT_OnOff
	Personal        *xmltypes.CT_OnOff
	PersonalCompose *xmltypes.CT_OnOff
	PersonalReply   *xmltypes.CT_OnOff
	Rsid            *xmltypes.CT_LongHexNumber
	PPr             *ppr.CT_PPrBase
	RPr             *rpr.CT_RPrBase
	TblPr           *table.CT_TblPr
	TrPr            *table.CT_TrPr
	TcPr            *table.CT_TcPr
	TblStylePr      []CT_TblStylePr
	Extra           []shared.RawXML
}

// CT_TblStylePr defines formatting overrides for conditional table regions.
type CT_TblStylePr struct {
	Type  string `xml:"type,attr"` // "firstRow"|"lastRow"|"band1Vert"|...
	PPr   *ppr.CT_PPrBase
	RPr   *rpr.CT_RPrBase
	TblPr *table.CT_TblPr
	TrPr  *table.CT_TrPr
	TcPr  *table.CT_TcPr
}
