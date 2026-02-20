// Package fonts implements the fontTable part (word/fontTable.xml) of an
// OOXML WordprocessingML package. It corresponds to contract C-24.
//
// Dependencies: xmltypes, wml/shared.
package fonts

import (
	"github.com/vortex/docx-go/wml/shared"
)

// ============================================================
// Field mapping — element order within CT_Font (xsd:sequence)
// ============================================================

// fieldMapping describes one struct field → XML local-name pair.
type fieldMapping struct {
	GoField  string // field name in the Go struct
	XMLLocal string // local element name in w: namespace
}

// fontChildOrder defines the strict XSD sequence for children of <w:font>.
var fontChildOrder = []fieldMapping{
	{"Panose1", "panose1"},
	{"Charset", "charset"},
	{"Family", "family"},
	{"Pitch", "pitch"},
	{"Sig", "sig"},
	{"EmbedRegular", "embedRegular"},
	{"EmbedBold", "embedBold"},
	{"EmbedItalic", "embedItalic"},
	{"EmbedBoldItalic", "embedBoldItalic"},
}

// ============================================================
// Root type
// ============================================================

// CT_FontsList is the root element of fontTable.xml (<w:fonts>).
type CT_FontsList struct {
	Font []CT_Font
}

// ============================================================
// CT_Font — one font definition
// ============================================================

// CT_Font represents a single <w:font> element within the font table.
type CT_Font struct {
	Name            string          `xml:"name,attr"`
	Panose1         *CT_Panose      // <w:panose1>
	Charset         *CT_Charset     // <w:charset>
	Family          *CT_FontFamily  // <w:family>
	Pitch           *CT_Pitch       // <w:pitch>
	Sig             *CT_FontSig     // <w:sig>
	EmbedRegular    *CT_FontRel     // <w:embedRegular>
	EmbedBold       *CT_FontRel     // <w:embedBold>
	EmbedItalic     *CT_FontRel     // <w:embedItalic>
	EmbedBoldItalic *CT_FontRel     // <w:embedBoldItalic>
	Extra           []shared.RawXML // unknown extension elements (round-trip)
}

// ============================================================
// Simple child types (attributes-only, self-closing elements)
// ============================================================

// CT_Panose represents <w:panose1 w:val="020F0502020204030204"/>.
type CT_Panose struct {
	Val string `xml:"val,attr"`
}

// CT_Charset represents <w:charset w:val="00"/>.
type CT_Charset struct {
	Val string `xml:"val,attr"`
}

// CT_FontFamily represents <w:family w:val="swiss"/>.
// Allowed values: roman, swiss, modern, script, decorative, auto.
type CT_FontFamily struct {
	Val string `xml:"val,attr"`
}

// CT_Pitch represents <w:pitch w:val="variable"/>.
// Allowed values: fixed, variable, default.
type CT_Pitch struct {
	Val string `xml:"val,attr"`
}

// CT_FontSig represents <w:sig> with USB and CSB bitmask attributes.
type CT_FontSig struct {
	Usb0 string `xml:"usb0,attr"`
	Usb1 string `xml:"usb1,attr"`
	Usb2 string `xml:"usb2,attr"`
	Usb3 string `xml:"usb3,attr"`
	Csb0 string `xml:"csb0,attr"`
	Csb1 string `xml:"csb1,attr"`
}

// CT_FontRel represents an embedded font reference (<w:embedRegular>, etc.).
// The ID attribute is in the relationships namespace (r:id).
type CT_FontRel struct {
	FontKey    *string `xml:"fontKey,attr,omitempty"`
	SubsetInfo *string `xml:"subsetted,attr,omitempty"`
	ID         string  `xml:"id,attr"` // r:id
}
