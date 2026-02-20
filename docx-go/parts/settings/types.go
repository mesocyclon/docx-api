// Package settings implements the document settings part (word/settings.xml).
//
// CT_Settings contains ~253 elements in the XSD. This package types ~20 of
// the most frequently used fields and stores the remaining 230+ elements as
// shared.RawXML for lossless round-trip fidelity.
//
// The original element order is tracked internally so that marshalling
// reproduces the input sequence, which is critical because the XSD defines
// settings as an xsd:sequence.
package settings

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_Settings is the root type for word/settings.xml (<w:settings>).
type CT_Settings struct {
	// Typed fields (frequently used).
	WriteProtection         *CT_WriteProtection
	Zoom                    *CT_Zoom
	ProofState              *CT_Proof
	DefaultTabStop          *xmltypes.CT_TwipsMeasure
	CharacterSpacingControl *xmltypes.CT_String
	EvenAndOddHeaders       *xmltypes.CT_OnOff
	MirrorMargins           *xmltypes.CT_OnOff
	TrackRevisions          *xmltypes.CT_OnOff
	DoNotTrackMoves         *xmltypes.CT_OnOff
	DoNotTrackFormatting    *xmltypes.CT_OnOff
	DocumentProtection      *CT_DocProtect
	Compat                  *CT_Compat
	Rsids                   *CT_DocRsids
	MathPr                  *shared.RawXML
	ThemeFontLang           *CT_ThemeFontLang
	ClrSchemeMapping        *CT_ClrSchemeMapping
	ShapeDefaults           *shared.RawXML
	DecimalSymbol           *xmltypes.CT_String
	ListSeparator           *xmltypes.CT_String
	DocId14                 *xmltypes.CT_LongHexNumber // w14:docId
	DocId15                 *xmltypes.CT_Guid          // w15:docId

	// All remaining 230+ elements are stored verbatim.
	Extra []shared.RawXML

	// Saved namespace declarations from the root <w:settings> element so
	// that a round-trip preserves every xmlns:* the original had.
	Namespaces []xml.Attr

	// elementOrder tracks the original sequence of child elements so that
	// marshal can reproduce the exact input order.  Each entry is either a
	// known-field key (e.g. "zoom", "compat") or a synthetic key for an
	// Extra element (e.g. "#extra:0", "#extra:1").
	elementOrder []string
}

// CT_Zoom represents <w:zoom>.
type CT_Zoom struct {
	Percent int    `xml:"percent,attr"`
	Val     string `xml:"val,attr,omitempty"`
}

// CT_Proof represents <w:proofState>.
type CT_Proof struct {
	Spelling *string `xml:"spelling,attr,omitempty"`
	Grammar  *string `xml:"grammar,attr,omitempty"`
}

// CT_Compat represents <w:compat>.
type CT_Compat struct {
	CompatSetting []CT_CompatSetting
	Extra         []shared.RawXML
}

// CT_CompatSetting represents <w:compatSetting>.
type CT_CompatSetting struct {
	Name string `xml:"name,attr"`
	URI  string `xml:"uri,attr"`
	Val  string `xml:"val,attr"`
}

// CT_DocRsids represents <w:rsids>.
type CT_DocRsids struct {
	RsidRoot *xmltypes.CT_LongHexNumber
	Rsid     []xmltypes.CT_LongHexNumber
}

// CT_ThemeFontLang represents <w:themeFontLang>.
type CT_ThemeFontLang struct {
	Val      *string `xml:"val,attr,omitempty"`
	EastAsia *string `xml:"eastAsia,attr,omitempty"`
	Bidi     *string `xml:"bidi,attr,omitempty"`
}

// CT_ClrSchemeMapping represents <w:clrSchemeMapping>.
type CT_ClrSchemeMapping struct {
	Bg1               string `xml:"bg1,attr"`
	T1                string `xml:"t1,attr"`
	Bg2               string `xml:"bg2,attr"`
	T2                string `xml:"t2,attr"`
	Accent1           string `xml:"accent1,attr"`
	Accent2           string `xml:"accent2,attr"`
	Accent3           string `xml:"accent3,attr"`
	Accent4           string `xml:"accent4,attr"`
	Accent5           string `xml:"accent5,attr"`
	Accent6           string `xml:"accent6,attr"`
	Hyperlink         string `xml:"hyperlink,attr"`
	FollowedHyperlink string `xml:"followedHyperlink,attr"`
}

// CT_WriteProtection represents <w:writeProtection>.
type CT_WriteProtection struct {
	Recommended   *bool   `xml:"recommended,attr,omitempty"`
	AlgorithmName *string `xml:"algorithmName,attr,omitempty"`
	HashValue     *string `xml:"hashValue,attr,omitempty"`
	SaltValue     *string `xml:"saltValue,attr,omitempty"`
	SpinCount     *int    `xml:"spinCount,attr,omitempty"`
}

// CT_DocProtect represents <w:documentProtection>.
type CT_DocProtect struct {
	Edit          *string `xml:"edit,attr,omitempty"`
	Enforcement   *bool   `xml:"enforcement,attr,omitempty"`
	AlgorithmName *string `xml:"algorithmName,attr,omitempty"`
	HashValue     *string `xml:"hashValue,attr,omitempty"`
	SaltValue     *string `xml:"saltValue,attr,omitempty"`
	SpinCount     *int    `xml:"spinCount,attr,omitempty"`
}
