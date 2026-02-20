// Package run implements the WML run element (CT_R) and its inline content
// types. A run is the fundamental unit of text content inside a paragraph.
//
// Contract: C-15 in contracts.md.
// Dependencies: xmltypes, wml/rpr, wml/shared, dml.
package run

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
)

// ---------------------------------------------------------------------------
// CT_Text — <w:t> and <w:delText>
// ---------------------------------------------------------------------------

// CT_Text represents a text run content element (<w:t> or <w:delText>).
// It automatically handles xml:space="preserve" for strings with leading or
// trailing whitespace.
type CT_Text struct {
	shared.RunContentMarker
	Space *string `xml:"space,attr,omitempty"` // xml:space
	Value string  `xml:",chardata"`
}

// NewText creates a CT_Text, automatically setting xml:space="preserve" when
// the value contains leading/trailing whitespace.
func NewText(s string) CT_Text {
	t := CT_Text{Value: s}
	if len(s) > 0 && (s[0] == ' ' || s[len(s)-1] == ' ' || s[0] == '\t') {
		preserve := "preserve"
		t.Space = &preserve
	}
	return t
}

const nsXMLSpace = "http://www.w3.org/XML/1998/namespace"

// MarshalXML writes the text element with an optional xml:space attribute.
func (t *CT_Text) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if t.Space != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: nsXMLSpace, Local: "space"},
			Value: *t.Space,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := e.EncodeToken(xml.CharData(t.Value)); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads a text element, extracting xml:space and chardata.
func (t *CT_Text) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "space" {
			s := attr.Value
			t.Space = &s
		}
	}
	type plain struct {
		Value string `xml:",chardata"`
	}
	var v plain
	if err := d.DecodeElement(&v, &start); err != nil {
		return err
	}
	t.Value = v.Value
	return nil
}

// ---------------------------------------------------------------------------
// CT_Br — <w:br>
// ---------------------------------------------------------------------------

// CT_Br represents a break element. Type specifies page/column/textWrapping;
// Clear specifies none/left/right/all.
type CT_Br struct {
	shared.RunContentMarker
	Type  *string `xml:"type,attr,omitempty"`
	Clear *string `xml:"clear,attr,omitempty"`
}

// ---------------------------------------------------------------------------
// CT_Drawing — <w:drawing>
// ---------------------------------------------------------------------------

// CT_Drawing represents a drawing container that holds inline and anchored
// images. For MVP, inline and anchor children are stored as raw XML via
// shared.RawXML to preserve full round-trip fidelity while the full dml
// module is being developed.
type CT_Drawing struct {
	shared.RunContentMarker
	Children []shared.RawXML // <wp:inline> and <wp:anchor> elements
}

// MarshalXML writes <w:drawing> with all child elements.
func (dr *CT_Drawing) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, ch := range dr.Children {
		if err := e.EncodeElement(ch, xml.StartElement{Name: ch.XMLName}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:drawing> children (wp:inline, wp:anchor, etc.) as
// raw XML for round-trip preservation.
func (dr *CT_Drawing) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			var raw shared.RawXML
			if err := d.DecodeElement(&raw, &t); err != nil {
				return err
			}
			dr.Children = append(dr.Children, raw)
		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// CT_FldChar — <w:fldChar>
// ---------------------------------------------------------------------------

// CT_FldChar represents a field character (begin, separate, end).
type CT_FldChar struct {
	shared.RunContentMarker
	FldCharType string `xml:"fldCharType,attr"`
	FldLock     *bool  `xml:"fldLock,attr,omitempty"`
	Dirty       *bool  `xml:"dirty,attr,omitempty"`
}

// ---------------------------------------------------------------------------
// CT_InstrText — <w:instrText>
// ---------------------------------------------------------------------------

// CT_InstrText represents a field instruction text element.
type CT_InstrText struct {
	shared.RunContentMarker
	Space *string `xml:"space,attr,omitempty"` // xml:space
	Value string  `xml:",chardata"`
}

// MarshalXML writes the instrText element with optional xml:space.
func (t *CT_InstrText) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if t.Space != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: nsXMLSpace, Local: "space"},
			Value: *t.Space,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := e.EncodeToken(xml.CharData(t.Value)); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads instrText with xml:space and chardata.
func (t *CT_InstrText) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "space" {
			s := attr.Value
			t.Space = &s
		}
	}
	type plain struct {
		Value string `xml:",chardata"`
	}
	var v plain
	if err := d.DecodeElement(&v, &start); err != nil {
		return err
	}
	t.Value = v.Value
	return nil
}

// ---------------------------------------------------------------------------
// CT_Sym — <w:sym>
// ---------------------------------------------------------------------------

// CT_Sym represents a symbol character.
type CT_Sym struct {
	shared.RunContentMarker
	Font string `xml:"font,attr"`
	Char string `xml:"char,attr"`
}

// ---------------------------------------------------------------------------
// CT_FtnEdnRef — <w:footnoteReference> / <w:endnoteReference>
// ---------------------------------------------------------------------------

// CT_FtnEdnRef represents a footnote or endnote reference.
type CT_FtnEdnRef struct {
	shared.RunContentMarker
	XMLName xml.Name // captures whether it is footnoteReference or endnoteReference
	ID      int      `xml:"id,attr"`
}

// ---------------------------------------------------------------------------
// CT_EmptyRunContent — empty elements: tab, cr, pgNum, noBreakHyphen,
// softHyphen, footnoteRef, endnoteRef, annotationRef, separator,
// continuationSeparator, lastRenderedPageBreak, etc.
// ---------------------------------------------------------------------------

// CT_EmptyRunContent represents a self-closing run content element.
type CT_EmptyRunContent struct {
	shared.RunContentMarker
	XMLName xml.Name
}

// emptyElementNames lists all known empty run-content element local names.
var emptyElementNames = map[string]bool{
	"tab":                   true,
	"cr":                    true,
	"pgNum":                 true,
	"noBreakHyphen":         true,
	"softHyphen":            true,
	"footnoteRef":           true,
	"endnoteRef":            true,
	"annotationRef":         true,
	"separator":             true,
	"continuationSeparator": true,
	"lastRenderedPageBreak": true,
	"dayShort":              true,
	"monthShort":            true,
	"yearShort":             true,
	"dayLong":               true,
	"monthLong":             true,
	"yearLong":              true,
}

// ---------------------------------------------------------------------------
// CT_RawRunContent — catch-all for unrecognised run content elements.
// ---------------------------------------------------------------------------

// CT_RawRunContent preserves an unknown run-content element for round-trip
// fidelity.
type CT_RawRunContent struct {
	shared.RunContentMarker
	Raw shared.RawXML
}

// Compile-time interface checks.
var (
	_ shared.RunContent = (*CT_Text)(nil)
	_ shared.RunContent = (*CT_Br)(nil)
	_ shared.RunContent = (*CT_Drawing)(nil)
	_ shared.RunContent = (*CT_FldChar)(nil)
	_ shared.RunContent = (*CT_InstrText)(nil)
	_ shared.RunContent = (*CT_Sym)(nil)
	_ shared.RunContent = (*CT_FtnEdnRef)(nil)
	_ shared.RunContent = (*CT_EmptyRunContent)(nil)
	_ shared.RunContent = (*CT_RawRunContent)(nil)

	_ xml.Marshaler   = (*CT_Text)(nil)
	_ xml.Unmarshaler = (*CT_Text)(nil)
	_ xml.Marshaler   = (*CT_InstrText)(nil)
	_ xml.Unmarshaler = (*CT_InstrText)(nil)
	_ xml.Marshaler   = (*CT_Drawing)(nil)
	_ xml.Unmarshaler = (*CT_Drawing)(nil)
)
