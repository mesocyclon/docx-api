// Package footnotes implements parsing and serialization of footnotes.xml and
// endnotes.xml parts (CT_Footnotes / CT_FtnEdn) in an OOXML word-processing
// document.
//
// Both footnotes and endnotes share the same schema; the only difference is
// the root element name (w:footnotes vs w:endnotes) and the child element
// name (w:footnote vs w:endnote). The exported types handle both cases.
//
// See contracts.md C-26 and reference-appendix.md §footnotes.
package footnotes

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// CT_Footnotes — root element of footnotes.xml (or endnotes.xml)
// ---------------------------------------------------------------------------

// CT_Footnotes represents the <w:footnotes> (or <w:endnotes>) root element.
type CT_Footnotes struct {
	Footnote []CT_FtnEdn

	// rootLocal is the local name of the root XML element that was parsed.
	// It is "footnotes" for footnotes.xml and "endnotes" for endnotes.xml.
	// When zero-valued, MarshalXML defaults to "footnotes".
	rootLocal string

	// childLocal is the local name of each child element.
	// It is "footnote" or "endnote".  Defaults to "footnote".
	childLocal string

	// namespaces preserves the original xmlns:* declarations on the root
	// element so they survive a round-trip without loss.
	// Attrs are stored in normalised form: {Local: "xmlns:w"} not
	// {Space: "xmlns", Local: "w"}, because Go's xml.Encoder would mangle
	// the latter into "_xmlns:w".
	namespaces []xml.Attr
}

// UnmarshalXML implements xml.Unmarshaler for CT_Footnotes.
func (fn *CT_Footnotes) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Remember root element name ("footnotes" or "endnotes").
	fn.rootLocal = start.Name.Local
	switch fn.rootLocal {
	case "footnotes":
		fn.childLocal = "footnote"
	case "endnotes":
		fn.childLocal = "endnote"
	default:
		fn.childLocal = "footnote"
	}

	// Preserve namespace declarations for round-trip fidelity.
	// Go's xml.Decoder exposes xmlns:prefix="…" as Attr with
	// Name.Space == "xmlns".  We normalise to {Local: "xmlns:prefix"}
	// so that the xml.Encoder writes them literally on output instead of
	// mangling them into "_xmlns:prefix".
	fn.namespaces = make([]xml.Attr, 0, len(start.Attr))
	for _, a := range start.Attr {
		if a.Name.Space == "xmlns" {
			// xmlns:prefix="uri" → {Local: "xmlns:prefix", Value: uri}
			fn.namespaces = append(fn.namespaces, xml.Attr{
				Name:  xml.Name{Local: "xmlns:" + a.Name.Local},
				Value: a.Value,
			})
		} else if a.Name.Space == "" && a.Name.Local == "xmlns" {
			// default namespace xmlns="uri"
			fn.namespaces = append(fn.namespaces, a)
		}
		// Non-namespace attributes on the root element are intentionally
		// dropped — the OOXML spec defines none.
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == fn.childLocal {
				var fe CT_FtnEdn
				if err := fe.unmarshal(d, t); err != nil {
					return err
				}
				fn.Footnote = append(fn.Footnote, fe)
			} else {
				// Unknown child — skip.
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML implements xml.Marshaler for CT_Footnotes.
func (fn CT_Footnotes) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	local := fn.rootLocal
	if local == "" {
		local = "footnotes"
	}
	childLocal := fn.childLocal
	if childLocal == "" {
		childLocal = "footnote"
	}

	start := xml.StartElement{
		Name: xml.Name{Local: "w:" + local},
	}

	if len(fn.namespaces) > 0 {
		start.Attr = fn.namespaces
	} else {
		start.Attr = defaultNamespaces()
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for i := range fn.Footnote {
		if err := fn.Footnote[i].marshal(e, childLocal); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// defaultNamespaces returns the minimal set of xmlns declarations for a
// newly-created footnotes/endnotes root element.
func defaultNamespaces() []xml.Attr {
	return []xml.Attr{
		{Name: xml.Name{Local: "xmlns:w"}, Value: xmltypes.NSw},
	}
}

// ---------------------------------------------------------------------------
// CT_FtnEdn — a single footnote or endnote entry
// ---------------------------------------------------------------------------

// CT_FtnEdn represents a <w:footnote> or <w:endnote> element.
type CT_FtnEdn struct {
	Type    *string `xml:"type,attr,omitempty"` // "normal"|"separator"|"continuationSeparator"
	ID      int     `xml:"id,attr"`
	Content []shared.BlockLevelElement
}

// unmarshal decodes a <w:footnote>/<w:endnote> element from the decoder.
func (fe *CT_FtnEdn) unmarshal(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes (w:type, w:id).
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "type":
			v := attr.Value
			fe.Type = &v
		case "id":
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("footnotes: invalid id %q: %w", attr.Value, err)
			}
			fe.ID = id
		}
	}

	// Parse child elements — block-level content (paragraphs, tables, …).
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if el := shared.CreateBlockElement(t.Name); el != nil {
				if err := d.DecodeElement(el, &t); err != nil {
					return err
				}
				fe.Content = append(fe.Content, el)
			} else {
				// Unknown element → preserve as RawXML for round-trip.
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				fe.Content = append(fe.Content, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// marshal encodes a <w:footnote>/<w:endnote> element into the encoder.
func (fe *CT_FtnEdn) marshal(e *xml.Encoder, childLocal string) error {
	start := xml.StartElement{
		Name: xml.Name{Local: "w:" + childLocal},
	}

	if fe.Type != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Local: "w:type"},
			Value: *fe.Type,
		})
	}
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Local: "w:id"},
		Value: strconv.Itoa(fe.ID),
	})

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, bl := range fe.Content {
		switch v := bl.(type) {
		case shared.RawXML:
			if err := encodeRawXML(e, v); err != nil {
				return err
			}
		default:
			if m, ok := bl.(xml.Marshaler); ok {
				if err := m.MarshalXML(e, xml.StartElement{}); err != nil {
					return err
				}
			}
		}
	}

	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// RawXML serialisation helpers
// ---------------------------------------------------------------------------

// encodeRawXML writes a shared.RawXML element through the encoder without
// replaying inner tokens through a fresh xml.Decoder.
//
// The standard token-replay approach (create a new xml.Decoder over Inner
// bytes, feed each token to the Encoder) is broken because the new Decoder
// lacks the parent's namespace context.  Prefixed names like "w:pPr" inside
// Inner resolve to {Space: "w", Local: "pPr"} (literal prefix as URI) and
// the Encoder then emits spurious xmlns="w" declarations that accumulate on
// every round-trip.
//
// Instead we use EncodeElement with a ",innerxml" struct field: the encoder
// writes the start/end tags (with correct prefix) and injects the Inner
// bytes verbatim between them — no re-parsing, no namespace corruption.
func encodeRawXML(e *xml.Encoder, r shared.RawXML) error {
	// Build the StartElement with prefixed names so the encoder writes
	// them literally (e.g. "w:p") without trying to declare namespaces.
	start := xml.StartElement{
		Name: xml.Name{Local: nsToPrefix(r.XMLName.Space, r.XMLName.Local)},
	}
	for _, a := range r.Attrs {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Local: nsToPrefix(a.Name.Space, a.Name.Local)},
			Value: a.Value,
		})
	}

	// The encoder writes <w:p attrs...>{Inner bytes verbatim}</w:p>.
	type rawContent struct {
		Inner []byte `xml:",innerxml"`
	}
	return e.EncodeElement(rawContent{Inner: r.Inner}, start)
}

// nsToPrefix builds a "prefix:local" string from an xml.Name.
// Known namespace URIs are mapped to their conventional prefix; short strings
// (like "w") that appear when inner bytes are re-parsed without context are
// passed through as-is.
func nsToPrefix(space, local string) string {
	if space == "" {
		return local
	}
	switch space {
	case xmltypes.NSw:
		return "w:" + local
	default:
		// For unresolved prefix strings ("w", "r", …) from innerxml
		// re-parsing, or other namespace URIs, use the space directly.
		return space + ":" + local
	}
}

// ---------------------------------------------------------------------------
// Parse / Serialize — top-level API
// ---------------------------------------------------------------------------

// Parse deserializes XML bytes into a CT_Footnotes structure.
func Parse(data []byte) (*CT_Footnotes, error) {
	var fn CT_Footnotes
	if err := xml.Unmarshal(data, &fn); err != nil {
		return nil, fmt.Errorf("footnotes.Parse: %w", err)
	}
	return &fn, nil
}

// Serialize serializes a CT_Footnotes structure back to XML bytes.
// The output includes the XML declaration.
//
// No indentation is applied because RawXML inner bytes are written verbatim:
// the encoder's indent whitespace before end-tags would be captured as part
// of Inner on re-parse, accumulating on every round-trip.  Compact output
// guarantees stable round-trips.
func Serialize(fn *CT_Footnotes) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)

	enc := xml.NewEncoder(&buf)

	if err := enc.Encode(fn); err != nil {
		return nil, fmt.Errorf("footnotes.Serialize: %w", err)
	}
	if err := enc.Flush(); err != nil {
		return nil, fmt.Errorf("footnotes.Serialize: %w", err)
	}

	return buf.Bytes(), nil
}
