package shared

import (
	"bytes"
	"encoding/xml"
)

// RawXML stores an unrecognised XML element verbatim so that it survives an
// unmarshal → marshal round-trip without data loss.
//
// During unmarshal, any element whose local name is not handled by the
// containing type's UnmarshalXML switch is decoded into a RawXML value.
// During marshal, the element is re-emitted in its original position.
//
// RawXML satisfies all three content interfaces so it can appear at any
// nesting level (block, paragraph-content, or run-content).
type RawXML struct {
	XMLName xml.Name
	Attrs   []xml.Attr `xml:",any,attr"`
	Inner   []byte     `xml:",innerxml"`
}

// Interface compliance.
func (RawXML) blockLevelElement() {}
func (RawXML) paragraphContent()  {}
func (RawXML) runContent()        {}

// MarshalXML writes the stored element back to the encoder, replaying all
// inner tokens so that encoding/xml produces well-formed output.
func (r RawXML) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Use the original element name and attributes, not whatever the caller
	// passed via `start`.
	st := xml.StartElement{
		Name: r.XMLName,
		Attr: r.Attrs,
	}
	if err := e.EncodeToken(st); err != nil {
		return err
	}

	// Replay inner content token-by-token through the encoder so that it
	// stays well-formed.
	if len(r.Inner) > 0 {
		dec := xml.NewDecoder(bytes.NewReader(r.Inner))
		for {
			tok, err := dec.Token()
			if err != nil {
				break // io.EOF or trailing garbage — stop replaying
			}
			if err := e.EncodeToken(xml.CopyToken(tok)); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(st.End())
}

// UnmarshalXML populates the RawXML from the decoder. encoding/xml fills
// XMLName, Attrs and Inner automatically thanks to the struct tags, but we
// provide an explicit implementation so we can handle edge-cases and make
// the intent clear.
func (r *RawXML) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	r.XMLName = start.Name

	// Copy attributes, but skip namespace declarations (xmlns:prefix="…"
	// and xmlns="…").  Go's xml.Decoder exposes them as Attr entries with
	// Name.Space == "xmlns" or Name.Local == "xmlns", but they are not
	// content attributes — they are namespace machinery.  The xml.Encoder
	// regenerates the required declarations from xml.Name.Space fields, so
	// keeping them would cause duplication on every round-trip.
	r.Attrs = make([]xml.Attr, 0, len(start.Attr))
	for _, a := range start.Attr {
		if a.Name.Space == "xmlns" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			continue
		}
		r.Attrs = append(r.Attrs, a)
	}

	// Collect the inner XML (everything between <start> and </start>).
	type raw struct {
		Inner []byte `xml:",innerxml"`
	}
	var v raw
	if err := d.DecodeElement(&v, &start); err != nil {
		return err
	}
	r.Inner = v.Inner
	return nil
}
