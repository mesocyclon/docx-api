// Package tracking implements OOXML track-change and annotation markup types
// defined in WML (ISO/IEC 29500-1 §17.13).
//
// Types provided:
//   - CT_RunTrackChange — wrapper for <w:ins>, <w:del>, <w:moveFrom>, <w:moveTo>
//   - CT_Markup         — base type for commentRangeStart/End, etc.
//   - CT_Bookmark       — for <w:bookmarkStart>
//   - CT_MarkupRange    — for <w:bookmarkEnd>
//   - CT_MoveBookmark   — for <w:moveFromRangeStart>, <w:moveToRangeStart>
package tracking

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// CT_RunTrackChange — <w:ins>, <w:del>, <w:moveFrom>, <w:moveTo>
// ---------------------------------------------------------------------------

// CT_RunTrackChange wraps tracked content such as insertions and deletions.
// Content holds the paragraph-level children (runs, etc.) that live inside
// the track-change element.  Items stored are concrete types that implement
// shared.ParagraphContent (or shared.RawXML as a fallback).
type CT_RunTrackChange struct {
	ID      int           `xml:"id,attr"`
	Author  string        `xml:"author,attr"`
	Date    *string       `xml:"date,attr,omitempty"`
	Content []interface{} // elements inside ins/del (ParagraphContent items)
}

// MarshalXML writes the CT_RunTrackChange and its children.
// The caller is responsible for setting start.Name (e.g. w:ins, w:del).
func (tc *CT_RunTrackChange) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Ensure namespace on the element.
	if start.Name.Space == "" {
		start.Name.Space = xmltypes.NSw
	}

	// Attributes: w:id, w:author, w:date
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "id", strconv.Itoa(tc.ID))
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "author", tc.Author)
	if tc.Date != nil {
		start.Attr = appendAttr(start.Attr, xmltypes.NSw, "date", *tc.Date)
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Encode children.
	for _, item := range tc.Content {
		switch v := item.(type) {
		case shared.RawXML:
			if err := encodeRawXML(e, v); err != nil {
				return err
			}
		case xml.Marshaler:
			// The concrete type (e.g. a run) knows its own start element.
			// We encode it without an explicit StartElement; EncodeElement will
			// call MarshalXML on the value which provides its own element name.
			if err := e.EncodeElement(v, xml.StartElement{}); err != nil {
				return err
			}
		default:
			// Best-effort: let encoding/xml figure it out.
			if err := e.EncodeElement(v, xml.StartElement{}); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads attributes and child elements from the decoder.
// Children are resolved via shared.CreateParagraphContent; unrecognised
// elements are preserved as shared.RawXML for lossless round-trip.
func (tc *CT_RunTrackChange) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: bad id %q: %w", attr.Value, err)
			}
			tc.ID = id
		case "author":
			tc.Author = attr.Value
		case "date":
			s := attr.Value
			tc.Date = &s
		}
	}

	// Parse children.
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			// Try the factory first.
			if el := shared.CreateParagraphContent(t.Name); el != nil {
				if err := d.DecodeElement(el, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, el)
			} else {
				// Unknown element → preserve as RawXML.
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// CT_Markup — base marker with just an ID (commentRangeStart/End, etc.)
// ---------------------------------------------------------------------------

// CT_Markup is a self-closing element carrying only a w:id attribute.
// Used for <w:commentRangeStart>, <w:commentRangeEnd>, <w:commentReference>.
type CT_Markup struct {
	ID int `xml:"id,attr"`
}

// MarshalXML writes CT_Markup as a self-closing element with w:id.
func (m *CT_Markup) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if start.Name.Space == "" {
		start.Name.Space = xmltypes.NSw
	}
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "id", strconv.Itoa(m.ID))
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads the w:id attribute and skips any content.
func (m *CT_Markup) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: CT_Markup bad id %q: %w", attr.Value, err)
			}
			m.ID = id
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_Bookmark — <w:bookmarkStart w:id="0" w:name="_Toc"/>
// ---------------------------------------------------------------------------

// CT_Bookmark represents a bookmark start marker.
type CT_Bookmark struct {
	ID       int    `xml:"id,attr"`
	Name     string `xml:"name,attr"`
	ColFirst *int   `xml:"colFirst,attr,omitempty"`
	ColLast  *int   `xml:"colLast,attr,omitempty"`
}

// MarshalXML writes CT_Bookmark as a self-closing element.
func (b *CT_Bookmark) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if start.Name.Space == "" {
		start.Name.Space = xmltypes.NSw
	}
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "id", strconv.Itoa(b.ID))
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "name", b.Name)
	if b.ColFirst != nil {
		start.Attr = appendAttr(start.Attr, xmltypes.NSw, "colFirst", strconv.Itoa(*b.ColFirst))
	}
	if b.ColLast != nil {
		start.Attr = appendAttr(start.Attr, xmltypes.NSw, "colLast", strconv.Itoa(*b.ColLast))
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads bookmark attributes and skips content.
func (b *CT_Bookmark) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: CT_Bookmark bad id %q: %w", attr.Value, err)
			}
			b.ID = id
		case "name":
			b.Name = attr.Value
		case "colFirst":
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: CT_Bookmark bad colFirst %q: %w", attr.Value, err)
			}
			b.ColFirst = &v
		case "colLast":
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: CT_Bookmark bad colLast %q: %w", attr.Value, err)
			}
			b.ColLast = &v
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_MarkupRange — <w:bookmarkEnd w:id="0"/>
// ---------------------------------------------------------------------------

// CT_MarkupRange is a self-closing element that ends a range (bookmark, permission, …).
type CT_MarkupRange struct {
	ID int `xml:"id,attr"`
}

// MarshalXML writes CT_MarkupRange as a self-closing element.
func (mr *CT_MarkupRange) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if start.Name.Space == "" {
		start.Name.Space = xmltypes.NSw
	}
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "id", strconv.Itoa(mr.ID))
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads the w:id attribute and skips content.
func (mr *CT_MarkupRange) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: CT_MarkupRange bad id %q: %w", attr.Value, err)
			}
			mr.ID = id
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_MoveBookmark — <w:moveFromRangeStart w:id="3" w:author="X" w:name="move1"/>
// ---------------------------------------------------------------------------

// CT_MoveBookmark is a bookmark-like marker used for move tracking.
type CT_MoveBookmark struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Name   string `xml:"name,attr"`
}

// MarshalXML writes CT_MoveBookmark as a self-closing element.
func (mb *CT_MoveBookmark) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if start.Name.Space == "" {
		start.Name.Space = xmltypes.NSw
	}
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "id", strconv.Itoa(mb.ID))
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "author", mb.Author)
	start.Attr = appendAttr(start.Attr, xmltypes.NSw, "name", mb.Name)
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads move-bookmark attributes and skips content.
func (mb *CT_MoveBookmark) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			id, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("tracking: CT_MoveBookmark bad id %q: %w", attr.Value, err)
			}
			mb.ID = id
		case "author":
			mb.Author = attr.Value
		case "name":
			mb.Name = attr.Value
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// appendAttr adds a namespaced attribute to the slice.
func appendAttr(attrs []xml.Attr, space, local, value string) []xml.Attr {
	return append(attrs, xml.Attr{
		Name:  xml.Name{Space: space, Local: local},
		Value: value,
	})
}

// encodeRawXML replays a RawXML element into the encoder, preserving its
// inner structure by re-parsing the stored inner bytes as XML tokens.
func encodeRawXML(e *xml.Encoder, raw shared.RawXML) error {
	start := xml.StartElement{Name: raw.XMLName, Attr: raw.Attrs}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// Re-parse inner XML tokens and replay them.
	if len(raw.Inner) > 0 {
		dec := xml.NewDecoder(bytes.NewReader(raw.Inner))
		for {
			tok, err := dec.Token()
			if err != nil {
				break // io.EOF or malformed — stop
			}
			if err := e.EncodeToken(xml.CopyToken(tok)); err != nil {
				return err
			}
		}
	}
	return e.EncodeToken(start.End())
}
