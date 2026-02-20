// Package comments implements parsing and serialization of the
// word/comments.xml part of an OOXML document.
//
// Contract: C-25 in contracts.md
// Dependencies: wml/shared only
package comments

import (
	"bytes"
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ============================================================
// Types
// ============================================================

// CT_Comments is the root element of word/comments.xml (<w:comments>).
type CT_Comments struct {
	Comment []CT_Comment

	// Namespaces preserves all xmlns declarations from the original document
	// for faithful round-trip serialization.
	Namespaces []xml.Attr
}

// CT_Comment represents a single <w:comment> element.
type CT_Comment struct {
	ID       int    `xml:"id,attr"`
	Author   string `xml:"author,attr"`
	Date     string `xml:"date,attr,omitempty"`
	Initials string `xml:"initials,attr,omitempty"`

	// Content holds the block-level children of the comment (paragraphs,
	// tables, etc.). Elements are created via shared.CreateBlockElement;
	// unrecognized elements are stored as shared.RawXML.
	Content []shared.BlockLevelElement
}

// ============================================================
// Parse / Serialize
// ============================================================

// Parse deserializes XML bytes into a CT_Comments structure.
func Parse(data []byte) (*CT_Comments, error) {
	var c CT_Comments
	if err := xml.Unmarshal(data, &c); err != nil {
		return nil, fmt.Errorf("comments: parse: %w", err)
	}
	return &c, nil
}

// Serialize serializes a CT_Comments structure back to XML bytes,
// including the XML declaration header.
func Serialize(c *CT_Comments) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)

	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")
	if err := enc.Encode(c); err != nil {
		return nil, fmt.Errorf("comments: serialize: %w", err)
	}
	if err := enc.Flush(); err != nil {
		return nil, fmt.Errorf("comments: serialize flush: %w", err)
	}
	return buf.Bytes(), nil
}

// ============================================================
// CT_Comments — custom XML marshalling
// ============================================================

// MarshalXML implements xml.Marshaler for CT_Comments.
// Emits <w:comments xmlns:w="…"> with preserved namespace declarations.
func (c *CT_Comments) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:comments"}

	if len(c.Namespaces) > 0 {
		start.Attr = make([]xml.Attr, len(c.Namespaces))
		copy(start.Attr, c.Namespaces)
	} else {
		start.Attr = defaultNamespaces()
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for i := range c.Comment {
		elemStart := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "comment"},
		}
		if err := e.EncodeElement(&c.Comment[i], elemStart); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Comments.
func (c *CT_Comments) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Preserve all namespace declarations from the root element.
	c.Namespaces = make([]xml.Attr, len(start.Attr))
	copy(c.Namespaces, start.Attr)

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "comment" {
				var cm CT_Comment
				if err := d.DecodeElement(&cm, &t); err != nil {
					return err
				}
				c.Comment = append(c.Comment, cm)
			} else {
				// Skip unknown children at the comments level.
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ============================================================
// CT_Comment — custom XML marshalling
// ============================================================

// MarshalXML implements xml.Marshaler for CT_Comment.
// Emits attributes in the w: namespace and marshals block-level content.
func (cm *CT_Comment) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Attributes — must be in w: namespace.
	start.Attr = append(start.Attr, xml.Attr{
		Name: xml.Name{Space: xmltypes.NSw, Local: "id"}, Value: fmt.Sprintf("%d", cm.ID),
	})
	start.Attr = append(start.Attr, xml.Attr{
		Name: xml.Name{Space: xmltypes.NSw, Local: "author"}, Value: cm.Author,
	})
	if cm.Date != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "date"}, Value: cm.Date,
		})
	}
	if cm.Initials != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "initials"}, Value: cm.Initials,
		})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Marshal block-level content.
	for _, el := range cm.Content {
		if err := marshalBlockElement(e, el); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler for CT_Comment.
// Reads attributes and then decodes block-level child elements.
func (cm *CT_Comment) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			n, err := parseInt(attr.Value)
			if err != nil {
				return fmt.Errorf("comments: bad id %q: %w", attr.Value, err)
			}
			cm.ID = n
		case "author":
			cm.Author = attr.Value
		case "date":
			cm.Date = attr.Value
		case "initials":
			cm.Initials = attr.Value
		}
	}

	// Parse child elements (block-level content).
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			el := shared.CreateBlockElement(t.Name)
			if el != nil {
				// A registered factory recognized this element.
				if err := d.DecodeElement(el, &t); err != nil {
					return err
				}
				cm.Content = append(cm.Content, el)
			} else {
				// Unknown element → preserve as RawXML for round-trip.
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				cm.Content = append(cm.Content, raw)
			}

		case xml.EndElement:
			return nil
		}
	}
}

// ============================================================
// Helpers
// ============================================================

// marshalBlockElement marshals a single BlockLevelElement.
// For RawXML, it replays the captured tokens. For typed elements,
// it delegates to EncodeElement with the element's original name.
func marshalBlockElement(e *xml.Encoder, el shared.BlockLevelElement) error {
	switch v := el.(type) {
	case shared.RawXML:
		return marshalRawXML(e, v)
	default:
		// Typed element — the element knows how to marshal itself.
		// We need to derive a start element. If the element implements
		// xml.Marshaler, EncodeElement will use it.
		return e.EncodeElement(v, xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "unknownBlock"}})
	}
}

// marshalRawXML faithfully replays a RawXML element to the encoder.
func marshalRawXML(e *xml.Encoder, raw shared.RawXML) error {
	start := xml.StartElement{Name: raw.XMLName, Attr: raw.Attrs}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Replay inner XML tokens.
	if len(raw.Inner) > 0 {
		if err := e.Flush(); err != nil {
			return err
		}
		innerDec := xml.NewDecoder(bytes.NewReader(raw.Inner))
		for {
			innerTok, err := innerDec.Token()
			if err != nil {
				break // io.EOF or parse error — done with inner content
			}
			if err := e.EncodeToken(xml.CopyToken(innerTok)); err != nil {
				return err
			}
		}
	}

	return e.EncodeToken(start.End())
}

// parseInt parses a decimal integer string.
func parseInt(s string) (int, error) {
	var n int
	_, err := fmt.Sscanf(s, "%d", &n)
	return n, err
}

// defaultNamespaces returns the minimal xmlns declarations for a new
// comments.xml document.
func defaultNamespaces() []xml.Attr {
	return []xml.Attr{
		{Name: xml.Name{Local: "xmlns:w"}, Value: xmltypes.NSw},
	}
}
