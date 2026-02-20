package rpr

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_RPr — full run properties including rPrChange for track changes.
type CT_RPr struct {
	Base      CT_RPrBase      // base formatting properties
	RPrChange *CT_RPrChange   // track-changes: previous formatting
	Extra     []shared.RawXML // w14:*, w15:*, unknown elements
}

// MarshalXML serialises CT_RPr. The outer element name is provided by the
// caller (typically <w:rPr>).
func (r *CT_RPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. Base fields (in strict XSD order).
	if err := r.Base.encodeFields(e); err != nil {
		return err
	}

	// 2. rPrChange (track changes).
	if r.RPrChange != nil {
		if err := encodeChild(e, "rPrChange", r.RPrChange); err != nil {
			return err
		}
	}

	// 3. Extension / unknown elements.
	for _, raw := range r.Extra {
		if err := encodeRawXML(e, raw); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_RPr children.
func (r *CT_RPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			// Try rPrChange first.
			if xmltypes.IsWNS(t.Name.Space) && t.Name.Local == "rPrChange" {
				r.RPrChange = &CT_RPrChange{}
				if err := d.DecodeElement(r.RPrChange, &t); err != nil {
					return err
				}
				continue
			}
			// Try base fields.
			if r.Base.decodeField(d, &t) {
				continue
			}
			// Unknown → RawXML.
			raw, err := decodeUnknown(d, &t)
			if err != nil {
				return err
			}
			r.Extra = append(r.Extra, raw)

		case xml.EndElement:
			return nil
		}
	}
}

// CT_RPrChange — track-changes wrapper for previous run formatting.
type CT_RPrChange struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Date   string `xml:"date,attr,omitempty"`
	RPr    *CT_RPrBase
}

// MarshalXML for CT_RPrChange: attrs + optional inner <w:rPr>.
func (c *CT_RPrChange) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = appendAttr(start.Attr, nsW, "id", intToStr(c.ID))
	start.Attr = appendAttr(start.Attr, nsW, "author", c.Author)
	if c.Date != "" {
		start.Attr = appendAttr(start.Attr, nsW, "date", c.Date)
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if c.RPr != nil {
		inner := xml.StartElement{Name: xml.Name{Space: nsW, Local: "rPr"}}
		if err := c.RPr.MarshalXML(e, inner); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML for CT_RPrChange.
func (c *CT_RPrChange) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			c.ID = strToInt(attr.Value)
		case "author":
			c.Author = attr.Value
		case "date":
			c.Date = attr.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "rPr" {
				c.RPr = &CT_RPrBase{}
				if err := d.DecodeElement(c.RPr, &t); err != nil {
					return err
				}
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// --- tiny helpers ---

func appendAttr(attrs []xml.Attr, space, local, value string) []xml.Attr {
	return append(attrs, xml.Attr{
		Name:  xml.Name{Space: space, Local: local},
		Value: value,
	})
}

func intToStr(v int) string {
	// Simple int→string without importing strconv (avoids dependency).
	if v == 0 {
		return "0"
	}
	neg := false
	if v < 0 {
		neg = true
		v = -v
	}
	buf := [20]byte{}
	i := len(buf) - 1
	for v > 0 {
		buf[i] = byte('0' + v%10)
		v /= 10
		i--
	}
	if neg {
		buf[i] = '-'
		i--
	}
	return string(buf[i+1:])
}

func strToInt(s string) int {
	n := 0
	neg := false
	i := 0
	if len(s) > 0 && s[0] == '-' {
		neg = true
		i = 1
	}
	for ; i < len(s); i++ {
		c := s[i]
		if c >= '0' && c <= '9' {
			n = n*10 + int(c-'0')
		}
	}
	if neg {
		return -n
	}
	return n
}

// Ensure CT_RPr satisfies xml.Marshaler and xml.Unmarshaler at compile time.
var (
	_ xml.Marshaler   = (*CT_RPr)(nil)
	_ xml.Unmarshaler = (*CT_RPr)(nil)
)
