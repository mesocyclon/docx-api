package dml

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// WP_PosH
// ---------------------------------------------------------------------------

// MarshalXML writes <wp:positionH relativeFrom="…">.
func (p *WP_PosH) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSwp, Local: "positionH"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "relativeFrom"}, Value: p.RelativeFrom},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.PosOffset != nil {
		if err := encodeTextElement(e, xmltypes.NSwp, "posOffset", strconv.FormatInt(*p.PosOffset, 10)); err != nil {
			return err
		}
	}
	if p.Align != nil {
		if err := encodeTextElement(e, xmltypes.NSwp, "align", *p.Align); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <wp:positionH>.
func (p *WP_PosH) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		if a.Name.Local == "relativeFrom" {
			p.RelativeFrom = a.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: WP_PosH: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "posOffset":
				var s string
				if err := d.DecodeElement(&s, &t); err != nil {
					return err
				}
				v, _ := strconv.ParseInt(s, 10, 64)
				p.PosOffset = &v
			case "align":
				var s string
				if err := d.DecodeElement(&s, &t); err != nil {
					return err
				}
				p.Align = &s
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// WP_PosV
// ---------------------------------------------------------------------------

// MarshalXML writes <wp:positionV relativeFrom="…">.
func (p *WP_PosV) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSwp, Local: "positionV"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "relativeFrom"}, Value: p.RelativeFrom},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.PosOffset != nil {
		if err := encodeTextElement(e, xmltypes.NSwp, "posOffset", strconv.FormatInt(*p.PosOffset, 10)); err != nil {
			return err
		}
	}
	if p.Align != nil {
		if err := encodeTextElement(e, xmltypes.NSwp, "align", *p.Align); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <wp:positionV>.
func (p *WP_PosV) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		if a.Name.Local == "relativeFrom" {
			p.RelativeFrom = a.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: WP_PosV: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "posOffset":
				var s string
				if err := d.DecodeElement(&s, &t); err != nil {
					return err
				}
				v, _ := strconv.ParseInt(s, 10, 64)
				p.PosOffset = &v
			case "align":
				var s string
				if err := d.DecodeElement(&s, &t); err != nil {
					return err
				}
				p.Align = &s
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// WP_EffectExtent
// ---------------------------------------------------------------------------

// MarshalXML writes <wp:effectExtent l="…" t="…" r="…" b="…"/>.
func (ee *WP_EffectExtent) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSwp, Local: "effectExtent"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "l"}, Value: strconv.FormatInt(ee.L, 10)},
			{Name: xml.Name{Local: "t"}, Value: strconv.FormatInt(ee.T, 10)},
			{Name: xml.Name{Local: "r"}, Value: strconv.FormatInt(ee.R, 10)},
			{Name: xml.Name{Local: "b"}, Value: strconv.FormatInt(ee.B, 10)},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <wp:effectExtent>.
func (ee *WP_EffectExtent) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "l":
			ee.L, _ = strconv.ParseInt(a.Value, 10, 64)
		case "t":
			ee.T, _ = strconv.ParseInt(a.Value, 10, 64)
		case "r":
			ee.R, _ = strconv.ParseInt(a.Value, 10, 64)
		case "b":
			ee.B, _ = strconv.ParseInt(a.Value, 10, 64)
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// WP_Point
// ---------------------------------------------------------------------------

// MarshalXML writes <wp:simplePos x="…" y="…"/>.
func (p *WP_Point) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSwp, Local: "simplePos"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "x"}, Value: strconv.FormatInt(p.X, 10)},
			{Name: xml.Name{Local: "y"}, Value: strconv.FormatInt(p.Y, 10)},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads a simplePos element.
func (p *WP_Point) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "x":
			p.X, _ = strconv.ParseInt(a.Value, 10, 64)
		case "y":
			p.Y, _ = strconv.ParseInt(a.Value, 10, 64)
		}
	}
	return d.Skip()
}
