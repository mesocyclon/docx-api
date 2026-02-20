package dml

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// A_SpPr
// ---------------------------------------------------------------------------

// MarshalXML writes <pic:spPr> (or <a:spPr> depending on context — the
// caller controls the element name via encodeElement).
func (sp *A_SpPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Use the name provided by the caller (typically pic:spPr).
	if start.Name.Local == "" {
		start.Name = xml.Name{Space: xmltypes.NSpic, Local: "spPr"}
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if sp.Xfrm != nil {
		if err := encodeElement(e, xmltypes.NSa, "xfrm", sp.Xfrm); err != nil {
			return err
		}
	}
	if sp.PrstGeom != nil {
		if err := encodeElement(e, xmltypes.NSa, "prstGeom", sp.PrstGeom); err != nil {
			return err
		}
	}
	if err := marshalExtras(e, sp.Extra); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <pic:spPr> (or <a:spPr>).
func (sp *A_SpPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: A_SpPr: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "xfrm":
				sp.Xfrm = &A_Xfrm{}
				if err := d.DecodeElement(sp.Xfrm, &t); err != nil {
					return err
				}
			case "prstGeom":
				sp.PrstGeom = &A_PrstGeom{}
				if err := d.DecodeElement(sp.PrstGeom, &t); err != nil {
					return err
				}
			default:
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				sp.Extra = append(sp.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// A_Xfrm
// ---------------------------------------------------------------------------

// MarshalXML writes <a:xfrm><a:off …/><a:ext …/></a:xfrm>.
func (x *A_Xfrm) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSa, Local: "xfrm"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// a:off
	off := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSa, Local: "off"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "x"}, Value: strconv.FormatInt(x.Off.X, 10)},
			{Name: xml.Name{Local: "y"}, Value: strconv.FormatInt(x.Off.Y, 10)},
		},
	}
	if err := e.EncodeToken(off); err != nil {
		return err
	}
	if err := e.EncodeToken(off.End()); err != nil {
		return err
	}
	// a:ext
	ext := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSa, Local: "ext"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "cx"}, Value: strconv.FormatInt(x.Ext.CX, 10)},
			{Name: xml.Name{Local: "cy"}, Value: strconv.FormatInt(x.Ext.CY, 10)},
		},
	}
	if err := e.EncodeToken(ext); err != nil {
		return err
	}
	if err := e.EncodeToken(ext.End()); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <a:xfrm>.
func (x *A_Xfrm) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: A_Xfrm: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "off":
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "x":
						x.Off.X, _ = strconv.ParseInt(a.Value, 10, 64)
					case "y":
						x.Off.Y, _ = strconv.ParseInt(a.Value, 10, 64)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
			case "ext":
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "cx":
						x.Ext.CX, _ = strconv.ParseInt(a.Value, 10, 64)
					case "cy":
						x.Ext.CY, _ = strconv.ParseInt(a.Value, 10, 64)
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
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
// A_PrstGeom
// ---------------------------------------------------------------------------

// MarshalXML writes <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>.
func (pg *A_PrstGeom) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSa, Local: "prstGeom"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "prst"}, Value: pg.Prst},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// a:avLst (always present, even if empty)
	avLst := xml.StartElement{Name: xml.Name{Space: xmltypes.NSa, Local: "avLst"}}
	if err := e.EncodeToken(avLst); err != nil {
		return err
	}
	if err := e.EncodeToken(avLst.End()); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <a:prstGeom>.
func (pg *A_PrstGeom) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		if a.Name.Local == "prst" {
			pg.Prst = a.Value
		}
	}
	// Consume children (avLst etc.) without storing them.
	return d.Skip()
}
