package dml

import (
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// A_Graphic
// ---------------------------------------------------------------------------

// MarshalXML writes <a:graphic><a:graphicData …>…</a:graphicData></a:graphic>.
func (g *A_Graphic) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSa, Local: "graphic"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := encodeElement(e, xmltypes.NSa, "graphicData", &g.GraphicData); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <a:graphic>.
func (g *A_Graphic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: A_Graphic: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "graphicData" {
				if err := d.DecodeElement(&g.GraphicData, &t); err != nil {
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

// ---------------------------------------------------------------------------
// A_GraphicData
// ---------------------------------------------------------------------------

// MarshalXML writes <a:graphicData uri="…">.
func (gd *A_GraphicData) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSa, Local: "graphicData"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "uri"}, Value: gd.URI},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if gd.Pic != nil {
		if err := encodeElement(e, xmltypes.NSpic, "pic", gd.Pic); err != nil {
			return err
		}
	} else if gd.RawData != nil {
		if err := e.EncodeElement(gd.RawData, xml.StartElement{Name: gd.RawData.XMLName}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <a:graphicData uri="…">.
func (gd *A_GraphicData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		if a.Name.Local == "uri" {
			gd.URI = a.Value
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: A_GraphicData: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pic" {
				gd.Pic = &PIC_Pic{}
				if err := d.DecodeElement(gd.Pic, &t); err != nil {
					return err
				}
			} else {
				// Non-picture graphic data → store as RawXML.
				raw := &shared.RawXML{}
				if err := d.DecodeElement(raw, &t); err != nil {
					return err
				}
				gd.RawData = raw
			}
		case xml.EndElement:
			return nil
		}
	}
}
