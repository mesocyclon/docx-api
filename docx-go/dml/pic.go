package dml

import (
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// PIC_Pic
// ---------------------------------------------------------------------------

// MarshalXML writes <pic:pic>.
func (p *PIC_Pic) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSpic, Local: "pic"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// pic:nvPicPr
	if err := encodeElement(e, xmltypes.NSpic, "nvPicPr", &p.NvPicPr); err != nil {
		return err
	}
	// pic:blipFill
	if err := encodeElement(e, xmltypes.NSpic, "blipFill", &p.BlipFill); err != nil {
		return err
	}
	// pic:spPr
	if err := encodeElement(e, xmltypes.NSpic, "spPr", &p.SpPr); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <pic:pic>.
func (p *PIC_Pic) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: PIC_Pic: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "nvPicPr":
				if err := d.DecodeElement(&p.NvPicPr, &t); err != nil {
					return err
				}
			case "blipFill":
				if err := d.DecodeElement(&p.BlipFill, &t); err != nil {
					return err
				}
			case "spPr":
				if err := d.DecodeElement(&p.SpPr, &t); err != nil {
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
// PIC_NvPicPr
// ---------------------------------------------------------------------------

// MarshalXML writes <pic:nvPicPr>.
func (nv *PIC_NvPicPr) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSpic, Local: "nvPicPr"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// pic:cNvPr
	if err := encodeElement(e, xmltypes.NSpic, "cNvPr", &nv.CNvPr); err != nil {
		return err
	}
	// pic:cNvPicPr (empty, but must be present)
	empty := xml.StartElement{Name: xml.Name{Space: xmltypes.NSpic, Local: "cNvPicPr"}}
	if err := e.EncodeToken(empty); err != nil {
		return err
	}
	if err := e.EncodeToken(empty.End()); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <pic:nvPicPr>.
func (nv *PIC_NvPicPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: PIC_NvPicPr: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cNvPr":
				if err := d.DecodeElement(&nv.CNvPr, &t); err != nil {
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
// PIC_BlipFill
// ---------------------------------------------------------------------------

// MarshalXML writes <pic:blipFill>.
func (bf *PIC_BlipFill) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSpic, Local: "blipFill"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// a:blip
	if err := encodeElement(e, xmltypes.NSa, "blip", &bf.Blip); err != nil {
		return err
	}
	// a:stretch
	if bf.Stretch != nil {
		if err := encodeElement(e, xmltypes.NSa, "stretch", bf.Stretch); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <pic:blipFill>.
func (bf *PIC_BlipFill) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: PIC_BlipFill: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "blip":
				if err := d.DecodeElement(&bf.Blip, &t); err != nil {
					return err
				}
			case "stretch":
				bf.Stretch = &A_Stretch{}
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
// A_Blip
// ---------------------------------------------------------------------------

// MarshalXML writes <a:blip r:embed="…"/>.
func (b *A_Blip) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSa, Local: "blip"},
		Attr: []xml.Attr{
			{Name: xml.Name{Space: xmltypes.NSr, Local: "embed"}, Value: b.Embed},
		},
	}
	if b.Link != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSr, Local: "link"}, Value: b.Link,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <a:blip>.
func (b *A_Blip) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "embed":
			b.Embed = a.Value
		case "link":
			b.Link = a.Value
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// A_Stretch
// ---------------------------------------------------------------------------

// MarshalXML writes <a:stretch><a:fillRect/></a:stretch>.
func (s *A_Stretch) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSa, Local: "stretch"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	// a:fillRect (always present, self-closing)
	fr := xml.StartElement{Name: xml.Name{Space: xmltypes.NSa, Local: "fillRect"}}
	if err := e.EncodeToken(fr); err != nil {
		return err
	}
	if err := e.EncodeToken(fr.End()); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}
