package ppr

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/rpr"
	"github.com/vortex/docx-go/xmltypes"
)

// MarshalXML encodes CT_PPr as <w:pPr> with base children, rPr, sectPr, pPrChange.
func (p *CT_PPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "pPr"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Base fields in strict order
	if err := marshalPPrBaseChildren(e, &p.Base); err != nil {
		return err
	}

	// rPr (paragraph default run properties)
	if p.RPr != nil {
		if err := encodeChild(e, "rPr", p.RPr); err != nil {
			return err
		}
	}

	// sectPr (raw reference)
	if p.SectPr != nil {
		if err := marshalSectPrRef(e, p.SectPr); err != nil {
			return err
		}
	}

	// pPrChange
	if p.PPrChange != nil {
		if err := marshalPPrChange(e, p.PPrChange); err != nil {
			return err
		}
	}

	// Extra (unknown elements at CT_PPr level)
	for _, raw := range p.Extra {
		if err := encodeRaw(e, raw); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML decodes CT_PPr: base fields, rPr, sectPr, pPrChange.
func (p *CT_PPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPr":
				p.RPr = &rpr.CT_ParaRPr{}
				if err := d.DecodeElement(p.RPr, &t); err != nil {
					return err
				}
			case "sectPr":
				p.SectPr = &CT_SectPrRef{}
				if err := d.DecodeElement(p.SectPr, &t); err != nil {
					return err
				}
			case "pPrChange":
				p.PPrChange = &CT_PPrChange{}
				if err := unmarshalPPrChange(d, t, p.PPrChange); err != nil {
					return err
				}
			default:
				// Delegate to CT_PPrBase child decoder
				if err := p.Base.decodeChild(d, t); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}
