package rpr

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_ParaRPr — default run properties inside pPr/rPr.
// Differs from CT_RPr by having ins/del tracking references.
type CT_ParaRPr struct {
	Base  CT_RPrBase         // base formatting fields
	Ins   *CT_TrackChangeRef // if rPr was inserted
	Del   *CT_TrackChangeRef // if rPr was deleted
	Extra []shared.RawXML
}

// MarshalXML writes CT_ParaRPr children. The outer element name is from the caller.
func (p *CT_ParaRPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. Base fields.
	if err := p.Base.encodeFields(e); err != nil {
		return err
	}

	// 2. ins / del.
	if p.Ins != nil {
		if err := encodeChild(e, "ins", p.Ins); err != nil {
			return err
		}
	}
	if p.Del != nil {
		if err := encodeChild(e, "del", p.Del); err != nil {
			return err
		}
	}

	// 3. Extra.
	for _, raw := range p.Extra {
		if err := encodeRawXML(e, raw); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_ParaRPr children.
func (p *CT_ParaRPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if xmltypes.IsWNS(t.Name.Space) {
				switch t.Name.Local {
				case "ins":
					p.Ins = &CT_TrackChangeRef{}
					if err := d.DecodeElement(p.Ins, &t); err != nil {
						return err
					}
					continue
				case "del":
					p.Del = &CT_TrackChangeRef{}
					if err := d.DecodeElement(p.Del, &t); err != nil {
						return err
					}
					continue
				}
			}
			// Try base fields.
			if p.Base.decodeField(d, &t) {
				continue
			}
			// Unknown → RawXML.
			raw, err := decodeUnknown(d, &t)
			if err != nil {
				return err
			}
			p.Extra = append(p.Extra, raw)

		case xml.EndElement:
			return nil
		}
	}
}

// Compile-time interface checks.
var (
	_ xml.Marshaler   = (*CT_ParaRPr)(nil)
	_ xml.Unmarshaler = (*CT_ParaRPr)(nil)
)
