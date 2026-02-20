package table

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
)

// MarshalXML serializes CT_Row as <w:tr>.
func (r *CT_Row) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Preserve existing attributes; add rsidR/rsidTr.
	if r.RsidR != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: nsw, Local: "rsidR"}, Value: *r.RsidR,
		})
	}
	if r.RsidTr != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: nsw, Local: "rsidTr"}, Value: *r.RsidTr,
		})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if r.TblPrEx != nil {
		if err := e.EncodeElement(r.TblPrEx, xml.StartElement{
			Name: xml.Name{Space: nsw, Local: "tblPrEx"},
		}); err != nil {
			return err
		}
	}
	if r.TrPr != nil {
		if err := e.EncodeElement(r.TrPr, xml.StartElement{
			Name: xml.Name{Space: nsw, Local: "trPr"},
		}); err != nil {
			return err
		}
	}

	// Content: cells and raw elements
	for _, c := range r.Content {
		switch v := c.(type) {
		case CT_Tc:
			if err := e.EncodeElement(&v, xml.StartElement{
				Name: xml.Name{Space: nsw, Local: "tc"},
			}); err != nil {
				return err
			}
		case *CT_Tc:
			if err := e.EncodeElement(v, xml.StartElement{
				Name: xml.Name{Space: nsw, Local: "tc"},
			}); err != nil {
				return err
			}
		case RawRowContent:
			if err := encodeRawSlice(e, []shared.RawXML{v.RawXML}); err != nil {
				return err
			}
		}
	}

	if err := encodeRawSlice(e, r.Extra); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes <w:tr>.
func (r *CT_Row) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			s := attr.Value
			r.RsidR = &s
		case "rsidTr":
			s := attr.Value
			r.RsidTr = &s
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblPrEx":
				r.TblPrEx = &CT_TblPrEx{}
				if err := d.DecodeElement(r.TblPrEx, &t); err != nil {
					return err
				}
			case "trPr":
				r.TrPr = &CT_TrPr{}
				if err := d.DecodeElement(r.TrPr, &t); err != nil {
					return err
				}
			case "tc":
				tc := CT_Tc{}
				if err := d.DecodeElement(&tc, &t); err != nil {
					return err
				}
				r.Content = append(r.Content, tc)
			default:
				raw, err := decodeRawElement(d, &t)
				if err != nil {
					return err
				}
				r.Content = append(r.Content, RawRowContent{raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}
