package table

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
)

// MarshalXML serializes CT_Tc as <w:tc>.
func (tc *CT_Tc) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if tc.TcPr != nil {
		if err := e.EncodeElement(tc.TcPr, xml.StartElement{
			Name: xml.Name{Space: nsw, Local: "tcPr"},
		}); err != nil {
			return err
		}
	}

	// Block-level content (paragraphs, nested tables, etc.)
	for _, bl := range tc.Content {
		switch v := bl.(type) {
		case *CT_Tbl:
			if err := e.EncodeElement(v, xml.StartElement{
				Name: xml.Name{Space: nsw, Local: "tbl"},
			}); err != nil {
				return err
			}
		case CT_Tbl:
			if err := e.EncodeElement(&v, xml.StartElement{
				Name: xml.Name{Space: nsw, Local: "tbl"},
			}); err != nil {
				return err
			}
		case shared.RawXML:
			if err := encodeRawSlice(e, []shared.RawXML{v}); err != nil {
				return err
			}
		default:
			// For BlockLevelElement types from other packages (e.g. para.CT_P)
			// we encode using the element's own MarshalXML.
			if m, ok := bl.(xml.Marshaler); ok {
				raw, merr := xml.Marshal(m)
				if merr != nil {
					return merr
				}
				innerDec := xml.NewDecoder(bytesReader(raw))
				for {
					innerTok, terr := innerDec.Token()
					if terr != nil {
						break
					}
					e.EncodeToken(xml.CopyToken(innerTok))
				}
			}
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes <w:tc>.
func (tc *CT_Tc) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tcPr":
				tc.TcPr = &CT_TcPr{}
				if err := d.DecodeElement(tc.TcPr, &t); err != nil {
					return err
				}
			case "tbl":
				// Nested table
				nestedTbl := &CT_Tbl{}
				if err := d.DecodeElement(nestedTbl, &t); err != nil {
					return err
				}
				tc.Content = append(tc.Content, nestedTbl)
			default:
				// Try block factory first (for paragraphs from other modules)
				el := shared.CreateBlockElement(t.Name)
				if el != nil {
					if err := d.DecodeElement(el, &t); err != nil {
						return err
					}
					tc.Content = append(tc.Content, el)
				} else {
					// Unknown element → RawXML
					raw, err := decodeRawElement(d, &t)
					if err != nil {
						return err
					}
					tc.Content = append(tc.Content, raw)
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}
