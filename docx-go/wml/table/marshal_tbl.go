package table

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
)

// MarshalXML serializes CT_Tbl as <w:tbl>.
func (tbl *CT_Tbl) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: nsw, Local: "tbl"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if tbl.TblPr != nil {
		if err := e.EncodeElement(tbl.TblPr, xml.StartElement{
			Name: xml.Name{Space: nsw, Local: "tblPr"},
		}); err != nil {
			return err
		}
	}
	if tbl.TblGrid != nil {
		if err := marshalTblGrid(e, tbl.TblGrid); err != nil {
			return err
		}
	}

	// Content: rows and raw elements
	for _, c := range tbl.Content {
		switch v := c.(type) {
		case CT_Row:
			if err := e.EncodeElement(&v, xml.StartElement{
				Name: xml.Name{Space: nsw, Local: "tr"},
			}); err != nil {
				return err
			}
		case *CT_Row:
			if err := e.EncodeElement(v, xml.StartElement{
				Name: xml.Name{Space: nsw, Local: "tr"},
			}); err != nil {
				return err
			}
		case RawTblContent:
			if err := encodeRawSlice(e, []shared.RawXML{v.RawXML}); err != nil {
				return err
			}
		}
	}

	if err := encodeRawSlice(e, tbl.Extra); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes <w:tbl>.
func (tbl *CT_Tbl) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblPr":
				tbl.TblPr = &CT_TblPr{}
				if err := d.DecodeElement(tbl.TblPr, &t); err != nil {
					return err
				}
			case "tblGrid":
				tbl.TblGrid = &CT_TblGrid{}
				if err := unmarshalTblGrid(d, &t, tbl.TblGrid); err != nil {
					return err
				}
			case "tr":
				row := CT_Row{}
				if err := d.DecodeElement(&row, &t); err != nil {
					return err
				}
				tbl.Content = append(tbl.Content, row)
			default:
				raw, err := decodeRawElement(d, &t)
				if err != nil {
					return err
				}
				tbl.Content = append(tbl.Content, RawTblContent{raw})
			}
		case xml.EndElement:
			return nil
		}
	}
}

func marshalTblGrid(e *xml.Encoder, g *CT_TblGrid) error {
	start := xml.StartElement{Name: xml.Name{Space: nsw, Local: "tblGrid"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, col := range g.GridCol {
		colStart := xml.StartElement{
			Name: xml.Name{Space: nsw, Local: "gridCol"},
			Attr: []xml.Attr{
				{Name: xml.Name{Space: nsw, Local: "w"}, Value: intToStr(col.W)},
			},
		}
		if err := e.EncodeToken(colStart); err != nil {
			return err
		}
		if err := e.EncodeToken(colStart.End()); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func unmarshalTblGrid(d *xml.Decoder, start *xml.StartElement, g *CT_TblGrid) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "gridCol" {
				col := CT_TblGridCol{}
				for _, attr := range t.Attr {
					if attr.Name.Local == "w" {
						col.W = strToInt(attr.Value)
					}
				}
				g.GridCol = append(g.GridCol, col)
				if err := d.Skip(); err != nil {
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
