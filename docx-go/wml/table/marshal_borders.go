package table

import "encoding/xml"

// marshalTblBorders encodes CT_TblBorders with children in order.
func marshalTblBorders(e *xml.Encoder, local string, b *CT_TblBorders) error {
	start := xml.StartElement{Name: xml.Name{Space: nsw, Local: local}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if b.Top != nil {
		if err := encodeChild(e, "top", b.Top); err != nil {
			return err
		}
	}
	if b.Start != nil {
		if err := encodeChild(e, "start", b.Start); err != nil {
			return err
		}
	}
	if b.Bottom != nil {
		if err := encodeChild(e, "bottom", b.Bottom); err != nil {
			return err
		}
	}
	if b.End != nil {
		if err := encodeChild(e, "end", b.End); err != nil {
			return err
		}
	}
	if b.InsideH != nil {
		if err := encodeChild(e, "insideH", b.InsideH); err != nil {
			return err
		}
	}
	if b.InsideV != nil {
		if err := encodeChild(e, "insideV", b.InsideV); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// unmarshalTblBorders parses <w:tblBorders> children.
func unmarshalTblBorders(d *xml.Decoder, start *xml.StartElement, b *CT_TblBorders) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "top":
				b.Top, err = decodeBorderChild(d, &t)
			case "start", "left":
				b.Start, err = decodeBorderChild(d, &t)
			case "bottom":
				b.Bottom, err = decodeBorderChild(d, &t)
			case "end", "right":
				b.End, err = decodeBorderChild(d, &t)
			case "insideH":
				b.InsideH, err = decodeBorderChild(d, &t)
			case "insideV":
				b.InsideV, err = decodeBorderChild(d, &t)
			default:
				err = d.Skip()
			}
			if err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

// marshalTblCellMar encodes CT_TblCellMar.
func marshalTblCellMar(e *xml.Encoder, m *CT_TblCellMar) error {
	start := xml.StartElement{Name: xml.Name{Space: nsw, Local: "tblCellMar"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if m.Top != nil {
		if err := encodeChild(e, "top", m.Top); err != nil {
			return err
		}
	}
	if m.Start != nil {
		if err := encodeChild(e, "start", m.Start); err != nil {
			return err
		}
	}
	if m.Bottom != nil {
		if err := encodeChild(e, "bottom", m.Bottom); err != nil {
			return err
		}
	}
	if m.End != nil {
		if err := encodeChild(e, "end", m.End); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// unmarshalTblCellMar parses <w:tblCellMar> children.
func unmarshalTblCellMar(d *xml.Decoder, start *xml.StartElement, m *CT_TblCellMar) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "top":
				m.Top, err = decodeTblWidthChild(d, &t)
			case "start", "left":
				m.Start, err = decodeTblWidthChild(d, &t)
			case "bottom":
				m.Bottom, err = decodeTblWidthChild(d, &t)
			case "end", "right":
				m.End, err = decodeTblWidthChild(d, &t)
			default:
				err = d.Skip()
			}
			if err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}
