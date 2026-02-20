package table

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// trPrFieldMap defines the order for CT_TrPr elements.
var trPrFieldMap = []fieldMapping{
	{"CnfStyle", "cnfStyle"},
	{"GridBefore", "gridBefore"},
	{"GridAfter", "gridAfter"},
	{"WBefore", "wBefore"},
	{"WAfter", "wAfter"},
	{"CantSplit", "cantSplit"},
	{"TrHeight", "trHeight"},
	{"TblHeader", "tblHeader"},
	{"TblCellSpacing", "tblCellSpacing"},
	{"Jc", "jc"},
	{"Hidden", "hidden"},
	{"TrPrChange", "trPrChange"},
}

// MarshalXML serializes CT_TrPr.
func (p *CT_TrPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if p.CnfStyle != nil {
		if err := encodeChild(e, "cnfStyle", p.CnfStyle); err != nil {
			return err
		}
	}
	if p.GridBefore != nil {
		if err := encodeChild(e, "gridBefore", p.GridBefore); err != nil {
			return err
		}
	}
	if p.GridAfter != nil {
		if err := encodeChild(e, "gridAfter", p.GridAfter); err != nil {
			return err
		}
	}
	if p.WBefore != nil {
		if err := encodeChild(e, "wBefore", p.WBefore); err != nil {
			return err
		}
	}
	if p.WAfter != nil {
		if err := encodeChild(e, "wAfter", p.WAfter); err != nil {
			return err
		}
	}
	if p.CantSplit != nil {
		if err := encodeChild(e, "cantSplit", p.CantSplit); err != nil {
			return err
		}
	}
	if p.TrHeight != nil {
		if err := encodeChild(e, "trHeight", p.TrHeight); err != nil {
			return err
		}
	}
	if p.TblHeader != nil {
		if err := encodeChild(e, "tblHeader", p.TblHeader); err != nil {
			return err
		}
	}
	if p.TblCellSpacing != nil {
		if err := encodeChild(e, "tblCellSpacing", p.TblCellSpacing); err != nil {
			return err
		}
	}
	if p.Jc != nil {
		if err := encodeChild(e, "jc", p.Jc); err != nil {
			return err
		}
	}
	if p.Hidden != nil {
		if err := encodeChild(e, "hidden", p.Hidden); err != nil {
			return err
		}
	}
	if p.TrPrChange != nil {
		if err := encodeChild(e, "trPrChange", p.TrPrChange); err != nil {
			return err
		}
	}

	if err := encodeRawSlice(e, p.Extra); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes <w:trPr>.
func (p *CT_TrPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cnfStyle":
				p.CnfStyle = &CT_Cnf{}
				err = d.DecodeElement(p.CnfStyle, &t)
			case "gridBefore":
				p.GridBefore = &xmltypes.CT_DecimalNumber{}
				err = d.DecodeElement(p.GridBefore, &t)
			case "gridAfter":
				p.GridAfter = &xmltypes.CT_DecimalNumber{}
				err = d.DecodeElement(p.GridAfter, &t)
			case "wBefore":
				p.WBefore = &CT_TblWidth{}
				err = d.DecodeElement(p.WBefore, &t)
			case "wAfter":
				p.WAfter = &CT_TblWidth{}
				err = d.DecodeElement(p.WAfter, &t)
			case "cantSplit":
				p.CantSplit = &xmltypes.CT_OnOff{}
				err = d.DecodeElement(p.CantSplit, &t)
			case "trHeight":
				p.TrHeight = &CT_Height{}
				err = d.DecodeElement(p.TrHeight, &t)
			case "tblHeader":
				p.TblHeader = &xmltypes.CT_OnOff{}
				err = d.DecodeElement(p.TblHeader, &t)
			case "tblCellSpacing":
				p.TblCellSpacing = &CT_TblWidth{}
				err = d.DecodeElement(p.TblCellSpacing, &t)
			case "jc":
				p.Jc = &CT_JcTable{}
				err = d.DecodeElement(p.Jc, &t)
			case "hidden":
				p.Hidden = &xmltypes.CT_OnOff{}
				err = d.DecodeElement(p.Hidden, &t)
			case "trPrChange":
				p.TrPrChange = &CT_TrPrChange{}
				err = d.DecodeElement(p.TrPrChange, &t)
			default:
				var raw shared.RawXML
				raw, err = decodeRawElement(d, &t)
				if err == nil {
					p.Extra = append(p.Extra, raw)
				}
			}
			if err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

// MarshalXML serializes CT_TblPrEx.
func (p *CT_TblPrEx) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.TblW != nil {
		if err := encodeChild(e, "tblW", p.TblW); err != nil {
			return err
		}
	}
	if p.Jc != nil {
		if err := encodeChild(e, "jc", p.Jc); err != nil {
			return err
		}
	}
	if p.TblBorders != nil {
		if err := marshalTblBorders(e, "tblBorders", p.TblBorders); err != nil {
			return err
		}
	}
	if p.Shd != nil {
		if err := encodeChild(e, "shd", p.Shd); err != nil {
			return err
		}
	}
	if p.TblLayout != nil {
		if err := encodeChild(e, "tblLayout", p.TblLayout); err != nil {
			return err
		}
	}
	if p.TblCellMar != nil {
		if err := marshalTblCellMar(e, p.TblCellMar); err != nil {
			return err
		}
	}
	if p.TblLook != nil {
		if err := encodeChild(e, "tblLook", p.TblLook); err != nil {
			return err
		}
	}
	if err := encodeRawSlice(e, p.Extra); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes <w:tblPrEx>.
func (p *CT_TblPrEx) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "tblW":
				p.TblW = &CT_TblWidth{}
				err = d.DecodeElement(p.TblW, &t)
			case "jc":
				p.Jc = &CT_JcTable{}
				err = d.DecodeElement(p.Jc, &t)
			case "tblBorders":
				p.TblBorders = &CT_TblBorders{}
				err = unmarshalTblBorders(d, &t, p.TblBorders)
			case "shd":
				p.Shd = &xmltypes.CT_Shd{}
				err = d.DecodeElement(p.Shd, &t)
			case "tblLayout":
				p.TblLayout = &CT_TblLayoutType{}
				err = d.DecodeElement(p.TblLayout, &t)
			case "tblCellMar":
				p.TblCellMar = &CT_TblCellMar{}
				err = unmarshalTblCellMar(d, &t, p.TblCellMar)
			case "tblLook":
				p.TblLook = &CT_TblLook{}
				err = d.DecodeElement(p.TblLook, &t)
			default:
				var raw shared.RawXML
				raw, err = decodeRawElement(d, &t)
				if err == nil {
					p.Extra = append(p.Extra, raw)
				}
			}
			if err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}
