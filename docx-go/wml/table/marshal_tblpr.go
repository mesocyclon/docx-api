package table

import (
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// tblPrFieldMap defines the STRICT XSD sequence order for CT_TblPr.
var tblPrFieldMap = []fieldMapping{
	{"TblStyle", "tblStyle"},
	{"TblpPr", "tblpPr"},
	{"TblOverlap", "tblOverlap"},
	{"BidiVisual", "bidiVisual"},
	{"TblStyleRowBandSize", "tblStyleRowBandSize"},
	{"TblStyleColBandSize", "tblStyleColBandSize"},
	{"TblW", "tblW"},
	{"Jc", "jc"},
	{"TblCellSpacing", "tblCellSpacing"},
	{"TblInd", "tblInd"},
	{"TblBorders", "tblBorders"},
	{"Shd", "shd"},
	{"TblLayout", "tblLayout"},
	{"TblCellMar", "tblCellMar"},
	{"TblLook", "tblLook"},
	{"TblCaption", "tblCaption"},
	{"TblDescription", "tblDescription"},
	{"TblPrChange", "tblPrChange"},
}

// MarshalXML serializes CT_TblPr with strict XSD element ordering.
func (p *CT_TblPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if p.TblStyle != nil {
		if err := encodeChild(e, "tblStyle", p.TblStyle); err != nil {
			return err
		}
	}
	if p.TblpPr != nil {
		if err := encodeChild(e, "tblpPr", p.TblpPr); err != nil {
			return err
		}
	}
	if p.TblOverlap != nil {
		if err := encodeChild(e, "tblOverlap", p.TblOverlap); err != nil {
			return err
		}
	}
	if p.BidiVisual != nil {
		if err := encodeChild(e, "bidiVisual", p.BidiVisual); err != nil {
			return err
		}
	}
	if p.TblStyleRowBandSize != nil {
		if err := encodeChild(e, "tblStyleRowBandSize", p.TblStyleRowBandSize); err != nil {
			return err
		}
	}
	if p.TblStyleColBandSize != nil {
		if err := encodeChild(e, "tblStyleColBandSize", p.TblStyleColBandSize); err != nil {
			return err
		}
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
	if p.TblCellSpacing != nil {
		if err := encodeChild(e, "tblCellSpacing", p.TblCellSpacing); err != nil {
			return err
		}
	}
	if p.TblInd != nil {
		if err := encodeChild(e, "tblInd", p.TblInd); err != nil {
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
	if p.TblCaption != nil {
		if err := encodeChild(e, "tblCaption", p.TblCaption); err != nil {
			return err
		}
	}
	if p.TblDescription != nil {
		if err := encodeChild(e, "tblDescription", p.TblDescription); err != nil {
			return err
		}
	}
	if p.TblPrChange != nil {
		if err := encodeChild(e, "tblPrChange", p.TblPrChange); err != nil {
			return err
		}
	}
	if err := encodeRawSlice(e, p.Extra); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes CT_TblPr.
func (p *CT_TblPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if err := unmarshalTblPrChild(d, &t, p); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

func unmarshalTblPrChild(d *xml.Decoder, t *xml.StartElement, p *CT_TblPr) error {
	switch t.Name.Local {
	case "tblStyle":
		p.TblStyle = &xmltypes.CT_String{}
		return d.DecodeElement(p.TblStyle, t)
	case "tblpPr":
		p.TblpPr = &CT_TblPPr{}
		return d.DecodeElement(p.TblpPr, t)
	case "tblOverlap":
		p.TblOverlap = &CT_TblOverlap{}
		return d.DecodeElement(p.TblOverlap, t)
	case "bidiVisual":
		p.BidiVisual = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.BidiVisual, t)
	case "tblStyleRowBandSize":
		p.TblStyleRowBandSize = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(p.TblStyleRowBandSize, t)
	case "tblStyleColBandSize":
		p.TblStyleColBandSize = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(p.TblStyleColBandSize, t)
	case "tblW":
		p.TblW = &CT_TblWidth{}
		return d.DecodeElement(p.TblW, t)
	case "jc":
		p.Jc = &CT_JcTable{}
		return d.DecodeElement(p.Jc, t)
	case "tblCellSpacing":
		p.TblCellSpacing = &CT_TblWidth{}
		return d.DecodeElement(p.TblCellSpacing, t)
	case "tblInd":
		p.TblInd = &CT_TblWidth{}
		return d.DecodeElement(p.TblInd, t)
	case "tblBorders":
		p.TblBorders = &CT_TblBorders{}
		return unmarshalTblBorders(d, t, p.TblBorders)
	case "shd":
		p.Shd = &xmltypes.CT_Shd{}
		return d.DecodeElement(p.Shd, t)
	case "tblLayout":
		p.TblLayout = &CT_TblLayoutType{}
		return d.DecodeElement(p.TblLayout, t)
	case "tblCellMar":
		p.TblCellMar = &CT_TblCellMar{}
		return unmarshalTblCellMar(d, t, p.TblCellMar)
	case "tblLook":
		p.TblLook = &CT_TblLook{}
		return d.DecodeElement(p.TblLook, t)
	case "tblCaption":
		p.TblCaption = &xmltypes.CT_String{}
		return d.DecodeElement(p.TblCaption, t)
	case "tblDescription":
		p.TblDescription = &xmltypes.CT_String{}
		return d.DecodeElement(p.TblDescription, t)
	case "tblPrChange":
		p.TblPrChange = &CT_TblPrChange{}
		return d.DecodeElement(p.TblPrChange, t)
	default:
		raw, err := decodeRawElement(d, t)
		if err != nil {
			return err
		}
		p.Extra = append(p.Extra, raw)
		return nil
	}
}
