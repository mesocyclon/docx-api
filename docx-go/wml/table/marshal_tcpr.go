package table

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// tcPrFieldMap defines the STRICT XSD sequence order for CT_TcPr.
var tcPrFieldMap = []fieldMapping{
	{"CnfStyle", "cnfStyle"},
	{"TcW", "tcW"},
	{"GridSpan", "gridSpan"},
	{"HMerge", "hMerge"},
	{"VMerge", "vMerge"},
	{"TcBorders", "tcBorders"},
	{"Shd", "shd"},
	{"NoWrap", "noWrap"},
	{"TcMar", "tcMar"},
	{"TextDirection", "textDirection"},
	{"TcFitText", "tcFitText"},
	{"VAlign", "vAlign"},
	{"HideMark", "hideMark"},
	{"TcPrChange", "tcPrChange"},
}

// MarshalXML serializes CT_TcPr with strict XSD element ordering.
func (p *CT_TcPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if p.CnfStyle != nil {
		if err := encodeChild(e, "cnfStyle", p.CnfStyle); err != nil {
			return err
		}
	}
	if p.TcW != nil {
		if err := encodeChild(e, "tcW", p.TcW); err != nil {
			return err
		}
	}
	if p.GridSpan != nil {
		if err := encodeChild(e, "gridSpan", p.GridSpan); err != nil {
			return err
		}
	}
	if p.HMerge != nil {
		if err := encodeChild(e, "hMerge", p.HMerge); err != nil {
			return err
		}
	}
	if p.VMerge != nil {
		if err := encodeChild(e, "vMerge", p.VMerge); err != nil {
			return err
		}
	}
	if p.TcBorders != nil {
		if err := marshalTcBorders(e, p.TcBorders); err != nil {
			return err
		}
	}
	if p.Shd != nil {
		if err := encodeChild(e, "shd", p.Shd); err != nil {
			return err
		}
	}
	if p.NoWrap != nil {
		if err := encodeChild(e, "noWrap", p.NoWrap); err != nil {
			return err
		}
	}
	if p.TcMar != nil {
		if err := marshalTcMar(e, p.TcMar); err != nil {
			return err
		}
	}
	if p.TextDirection != nil {
		if err := encodeChild(e, "textDirection", p.TextDirection); err != nil {
			return err
		}
	}
	if p.TcFitText != nil {
		if err := encodeChild(e, "tcFitText", p.TcFitText); err != nil {
			return err
		}
	}
	if p.VAlign != nil {
		if err := encodeChild(e, "vAlign", p.VAlign); err != nil {
			return err
		}
	}
	if p.HideMark != nil {
		if err := encodeChild(e, "hideMark", p.HideMark); err != nil {
			return err
		}
	}
	if p.TcPrChange != nil {
		if err := encodeChild(e, "tcPrChange", p.TcPrChange); err != nil {
			return err
		}
	}

	if err := encodeRawSlice(e, p.Extra); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML deserializes <w:tcPr>.
func (p *CT_TcPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
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
			case "tcW":
				p.TcW = &CT_TblWidth{}
				err = d.DecodeElement(p.TcW, &t)
			case "gridSpan":
				p.GridSpan = &xmltypes.CT_DecimalNumber{}
				err = d.DecodeElement(p.GridSpan, &t)
			case "hMerge":
				p.HMerge = &CT_HMerge{}
				err = d.DecodeElement(p.HMerge, &t)
			case "vMerge":
				p.VMerge = &CT_VMerge{}
				err = d.DecodeElement(p.VMerge, &t)
			case "tcBorders":
				p.TcBorders = &CT_TcBorders{}
				err = unmarshalTcBorders(d, &t, p.TcBorders)
			case "shd":
				p.Shd = &xmltypes.CT_Shd{}
				err = d.DecodeElement(p.Shd, &t)
			case "noWrap":
				p.NoWrap = &xmltypes.CT_OnOff{}
				err = d.DecodeElement(p.NoWrap, &t)
			case "tcMar":
				p.TcMar = &CT_TblCellMar{}
				err = unmarshalTblCellMar(d, &t, p.TcMar)
			case "textDirection":
				p.TextDirection = &CT_TextDirection{}
				err = d.DecodeElement(p.TextDirection, &t)
			case "tcFitText":
				p.TcFitText = &xmltypes.CT_OnOff{}
				err = d.DecodeElement(p.TcFitText, &t)
			case "vAlign":
				p.VAlign = &CT_VerticalJc{}
				err = d.DecodeElement(p.VAlign, &t)
			case "hideMark":
				p.HideMark = &xmltypes.CT_OnOff{}
				err = d.DecodeElement(p.HideMark, &t)
			case "tcPrChange":
				p.TcPrChange = &CT_TcPrChange{}
				err = d.DecodeElement(p.TcPrChange, &t)
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

// marshalTcBorders encodes CT_TcBorders.
func marshalTcBorders(e *xml.Encoder, b *CT_TcBorders) error {
	start := xml.StartElement{Name: xml.Name{Space: nsw, Local: "tcBorders"}}
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
	if b.Tl2br != nil {
		if err := encodeChild(e, "tl2br", b.Tl2br); err != nil {
			return err
		}
	}
	if b.Tr2bl != nil {
		if err := encodeChild(e, "tr2bl", b.Tr2bl); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// unmarshalTcBorders parses <w:tcBorders>.
func unmarshalTcBorders(d *xml.Decoder, start *xml.StartElement, b *CT_TcBorders) error {
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
			case "tl2br":
				b.Tl2br, err = decodeBorderChild(d, &t)
			case "tr2bl":
				b.Tr2bl, err = decodeBorderChild(d, &t)
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

// marshalTcMar encodes CT_TblCellMar specifically for tcMar context.
func marshalTcMar(e *xml.Encoder, m *CT_TblCellMar) error {
	start := xml.StartElement{Name: xml.Name{Space: nsw, Local: "tcMar"}}
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
