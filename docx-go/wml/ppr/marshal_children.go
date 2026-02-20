package ppr

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// ============================================================
// CT_NumPr
// ============================================================

func marshalNumPr(e *xml.Encoder, np *CT_NumPr) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "numPr"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if np.Ilvl != nil {
		if err := encodeChild(e, "ilvl", np.Ilvl); err != nil {
			return err
		}
	}
	if np.NumId != nil {
		if err := encodeChild(e, "numId", np.NumId); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func unmarshalNumPr(d *xml.Decoder, _ xml.StartElement, np *CT_NumPr) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "ilvl":
				np.Ilvl = &xmltypes.CT_DecimalNumber{}
				if err := d.DecodeElement(np.Ilvl, &t); err != nil {
					return err
				}
			case "numId":
				np.NumId = &xmltypes.CT_DecimalNumber{}
				if err := d.DecodeElement(np.NumId, &t); err != nil {
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

// ============================================================
// CT_PBdr
// ============================================================

func marshalPBdr(e *xml.Encoder, pb *CT_PBdr) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "pBdr"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if pb.Top != nil {
		if err := encodeChild(e, "top", pb.Top); err != nil {
			return err
		}
	}
	if pb.Left != nil {
		if err := encodeChild(e, "left", pb.Left); err != nil {
			return err
		}
	}
	if pb.Bottom != nil {
		if err := encodeChild(e, "bottom", pb.Bottom); err != nil {
			return err
		}
	}
	if pb.Right != nil {
		if err := encodeChild(e, "right", pb.Right); err != nil {
			return err
		}
	}
	if pb.Between != nil {
		if err := encodeChild(e, "between", pb.Between); err != nil {
			return err
		}
	}
	if pb.Bar != nil {
		if err := encodeChild(e, "bar", pb.Bar); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func unmarshalPBdr(d *xml.Decoder, _ xml.StartElement, pb *CT_PBdr) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "top":
				pb.Top = &xmltypes.CT_Border{}
				if err := d.DecodeElement(pb.Top, &t); err != nil {
					return err
				}
			case "left":
				pb.Left = &xmltypes.CT_Border{}
				if err := d.DecodeElement(pb.Left, &t); err != nil {
					return err
				}
			case "bottom":
				pb.Bottom = &xmltypes.CT_Border{}
				if err := d.DecodeElement(pb.Bottom, &t); err != nil {
					return err
				}
			case "right":
				pb.Right = &xmltypes.CT_Border{}
				if err := d.DecodeElement(pb.Right, &t); err != nil {
					return err
				}
			case "between":
				pb.Between = &xmltypes.CT_Border{}
				if err := d.DecodeElement(pb.Between, &t); err != nil {
					return err
				}
			case "bar":
				pb.Bar = &xmltypes.CT_Border{}
				if err := d.DecodeElement(pb.Bar, &t); err != nil {
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

// ============================================================
// CT_Tabs
// ============================================================

func marshalTabs(e *xml.Encoder, tabs *CT_Tabs) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "tabs"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for i := range tabs.Tab {
		if err := encodeChild(e, "tab", &tabs.Tab[i]); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func unmarshalTabs(d *xml.Decoder, _ xml.StartElement, tabs *CT_Tabs) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "tab" {
				var ts CT_TabStop
				if err := d.DecodeElement(&ts, &t); err != nil {
					return err
				}
				tabs.Tab = append(tabs.Tab, ts)
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

// ============================================================
// CT_SectPrRef (raw inner XML pass-through)
// ============================================================

func marshalSectPrRef(e *xml.Encoder, ref *CT_SectPrRef) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "sectPr"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if len(ref.InnerXML) > 0 {
		e.Flush()
		innerDec := xml.NewDecoder(bytes.NewReader(ref.InnerXML))
		for {
			innerTok, err := innerDec.Token()
			if err != nil {
				break
			}
			e.EncodeToken(xml.CopyToken(innerTok))
		}
	}
	return e.EncodeToken(start.End())
}

// ============================================================
// CT_PPrChange
// ============================================================

func marshalPPrChange(e *xml.Encoder, pc *CT_PPrChange) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: "pPrChange"},
		Attr: []xml.Attr{
			{Name: xml.Name{Space: xmltypes.NSw, Local: "id"}, Value: intToStr(pc.ID)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "author"}, Value: pc.Author},
		},
	}
	if pc.Date != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "date"}, Value: pc.Date,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if pc.PPr != nil {
		pprStart := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "pPr"},
		}
		if err := pc.PPr.MarshalXML(e, pprStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func unmarshalPPrChange(d *xml.Decoder, start xml.StartElement, pc *CT_PPrChange) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "id":
			pc.ID = strToInt(attr.Value)
		case "author":
			pc.Author = attr.Value
		case "date":
			pc.Date = attr.Value
		}
	}
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pPr" {
				pc.PPr = &CT_PPrBase{}
				if err := pc.PPr.UnmarshalXML(d, t); err != nil {
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
