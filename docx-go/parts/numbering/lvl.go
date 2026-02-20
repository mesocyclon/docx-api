package numbering

import (
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/wml/ppr"
	"github.com/vortex/docx-go/wml/rpr"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_Lvl represents a single numbering level definition (w:lvl).
//
// Child elements follow a STRICT XSD sequence order; custom MarshalXML
// and UnmarshalXML ensure correct serialisation and round-trip fidelity.
type CT_Lvl struct {
	// Attributes
	Ilvl      int     `xml:"ilvl,attr"`
	Tplc      *string `xml:"tplc,attr,omitempty"`
	Tentative *bool   `xml:"tentative,attr,omitempty"`

	// Elements — strict order per XSD (see patterns.md §2.9)
	Start          *xmltypes.CT_DecimalNumber
	NumFmt         *xmltypes.CT_String
	LvlRestart     *xmltypes.CT_DecimalNumber
	PStyle         *xmltypes.CT_String
	IsLgl          *xmltypes.CT_OnOff
	Suff           *xmltypes.CT_String
	LvlText        *xmltypes.CT_String
	LvlPicBulletId *xmltypes.CT_DecimalNumber
	LvlJc          *ppr.CT_Jc
	PPr            *ppr.CT_PPrBase
	RPr            *rpr.CT_RPrBase
	Extra          []shared.RawXML
}

// MarshalXML serialises CT_Lvl with elements in the strict XSD order.
func (l *CT_Lvl) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "lvl"}

	// Attributes
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Local: "ilvl"},
		Value: strconv.Itoa(l.Ilvl),
	})
	if l.Tplc != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Local: "tplc"},
			Value: *l.Tplc,
		})
	}
	if l.Tentative != nil && *l.Tentative {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Local: "tentative"},
			Value: "1",
		})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Child elements in XSD sequence order
	if l.Start != nil {
		if err := encodeChild(e, "start", l.Start); err != nil {
			return err
		}
	}
	if l.NumFmt != nil {
		if err := encodeChild(e, "numFmt", l.NumFmt); err != nil {
			return err
		}
	}
	if l.LvlRestart != nil {
		if err := encodeChild(e, "lvlRestart", l.LvlRestart); err != nil {
			return err
		}
	}
	if l.PStyle != nil {
		if err := encodeChild(e, "pStyle", l.PStyle); err != nil {
			return err
		}
	}
	if l.IsLgl != nil {
		if err := encodeChild(e, "isLgl", l.IsLgl); err != nil {
			return err
		}
	}
	if l.Suff != nil {
		if err := encodeChild(e, "suff", l.Suff); err != nil {
			return err
		}
	}
	if l.LvlText != nil {
		if err := encodeChild(e, "lvlText", l.LvlText); err != nil {
			return err
		}
	}
	if l.LvlPicBulletId != nil {
		if err := encodeChild(e, "lvlPicBulletId", l.LvlPicBulletId); err != nil {
			return err
		}
	}
	if l.LvlJc != nil {
		if err := encodeChild(e, "lvlJc", l.LvlJc); err != nil {
			return err
		}
	}
	if l.PPr != nil {
		if err := encodeChild(e, "pPr", l.PPr); err != nil {
			return err
		}
	}
	if l.RPr != nil {
		if err := encodeChild(e, "rPr", l.RPr); err != nil {
			return err
		}
	}

	// Extra (unrecognised elements) at the end
	if err := encodeRawExtras(e, l.Extra); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_Lvl, dispatching known children to typed fields
// and capturing unknown elements as RawXML for round-trip preservation.
func (l *CT_Lvl) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "ilvl":
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			l.Ilvl = v
		case "tplc":
			s := attr.Value
			l.Tplc = &s
		case "tentative":
			b := attr.Value == "1" || attr.Value == "true" || attr.Value == "on"
			l.Tentative = &b
		}
	}

	// Parse child elements
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if err := l.decodeChild(d, t); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

func (l *CT_Lvl) decodeChild(d *xml.Decoder, start xml.StartElement) error {
	switch start.Name.Local {
	case "start":
		l.Start = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(l.Start, &start)
	case "numFmt":
		l.NumFmt = &xmltypes.CT_String{}
		return d.DecodeElement(l.NumFmt, &start)
	case "lvlRestart":
		l.LvlRestart = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(l.LvlRestart, &start)
	case "pStyle":
		l.PStyle = &xmltypes.CT_String{}
		return d.DecodeElement(l.PStyle, &start)
	case "isLgl":
		l.IsLgl = &xmltypes.CT_OnOff{}
		return d.DecodeElement(l.IsLgl, &start)
	case "suff":
		l.Suff = &xmltypes.CT_String{}
		return d.DecodeElement(l.Suff, &start)
	case "lvlText":
		l.LvlText = &xmltypes.CT_String{}
		return d.DecodeElement(l.LvlText, &start)
	case "lvlPicBulletId":
		l.LvlPicBulletId = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(l.LvlPicBulletId, &start)
	case "lvlJc":
		l.LvlJc = &ppr.CT_Jc{}
		return d.DecodeElement(l.LvlJc, &start)
	case "pPr":
		l.PPr = &ppr.CT_PPrBase{}
		return d.DecodeElement(l.PPr, &start)
	case "rPr":
		l.RPr = &rpr.CT_RPrBase{}
		return d.DecodeElement(l.RPr, &start)
	default:
		raw, err := decodeRawXML(d, start)
		if err != nil {
			return err
		}
		l.Extra = append(l.Extra, raw)
		return nil
	}
}
