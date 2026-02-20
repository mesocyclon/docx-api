package numbering

import (
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/xmltypes"
)

// CT_Num represents a concrete numbering instance (w:num). It references
// an abstract numbering definition and may contain level overrides.
type CT_Num struct {
	NumID         int `xml:"numId,attr"`
	AbstractNumID xmltypes.CT_DecimalNumber
	LvlOverride   []CT_NumLvl
}

// MarshalXML serialises CT_Num.
func (n *CT_Num) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "num"}
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Local: "numId"},
		Value: strconv.Itoa(n.NumID),
	})

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if err := encodeChild(e, "abstractNumId", &n.AbstractNumID); err != nil {
		return err
	}
	for i := range n.LvlOverride {
		if err := n.LvlOverride[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_Num from XML.
func (n *CT_Num) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "numId" {
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			n.NumID = v
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
			case "abstractNumId":
				if err := d.DecodeElement(&n.AbstractNumID, &t); err != nil {
					return err
				}
			case "lvlOverride":
				var lo CT_NumLvl
				if err := lo.UnmarshalXML(d, t); err != nil {
					return err
				}
				n.LvlOverride = append(n.LvlOverride, lo)
			default:
				if err := skipToEnd(d, t); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// CT_NumLvl represents a level override inside a w:num (w:lvlOverride).
type CT_NumLvl struct {
	Ilvl          int `xml:"ilvl,attr"`
	StartOverride *xmltypes.CT_DecimalNumber
	Lvl           *CT_Lvl
}

// MarshalXML serialises CT_NumLvl.
func (nl *CT_NumLvl) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "lvlOverride"}
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Local: "ilvl"},
		Value: strconv.Itoa(nl.Ilvl),
	})

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if nl.StartOverride != nil {
		if err := encodeChild(e, "startOverride", nl.StartOverride); err != nil {
			return err
		}
	}
	if nl.Lvl != nil {
		if err := nl.Lvl.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_NumLvl from XML.
func (nl *CT_NumLvl) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "ilvl" {
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			nl.Ilvl = v
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
			case "startOverride":
				nl.StartOverride = &xmltypes.CT_DecimalNumber{}
				if err := d.DecodeElement(nl.StartOverride, &t); err != nil {
					return err
				}
			case "lvl":
				nl.Lvl = &CT_Lvl{}
				if err := nl.Lvl.UnmarshalXML(d, t); err != nil {
					return err
				}
			default:
				if err := skipToEnd(d, t); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}
