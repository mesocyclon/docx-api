package numbering

import (
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/xmltypes"
)

// CT_AbstractNum represents an abstract numbering definition (w:abstractNum).
//
// It groups level definitions (CT_Lvl) under a unique abstractNumId. Concrete
// numbering instances (CT_Num) reference an abstract definition by ID.
type CT_AbstractNum struct {
	AbstractNumID  int
	Nsid           *xmltypes.CT_LongHexNumber
	MultiLevelType *xmltypes.CT_String
	Tmpl           *xmltypes.CT_LongHexNumber
	Name           *xmltypes.CT_String
	StyleLink      *xmltypes.CT_String
	NumStyleLink   *xmltypes.CT_String
	Lvl            []CT_Lvl
}

// MarshalXML serialises CT_AbstractNum with its children in XSD order.
func (a *CT_AbstractNum) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "abstractNum"}
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Local: "abstractNumId"},
		Value: strconv.Itoa(a.AbstractNumID),
	})

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if a.Nsid != nil {
		if err := encodeChild(e, "nsid", a.Nsid); err != nil {
			return err
		}
	}
	if a.MultiLevelType != nil {
		if err := encodeChild(e, "multiLevelType", a.MultiLevelType); err != nil {
			return err
		}
	}
	if a.Tmpl != nil {
		if err := encodeChild(e, "tmpl", a.Tmpl); err != nil {
			return err
		}
	}
	if a.Name != nil {
		if err := encodeChild(e, "name", a.Name); err != nil {
			return err
		}
	}
	if a.StyleLink != nil {
		if err := encodeChild(e, "styleLink", a.StyleLink); err != nil {
			return err
		}
	}
	if a.NumStyleLink != nil {
		if err := encodeChild(e, "numStyleLink", a.NumStyleLink); err != nil {
			return err
		}
	}
	for i := range a.Lvl {
		if err := a.Lvl[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_AbstractNum from XML.
func (a *CT_AbstractNum) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes
	for _, attr := range start.Attr {
		if attr.Name.Local == "abstractNumId" {
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
			a.AbstractNumID = v
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
			if err := a.decodeChild(d, t); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

func (a *CT_AbstractNum) decodeChild(d *xml.Decoder, start xml.StartElement) error {
	switch start.Name.Local {
	case "nsid":
		a.Nsid = &xmltypes.CT_LongHexNumber{}
		return d.DecodeElement(a.Nsid, &start)
	case "multiLevelType":
		a.MultiLevelType = &xmltypes.CT_String{}
		return d.DecodeElement(a.MultiLevelType, &start)
	case "tmpl":
		a.Tmpl = &xmltypes.CT_LongHexNumber{}
		return d.DecodeElement(a.Tmpl, &start)
	case "name":
		a.Name = &xmltypes.CT_String{}
		return d.DecodeElement(a.Name, &start)
	case "styleLink":
		a.StyleLink = &xmltypes.CT_String{}
		return d.DecodeElement(a.StyleLink, &start)
	case "numStyleLink":
		a.NumStyleLink = &xmltypes.CT_String{}
		return d.DecodeElement(a.NumStyleLink, &start)
	case "lvl":
		var lvl CT_Lvl
		if err := lvl.UnmarshalXML(d, start); err != nil {
			return err
		}
		a.Lvl = append(a.Lvl, lvl)
		return nil
	default:
		// Skip unknown children inside abstractNum (unlikely but safe)
		return skipToEnd(d, start)
	}
}
