package dml

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// MarshalXML serialises WP_Anchor as <wp:anchor> with children in
// xsd:sequence order.
func (a *WP_Anchor) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSwp, Local: "anchor"},
		Attr: []xml.Attr{
			boolAttr("behindDoc", a.BehindDoc),
			intAttr("distT", a.DistT),
			intAttr("distB", a.DistB),
			intAttr("distL", a.DistL),
			intAttr("distR", a.DistR),
			intAttr("relativeHeight", a.RelativeHeight),
			boolAttr("simplePos", a.SimplePos),
			boolAttr("locked", a.Locked),
			boolAttr("layoutInCell", a.LayoutInCell),
			boolAttr("allowOverlap", a.AllowOverlap),
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. wp:simplePos
	if a.SimplePosXY != nil {
		if err := encodeElement(e, xmltypes.NSwp, "simplePos", a.SimplePosXY); err != nil {
			return err
		}
	} else {
		// Word always emits simplePos even when unused (x="0" y="0").
		zero := WP_Point{}
		if err := encodeElement(e, xmltypes.NSwp, "simplePos", &zero); err != nil {
			return err
		}
	}
	// 2. wp:positionH
	if err := encodeElement(e, xmltypes.NSwp, "positionH", &a.PositionH); err != nil {
		return err
	}
	// 3. wp:positionV
	if err := encodeElement(e, xmltypes.NSwp, "positionV", &a.PositionV); err != nil {
		return err
	}
	// 4. wp:extent
	if err := encodeElement(e, xmltypes.NSwp, "extent", &a.Extent); err != nil {
		return err
	}
	// 5. wp:effectExtent
	if a.EffectExtent != nil {
		if err := encodeElement(e, xmltypes.NSwp, "effectExtent", a.EffectExtent); err != nil {
			return err
		}
	}
	// 6. wrap element
	if err := marshalWrapType(e, a.WrapType); err != nil {
		return err
	}
	// 7. wp:docPr
	if err := encodeElement(e, xmltypes.NSwp, "docPr", &a.DocPr); err != nil {
		return err
	}
	// 8. Extra
	if err := marshalExtras(e, a.Extra); err != nil {
		return err
	}
	// 9. a:graphic
	if err := encodeElement(e, xmltypes.NSa, "graphic", &a.Graphic); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads a <wp:anchor> element.
func (a *WP_Anchor) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "behindDoc":
			a.BehindDoc = parseBool(attr.Value)
		case "distT":
			a.DistT, _ = strconv.Atoi(attr.Value)
		case "distB":
			a.DistB, _ = strconv.Atoi(attr.Value)
		case "distL":
			a.DistL, _ = strconv.Atoi(attr.Value)
		case "distR":
			a.DistR, _ = strconv.Atoi(attr.Value)
		case "relativeHeight":
			a.RelativeHeight, _ = strconv.Atoi(attr.Value)
		case "simplePos":
			a.SimplePos = parseBool(attr.Value)
		case "locked":
			a.Locked = parseBool(attr.Value)
		case "layoutInCell":
			a.LayoutInCell = parseBool(attr.Value)
		case "allowOverlap":
			a.AllowOverlap = parseBool(attr.Value)
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: WP_Anchor: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "simplePos":
				a.SimplePosXY = &WP_Point{}
				if err := d.DecodeElement(a.SimplePosXY, &t); err != nil {
					return err
				}
			case "positionH":
				if err := d.DecodeElement(&a.PositionH, &t); err != nil {
					return err
				}
			case "positionV":
				if err := d.DecodeElement(&a.PositionV, &t); err != nil {
					return err
				}
			case "extent":
				if err := d.DecodeElement(&a.Extent, &t); err != nil {
					return err
				}
			case "effectExtent":
				a.EffectExtent = &WP_EffectExtent{}
				if err := d.DecodeElement(a.EffectExtent, &t); err != nil {
					return err
				}
			case "wrapNone":
				a.WrapType = WP_WrapNone{}
				if err := d.Skip(); err != nil {
					return err
				}
			case "wrapSquare":
				w := WP_WrapSquare{}
				for _, wa := range t.Attr {
					if wa.Name.Local == "wrapText" {
						w.WrapText = wa.Value
					}
				}
				a.WrapType = w
				if err := d.Skip(); err != nil {
					return err
				}
			case "wrapTight":
				w := WP_WrapTight{}
				for _, wa := range t.Attr {
					if wa.Name.Local == "wrapText" {
						w.WrapText = wa.Value
					}
				}
				a.WrapType = w
				if err := d.Skip(); err != nil {
					return err
				}
			case "wrapTopAndBottom":
				a.WrapType = WP_WrapTopAndBottom{}
				if err := d.Skip(); err != nil {
					return err
				}
			case "docPr":
				if err := d.DecodeElement(&a.DocPr, &t); err != nil {
					return err
				}
			case "graphic":
				if err := d.DecodeElement(&a.Graphic, &t); err != nil {
					return err
				}
			default:
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				a.Extra = append(a.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// marshalWrapType serialises the concrete wrap variant.
func marshalWrapType(e *xml.Encoder, wt interface{}) error {
	if wt == nil {
		return nil
	}
	switch w := wt.(type) {
	case WP_WrapNone:
		st := xml.StartElement{Name: xml.Name{Space: xmltypes.NSwp, Local: "wrapNone"}}
		if err := e.EncodeToken(st); err != nil {
			return err
		}
		return e.EncodeToken(st.End())
	case WP_WrapSquare:
		st := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSwp, Local: "wrapSquare"},
			Attr: []xml.Attr{{Name: xml.Name{Local: "wrapText"}, Value: w.WrapText}},
		}
		if err := e.EncodeToken(st); err != nil {
			return err
		}
		return e.EncodeToken(st.End())
	case WP_WrapTight:
		st := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSwp, Local: "wrapTight"},
			Attr: []xml.Attr{{Name: xml.Name{Local: "wrapText"}, Value: w.WrapText}},
		}
		if err := e.EncodeToken(st); err != nil {
			return err
		}
		return e.EncodeToken(st.End())
	case WP_WrapTopAndBottom:
		st := xml.StartElement{Name: xml.Name{Space: xmltypes.NSwp, Local: "wrapTopAndBottom"}}
		if err := e.EncodeToken(st); err != nil {
			return err
		}
		return e.EncodeToken(st.End())
	default:
		return fmt.Errorf("dml: unknown wrap type %T", wt)
	}
}
