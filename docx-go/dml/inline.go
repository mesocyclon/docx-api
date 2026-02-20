package dml

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// MarshalXML serialises WP_Inline as <wp:inline distT="…" …> with children
// in xsd:sequence order.
func (in *WP_Inline) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSwp, Local: "inline"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "distT"}, Value: strconv.Itoa(in.DistT)},
			{Name: xml.Name{Local: "distB"}, Value: strconv.Itoa(in.DistB)},
			{Name: xml.Name{Local: "distL"}, Value: strconv.Itoa(in.DistL)},
			{Name: xml.Name{Local: "distR"}, Value: strconv.Itoa(in.DistR)},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. wp:extent
	if err := encodeElement(e, xmltypes.NSwp, "extent", &in.Extent); err != nil {
		return err
	}
	// 2. wp:effectExtent
	if in.EffectExtent != nil {
		if err := encodeElement(e, xmltypes.NSwp, "effectExtent", in.EffectExtent); err != nil {
			return err
		}
	}
	// 3. wp:docPr
	if err := encodeElement(e, xmltypes.NSwp, "docPr", &in.DocPr); err != nil {
		return err
	}
	// 4. Extra (unknown elements like wp:cNvGraphicFramePr)
	if err := marshalExtras(e, in.Extra); err != nil {
		return err
	}
	// 5. a:graphic
	if err := encodeElement(e, xmltypes.NSa, "graphic", &in.Graphic); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads a <wp:inline> element.
func (in *WP_Inline) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "distT":
			in.DistT, _ = strconv.Atoi(a.Value)
		case "distB":
			in.DistB, _ = strconv.Atoi(a.Value)
		case "distL":
			in.DistL, _ = strconv.Atoi(a.Value)
		case "distR":
			in.DistR, _ = strconv.Atoi(a.Value)
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("dml: WP_Inline: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "extent":
				if err := d.DecodeElement(&in.Extent, &t); err != nil {
					return err
				}
			case "effectExtent":
				in.EffectExtent = &WP_EffectExtent{}
				if err := d.DecodeElement(in.EffectExtent, &t); err != nil {
					return err
				}
			case "docPr":
				if err := d.DecodeElement(&in.DocPr, &t); err != nil {
					return err
				}
			case "graphic":
				if err := d.DecodeElement(&in.Graphic, &t); err != nil {
					return err
				}
			default:
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				in.Extra = append(in.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}
