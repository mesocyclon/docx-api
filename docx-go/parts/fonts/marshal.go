package fonts

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ============================================================
// CT_FontsList — MarshalXML / UnmarshalXML
// ============================================================

// MarshalXML encodes CT_FontsList as <w:fonts>.
func (fl *CT_FontsList) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "fonts"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for i := range fl.Font {
		fontStart := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "font"},
		}
		if err := e.EncodeElement(&fl.Font[i], fontStart); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML decodes <w:fonts> into CT_FontsList.
func (fl *CT_FontsList) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "font" {
				var f CT_Font
				if err := d.DecodeElement(&f, &t); err != nil {
					return err
				}
				fl.Font = append(fl.Font, f)
			} else {
				// Skip unknown children of <w:fonts>.
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
// CT_Font — MarshalXML / UnmarshalXML
// ============================================================

// MarshalXML encodes CT_Font as <w:font w:name="..."> with children in
// strict XSD sequence order.
func (f *CT_Font) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "font"}
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Space: xmltypes.NSw, Local: "name"},
		Value: f.Name,
	})
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Children in strict XSD sequence order.
	if f.Panose1 != nil {
		if err := encodeChild(e, "panose1", f.Panose1); err != nil {
			return err
		}
	}
	if f.Charset != nil {
		if err := encodeChild(e, "charset", f.Charset); err != nil {
			return err
		}
	}
	if f.Family != nil {
		if err := encodeChild(e, "family", f.Family); err != nil {
			return err
		}
	}
	if f.Pitch != nil {
		if err := encodeChild(e, "pitch", f.Pitch); err != nil {
			return err
		}
	}
	if f.Sig != nil {
		if err := encodeChild(e, "sig", f.Sig); err != nil {
			return err
		}
	}
	if f.EmbedRegular != nil {
		if err := encodeFontRel(e, "embedRegular", f.EmbedRegular); err != nil {
			return err
		}
	}
	if f.EmbedBold != nil {
		if err := encodeFontRel(e, "embedBold", f.EmbedBold); err != nil {
			return err
		}
	}
	if f.EmbedItalic != nil {
		if err := encodeFontRel(e, "embedItalic", f.EmbedItalic); err != nil {
			return err
		}
	}
	if f.EmbedBoldItalic != nil {
		if err := encodeFontRel(e, "embedBoldItalic", f.EmbedBoldItalic); err != nil {
			return err
		}
	}

	// Extra — unknown extension elements at the end.
	for _, raw := range f.Extra {
		if err := e.EncodeElement(raw, xml.StartElement{Name: raw.XMLName}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML decodes <w:font> with all known children and captures unknown
// elements into Extra for round-trip fidelity.
func (f *CT_Font) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, attr := range start.Attr {
		if attr.Name.Local == "name" {
			f.Name = attr.Value
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
			case "panose1":
				f.Panose1 = &CT_Panose{}
				if err := d.DecodeElement(f.Panose1, &t); err != nil {
					return err
				}
			case "charset":
				f.Charset = &CT_Charset{}
				if err := d.DecodeElement(f.Charset, &t); err != nil {
					return err
				}
			case "family":
				f.Family = &CT_FontFamily{}
				if err := d.DecodeElement(f.Family, &t); err != nil {
					return err
				}
			case "pitch":
				f.Pitch = &CT_Pitch{}
				if err := d.DecodeElement(f.Pitch, &t); err != nil {
					return err
				}
			case "sig":
				f.Sig = &CT_FontSig{}
				if err := d.DecodeElement(f.Sig, &t); err != nil {
					return err
				}
			case "embedRegular":
				f.EmbedRegular = &CT_FontRel{}
				if err := decodeFontRel(d, &t, f.EmbedRegular); err != nil {
					return err
				}
			case "embedBold":
				f.EmbedBold = &CT_FontRel{}
				if err := decodeFontRel(d, &t, f.EmbedBold); err != nil {
					return err
				}
			case "embedItalic":
				f.EmbedItalic = &CT_FontRel{}
				if err := decodeFontRel(d, &t, f.EmbedItalic); err != nil {
					return err
				}
			case "embedBoldItalic":
				f.EmbedBoldItalic = &CT_FontRel{}
				if err := decodeFontRel(d, &t, f.EmbedBoldItalic); err != nil {
					return err
				}
			default:
				// Unknown element → save as RawXML for round-trip.
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				f.Extra = append(f.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ============================================================
// Helpers
// ============================================================

// encodeChild encodes an element in w: namespace.
func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

// encodeFontRel encodes a CT_FontRel element. The ID attribute needs the
// r: (relationships) namespace prefix.
func encodeFontRel(e *xml.Encoder, local string, fr *CT_FontRel) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	}
	if fr.FontKey != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: xmltypes.NSw, Local: "fontKey"},
			Value: *fr.FontKey,
		})
	}
	if fr.SubsetInfo != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: xmltypes.NSw, Local: "subsetted"},
			Value: *fr.SubsetInfo,
		})
	}
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Space: xmltypes.NSr, Local: "id"},
		Value: fr.ID,
	})
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// decodeFontRel decodes a CT_FontRel from attributes. The r:id attribute
// may appear as Space="…/relationships" + Local="id", or just Local="id"
// depending on the decoder.
func decodeFontRel(d *xml.Decoder, start *xml.StartElement, fr *CT_FontRel) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "fontKey":
			s := attr.Value
			fr.FontKey = &s
		case "subsetted":
			s := attr.Value
			fr.SubsetInfo = &s
		case "id":
			fr.ID = attr.Value
		}
	}
	// Self-closing element; skip any content.
	return d.Skip()
}
