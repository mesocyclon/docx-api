package rpr

import (
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// UnmarshalXML reads CT_RPrBase children, dispatching by local name.
func (b *CT_RPrBase) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	return b.decodeChildren(d)
}

// decodeChildren consumes tokens until the matching EndElement.
// Extracted so callers (CT_RPr, CT_ParaRPr) can delegate base-field
// decoding while still handling their own extra children.
func (b *CT_RPrBase) decodeChildren(d *xml.Decoder) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if !b.decodeField(d, &t) {
				raw, err := decodeUnknown(d, &t)
				if err != nil {
					return err
				}
				b.Extra = append(b.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// decodeField decodes a known child element. Returns true if handled.
func (b *CT_RPrBase) decodeField(d *xml.Decoder, t *xml.StartElement) bool {
	if !xmltypes.IsWNS(t.Name.Space) {
		return false
	}
	switch t.Name.Local {
	case "rStyle":
		b.RStyle = &xmltypes.CT_String{}
		d.DecodeElement(b.RStyle, t)
	case "rFonts":
		b.RFonts = &xmltypes.CT_Fonts{}
		d.DecodeElement(b.RFonts, t)
	case "b":
		b.B = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.B, t)
	case "bCs":
		b.BCs = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.BCs, t)
	case "i":
		b.I = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.I, t)
	case "iCs":
		b.ICs = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.ICs, t)
	case "caps":
		b.Caps = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Caps, t)
	case "smallCaps":
		b.SmallCaps = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.SmallCaps, t)
	case "strike":
		b.Strike = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Strike, t)
	case "dstrike":
		b.Dstrike = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Dstrike, t)
	case "outline":
		b.Outline = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Outline, t)
	case "shadow":
		b.Shadow = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Shadow, t)
	case "emboss":
		b.Emboss = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Emboss, t)
	case "imprint":
		b.Imprint = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Imprint, t)
	case "noProof":
		b.NoProof = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.NoProof, t)
	case "snapToGrid":
		b.SnapToGrid = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.SnapToGrid, t)
	case "vanish":
		b.Vanish = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Vanish, t)
	case "webHidden":
		b.WebHidden = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.WebHidden, t)
	case "color":
		b.Color = &xmltypes.CT_Color{}
		d.DecodeElement(b.Color, t)
	case "spacing":
		b.Spacing = &xmltypes.CT_SignedTwipsMeasure{}
		d.DecodeElement(b.Spacing, t)
	case "w":
		b.W = &xmltypes.CT_TextScale{}
		d.DecodeElement(b.W, t)
	case "kern":
		b.Kern = &xmltypes.CT_HpsMeasure{}
		d.DecodeElement(b.Kern, t)
	case "position":
		b.Position = &xmltypes.CT_SignedHpsMeasure{}
		d.DecodeElement(b.Position, t)
	case "sz":
		b.Sz = &xmltypes.CT_HpsMeasure{}
		d.DecodeElement(b.Sz, t)
	case "szCs":
		b.SzCs = &xmltypes.CT_HpsMeasure{}
		d.DecodeElement(b.SzCs, t)
	case "highlight":
		b.Highlight = &xmltypes.CT_Highlight{}
		d.DecodeElement(b.Highlight, t)
	case "u":
		b.U = &xmltypes.CT_Underline{}
		d.DecodeElement(b.U, t)
	case "effect":
		b.Effect = &CT_TextEffect{}
		d.DecodeElement(b.Effect, t)
	case "bdr":
		b.Bdr = &xmltypes.CT_Border{}
		d.DecodeElement(b.Bdr, t)
	case "shd":
		b.Shd = &xmltypes.CT_Shd{}
		d.DecodeElement(b.Shd, t)
	case "fitText":
		b.FitText = &CT_FitText{}
		d.DecodeElement(b.FitText, t)
	case "vertAlign":
		b.VertAlign = &CT_VerticalAlignRun{}
		d.DecodeElement(b.VertAlign, t)
	case "rtl":
		b.Rtl = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Rtl, t)
	case "cs":
		b.Cs = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.Cs, t)
	case "em":
		b.Em = &CT_Em{}
		d.DecodeElement(b.Em, t)
	case "lang":
		b.Lang = &xmltypes.CT_Language{}
		d.DecodeElement(b.Lang, t)
	case "eastAsianLayout":
		b.EastAsianLayout = &CT_EastAsianLayout{}
		d.DecodeElement(b.EastAsianLayout, t)
	case "specVanish":
		b.SpecVanish = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.SpecVanish, t)
	case "oMath":
		b.OMath = &xmltypes.CT_OnOff{}
		d.DecodeElement(b.OMath, t)
	default:
		return false
	}
	return true
}
