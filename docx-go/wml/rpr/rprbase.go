package rpr

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_RPrBase — base character formatting properties.
// Field order matches EG_RPrBase in the XSD (Word writes in this fixed order).
type CT_RPrBase struct {
	RStyle          *xmltypes.CT_String             // w:rStyle
	RFonts          *xmltypes.CT_Fonts              // w:rFonts
	B               *xmltypes.CT_OnOff              // w:b
	BCs             *xmltypes.CT_OnOff              // w:bCs
	I               *xmltypes.CT_OnOff              // w:i
	ICs             *xmltypes.CT_OnOff              // w:iCs
	Caps            *xmltypes.CT_OnOff              // w:caps
	SmallCaps       *xmltypes.CT_OnOff              // w:smallCaps
	Strike          *xmltypes.CT_OnOff              // w:strike
	Dstrike         *xmltypes.CT_OnOff              // w:dstrike
	Outline         *xmltypes.CT_OnOff              // w:outline
	Shadow          *xmltypes.CT_OnOff              // w:shadow
	Emboss          *xmltypes.CT_OnOff              // w:emboss
	Imprint         *xmltypes.CT_OnOff              // w:imprint
	NoProof         *xmltypes.CT_OnOff              // w:noProof
	SnapToGrid      *xmltypes.CT_OnOff              // w:snapToGrid
	Vanish          *xmltypes.CT_OnOff              // w:vanish
	WebHidden       *xmltypes.CT_OnOff              // w:webHidden
	Color           *xmltypes.CT_Color              // w:color
	Spacing         *xmltypes.CT_SignedTwipsMeasure // w:spacing
	W               *xmltypes.CT_TextScale          // w:w
	Kern            *xmltypes.CT_HpsMeasure         // w:kern
	Position        *xmltypes.CT_SignedHpsMeasure   // w:position
	Sz              *xmltypes.CT_HpsMeasure         // w:sz
	SzCs            *xmltypes.CT_HpsMeasure         // w:szCs
	Highlight       *xmltypes.CT_Highlight          // w:highlight
	U               *xmltypes.CT_Underline          // w:u
	Effect          *CT_TextEffect                  // w:effect
	Bdr             *xmltypes.CT_Border             // w:bdr
	Shd             *xmltypes.CT_Shd                // w:shd
	FitText         *CT_FitText                     // w:fitText
	VertAlign       *CT_VerticalAlignRun            // w:vertAlign
	Rtl             *xmltypes.CT_OnOff              // w:rtl
	Cs              *xmltypes.CT_OnOff              // w:cs
	Em              *CT_Em                          // w:em
	Lang            *xmltypes.CT_Language           // w:lang
	EastAsianLayout *CT_EastAsianLayout             // w:eastAsianLayout
	SpecVanish      *xmltypes.CT_OnOff              // w:specVanish
	OMath           *xmltypes.CT_OnOff              // w:oMath
	Extra           []shared.RawXML                 // w14:*, w15:*, unknown
}

// MarshalXML writes CT_RPrBase children in the strict XSD order.
func (b *CT_RPrBase) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := b.encodeFields(e); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// encodeFields writes all children in XSD order. Package-internal helper so
// CT_RPr and CT_ParaRPr can embed base fields into their own sequences.
func (b *CT_RPrBase) encodeFields(e *xml.Encoder) error {
	enc := func(local string, v interface{}) error { return encodeChild(e, local, v) }

	if b.RStyle != nil {
		if err := enc("rStyle", b.RStyle); err != nil {
			return err
		}
	}
	if b.RFonts != nil {
		if err := enc("rFonts", b.RFonts); err != nil {
			return err
		}
	}
	if b.B != nil {
		if err := enc("b", b.B); err != nil {
			return err
		}
	}
	if b.BCs != nil {
		if err := enc("bCs", b.BCs); err != nil {
			return err
		}
	}
	if b.I != nil {
		if err := enc("i", b.I); err != nil {
			return err
		}
	}
	if b.ICs != nil {
		if err := enc("iCs", b.ICs); err != nil {
			return err
		}
	}
	if b.Caps != nil {
		if err := enc("caps", b.Caps); err != nil {
			return err
		}
	}
	if b.SmallCaps != nil {
		if err := enc("smallCaps", b.SmallCaps); err != nil {
			return err
		}
	}
	if b.Strike != nil {
		if err := enc("strike", b.Strike); err != nil {
			return err
		}
	}
	if b.Dstrike != nil {
		if err := enc("dstrike", b.Dstrike); err != nil {
			return err
		}
	}
	if b.Outline != nil {
		if err := enc("outline", b.Outline); err != nil {
			return err
		}
	}
	if b.Shadow != nil {
		if err := enc("shadow", b.Shadow); err != nil {
			return err
		}
	}
	if b.Emboss != nil {
		if err := enc("emboss", b.Emboss); err != nil {
			return err
		}
	}
	if b.Imprint != nil {
		if err := enc("imprint", b.Imprint); err != nil {
			return err
		}
	}
	if b.NoProof != nil {
		if err := enc("noProof", b.NoProof); err != nil {
			return err
		}
	}
	if b.SnapToGrid != nil {
		if err := enc("snapToGrid", b.SnapToGrid); err != nil {
			return err
		}
	}
	if b.Vanish != nil {
		if err := enc("vanish", b.Vanish); err != nil {
			return err
		}
	}
	if b.WebHidden != nil {
		if err := enc("webHidden", b.WebHidden); err != nil {
			return err
		}
	}
	if b.Color != nil {
		if err := enc("color", b.Color); err != nil {
			return err
		}
	}
	if b.Spacing != nil {
		if err := enc("spacing", b.Spacing); err != nil {
			return err
		}
	}
	if b.W != nil {
		if err := enc("w", b.W); err != nil {
			return err
		}
	}
	if b.Kern != nil {
		if err := enc("kern", b.Kern); err != nil {
			return err
		}
	}
	if b.Position != nil {
		if err := enc("position", b.Position); err != nil {
			return err
		}
	}
	if b.Sz != nil {
		if err := enc("sz", b.Sz); err != nil {
			return err
		}
	}
	if b.SzCs != nil {
		if err := enc("szCs", b.SzCs); err != nil {
			return err
		}
	}
	if b.Highlight != nil {
		if err := enc("highlight", b.Highlight); err != nil {
			return err
		}
	}
	if b.U != nil {
		if err := enc("u", b.U); err != nil {
			return err
		}
	}
	if b.Effect != nil {
		if err := enc("effect", b.Effect); err != nil {
			return err
		}
	}
	if b.Bdr != nil {
		if err := enc("bdr", b.Bdr); err != nil {
			return err
		}
	}
	if b.Shd != nil {
		if err := enc("shd", b.Shd); err != nil {
			return err
		}
	}
	if b.FitText != nil {
		if err := enc("fitText", b.FitText); err != nil {
			return err
		}
	}
	if b.VertAlign != nil {
		if err := enc("vertAlign", b.VertAlign); err != nil {
			return err
		}
	}
	if b.Rtl != nil {
		if err := enc("rtl", b.Rtl); err != nil {
			return err
		}
	}
	if b.Cs != nil {
		if err := enc("cs", b.Cs); err != nil {
			return err
		}
	}
	if b.Em != nil {
		if err := enc("em", b.Em); err != nil {
			return err
		}
	}
	if b.Lang != nil {
		if err := enc("lang", b.Lang); err != nil {
			return err
		}
	}
	if b.EastAsianLayout != nil {
		if err := enc("eastAsianLayout", b.EastAsianLayout); err != nil {
			return err
		}
	}
	if b.SpecVanish != nil {
		if err := enc("specVanish", b.SpecVanish); err != nil {
			return err
		}
	}
	if b.OMath != nil {
		if err := enc("oMath", b.OMath); err != nil {
			return err
		}
	}
	for _, raw := range b.Extra {
		if err := encodeRawXML(e, raw); err != nil {
			return err
		}
	}
	return nil
}
