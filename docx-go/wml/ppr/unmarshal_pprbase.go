package ppr

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// UnmarshalXML decodes CT_PPrBase, handling known and unknown elements.
func (p *CT_PPrBase) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if err := p.decodeChild(d, t); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

// decodeChild routes a child element to the appropriate field or Extra.
func (p *CT_PPrBase) decodeChild(d *xml.Decoder, t xml.StartElement) error {
	switch t.Name.Local {
	case "pStyle":
		p.PStyle = &xmltypes.CT_String{}
		return d.DecodeElement(p.PStyle, &t)
	case "keepNext":
		p.KeepNext = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.KeepNext, &t)
	case "keepLines":
		p.KeepLines = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.KeepLines, &t)
	case "pageBreakBefore":
		p.PageBreakBefore = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.PageBreakBefore, &t)
	case "framePr":
		p.FramePr = &CT_FramePr{}
		return d.DecodeElement(p.FramePr, &t)
	case "widowControl":
		p.WidowControl = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.WidowControl, &t)
	case "numPr":
		p.NumPr = &CT_NumPr{}
		return unmarshalNumPr(d, t, p.NumPr)
	case "suppressLineNumbers":
		p.SuppressLineNumbers = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.SuppressLineNumbers, &t)
	case "pBdr":
		p.PBdr = &CT_PBdr{}
		return unmarshalPBdr(d, t, p.PBdr)
	case "shd":
		p.Shd = &xmltypes.CT_Shd{}
		return d.DecodeElement(p.Shd, &t)
	case "tabs":
		p.Tabs = &CT_Tabs{}
		return unmarshalTabs(d, t, p.Tabs)
	case "suppressAutoHyphens":
		p.SuppressAutoHyphens = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.SuppressAutoHyphens, &t)
	case "kinsoku":
		p.Kinsoku = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.Kinsoku, &t)
	case "wordWrap":
		p.WordWrap = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.WordWrap, &t)
	case "overflowPunct":
		p.OverflowPunct = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.OverflowPunct, &t)
	case "topLinePunct":
		p.TopLinePunct = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.TopLinePunct, &t)
	case "autoSpaceDE":
		p.AutoSpaceDE = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.AutoSpaceDE, &t)
	case "autoSpaceDN":
		p.AutoSpaceDN = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.AutoSpaceDN, &t)
	case "bidi":
		p.Bidi = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.Bidi, &t)
	case "adjustRightInd":
		p.AdjustRightInd = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.AdjustRightInd, &t)
	case "snapToGrid":
		p.SnapToGrid = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.SnapToGrid, &t)
	case "spacing":
		p.Spacing = &CT_Spacing{}
		return d.DecodeElement(p.Spacing, &t)
	case "ind":
		p.Ind = &CT_Ind{}
		return d.DecodeElement(p.Ind, &t)
	case "contextualSpacing":
		p.ContextualSpacing = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.ContextualSpacing, &t)
	case "mirrorIndents":
		p.MirrorIndents = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.MirrorIndents, &t)
	case "suppressOverlap":
		p.SuppressOverlap = &xmltypes.CT_OnOff{}
		return d.DecodeElement(p.SuppressOverlap, &t)
	case "jc":
		p.Jc = &CT_Jc{}
		return d.DecodeElement(p.Jc, &t)
	case "textDirection":
		p.TextDirection = &CT_TextDirection{}
		return d.DecodeElement(p.TextDirection, &t)
	case "textAlignment":
		p.TextAlignment = &CT_TextAlignment{}
		return d.DecodeElement(p.TextAlignment, &t)
	case "textboxTightWrap":
		p.TextboxTightWrap = &CT_TextboxTightWrap{}
		return d.DecodeElement(p.TextboxTightWrap, &t)
	case "outlineLvl":
		p.OutlineLvl = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(p.OutlineLvl, &t)
	case "divId":
		p.DivId = &xmltypes.CT_DecimalNumber{}
		return d.DecodeElement(p.DivId, &t)
	case "cnfStyle":
		p.CnfStyle = &CT_Cnf{}
		return d.DecodeElement(p.CnfStyle, &t)
	default:
		// Unknown element → save as RawXML for round-trip
		var raw shared.RawXML
		if err := d.DecodeElement(&raw, &t); err != nil {
			return err
		}
		p.Extra = append(p.Extra, raw)
		return nil
	}
}
