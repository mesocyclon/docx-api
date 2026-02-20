package ppr

import (
	"encoding/xml"
)

// MarshalXML encodes CT_PPrBase with strict xsd:sequence element order.
// Violation of this order causes Word to report "file is corrupted".
func (p *CT_PPrBase) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := marshalPPrBaseChildren(e, p); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// marshalPPrBaseChildren marshals CT_PPrBase children without a wrapping element.
// Used both by CT_PPrBase.MarshalXML and CT_PPr.MarshalXML.
func marshalPPrBaseChildren(e *xml.Encoder, p *CT_PPrBase) error {
	if p.PStyle != nil {
		if err := encodeChild(e, "pStyle", p.PStyle); err != nil {
			return err
		}
	}
	if p.KeepNext != nil {
		if err := encodeChild(e, "keepNext", p.KeepNext); err != nil {
			return err
		}
	}
	if p.KeepLines != nil {
		if err := encodeChild(e, "keepLines", p.KeepLines); err != nil {
			return err
		}
	}
	if p.PageBreakBefore != nil {
		if err := encodeChild(e, "pageBreakBefore", p.PageBreakBefore); err != nil {
			return err
		}
	}
	if p.FramePr != nil {
		if err := encodeChild(e, "framePr", p.FramePr); err != nil {
			return err
		}
	}
	if p.WidowControl != nil {
		if err := encodeChild(e, "widowControl", p.WidowControl); err != nil {
			return err
		}
	}
	if p.NumPr != nil {
		if err := marshalNumPr(e, p.NumPr); err != nil {
			return err
		}
	}
	if p.SuppressLineNumbers != nil {
		if err := encodeChild(e, "suppressLineNumbers", p.SuppressLineNumbers); err != nil {
			return err
		}
	}
	if p.PBdr != nil {
		if err := marshalPBdr(e, p.PBdr); err != nil {
			return err
		}
	}
	if p.Shd != nil {
		if err := encodeChild(e, "shd", p.Shd); err != nil {
			return err
		}
	}
	if p.Tabs != nil {
		if err := marshalTabs(e, p.Tabs); err != nil {
			return err
		}
	}
	if p.SuppressAutoHyphens != nil {
		if err := encodeChild(e, "suppressAutoHyphens", p.SuppressAutoHyphens); err != nil {
			return err
		}
	}
	if p.Kinsoku != nil {
		if err := encodeChild(e, "kinsoku", p.Kinsoku); err != nil {
			return err
		}
	}
	if p.WordWrap != nil {
		if err := encodeChild(e, "wordWrap", p.WordWrap); err != nil {
			return err
		}
	}
	if p.OverflowPunct != nil {
		if err := encodeChild(e, "overflowPunct", p.OverflowPunct); err != nil {
			return err
		}
	}
	if p.TopLinePunct != nil {
		if err := encodeChild(e, "topLinePunct", p.TopLinePunct); err != nil {
			return err
		}
	}
	if p.AutoSpaceDE != nil {
		if err := encodeChild(e, "autoSpaceDE", p.AutoSpaceDE); err != nil {
			return err
		}
	}
	if p.AutoSpaceDN != nil {
		if err := encodeChild(e, "autoSpaceDN", p.AutoSpaceDN); err != nil {
			return err
		}
	}
	if p.Bidi != nil {
		if err := encodeChild(e, "bidi", p.Bidi); err != nil {
			return err
		}
	}
	if p.AdjustRightInd != nil {
		if err := encodeChild(e, "adjustRightInd", p.AdjustRightInd); err != nil {
			return err
		}
	}
	if p.SnapToGrid != nil {
		if err := encodeChild(e, "snapToGrid", p.SnapToGrid); err != nil {
			return err
		}
	}
	if p.Spacing != nil {
		if err := encodeChild(e, "spacing", p.Spacing); err != nil {
			return err
		}
	}
	if p.Ind != nil {
		if err := encodeChild(e, "ind", p.Ind); err != nil {
			return err
		}
	}
	if p.ContextualSpacing != nil {
		if err := encodeChild(e, "contextualSpacing", p.ContextualSpacing); err != nil {
			return err
		}
	}
	if p.MirrorIndents != nil {
		if err := encodeChild(e, "mirrorIndents", p.MirrorIndents); err != nil {
			return err
		}
	}
	if p.SuppressOverlap != nil {
		if err := encodeChild(e, "suppressOverlap", p.SuppressOverlap); err != nil {
			return err
		}
	}
	if p.Jc != nil {
		if err := encodeChild(e, "jc", p.Jc); err != nil {
			return err
		}
	}
	if p.TextDirection != nil {
		if err := encodeChild(e, "textDirection", p.TextDirection); err != nil {
			return err
		}
	}
	if p.TextAlignment != nil {
		if err := encodeChild(e, "textAlignment", p.TextAlignment); err != nil {
			return err
		}
	}
	if p.TextboxTightWrap != nil {
		if err := encodeChild(e, "textboxTightWrap", p.TextboxTightWrap); err != nil {
			return err
		}
	}
	if p.OutlineLvl != nil {
		if err := encodeChild(e, "outlineLvl", p.OutlineLvl); err != nil {
			return err
		}
	}
	if p.DivId != nil {
		if err := encodeChild(e, "divId", p.DivId); err != nil {
			return err
		}
	}
	if p.CnfStyle != nil {
		if err := encodeChild(e, "cnfStyle", p.CnfStyle); err != nil {
			return err
		}
	}
	for _, raw := range p.Extra {
		if err := encodeRaw(e, raw); err != nil {
			return err
		}
	}
	return nil
}
