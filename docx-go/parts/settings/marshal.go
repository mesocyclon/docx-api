package settings

import (
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// CT_Settings – UnmarshalXML
// ---------------------------------------------------------------------------

func (s *CT_Settings) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Preserve all root-element attributes (namespace declarations, etc.).
	s.Namespaces = make([]xml.Attr, len(start.Attr))
	copy(s.Namespaces, start.Attr)

	extraIdx := 0

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch {
			// --- w: namespace typed elements ---
			case t.Name.Local == "writeProtection" && xmltypes.IsWNS(t.Name.Space):
				s.WriteProtection = &CT_WriteProtection{}
				if err := d.DecodeElement(s.WriteProtection, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "writeProtection")

			case t.Name.Local == "zoom" && xmltypes.IsWNS(t.Name.Space):
				s.Zoom = &CT_Zoom{}
				if err := d.DecodeElement(s.Zoom, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "zoom")

			case t.Name.Local == "proofState" && xmltypes.IsWNS(t.Name.Space):
				s.ProofState = &CT_Proof{}
				if err := d.DecodeElement(s.ProofState, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "proofState")

			case t.Name.Local == "defaultTabStop" && xmltypes.IsWNS(t.Name.Space):
				s.DefaultTabStop = &xmltypes.CT_TwipsMeasure{}
				if err := d.DecodeElement(s.DefaultTabStop, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "defaultTabStop")

			case t.Name.Local == "characterSpacingControl" && xmltypes.IsWNS(t.Name.Space):
				s.CharacterSpacingControl = &xmltypes.CT_String{}
				if err := d.DecodeElement(s.CharacterSpacingControl, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "characterSpacingControl")

			case t.Name.Local == "evenAndOddHeaders" && xmltypes.IsWNS(t.Name.Space):
				s.EvenAndOddHeaders = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(s.EvenAndOddHeaders, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "evenAndOddHeaders")

			case t.Name.Local == "mirrorMargins" && xmltypes.IsWNS(t.Name.Space):
				s.MirrorMargins = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(s.MirrorMargins, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "mirrorMargins")

			case t.Name.Local == "trackRevisions" && xmltypes.IsWNS(t.Name.Space):
				s.TrackRevisions = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(s.TrackRevisions, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "trackRevisions")

			case t.Name.Local == "doNotTrackMoves" && xmltypes.IsWNS(t.Name.Space):
				s.DoNotTrackMoves = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(s.DoNotTrackMoves, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "doNotTrackMoves")

			case t.Name.Local == "doNotTrackFormatting" && xmltypes.IsWNS(t.Name.Space):
				s.DoNotTrackFormatting = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(s.DoNotTrackFormatting, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "doNotTrackFormatting")

			case t.Name.Local == "documentProtection" && xmltypes.IsWNS(t.Name.Space):
				s.DocumentProtection = &CT_DocProtect{}
				if err := d.DecodeElement(s.DocumentProtection, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "documentProtection")

			case t.Name.Local == "compat" && xmltypes.IsWNS(t.Name.Space):
				s.Compat = &CT_Compat{}
				if err := d.DecodeElement(s.Compat, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "compat")

			case t.Name.Local == "rsids" && xmltypes.IsWNS(t.Name.Space):
				s.Rsids = &CT_DocRsids{}
				if err := d.DecodeElement(s.Rsids, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "rsids")

			case t.Name.Local == "mathPr" && t.Name.Space == xmltypes.NSm:
				raw := &shared.RawXML{}
				if err := d.DecodeElement(raw, &t); err != nil {
					return err
				}
				s.MathPr = raw
				s.elementOrder = append(s.elementOrder, "m:mathPr")

			case t.Name.Local == "themeFontLang" && xmltypes.IsWNS(t.Name.Space):
				s.ThemeFontLang = &CT_ThemeFontLang{}
				if err := d.DecodeElement(s.ThemeFontLang, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "themeFontLang")

			case t.Name.Local == "clrSchemeMapping" && xmltypes.IsWNS(t.Name.Space):
				s.ClrSchemeMapping = &CT_ClrSchemeMapping{}
				if err := d.DecodeElement(s.ClrSchemeMapping, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "clrSchemeMapping")

			case t.Name.Local == "shapeDefaults" && xmltypes.IsWNS(t.Name.Space):
				raw := &shared.RawXML{}
				if err := d.DecodeElement(raw, &t); err != nil {
					return err
				}
				s.ShapeDefaults = raw
				s.elementOrder = append(s.elementOrder, "shapeDefaults")

			case t.Name.Local == "decimalSymbol" && xmltypes.IsWNS(t.Name.Space):
				s.DecimalSymbol = &xmltypes.CT_String{}
				if err := d.DecodeElement(s.DecimalSymbol, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "decimalSymbol")

			case t.Name.Local == "listSeparator" && xmltypes.IsWNS(t.Name.Space):
				s.ListSeparator = &xmltypes.CT_String{}
				if err := d.DecodeElement(s.ListSeparator, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "listSeparator")

			case t.Name.Local == "docId" && t.Name.Space == xmltypes.NSw14:
				s.DocId14 = &xmltypes.CT_LongHexNumber{}
				if err := d.DecodeElement(s.DocId14, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "w14:docId")

			case t.Name.Local == "docId" && t.Name.Space == xmltypes.NSw15:
				s.DocId15 = &xmltypes.CT_Guid{}
				if err := d.DecodeElement(s.DocId15, &t); err != nil {
					return err
				}
				s.elementOrder = append(s.elementOrder, "w15:docId")

			default:
				// Unknown element → preserve as RawXML for round-trip.
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				s.Extra = append(s.Extra, raw)
				s.elementOrder = append(s.elementOrder, fmt.Sprintf("#extra:%d", extraIdx))
				extraIdx++
			}

		case xml.EndElement:
			return nil // end of <w:settings>
		}
	}
}

// ---------------------------------------------------------------------------
// CT_Settings – MarshalXML
// ---------------------------------------------------------------------------

func (s *CT_Settings) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:settings"}

	// Restore original namespace declarations, or use defaults.
	if len(s.Namespaces) > 0 {
		start.Attr = s.Namespaces
	} else {
		start.Attr = defaultSettingsNamespaces()
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if len(s.elementOrder) > 0 {
		// Reproduce original element order.
		extraIdx := 0
		for _, key := range s.elementOrder {
			if err := s.marshalByKey(e, key, &extraIdx); err != nil {
				return err
			}
		}
	} else {
		// No recorded order: emit typed fields in a sensible default sequence,
		// then extras at the end.
		if err := s.marshalDefaultOrder(e); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// marshalByKey emits a single element identified by its order key.
func (s *CT_Settings) marshalByKey(e *xml.Encoder, key string, extraIdx *int) error {
	switch key {
	case "writeProtection":
		if s.WriteProtection != nil {
			return encodeChild(e, xmltypes.NSw, "writeProtection", s.WriteProtection)
		}
	case "zoom":
		if s.Zoom != nil {
			return encodeChild(e, xmltypes.NSw, "zoom", s.Zoom)
		}
	case "proofState":
		if s.ProofState != nil {
			return encodeChild(e, xmltypes.NSw, "proofState", s.ProofState)
		}
	case "defaultTabStop":
		if s.DefaultTabStop != nil {
			return encodeChild(e, xmltypes.NSw, "defaultTabStop", s.DefaultTabStop)
		}
	case "characterSpacingControl":
		if s.CharacterSpacingControl != nil {
			return encodeChild(e, xmltypes.NSw, "characterSpacingControl", s.CharacterSpacingControl)
		}
	case "evenAndOddHeaders":
		return encodeOnOff(e, "evenAndOddHeaders", s.EvenAndOddHeaders)
	case "mirrorMargins":
		return encodeOnOff(e, "mirrorMargins", s.MirrorMargins)
	case "trackRevisions":
		return encodeOnOff(e, "trackRevisions", s.TrackRevisions)
	case "doNotTrackMoves":
		return encodeOnOff(e, "doNotTrackMoves", s.DoNotTrackMoves)
	case "doNotTrackFormatting":
		return encodeOnOff(e, "doNotTrackFormatting", s.DoNotTrackFormatting)
	case "documentProtection":
		if s.DocumentProtection != nil {
			return encodeChild(e, xmltypes.NSw, "documentProtection", s.DocumentProtection)
		}
	case "compat":
		if s.Compat != nil {
			return encodeChild(e, xmltypes.NSw, "compat", s.Compat)
		}
	case "rsids":
		if s.Rsids != nil {
			return encodeChild(e, xmltypes.NSw, "rsids", s.Rsids)
		}
	case "m:mathPr":
		return encodeRawXML(e, s.MathPr)
	case "themeFontLang":
		if s.ThemeFontLang != nil {
			return encodeChild(e, xmltypes.NSw, "themeFontLang", s.ThemeFontLang)
		}
	case "clrSchemeMapping":
		if s.ClrSchemeMapping != nil {
			return encodeChild(e, xmltypes.NSw, "clrSchemeMapping", s.ClrSchemeMapping)
		}
	case "shapeDefaults":
		return encodeRawXML(e, s.ShapeDefaults)
	case "decimalSymbol":
		if s.DecimalSymbol != nil {
			return encodeChild(e, xmltypes.NSw, "decimalSymbol", s.DecimalSymbol)
		}
	case "listSeparator":
		if s.ListSeparator != nil {
			return encodeChild(e, xmltypes.NSw, "listSeparator", s.ListSeparator)
		}
	case "w14:docId":
		if s.DocId14 != nil {
			return encodeChild(e, xmltypes.NSw14, "docId", s.DocId14)
		}
	case "w15:docId":
		if s.DocId15 != nil {
			return encodeChild(e, xmltypes.NSw15, "docId", s.DocId15)
		}
	default:
		// Must be an Extra element.
		if *extraIdx < len(s.Extra) {
			raw := s.Extra[*extraIdx]
			*extraIdx++
			return e.EncodeElement(raw, xml.StartElement{Name: raw.XMLName})
		}
	}
	return nil
}

// marshalDefaultOrder writes all non-nil fields in a sensible default order,
// used when no elementOrder was recorded (e.g. programmatically built settings).
func (s *CT_Settings) marshalDefaultOrder(e *xml.Encoder) error {
	if s.WriteProtection != nil {
		encodeChild(e, xmltypes.NSw, "writeProtection", s.WriteProtection)
	}
	if s.Zoom != nil {
		encodeChild(e, xmltypes.NSw, "zoom", s.Zoom)
	}
	if s.ProofState != nil {
		encodeChild(e, xmltypes.NSw, "proofState", s.ProofState)
	}
	if s.DefaultTabStop != nil {
		encodeChild(e, xmltypes.NSw, "defaultTabStop", s.DefaultTabStop)
	}
	if s.CharacterSpacingControl != nil {
		encodeChild(e, xmltypes.NSw, "characterSpacingControl", s.CharacterSpacingControl)
	}
	encodeOnOff(e, "evenAndOddHeaders", s.EvenAndOddHeaders)
	encodeOnOff(e, "mirrorMargins", s.MirrorMargins)
	encodeOnOff(e, "trackRevisions", s.TrackRevisions)
	encodeOnOff(e, "doNotTrackMoves", s.DoNotTrackMoves)
	encodeOnOff(e, "doNotTrackFormatting", s.DoNotTrackFormatting)
	if s.DocumentProtection != nil {
		encodeChild(e, xmltypes.NSw, "documentProtection", s.DocumentProtection)
	}
	if s.Compat != nil {
		encodeChild(e, xmltypes.NSw, "compat", s.Compat)
	}
	if s.Rsids != nil {
		encodeChild(e, xmltypes.NSw, "rsids", s.Rsids)
	}
	encodeRawXML(e, s.MathPr)
	if s.ThemeFontLang != nil {
		encodeChild(e, xmltypes.NSw, "themeFontLang", s.ThemeFontLang)
	}
	if s.ClrSchemeMapping != nil {
		encodeChild(e, xmltypes.NSw, "clrSchemeMapping", s.ClrSchemeMapping)
	}
	encodeRawXML(e, s.ShapeDefaults)
	if s.DecimalSymbol != nil {
		encodeChild(e, xmltypes.NSw, "decimalSymbol", s.DecimalSymbol)
	}
	if s.ListSeparator != nil {
		encodeChild(e, xmltypes.NSw, "listSeparator", s.ListSeparator)
	}
	if s.DocId14 != nil {
		encodeChild(e, xmltypes.NSw14, "docId", s.DocId14)
	}
	if s.DocId15 != nil {
		encodeChild(e, xmltypes.NSw15, "docId", s.DocId15)
	}

	for _, raw := range s.Extra {
		e.EncodeElement(raw, xml.StartElement{Name: raw.XMLName})
	}
	return nil
}

// ---------------------------------------------------------------------------
// CT_Compat – UnmarshalXML / MarshalXML
// ---------------------------------------------------------------------------

func (c *CT_Compat) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "compatSetting" && xmltypes.IsWNS(t.Name.Space) {
				var cs CT_CompatSetting
				if err := d.DecodeElement(&cs, &t); err != nil {
					return err
				}
				c.CompatSetting = append(c.CompatSetting, cs)
			} else {
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				c.Extra = append(c.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

func (c *CT_Compat) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, cs := range c.CompatSetting {
		csStart := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "compatSetting"},
			Attr: []xml.Attr{
				{Name: xml.Name{Space: xmltypes.NSw, Local: "name"}, Value: cs.Name},
				{Name: xml.Name{Space: xmltypes.NSw, Local: "uri"}, Value: cs.URI},
				{Name: xml.Name{Space: xmltypes.NSw, Local: "val"}, Value: cs.Val},
			},
		}
		if err := e.EncodeToken(csStart); err != nil {
			return err
		}
		if err := e.EncodeToken(csStart.End()); err != nil {
			return err
		}
	}
	for _, raw := range c.Extra {
		if err := e.EncodeElement(raw, xml.StartElement{Name: raw.XMLName}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// CT_DocRsids – UnmarshalXML / MarshalXML
// ---------------------------------------------------------------------------

func (r *CT_DocRsids) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rsidRoot":
				r.RsidRoot = &xmltypes.CT_LongHexNumber{}
				if err := d.DecodeElement(r.RsidRoot, &t); err != nil {
					return err
				}
			case "rsid":
				var v xmltypes.CT_LongHexNumber
				if err := d.DecodeElement(&v, &t); err != nil {
					return err
				}
				r.Rsid = append(r.Rsid, v)
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

func (r *CT_DocRsids) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if r.RsidRoot != nil {
		if err := e.EncodeElement(r.RsidRoot, xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "rsidRoot"},
		}); err != nil {
			return err
		}
	}
	for _, rsid := range r.Rsid {
		if err := e.EncodeElement(rsid, xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "rsid"},
		}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// CT_ThemeFontLang – MarshalXML (attribute-only, self-closing)
// ---------------------------------------------------------------------------

func (t *CT_ThemeFontLang) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if t.Val != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "val"}, Value: *t.Val,
		})
	}
	if t.EastAsia != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "eastAsia"}, Value: *t.EastAsia,
		})
	}
	if t.Bidi != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "bidi"}, Value: *t.Bidi,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func (t *CT_ThemeFontLang) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "val":
			s := a.Value
			t.Val = &s
		case "eastAsia":
			s := a.Value
			t.EastAsia = &s
		case "bidi":
			s := a.Value
			t.Bidi = &s
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_ClrSchemeMapping – MarshalXML (all attributes, self-closing)
// ---------------------------------------------------------------------------

func (c *CT_ClrSchemeMapping) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr,
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "bg1"}, Value: c.Bg1},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "t1"}, Value: c.T1},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "bg2"}, Value: c.Bg2},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "t2"}, Value: c.T2},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "accent1"}, Value: c.Accent1},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "accent2"}, Value: c.Accent2},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "accent3"}, Value: c.Accent3},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "accent4"}, Value: c.Accent4},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "accent5"}, Value: c.Accent5},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "accent6"}, Value: c.Accent6},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "hyperlink"}, Value: c.Hyperlink},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "followedHyperlink"}, Value: c.FollowedHyperlink},
	)
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func (c *CT_ClrSchemeMapping) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "bg1":
			c.Bg1 = a.Value
		case "t1":
			c.T1 = a.Value
		case "bg2":
			c.Bg2 = a.Value
		case "t2":
			c.T2 = a.Value
		case "accent1":
			c.Accent1 = a.Value
		case "accent2":
			c.Accent2 = a.Value
		case "accent3":
			c.Accent3 = a.Value
		case "accent4":
			c.Accent4 = a.Value
		case "accent5":
			c.Accent5 = a.Value
		case "accent6":
			c.Accent6 = a.Value
		case "hyperlink":
			c.Hyperlink = a.Value
		case "followedHyperlink":
			c.FollowedHyperlink = a.Value
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_WriteProtection / CT_DocProtect – MarshalXML
// ---------------------------------------------------------------------------

func (w *CT_WriteProtection) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if w.Recommended != nil && *w.Recommended {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "recommended"}, Value: "1",
		})
	}
	appendOptionalStringAttr(&start, xmltypes.NSw, "algorithmName", w.AlgorithmName)
	appendOptionalStringAttr(&start, xmltypes.NSw, "hashValue", w.HashValue)
	appendOptionalStringAttr(&start, xmltypes.NSw, "saltValue", w.SaltValue)
	if w.SpinCount != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "spinCount"}, Value: fmt.Sprintf("%d", *w.SpinCount),
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func (w *CT_WriteProtection) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "recommended":
			b := a.Value == "1" || a.Value == "true"
			w.Recommended = &b
		case "algorithmName":
			s := a.Value
			w.AlgorithmName = &s
		case "hashValue":
			s := a.Value
			w.HashValue = &s
		case "saltValue":
			s := a.Value
			w.SaltValue = &s
		case "spinCount":
			var n int
			fmt.Sscanf(a.Value, "%d", &n)
			w.SpinCount = &n
		}
	}
	return d.Skip()
}

func (p *CT_DocProtect) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	appendOptionalStringAttr(&start, xmltypes.NSw, "edit", p.Edit)
	if p.Enforcement != nil {
		val := "0"
		if *p.Enforcement {
			val = "1"
		}
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "enforcement"}, Value: val,
		})
	}
	appendOptionalStringAttr(&start, xmltypes.NSw, "algorithmName", p.AlgorithmName)
	appendOptionalStringAttr(&start, xmltypes.NSw, "hashValue", p.HashValue)
	appendOptionalStringAttr(&start, xmltypes.NSw, "saltValue", p.SaltValue)
	if p.SpinCount != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "spinCount"}, Value: fmt.Sprintf("%d", *p.SpinCount),
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func (p *CT_DocProtect) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "edit":
			s := a.Value
			p.Edit = &s
		case "enforcement":
			b := a.Value == "1" || a.Value == "true"
			p.Enforcement = &b
		case "algorithmName":
			s := a.Value
			p.AlgorithmName = &s
		case "hashValue":
			s := a.Value
			p.HashValue = &s
		case "saltValue":
			s := a.Value
			p.SaltValue = &s
		case "spinCount":
			var n int
			fmt.Sscanf(a.Value, "%d", &n)
			p.SpinCount = &n
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_Zoom – MarshalXML / UnmarshalXML
// ---------------------------------------------------------------------------

func (z *CT_Zoom) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{
		Name: xml.Name{Space: xmltypes.NSw, Local: "percent"}, Value: fmt.Sprintf("%d", z.Percent),
	})
	if z.Val != "" {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "val"}, Value: z.Val,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func (z *CT_Zoom) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "percent":
			fmt.Sscanf(a.Value, "%d", &z.Percent)
		case "val":
			z.Val = a.Value
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// CT_Proof – MarshalXML / UnmarshalXML
// ---------------------------------------------------------------------------

func (p *CT_Proof) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	appendOptionalStringAttr(&start, xmltypes.NSw, "spelling", p.Spelling)
	appendOptionalStringAttr(&start, xmltypes.NSw, "grammar", p.Grammar)
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func (p *CT_Proof) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "spelling":
			s := a.Value
			p.Spelling = &s
		case "grammar":
			s := a.Value
			p.Grammar = &s
		}
	}
	return d.Skip()
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// encodeChild encodes v as a child element. The caller must ensure v is
// non-nil before calling this function.
func encodeChild(e *xml.Encoder, space, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: space, Local: local},
	})
}

// encodeOnOff encodes a CT_OnOff element only when the pointer is non-nil.
func encodeOnOff(e *xml.Encoder, local string, o *xmltypes.CT_OnOff) error {
	if o == nil {
		return nil
	}
	return o.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

// encodeRawXML encodes a RawXML element when the pointer is non-nil.
func encodeRawXML(e *xml.Encoder, raw *shared.RawXML) error {
	if raw == nil {
		return nil
	}
	return e.EncodeElement(*raw, xml.StartElement{Name: raw.XMLName})
}

// appendOptionalStringAttr appends a string attribute only when the pointer
// is non-nil.
func appendOptionalStringAttr(start *xml.StartElement, space, local string, v *string) {
	if v != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: space, Local: local}, Value: *v,
		})
	}
}

// defaultSettingsNamespaces returns the namespace declarations for a
// programmatically created settings.xml.
func defaultSettingsNamespaces() []xml.Attr {
	return []xml.Attr{
		{Name: xml.Name{Local: "xmlns:mc"}, Value: xmltypes.NSmc},
		{Name: xml.Name{Local: "xmlns:o"}, Value: xmltypes.NSo},
		{Name: xml.Name{Local: "xmlns:r"}, Value: xmltypes.NSr},
		{Name: xml.Name{Local: "xmlns:m"}, Value: xmltypes.NSm},
		{Name: xml.Name{Local: "xmlns:v"}, Value: xmltypes.NSv},
		{Name: xml.Name{Local: "xmlns:w10"}, Value: xmltypes.NSw10},
		{Name: xml.Name{Local: "xmlns:w"}, Value: xmltypes.NSw},
		{Name: xml.Name{Local: "xmlns:w14"}, Value: xmltypes.NSw14},
		{Name: xml.Name{Local: "xmlns:w15"}, Value: xmltypes.NSw15},
		{Name: xml.Name{Local: "xmlns:sl"}, Value: xmltypes.NSsl},
		{Name: xml.Name{Local: "mc:Ignorable"}, Value: "w14 w15"},
	}
}
