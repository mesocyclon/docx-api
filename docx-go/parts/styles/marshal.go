package styles

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/ppr"
	"github.com/vortex/docx-go/wml/rpr"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
	"github.com/vortex/docx-go/xmltypes"
)

// encodeChild encodes v as a child element with the given local name in the w: namespace.
func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

// encodeRawExtras writes all RawXML extras to the encoder.
func encodeRawExtras(e *xml.Encoder, extras []shared.RawXML) error {
	for _, raw := range extras {
		if err := raw.MarshalXML(e, xml.StartElement{Name: raw.XMLName}); err != nil {
			return err
		}
	}
	return nil
}

// defaultStylesNamespaces returns the default namespace declarations for a
// new styles.xml document.
func defaultStylesNamespaces() []xml.Attr {
	return []xml.Attr{
		{Name: xml.Name{Local: "xmlns:w"}, Value: xmltypes.NSw},
		{Name: xml.Name{Local: "xmlns:mc"}, Value: xmltypes.NSmc},
		{Name: xml.Name{Local: "xmlns:r"}, Value: xmltypes.NSr},
		{Name: xml.Name{Local: "xmlns:w14"}, Value: xmltypes.NSw14},
		{Name: xml.Name{Local: "xmlns:w15"}, Value: xmltypes.NSw15},
		{Name: xml.Name{Local: "mc:Ignorable"}, Value: "w14 w15"},
	}
}

// =========================================================================
// CT_Styles
// =========================================================================

// MarshalXML writes <w:styles> with all children in the correct order.
func (s *CT_Styles) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:styles"}

	if len(s.Namespaces) > 0 {
		start.Attr = s.Namespaces
	} else {
		start.Attr = defaultStylesNamespaces()
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if s.DocDefaults != nil {
		if err := encodeChild(e, "docDefaults", s.DocDefaults); err != nil {
			return err
		}
	}
	if s.LatentStyles != nil {
		if err := encodeChild(e, "latentStyles", s.LatentStyles); err != nil {
			return err
		}
	}
	for i := range s.Style {
		if err := encodeChild(e, "style", &s.Style[i]); err != nil {
			return err
		}
	}
	if err := encodeRawExtras(e, s.Extra); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:styles> and its children.
func (s *CT_Styles) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Preserve all namespace declarations from the root element.
	s.Namespaces = make([]xml.Attr, len(start.Attr))
	copy(s.Namespaces, start.Attr)

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "docDefaults":
				s.DocDefaults = &CT_DocDefaults{}
				if err := d.DecodeElement(s.DocDefaults, &t); err != nil {
					return err
				}
			case "latentStyles":
				s.LatentStyles = &CT_LatentStyles{}
				if err := d.DecodeElement(s.LatentStyles, &t); err != nil {
					return err
				}
			case "style":
				var st CT_Style
				if err := d.DecodeElement(&st, &t); err != nil {
					return err
				}
				s.Style = append(s.Style, st)
			default:
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				s.Extra = append(s.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// =========================================================================
// CT_DocDefaults
// =========================================================================

// MarshalXML writes <w:docDefaults>.
func (dd *CT_DocDefaults) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if dd.RPrDefault != nil {
		if err := encodeChild(e, "rPrDefault", dd.RPrDefault); err != nil {
			return err
		}
	}
	if dd.PPrDefault != nil {
		if err := encodeChild(e, "pPrDefault", dd.PPrDefault); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:docDefaults>.
func (dd *CT_DocDefaults) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPrDefault":
				dd.RPrDefault = &CT_RPrDefault{}
				if err := d.DecodeElement(dd.RPrDefault, &t); err != nil {
					return err
				}
			case "pPrDefault":
				dd.PPrDefault = &CT_PPrDefault{}
				if err := d.DecodeElement(dd.PPrDefault, &t); err != nil {
					return err
				}
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

// =========================================================================
// CT_RPrDefault
// =========================================================================

// MarshalXML writes <w:rPrDefault>.
func (rd *CT_RPrDefault) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if rd.RPr != nil {
		if err := encodeChild(e, "rPr", rd.RPr); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:rPrDefault>.
func (rd *CT_RPrDefault) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "rPr" {
				rd.RPr = &rpr.CT_RPr{}
				if err := d.DecodeElement(rd.RPr, &t); err != nil {
					return err
				}
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// =========================================================================
// CT_PPrDefault
// =========================================================================

// MarshalXML writes <w:pPrDefault>.
func (pd *CT_PPrDefault) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if pd.PPr != nil {
		if err := encodeChild(e, "pPr", pd.PPr); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:pPrDefault>.
func (pd *CT_PPrDefault) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pPr" {
				pd.PPr = &ppr.CT_PPrBase{}
				if err := d.DecodeElement(pd.PPr, &t); err != nil {
					return err
				}
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// =========================================================================
// CT_LatentStyles
// =========================================================================

// MarshalXML writes <w:latentStyles> with attributes and lsdException children.
func (ls *CT_LatentStyles) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Add attributes
	if ls.DefLockedState != nil {
		start.Attr = append(start.Attr, boolAttr(xmltypes.NSw, "defLockedState", *ls.DefLockedState))
	}
	if ls.DefUIPriority != nil {
		start.Attr = append(start.Attr, intAttr(xmltypes.NSw, "defUIPriority", *ls.DefUIPriority))
	}
	if ls.DefSemiHidden != nil {
		start.Attr = append(start.Attr, boolAttr(xmltypes.NSw, "defSemiHidden", *ls.DefSemiHidden))
	}
	if ls.DefUnhideWhenUsed != nil {
		start.Attr = append(start.Attr, boolAttr(xmltypes.NSw, "defUnhideWhenUsed", *ls.DefUnhideWhenUsed))
	}
	if ls.DefQFormat != nil {
		start.Attr = append(start.Attr, boolAttr(xmltypes.NSw, "defQFormat", *ls.DefQFormat))
	}
	if ls.Count != nil {
		start.Attr = append(start.Attr, intAttr(xmltypes.NSw, "count", *ls.Count))
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for i := range ls.LsdException {
		if err := encodeChild(e, "lsdException", &ls.LsdException[i]); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:latentStyles>.
func (ls *CT_LatentStyles) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Read attributes
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "defLockedState":
			v := parseBool(attr.Value)
			ls.DefLockedState = &v
		case "defUIPriority":
			v := parseInt(attr.Value)
			ls.DefUIPriority = &v
		case "defSemiHidden":
			v := parseBool(attr.Value)
			ls.DefSemiHidden = &v
		case "defUnhideWhenUsed":
			v := parseBool(attr.Value)
			ls.DefUnhideWhenUsed = &v
		case "defQFormat":
			v := parseBool(attr.Value)
			ls.DefQFormat = &v
		case "count":
			v := parseInt(attr.Value)
			ls.Count = &v
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "lsdException" {
				var exc CT_LsdException
				if err := d.DecodeElement(&exc, &t); err != nil {
					return err
				}
				ls.LsdException = append(ls.LsdException, exc)
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// =========================================================================
// CT_Style
// =========================================================================

// MarshalXML writes <w:style> with children in STRICT XSD sequence order.
func (st *CT_Style) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Attributes
	start.Attr = appendNonEmpty(start.Attr, xmltypes.NSw, "type", st.Type)
	if st.Default != nil {
		start.Attr = append(start.Attr, boolAttr(xmltypes.NSw, "default", *st.Default))
	}
	if st.CustomStyle != nil {
		start.Attr = append(start.Attr, boolAttr(xmltypes.NSw, "customStyle", *st.CustomStyle))
	}
	start.Attr = appendNonEmpty(start.Attr, xmltypes.NSw, "styleId", st.StyleID)

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Children in strict XSD sequence order (patterns.md 2.8)
	if st.Name != nil {
		encodeChild(e, "name", st.Name)
	}
	if st.Aliases != nil {
		encodeChild(e, "aliases", st.Aliases)
	}
	if st.BasedOn != nil {
		encodeChild(e, "basedOn", st.BasedOn)
	}
	if st.Next != nil {
		encodeChild(e, "next", st.Next)
	}
	if st.Link != nil {
		encodeChild(e, "link", st.Link)
	}
	if st.AutoRedefine != nil {
		encodeChild(e, "autoRedefine", st.AutoRedefine)
	}
	if st.Hidden != nil {
		encodeChild(e, "hidden", st.Hidden)
	}
	if st.UIpriority != nil {
		encodeChild(e, "uiPriority", st.UIpriority)
	}
	if st.SemiHidden != nil {
		encodeChild(e, "semiHidden", st.SemiHidden)
	}
	if st.UnhideWhenUsed != nil {
		encodeChild(e, "unhideWhenUsed", st.UnhideWhenUsed)
	}
	if st.QFormat != nil {
		encodeChild(e, "qFormat", st.QFormat)
	}
	if st.Locked != nil {
		encodeChild(e, "locked", st.Locked)
	}
	if st.Personal != nil {
		encodeChild(e, "personal", st.Personal)
	}
	if st.PersonalCompose != nil {
		encodeChild(e, "personalCompose", st.PersonalCompose)
	}
	if st.PersonalReply != nil {
		encodeChild(e, "personalReply", st.PersonalReply)
	}
	if st.Rsid != nil {
		encodeChild(e, "rsid", st.Rsid)
	}
	if st.PPr != nil {
		encodeChild(e, "pPr", st.PPr)
	}
	if st.RPr != nil {
		encodeChild(e, "rPr", st.RPr)
	}
	if st.TblPr != nil {
		encodeChild(e, "tblPr", st.TblPr)
	}
	if st.TrPr != nil {
		encodeChild(e, "trPr", st.TrPr)
	}
	if st.TcPr != nil {
		encodeChild(e, "tcPr", st.TcPr)
	}
	for i := range st.TblStylePr {
		if err := encodeChild(e, "tblStylePr", &st.TblStylePr[i]); err != nil {
			return err
		}
	}

	if err := encodeRawExtras(e, st.Extra); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:style> and its children.
func (st *CT_Style) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Read attributes
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "type":
			st.Type = attr.Value
		case "default":
			v := parseBool(attr.Value)
			st.Default = &v
		case "customStyle":
			v := parseBool(attr.Value)
			st.CustomStyle = &v
		case "styleId":
			st.StyleID = attr.Value
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
			case "name":
				st.Name = &xmltypes.CT_String{}
				if err := d.DecodeElement(st.Name, &t); err != nil {
					return err
				}
			case "aliases":
				st.Aliases = &xmltypes.CT_String{}
				if err := d.DecodeElement(st.Aliases, &t); err != nil {
					return err
				}
			case "basedOn":
				st.BasedOn = &xmltypes.CT_String{}
				if err := d.DecodeElement(st.BasedOn, &t); err != nil {
					return err
				}
			case "next":
				st.Next = &xmltypes.CT_String{}
				if err := d.DecodeElement(st.Next, &t); err != nil {
					return err
				}
			case "link":
				st.Link = &xmltypes.CT_String{}
				if err := d.DecodeElement(st.Link, &t); err != nil {
					return err
				}
			case "autoRedefine":
				st.AutoRedefine = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.AutoRedefine, &t); err != nil {
					return err
				}
			case "hidden":
				st.Hidden = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.Hidden, &t); err != nil {
					return err
				}
			case "uiPriority":
				st.UIpriority = &xmltypes.CT_DecimalNumber{}
				if err := d.DecodeElement(st.UIpriority, &t); err != nil {
					return err
				}
			case "semiHidden":
				st.SemiHidden = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.SemiHidden, &t); err != nil {
					return err
				}
			case "unhideWhenUsed":
				st.UnhideWhenUsed = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.UnhideWhenUsed, &t); err != nil {
					return err
				}
			case "qFormat":
				st.QFormat = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.QFormat, &t); err != nil {
					return err
				}
			case "locked":
				st.Locked = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.Locked, &t); err != nil {
					return err
				}
			case "personal":
				st.Personal = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.Personal, &t); err != nil {
					return err
				}
			case "personalCompose":
				st.PersonalCompose = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.PersonalCompose, &t); err != nil {
					return err
				}
			case "personalReply":
				st.PersonalReply = &xmltypes.CT_OnOff{}
				if err := d.DecodeElement(st.PersonalReply, &t); err != nil {
					return err
				}
			case "rsid":
				st.Rsid = &xmltypes.CT_LongHexNumber{}
				if err := d.DecodeElement(st.Rsid, &t); err != nil {
					return err
				}
			case "pPr":
				st.PPr = &ppr.CT_PPrBase{}
				if err := d.DecodeElement(st.PPr, &t); err != nil {
					return err
				}
			case "rPr":
				st.RPr = &rpr.CT_RPrBase{}
				if err := d.DecodeElement(st.RPr, &t); err != nil {
					return err
				}
			case "tblPr":
				st.TblPr = &table.CT_TblPr{}
				if err := d.DecodeElement(st.TblPr, &t); err != nil {
					return err
				}
			case "trPr":
				st.TrPr = &table.CT_TrPr{}
				if err := d.DecodeElement(st.TrPr, &t); err != nil {
					return err
				}
			case "tcPr":
				st.TcPr = &table.CT_TcPr{}
				if err := d.DecodeElement(st.TcPr, &t); err != nil {
					return err
				}
			case "tblStylePr":
				var tsp CT_TblStylePr
				if err := d.DecodeElement(&tsp, &t); err != nil {
					return err
				}
				st.TblStylePr = append(st.TblStylePr, tsp)
			default:
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				st.Extra = append(st.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// =========================================================================
// CT_TblStylePr
// =========================================================================

// MarshalXML writes <w:tblStylePr> with children in order.
func (tsp *CT_TblStylePr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = appendNonEmpty(start.Attr, xmltypes.NSw, "type", tsp.Type)

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if tsp.PPr != nil {
		encodeChild(e, "pPr", tsp.PPr)
	}
	if tsp.RPr != nil {
		encodeChild(e, "rPr", tsp.RPr)
	}
	if tsp.TblPr != nil {
		encodeChild(e, "tblPr", tsp.TblPr)
	}
	if tsp.TrPr != nil {
		encodeChild(e, "trPr", tsp.TrPr)
	}
	if tsp.TcPr != nil {
		encodeChild(e, "tcPr", tsp.TcPr)
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads <w:tblStylePr>.
func (tsp *CT_TblStylePr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "type" {
			tsp.Type = attr.Value
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
			case "pPr":
				tsp.PPr = &ppr.CT_PPrBase{}
				if err := d.DecodeElement(tsp.PPr, &t); err != nil {
					return err
				}
			case "rPr":
				tsp.RPr = &rpr.CT_RPrBase{}
				if err := d.DecodeElement(tsp.RPr, &t); err != nil {
					return err
				}
			case "tblPr":
				tsp.TblPr = &table.CT_TblPr{}
				if err := d.DecodeElement(tsp.TblPr, &t); err != nil {
					return err
				}
			case "trPr":
				tsp.TrPr = &table.CT_TrPr{}
				if err := d.DecodeElement(tsp.TrPr, &t); err != nil {
					return err
				}
			case "tcPr":
				tsp.TcPr = &table.CT_TcPr{}
				if err := d.DecodeElement(tsp.TcPr, &t); err != nil {
					return err
				}
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

// =========================================================================
// Attribute helpers
// =========================================================================

func appendNonEmpty(attrs []xml.Attr, space, local, val string) []xml.Attr {
	if val == "" {
		return attrs
	}
	return append(attrs, xml.Attr{
		Name:  xml.Name{Space: space, Local: local},
		Value: val,
	})
}

func boolAttr(space, local string, v bool) xml.Attr {
	val := "0"
	if v {
		val = "1"
	}
	return xml.Attr{
		Name:  xml.Name{Space: space, Local: local},
		Value: val,
	}
}

func intAttr(space, local string, v int) xml.Attr {
	return xml.Attr{
		Name:  xml.Name{Space: space, Local: local},
		Value: intToStr(v),
	}
}

func parseBool(s string) bool {
	switch s {
	case "1", "true", "on":
		return true
	default:
		return false
	}
}

func parseInt(s string) int {
	var v int
	for _, c := range s {
		if c >= '0' && c <= '9' {
			v = v*10 + int(c-'0')
		}
	}
	// Handle negative
	if len(s) > 0 && s[0] == '-' {
		v = -v
	}
	return v
}

func intToStr(v int) string {
	if v == 0 {
		return "0"
	}
	neg := false
	if v < 0 {
		neg = true
		v = -v
	}
	buf := make([]byte, 0, 10)
	for v > 0 {
		buf = append(buf, byte('0'+v%10))
		v /= 10
	}
	if neg {
		buf = append(buf, '-')
	}
	// reverse
	for i, j := 0, len(buf)-1; i < j; i, j = i+1, j-1 {
		buf[i], buf[j] = buf[j], buf[i]
	}
	return string(buf)
}
