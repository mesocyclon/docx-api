package sectpr

import (
	"bytes"
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// MarshalXML writes CT_SectPr in strict xsd:sequence order.
func (s *CT_SectPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "sectPr"}

	// Attributes
	if s.RsidR != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "rsidR"}, Value: *s.RsidR,
		})
	}
	if s.RsidSect != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Space: xmltypes.NSw, Local: "rsidSect"}, Value: *s.RsidSect,
		})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 1. EG_HdrFtrReferences — headers first, then footers
	for i := range s.HeaderRefs {
		if err := marshalHdrFtrRef(e, "headerReference", &s.HeaderRefs[i]); err != nil {
			return err
		}
	}
	for i := range s.FooterRefs {
		if err := marshalHdrFtrRef(e, "footerReference", &s.FooterRefs[i]); err != nil {
			return err
		}
	}

	// 2. EG_SectPrContents — strict order
	if s.FootnotePr != nil {
		if err := marshalFtnProps(e, s.FootnotePr); err != nil {
			return err
		}
	}
	if s.EndnotePr != nil {
		if err := marshalEdnProps(e, s.EndnotePr); err != nil {
			return err
		}
	}
	if s.Type != nil {
		if err := encodeChild(e, "type", s.Type); err != nil {
			return err
		}
	}
	if s.PgSz != nil {
		if err := marshalPgSz(e, s.PgSz); err != nil {
			return err
		}
	}
	if s.PgMar != nil {
		if err := marshalPgMar(e, s.PgMar); err != nil {
			return err
		}
	}
	if s.PaperSrc != nil {
		if err := encodeChild(e, "paperSrc", s.PaperSrc); err != nil {
			return err
		}
	}
	if s.PgBorders != nil {
		if err := marshalPgBorders(e, s.PgBorders); err != nil {
			return err
		}
	}
	if s.LnNumType != nil {
		if err := encodeChild(e, "lnNumType", s.LnNumType); err != nil {
			return err
		}
	}
	if s.PgNumType != nil {
		if err := encodeChild(e, "pgNumType", s.PgNumType); err != nil {
			return err
		}
	}
	if s.Cols != nil {
		if err := marshalCols(e, s.Cols); err != nil {
			return err
		}
	}
	if s.FormProt != nil {
		if err := encodeOnOff(e, "formProt", s.FormProt); err != nil {
			return err
		}
	}
	if s.VAlign != nil {
		if err := encodeChild(e, "vAlign", s.VAlign); err != nil {
			return err
		}
	}
	if s.NoEndnote != nil {
		if err := encodeOnOff(e, "noEndnote", s.NoEndnote); err != nil {
			return err
		}
	}
	if s.TitlePg != nil {
		if err := encodeOnOff(e, "titlePg", s.TitlePg); err != nil {
			return err
		}
	}
	if s.TextDirection != nil {
		if err := encodeChild(e, "textDirection", s.TextDirection); err != nil {
			return err
		}
	}
	if s.Bidi != nil {
		if err := encodeOnOff(e, "bidi", s.Bidi); err != nil {
			return err
		}
	}
	if s.RtlGutter != nil {
		if err := encodeOnOff(e, "rtlGutter", s.RtlGutter); err != nil {
			return err
		}
	}
	if s.DocGrid != nil {
		if err := marshalDocGrid(e, s.DocGrid); err != nil {
			return err
		}
	}

	// 3. Extra (unknown/extension elements) — at the end
	for _, raw := range s.Extra {
		if err := marshalRawXML(e, raw); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_SectPr, preserving unknown elements as RawXML.
func (s *CT_SectPr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			v := attr.Value
			s.RsidR = &v
		case "rsidSect":
			v := attr.Value
			s.RsidSect = &v
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if err := s.decodeChild(d, t); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

func (s *CT_SectPr) decodeChild(d *xml.Decoder, t xml.StartElement) error {
	switch t.Name.Local {
	case "headerReference":
		var ref CT_HdrFtrRef
		if err := unmarshalHdrFtrRef(d, t, &ref); err != nil {
			return err
		}
		s.HeaderRefs = append(s.HeaderRefs, ref)
	case "footerReference":
		var ref CT_HdrFtrRef
		if err := unmarshalHdrFtrRef(d, t, &ref); err != nil {
			return err
		}
		s.FooterRefs = append(s.FooterRefs, ref)
	case "footnotePr":
		s.FootnotePr = &CT_FtnProps{}
		return d.DecodeElement(s.FootnotePr, &t)
	case "endnotePr":
		s.EndnotePr = &CT_EdnProps{}
		return d.DecodeElement(s.EndnotePr, &t)
	case "type":
		s.Type = &CT_SectType{}
		return d.DecodeElement(s.Type, &t)
	case "pgSz":
		s.PgSz = &CT_PageSz{}
		return unmarshalPgSz(d, t, s.PgSz)
	case "pgMar":
		s.PgMar = &CT_PageMar{}
		return unmarshalPgMar(d, t, s.PgMar)
	case "paperSrc":
		s.PaperSrc = &CT_PaperSource{}
		return d.DecodeElement(s.PaperSrc, &t)
	case "pgBorders":
		s.PgBorders = &CT_PageBorders{}
		return d.DecodeElement(s.PgBorders, &t)
	case "lnNumType":
		s.LnNumType = &CT_LineNumber{}
		return d.DecodeElement(s.LnNumType, &t)
	case "pgNumType":
		s.PgNumType = &CT_PageNumber{}
		return d.DecodeElement(s.PgNumType, &t)
	case "cols":
		s.Cols = &CT_Columns{}
		return unmarshalCols(d, t, s.Cols)
	case "formProt":
		s.FormProt = &xmltypes.CT_OnOff{}
		return d.DecodeElement(s.FormProt, &t)
	case "vAlign":
		s.VAlign = &CT_VerticalJc{}
		return d.DecodeElement(s.VAlign, &t)
	case "noEndnote":
		s.NoEndnote = &xmltypes.CT_OnOff{}
		return d.DecodeElement(s.NoEndnote, &t)
	case "titlePg":
		s.TitlePg = &xmltypes.CT_OnOff{}
		return d.DecodeElement(s.TitlePg, &t)
	case "textDirection":
		s.TextDirection = &CT_TextDirection{}
		return d.DecodeElement(s.TextDirection, &t)
	case "bidi":
		s.Bidi = &xmltypes.CT_OnOff{}
		return d.DecodeElement(s.Bidi, &t)
	case "rtlGutter":
		s.RtlGutter = &xmltypes.CT_OnOff{}
		return d.DecodeElement(s.RtlGutter, &t)
	case "docGrid":
		s.DocGrid = &CT_DocGrid{}
		return unmarshalDocGrid(d, t, s.DocGrid)
	default:
		// Unknown element → preserve as RawXML for round-trip
		var raw shared.RawXML
		if err := d.DecodeElement(&raw, &t); err != nil {
			return err
		}
		s.Extra = append(s.Extra, raw)
	}
	return nil
}

// --- Marshal helpers ---

func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

func encodeOnOff(e *xml.Encoder, local string, o *xmltypes.CT_OnOff) error {
	return o.MarshalXML(e, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

func marshalHdrFtrRef(e *xml.Encoder, local string, ref *CT_HdrFtrRef) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
		Attr: []xml.Attr{
			{Name: xml.Name{Space: xmltypes.NSw, Local: "type"}, Value: ref.Type},
			{Name: xml.Name{Space: xmltypes.NSr, Local: "id"}, Value: ref.RID},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func unmarshalHdrFtrRef(d *xml.Decoder, t xml.StartElement, ref *CT_HdrFtrRef) error {
	for _, attr := range t.Attr {
		switch {
		case attr.Name.Local == "type":
			ref.Type = attr.Value
		case attr.Name.Local == "id":
			ref.RID = attr.Value
		}
	}
	return d.Skip()
}

func marshalPgSz(e *xml.Encoder, p *CT_PageSz) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: "pgSz"},
	}
	start.Attr = append(start.Attr,
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "w"}, Value: strconv.Itoa(p.W)},
		xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "h"}, Value: strconv.Itoa(p.H)},
	)
	if p.Orient != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "orient"}, Value: *p.Orient},
		)
	}
	if p.Code != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "code"}, Value: strconv.Itoa(*p.Code)},
		)
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func unmarshalPgSz(d *xml.Decoder, t xml.StartElement, p *CT_PageSz) error {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "w":
			v, _ := strconv.Atoi(attr.Value)
			p.W = v
		case "h":
			v, _ := strconv.Atoi(attr.Value)
			p.H = v
		case "orient":
			v := attr.Value
			p.Orient = &v
		case "code":
			v, _ := strconv.Atoi(attr.Value)
			p.Code = &v
		}
	}
	return d.Skip()
}

func marshalPgMar(e *xml.Encoder, p *CT_PageMar) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: "pgMar"},
		Attr: []xml.Attr{
			{Name: xml.Name{Space: xmltypes.NSw, Local: "top"}, Value: strconv.Itoa(p.Top)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "right"}, Value: strconv.Itoa(p.Right)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "bottom"}, Value: strconv.Itoa(p.Bottom)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "left"}, Value: strconv.Itoa(p.Left)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "header"}, Value: strconv.Itoa(p.Header)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "footer"}, Value: strconv.Itoa(p.Footer)},
			{Name: xml.Name{Space: xmltypes.NSw, Local: "gutter"}, Value: strconv.Itoa(p.Gutter)},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func unmarshalPgMar(d *xml.Decoder, t xml.StartElement, p *CT_PageMar) error {
	for _, attr := range t.Attr {
		v, _ := strconv.Atoi(attr.Value)
		switch attr.Name.Local {
		case "top":
			p.Top = v
		case "right":
			p.Right = v
		case "bottom":
			p.Bottom = v
		case "left":
			p.Left = v
		case "header":
			p.Header = v
		case "footer":
			p.Footer = v
		case "gutter":
			p.Gutter = v
		}
	}
	return d.Skip()
}

func marshalDocGrid(e *xml.Encoder, g *CT_DocGrid) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: "docGrid"},
	}
	if g.Type != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "type"}, Value: *g.Type},
		)
	}
	if g.LinePitch != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "linePitch"}, Value: strconv.Itoa(*g.LinePitch)},
		)
	}
	if g.CharSpace != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "charSpace"}, Value: strconv.Itoa(*g.CharSpace)},
		)
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func unmarshalDocGrid(d *xml.Decoder, t xml.StartElement, g *CT_DocGrid) error {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "type":
			v := attr.Value
			g.Type = &v
		case "linePitch":
			v, _ := strconv.Atoi(attr.Value)
			g.LinePitch = &v
		case "charSpace":
			v, _ := strconv.Atoi(attr.Value)
			g.CharSpace = &v
		}
	}
	return d.Skip()
}

func marshalCols(e *xml.Encoder, c *CT_Columns) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: "cols"},
	}
	if c.EqualWidth != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "equalWidth"}, Value: boolStr(*c.EqualWidth)},
		)
	}
	if c.Space != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "space"}, Value: strconv.Itoa(*c.Space)},
		)
	}
	if c.Num != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "num"}, Value: strconv.Itoa(*c.Num)},
		)
	}
	if c.Sep != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "sep"}, Value: boolStr(*c.Sep)},
		)
	}

	if len(c.Col) == 0 {
		// Self-closing
		if err := e.EncodeToken(start); err != nil {
			return err
		}
		return e.EncodeToken(start.End())
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for i := range c.Col {
		if err := marshalCol(e, &c.Col[i]); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func marshalCol(e *xml.Encoder, c *CT_Column) error {
	start := xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: "col"},
	}
	if c.W != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "w"}, Value: strconv.Itoa(*c.W)},
		)
	}
	if c.Space != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "space"}, Value: strconv.Itoa(*c.Space)},
		)
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

func unmarshalCols(d *xml.Decoder, t xml.StartElement, c *CT_Columns) error {
	for _, attr := range t.Attr {
		switch attr.Name.Local {
		case "equalWidth":
			v := attr.Value == "1" || attr.Value == "true"
			c.EqualWidth = &v
		case "space":
			v, _ := strconv.Atoi(attr.Value)
			c.Space = &v
		case "num":
			v, _ := strconv.Atoi(attr.Value)
			c.Num = &v
		case "sep":
			v := attr.Value == "1" || attr.Value == "true"
			c.Sep = &v
		}
	}

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch tt := tok.(type) {
		case xml.StartElement:
			if tt.Name.Local == "col" {
				var col CT_Column
				for _, a := range tt.Attr {
					switch a.Name.Local {
					case "w":
						v, _ := strconv.Atoi(a.Value)
						col.W = &v
					case "space":
						v, _ := strconv.Atoi(a.Value)
						col.Space = &v
					}
				}
				if err := d.Skip(); err != nil {
					return err
				}
				c.Col = append(c.Col, col)
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

func marshalFtnProps(e *xml.Encoder, p *CT_FtnProps) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "footnotePr"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.Pos != nil {
		if err := encodeChild(e, "pos", p.Pos); err != nil {
			return err
		}
	}
	if p.NumFmt != nil {
		if err := encodeChild(e, "numFmt", p.NumFmt); err != nil {
			return err
		}
	}
	if p.NumStart != nil {
		if err := encodeChild(e, "numStart", p.NumStart); err != nil {
			return err
		}
	}
	if p.NumRestart != nil {
		if err := encodeChild(e, "numRestart", p.NumRestart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func marshalEdnProps(e *xml.Encoder, p *CT_EdnProps) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "endnotePr"}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if p.Pos != nil {
		if err := encodeChild(e, "pos", p.Pos); err != nil {
			return err
		}
	}
	if p.NumFmt != nil {
		if err := encodeChild(e, "numFmt", p.NumFmt); err != nil {
			return err
		}
	}
	if p.NumStart != nil {
		if err := encodeChild(e, "numStart", p.NumStart); err != nil {
			return err
		}
	}
	if p.NumRestart != nil {
		if err := encodeChild(e, "numRestart", p.NumRestart); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func marshalPgBorders(e *xml.Encoder, b *CT_PageBorders) error {
	start := xml.StartElement{Name: xml.Name{Space: xmltypes.NSw, Local: "pgBorders"}}
	if b.OffsetFrom != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "offsetFrom"}, Value: *b.OffsetFrom},
		)
	}
	if b.ZOrder != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "zOrder"}, Value: *b.ZOrder},
		)
	}
	if b.Display != nil {
		start.Attr = append(start.Attr,
			xml.Attr{Name: xml.Name{Space: xmltypes.NSw, Local: "display"}, Value: *b.Display},
		)
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if b.Top != nil {
		if err := encodeChild(e, "top", b.Top); err != nil {
			return err
		}
	}
	if b.Left != nil {
		if err := encodeChild(e, "left", b.Left); err != nil {
			return err
		}
	}
	if b.Bottom != nil {
		if err := encodeChild(e, "bottom", b.Bottom); err != nil {
			return err
		}
	}
	if b.Right != nil {
		if err := encodeChild(e, "right", b.Right); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// marshalRawXML restores an unknown element by re-playing its tokens.
func marshalRawXML(e *xml.Encoder, raw shared.RawXML) error {
	start := xml.StartElement{Name: raw.XMLName, Attr: raw.Attrs}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if len(raw.Inner) > 0 {
		if err := e.Flush(); err != nil {
			return err
		}
		innerDec := xml.NewDecoder(bytes.NewReader(raw.Inner))
		for {
			innerTok, err := innerDec.Token()
			if err != nil {
				break
			}
			if err := e.EncodeToken(xml.CopyToken(innerTok)); err != nil {
				return err
			}
		}
	}
	return e.EncodeToken(start.End())
}

func boolStr(b bool) string {
	if b {
		return "1"
	}
	return "0"
}
