package para

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// CT_Hyperlink — <w:hyperlink>
// ---------------------------------------------------------------------------

// CT_Hyperlink represents a hyperlink element inside a paragraph.
// It may reference an external URL via r:id or an internal anchor.
type CT_Hyperlink struct {
	RID     *string                   `xml:"id,attr,omitempty"`     // r:id for external hyperlinks
	Anchor  *string                   `xml:"anchor,attr,omitempty"` // for internal bookmarks
	Content []shared.ParagraphContent // runs, fields inside hyperlink
}

func (h *CT_Hyperlink) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if h.RID != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: xmltypes.NSr, Local: "id"},
			Value: *h.RID,
		})
	}
	if h.Anchor != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: xmltypes.NSw, Local: "anchor"},
			Value: *h.Anchor,
		})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if err := marshalParagraphContent(e, h.Content); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

func (h *CT_Hyperlink) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch {
		case attr.Name.Local == "id":
			s := attr.Value
			h.RID = &s
		case attr.Name.Local == "anchor":
			s := attr.Value
			h.Anchor = &s
		}
	}

	var err error
	h.Content, err = unmarshalParagraphContent(d)
	return err
}

// ---------------------------------------------------------------------------
// CT_SimpleField — <w:fldSimple>
// ---------------------------------------------------------------------------

// CT_SimpleField represents a simple field code.
type CT_SimpleField struct {
	Instr   string `xml:"instr,attr"`
	Content []shared.ParagraphContent
}

func (f *CT_SimpleField) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Attr = append(start.Attr, xml.Attr{
		Name:  xml.Name{Space: xmltypes.NSw, Local: "instr"},
		Value: f.Instr,
	})

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if err := marshalParagraphContent(e, f.Content); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

func (f *CT_SimpleField) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "instr" {
			f.Instr = attr.Value
		}
	}

	var err error
	f.Content, err = unmarshalParagraphContent(d)
	return err
}

// ---------------------------------------------------------------------------
// CT_SdtRun — <w:sdt> at paragraph (inline) level
// ---------------------------------------------------------------------------

// CT_SdtRun represents a structured document tag (content control) at the
// paragraph level.  sdtPr and sdtEndPr are stored as raw XML because the
// full SDT property model is very complex.
type CT_SdtRun struct {
	SdtPr      *shared.RawXML
	SdtEndPr   *shared.RawXML
	SdtContent []shared.ParagraphContent
}

func (s *CT_SdtRun) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if s.SdtPr != nil {
		if err := e.EncodeElement(s.SdtPr, xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "sdtPr"},
		}); err != nil {
			return err
		}
	}
	if s.SdtEndPr != nil {
		if err := e.EncodeElement(s.SdtEndPr, xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "sdtEndPr"},
		}); err != nil {
			return err
		}
	}

	// sdtContent wrapper
	if len(s.SdtContent) > 0 {
		contentStart := xml.StartElement{
			Name: xml.Name{Space: xmltypes.NSw, Local: "sdtContent"},
		}
		if err := e.EncodeToken(contentStart); err != nil {
			return err
		}
		if err := marshalParagraphContent(e, s.SdtContent); err != nil {
			return err
		}
		if err := e.EncodeToken(contentStart.End()); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

func (s *CT_SdtRun) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sdtPr":
				raw := &shared.RawXML{}
				if err := d.DecodeElement(raw, &t); err != nil {
					return err
				}
				s.SdtPr = raw
			case "sdtEndPr":
				raw := &shared.RawXML{}
				if err := d.DecodeElement(raw, &t); err != nil {
					return err
				}
				s.SdtEndPr = raw
			case "sdtContent":
				s.SdtContent, err = unmarshalParagraphContent(d)
				if err != nil {
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
