package run

import (
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/wml/rpr"
	"github.com/vortex/docx-go/wml/shared"
)

// CT_R represents a run element (<w:r>), the atomic unit of text content
// inside a paragraph. A run contains optional run properties (rPr) followed
// by a sequence of inline content elements (text, breaks, drawings, etc.).
type CT_R struct {
	shared.ParagraphContentMarker

	RPr     *rpr.CT_RPr
	Content []shared.RunContent

	// Attributes
	RsidR   *string `xml:"rsidR,attr,omitempty"`
	RsidRPr *string `xml:"rsidRPr,attr,omitempty"`
	RsidDel *string `xml:"rsidDel,attr,omitempty"`
}

// MarshalXML serialises CT_R as <w:r> with attributes, optional <w:rPr>, and
// all run content elements in document order.
func (r *CT_R) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:r"}

	// Marshal rsid* attributes in the w: namespace.
	if r.RsidR != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Local: "w:rsidR"}, Value: *r.RsidR,
		})
	}
	if r.RsidRPr != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Local: "w:rsidRPr"}, Value: *r.RsidRPr,
		})
	}
	if r.RsidDel != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name: xml.Name{Local: "w:rsidDel"}, Value: *r.RsidDel,
		})
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// <w:rPr> (optional, always first child when present)
	if r.RPr != nil {
		if err := e.EncodeElement(r.RPr, xml.StartElement{
			Name: xml.Name{Local: "w:rPr"},
		}); err != nil {
			return err
		}
	}

	// Run content elements in order.
	for _, c := range r.Content {
		if err := marshalRunContent(e, c); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// marshalRunContent encodes a single RunContent element with the correct XML
// element name.
func marshalRunContent(e *xml.Encoder, c shared.RunContent) error {
	switch v := c.(type) {
	case *CT_Text:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:t"},
		})
	case *CT_Br:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:br"},
		})
	case *CT_Drawing:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:drawing"},
		})
	case *CT_FldChar:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:fldChar"},
		})
	case *CT_InstrText:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:instrText"},
		})
	case *CT_Sym:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:sym"},
		})
	case *CT_FtnEdnRef:
		// Use the original XMLName to distinguish footnoteReference/endnoteReference.
		name := v.XMLName
		if name.Local == "" {
			name.Local = "footnoteReference"
		}
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:" + name.Local},
		})
	case *CT_EmptyRunContent:
		return e.EncodeElement(v, xml.StartElement{
			Name: xml.Name{Local: "w:" + v.XMLName.Local},
		})
	case *CT_RawRunContent:
		return e.EncodeElement(v.Raw, xml.StartElement{Name: v.Raw.XMLName})
	case shared.RawXML:
		return e.EncodeElement(v, xml.StartElement{Name: v.XMLName})
	default:
		return fmt.Errorf("run: unknown RunContent type %T", c)
	}
}

// UnmarshalXML decodes <w:r> from XML. It reads attributes, optional <w:rPr>,
// and all inline content children.
func (r *CT_R) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Parse attributes.
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "rsidR":
			s := attr.Value
			r.RsidR = &s
		case "rsidRPr":
			s := attr.Value
			r.RsidRPr = &s
		case "rsidDel":
			s := attr.Value
			r.RsidDel = &s
		}
	}

	// Parse child elements.
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			child, err := decodeRunChild(d, t)
			if err != nil {
				return err
			}
			if child != nil {
				if rPr, ok := child.(*rpr.CT_RPr); ok {
					r.RPr = rPr
				} else if rc, ok := child.(shared.RunContent); ok {
					r.Content = append(r.Content, rc)
				}
			}

		case xml.EndElement:
			return nil // end of <w:r>
		}
	}
}

// decodeRunChild decodes a single child element of <w:r>.
// It returns either a *rpr.CT_RPr or a shared.RunContent.
func decodeRunChild(d *xml.Decoder, start xml.StartElement) (interface{}, error) {
	local := start.Name.Local

	switch local {
	case "rPr":
		rPr := &rpr.CT_RPr{}
		if err := d.DecodeElement(rPr, &start); err != nil {
			return nil, err
		}
		return rPr, nil

	case "t", "delText":
		t := &CT_Text{}
		if err := d.DecodeElement(t, &start); err != nil {
			return nil, err
		}
		return t, nil

	case "br":
		br := &CT_Br{}
		if err := d.DecodeElement(br, &start); err != nil {
			return nil, err
		}
		return br, nil

	case "drawing":
		dr := &CT_Drawing{}
		if err := d.DecodeElement(dr, &start); err != nil {
			return nil, err
		}
		return dr, nil

	case "fldChar":
		fc := &CT_FldChar{}
		if err := d.DecodeElement(fc, &start); err != nil {
			return nil, err
		}
		return fc, nil

	case "instrText", "delInstrText":
		it := &CT_InstrText{}
		if err := d.DecodeElement(it, &start); err != nil {
			return nil, err
		}
		return it, nil

	case "sym":
		s := &CT_Sym{}
		if err := d.DecodeElement(s, &start); err != nil {
			return nil, err
		}
		return s, nil

	case "footnoteReference", "endnoteReference":
		ref := &CT_FtnEdnRef{}
		ref.XMLName = xml.Name{Space: start.Name.Space, Local: local}
		if err := d.DecodeElement(ref, &start); err != nil {
			return nil, err
		}
		return ref, nil

	case "commentReference":
		ref := &CT_FtnEdnRef{}
		ref.XMLName = xml.Name{Space: start.Name.Space, Local: local}
		if err := d.DecodeElement(ref, &start); err != nil {
			return nil, err
		}
		return ref, nil

	default:
		// Check if this is a known empty element.
		if emptyElementNames[local] {
			empty := &CT_EmptyRunContent{
				XMLName: xml.Name{Space: start.Name.Space, Local: local},
			}
			if err := d.Skip(); err != nil {
				return nil, err
			}
			return empty, nil
		}

		// Unknown element — preserve as RawXML for round-trip fidelity.
		var raw shared.RawXML
		if err := d.DecodeElement(&raw, &start); err != nil {
			return nil, err
		}
		return &CT_RawRunContent{Raw: raw}, nil
	}
}

// Compile-time interface check.
var _ xml.Marshaler = (*CT_R)(nil)
var _ xml.Unmarshaler = (*CT_R)(nil)
