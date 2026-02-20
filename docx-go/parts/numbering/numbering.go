// Package numbering implements the w:numbering part of an OOXML word-processing
// document (numbering.xml). It provides typed access to abstract numbering
// definitions, concrete numbering instances, and level overrides while
// preserving unrecognised XML elements for round-trip fidelity.
//
// See contracts.md C-22 for the public API contract and patterns.md §2.9, §3,
// §4 for implementation patterns.
package numbering

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_Numbering is the root element of the numbering part (w:numbering).
type CT_Numbering struct {
	AbstractNum []CT_AbstractNum
	Num         []CT_Num
	Extra       []shared.RawXML
}

// MarshalXML serialises CT_Numbering. AbstractNum elements are written before
// Num elements, followed by any Extra (unknown) elements, matching the XSD
// sequence order.
func (n *CT_Numbering) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Space: xmltypes.NSw, Local: "numbering"}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for i := range n.AbstractNum {
		if err := n.AbstractNum[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	for i := range n.Num {
		if err := n.Num[i].MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	if err := encodeRawExtras(e, n.Extra); err != nil {
		return err
	}

	return e.EncodeToken(start.End())
}

// UnmarshalXML reads CT_Numbering from XML, dispatching known children to
// typed slices and capturing unknown elements as RawXML.
func (n *CT_Numbering) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "abstractNum":
				var an CT_AbstractNum
				if err := an.UnmarshalXML(d, t); err != nil {
					return err
				}
				n.AbstractNum = append(n.AbstractNum, an)
			case "num":
				var num CT_Num
				if err := num.UnmarshalXML(d, t); err != nil {
					return err
				}
				n.Num = append(n.Num, num)
			default:
				raw, err := decodeRawXML(d, t)
				if err != nil {
					return err
				}
				n.Extra = append(n.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// Parse decodes numbering.xml bytes into a CT_Numbering structure.
func Parse(data []byte) (*CT_Numbering, error) {
	var n CT_Numbering
	if err := xml.NewDecoder(bytes.NewReader(data)).Decode(&n); err != nil {
		return nil, err
	}
	return &n, nil
}

// Serialize encodes a CT_Numbering structure back to XML bytes,
// including the standard XML declaration header.
func Serialize(n *CT_Numbering) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")
	if err := enc.Encode(n); err != nil {
		return nil, err
	}
	if err := enc.Flush(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
