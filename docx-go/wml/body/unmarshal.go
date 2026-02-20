package body

import (
	"encoding/xml"

	"github.com/vortex/docx-go/wml/para"
	"github.com/vortex/docx-go/wml/sectpr"
	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/wml/table"
	"github.com/vortex/docx-go/xmltypes"
)

// ---------------------------------------------------------------------------
// CT_Document — UnmarshalXML
// ---------------------------------------------------------------------------

// UnmarshalXML parses the <w:document> root element, preserving all namespace
// declarations from the start tag.
func (doc *CT_Document) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Preserve all attributes (including xmlns:* declarations) so they
	// survive a round-trip.  See patterns.md §6.
	doc.Namespaces = make([]xml.Attr, len(start.Attr))
	copy(doc.Namespaces, start.Attr)

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "body":
				doc.Body = &CT_Body{}
				if err := d.DecodeElement(doc.Body, &t); err != nil {
					return err
				}
			default:
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				doc.Extra = append(doc.Extra, raw)
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ---------------------------------------------------------------------------
// CT_Body — UnmarshalXML
// ---------------------------------------------------------------------------

// UnmarshalXML parses the <w:body> element.  It dispatches child elements to
// typed wrappers for known elements (p, tbl, sdt) and stores everything else
// as RawBlockElement for round-trip fidelity.
//
// The last <w:sectPr> child is separated out into CT_Body.SectPr.
func (b *CT_Body) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if err := b.decodeChild(d, t); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

// decodeChild handles a single child element of <w:body>.
func (b *CT_Body) decodeChild(d *xml.Decoder, t xml.StartElement) error {
	// Only match elements in the w: namespace (or no namespace for
	// compatibility).
	if isWML(t.Name) {
		switch t.Name.Local {
		case "p":
			p := &para.CT_P{}
			if err := d.DecodeElement(p, &t); err != nil {
				return err
			}
			b.Content = append(b.Content, ParagraphElement{P: p})
			return nil

		case "tbl":
			tbl := &table.CT_Tbl{}
			if err := d.DecodeElement(tbl, &t); err != nil {
				return err
			}
			b.Content = append(b.Content, TableElement{T: tbl})
			return nil

		case "sdt":
			sdt := &CT_SdtBlock{}
			if err := d.DecodeElement(sdt, &t); err != nil {
				return err
			}
			b.Content = append(b.Content, SdtBlockElement{Sdt: sdt})
			return nil

		case "sectPr":
			// Body-level sectPr — always the last element, stored separately.
			b.SectPr = &sectpr.CT_SectPr{}
			return d.DecodeElement(b.SectPr, &t)
		}
	}

	// Unknown / extension element → preserve as RawXML.
	var raw shared.RawXML
	if err := d.DecodeElement(&raw, &t); err != nil {
		return err
	}
	b.Content = append(b.Content, RawBlockElement{Raw: raw})
	return nil
}

// ---------------------------------------------------------------------------
// CT_SdtBlock — UnmarshalXML
// ---------------------------------------------------------------------------

// UnmarshalXML parses a block-level <w:sdt> element.
func (sdt *CT_SdtBlock) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if isWML(t.Name) {
				switch t.Name.Local {
				case "sdtPr":
					pr := &shared.RawXML{}
					if err := d.DecodeElement(pr, &t); err != nil {
						return err
					}
					sdt.SdtPr = pr
					continue
				case "sdtEndPr":
					epr := &shared.RawXML{}
					if err := d.DecodeElement(epr, &t); err != nil {
						return err
					}
					sdt.SdtEndPr = epr
					continue
				case "sdtContent":
					if err := decodeSdtContent(d, sdt); err != nil {
						return err
					}
					continue
				}
			}
			// Unknown child of <w:sdt> — skip gracefully.
			if err := d.Skip(); err != nil {
				return err
			}
		case xml.EndElement:
			return nil
		}
	}
}

// decodeSdtContent reads the children of <w:sdtContent> and appends them to
// sdt.SdtContent.  The child elements are the same as those in <w:body>.
func decodeSdtContent(d *xml.Decoder, sdt *CT_SdtBlock) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			el, err := decodeBlockChild(d, t)
			if err != nil {
				return err
			}
			sdt.SdtContent = append(sdt.SdtContent, el)
		case xml.EndElement:
			return nil // end of <w:sdtContent>
		}
	}
}

// decodeBlockChild decodes a single block-level child (p, tbl, sdt, or raw).
func decodeBlockChild(d *xml.Decoder, t xml.StartElement) (shared.BlockLevelElement, error) {
	if isWML(t.Name) {
		switch t.Name.Local {
		case "p":
			p := &para.CT_P{}
			if err := d.DecodeElement(p, &t); err != nil {
				return nil, err
			}
			return ParagraphElement{P: p}, nil
		case "tbl":
			tbl := &table.CT_Tbl{}
			if err := d.DecodeElement(tbl, &t); err != nil {
				return nil, err
			}
			return TableElement{T: tbl}, nil
		case "sdt":
			sdt := &CT_SdtBlock{}
			if err := d.DecodeElement(sdt, &t); err != nil {
				return nil, err
			}
			return SdtBlockElement{Sdt: sdt}, nil
		}
	}

	var raw shared.RawXML
	if err := d.DecodeElement(&raw, &t); err != nil {
		return nil, err
	}
	return RawBlockElement{Raw: raw}, nil
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// isWML returns true if the name is in the WML namespace (or empty namespace,
// which happens when encoding/xml strips the prefix).
func isWML(name xml.Name) bool {
	return name.Space == xmltypes.NSw || name.Space == ""
}
