// Package hdft implements CT_HdrFtr — the content model shared by
// <w:hdr> (header) and <w:ftr> (footer) parts in OOXML documents.
//
// CT_HdrFtr is a simple container of block-level elements (paragraphs,
// tables, SDTs, …).  The package relies on wml/shared factories to
// decode concrete block-level types and preserves unrecognised elements
// as shared.RawXML for lossless round-trip fidelity.
//
// Contract: C-18 in contracts.md.
package hdft

import (
	"bytes"
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// CT_HdrFtr is the complex type for both <w:hdr> and <w:ftr>.
// Per the spec a header/footer must contain at least one <w:p/>.
type CT_HdrFtr struct {
	// Content holds the block-level children (paragraphs, tables, etc.)
	// in document order.
	Content []shared.BlockLevelElement

	// Namespaces stores the xmlns attributes from the root element so
	// they survive a round-trip without loss.
	Namespaces []xml.Attr
}

// ---------------------------------------------------------------------------
// Parse / Serialize — public API
// ---------------------------------------------------------------------------

// Parse decodes the raw XML bytes of a header or footer part
// (word/header*.xml or word/footer*.xml) into a CT_HdrFtr value.
func Parse(data []byte) (*CT_HdrFtr, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var hf CT_HdrFtr
	if err := dec.Decode(&hf); err != nil {
		return nil, fmt.Errorf("hdft.Parse: %w", err)
	}
	return &hf, nil
}

// Serialize encodes the CT_HdrFtr back to XML bytes.
// rootName must be "w:hdr" or "w:ftr".
func Serialize(hf *CT_HdrFtr, rootName string) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	start := xml.StartElement{
		Name: xml.Name{Local: rootName},
	}

	if err := enc.EncodeElement(hf, start); err != nil {
		return nil, fmt.Errorf("hdft.Serialize: %w", err)
	}
	if err := enc.Flush(); err != nil {
		return nil, fmt.Errorf("hdft.Serialize: flush: %w", err)
	}
	return buf.Bytes(), nil
}

// ---------------------------------------------------------------------------
// XML codec
// ---------------------------------------------------------------------------

// UnmarshalXML decodes the root <w:hdr> or <w:ftr> element and all its
// block-level children.  Known element types are produced by
// shared.CreateBlockElement; anything unrecognised is stored as
// shared.RawXML.
func (hf *CT_HdrFtr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// Preserve every attribute of the root element, including xmlns
	// declarations, so we can replay them on marshal.
	hf.Namespaces = make([]xml.Attr, len(start.Attr))
	copy(hf.Namespaces, start.Attr)

	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			elem := shared.CreateBlockElement(t.Name)
			if elem != nil {
				// The factory gave us a typed value; let the
				// decoder fill it in.
				if err := d.DecodeElement(elem, &t); err != nil {
					return err
				}
			} else {
				// Unknown element — capture as RawXML.
				var raw shared.RawXML
				if err := d.DecodeElement(&raw, &t); err != nil {
					return err
				}
				elem = raw
			}
			hf.Content = append(hf.Content, elem)

		case xml.EndElement:
			return nil // closing </w:hdr> or </w:ftr>
		}
	}
}

// MarshalXML writes the <w:hdr> or <w:ftr> element.  The actual element
// name is determined by the caller-supplied start (overridden by
// Serialize).  Namespace declarations are restored from the value
// captured during unmarshal; for brand-new documents a minimal set of
// defaults is emitted.
func (hf *CT_HdrFtr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Restore namespace declarations from the original document when
	// available; otherwise provide a minimal default set.
	if len(hf.Namespaces) > 0 {
		start.Attr = hf.Namespaces
	} else {
		start.Attr = defaultNamespaces()
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	for _, elem := range hf.Content {
		if err := e.Encode(elem); err != nil {
			return fmt.Errorf("hdft: marshal block element: %w", err)
		}
	}

	return e.EncodeToken(start.End())
}

// defaultNamespaces returns the minimal xmlns attributes for a new
// header/footer part.
func defaultNamespaces() []xml.Attr {
	return []xml.Attr{
		{Name: xml.Name{Local: "xmlns:w"}, Value: xmltypes.NSw},
		{Name: xml.Name{Local: "xmlns:r"}, Value: xmltypes.NSr},
	}
}
