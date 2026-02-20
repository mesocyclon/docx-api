package table

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// fieldMapping describes a Go struct field → XML element name mapping.
type fieldMapping struct {
	GoField  string
	XMLLocal string
}

// nsw is a shorthand for the main WML namespace.
var nsw = xmltypes.NSw

// encodeChild marshals a non-nil value as a child element in the w: namespace.
func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: nsw, Local: local},
	})
}

// encodeRawSlice marshals a slice of RawXML elements, preserving inner content.
func encodeRawSlice(e *xml.Encoder, extras []shared.RawXML) error {
	for _, raw := range extras {
		start := xml.StartElement{Name: raw.XMLName, Attr: raw.Attrs}
		if err := e.EncodeToken(start); err != nil {
			return err
		}
		if len(raw.Inner) > 0 {
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
		if err := e.EncodeToken(start.End()); err != nil {
			return err
		}
	}
	return nil
}

// decodeRawElement captures an unknown XML element as shared.RawXML.
func decodeRawElement(d *xml.Decoder, start *xml.StartElement) (shared.RawXML, error) {
	var raw shared.RawXML
	if err := d.DecodeElement(&raw, start); err != nil {
		return raw, err
	}
	return raw, nil
}

// decodeBorderChild decodes a CT_Border child element.
func decodeBorderChild(d *xml.Decoder, t *xml.StartElement) (*xmltypes.CT_Border, error) {
	b := &xmltypes.CT_Border{}
	if err := d.DecodeElement(b, t); err != nil {
		return nil, err
	}
	return b, nil
}

// decodeTblWidthChild decodes a CT_TblWidth child element.
func decodeTblWidthChild(d *xml.Decoder, t *xml.StartElement) (*CT_TblWidth, error) {
	w := &CT_TblWidth{}
	if err := d.DecodeElement(w, t); err != nil {
		return nil, err
	}
	return w, nil
}
