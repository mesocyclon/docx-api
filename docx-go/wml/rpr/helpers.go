package rpr

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// nsW is a shorthand for the WML namespace used throughout this package.
var nsW = xmltypes.NSw

// encodeChild marshals a non-nil child element with the given local name in the w: namespace.
func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: nsW, Local: local},
	})
}

// encodeRawXML writes a shared.RawXML element, replaying inner tokens
// through the encoder to produce valid XML.
func encodeRawXML(e *xml.Encoder, raw shared.RawXML) error {
	start := xml.StartElement{Name: raw.XMLName, Attr: raw.Attrs}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if len(raw.Inner) > 0 {
		innerDec := xml.NewDecoder(bytes.NewReader(raw.Inner))
		for {
			tok, err := innerDec.Token()
			if err != nil {
				break // EOF or error — stop replaying
			}
			if err := e.EncodeToken(xml.CopyToken(tok)); err != nil {
				return err
			}
		}
	}
	return e.EncodeToken(start.End())
}

// decodeUnknown captures an unrecognized child element as shared.RawXML.
func decodeUnknown(d *xml.Decoder, start *xml.StartElement) (shared.RawXML, error) {
	var raw shared.RawXML
	if err := d.DecodeElement(&raw, start); err != nil {
		return raw, err
	}
	return raw, nil
}
