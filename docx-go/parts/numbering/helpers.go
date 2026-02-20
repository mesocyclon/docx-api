package numbering

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// encodeChild encodes a non-nil child element with the given local name
// under the WML namespace.
func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

// encodeRawExtras writes all Extra (RawXML) elements to the encoder.
func encodeRawExtras(e *xml.Encoder, extras []shared.RawXML) error {
	for _, raw := range extras {
		if err := raw.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}
	return nil
}

// decodeRawXML captures an unrecognised child element as shared.RawXML.
func decodeRawXML(d *xml.Decoder, start xml.StartElement) (shared.RawXML, error) {
	var raw shared.RawXML
	if err := raw.UnmarshalXML(d, start); err != nil {
		return raw, err
	}
	return raw, nil
}

// skipToEnd consumes all remaining tokens until the corresponding EndElement.
// Used when a decoder needs to drain the current element without processing it.
func skipToEnd(d *xml.Decoder, start xml.StartElement) error {
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)
	depth := 1
	for depth > 0 {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		if err := enc.EncodeToken(xml.CopyToken(tok)); err != nil {
			return err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return enc.Flush()
}
