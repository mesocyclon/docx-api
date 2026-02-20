package dml

import (
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
)

// encodeElement is a convenience wrapper that marshals v inside a StartElement
// with the given namespace and local name.
func encodeElement(e *xml.Encoder, space, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: space, Local: local},
	})
}

// encodeTextElement writes <ns:local>text</ns:local>.
func encodeTextElement(e *xml.Encoder, space, local, text string) error {
	start := xml.StartElement{Name: xml.Name{Space: space, Local: local}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := e.EncodeToken(xml.CharData([]byte(text))); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// marshalExtras writes a slice of RawXML elements to the encoder.
func marshalExtras(e *xml.Encoder, extras []shared.RawXML) error {
	for i := range extras {
		if err := e.EncodeElement(&extras[i], xml.StartElement{Name: extras[i].XMLName}); err != nil {
			return err
		}
	}
	return nil
}

// boolAttr creates an xml.Attr with "0" or "1".
func boolAttr(name string, v bool) xml.Attr {
	val := "0"
	if v {
		val = "1"
	}
	return xml.Attr{Name: xml.Name{Local: name}, Value: val}
}

// intAttr creates an xml.Attr from an int.
func intAttr(name string, v int) xml.Attr {
	return xml.Attr{Name: xml.Name{Local: name}, Value: strconv.Itoa(v)}
}

// parseBool interprets "1", "true", "on" as true; everything else as false.
func parseBool(s string) bool {
	switch s {
	case "1", "true", "on":
		return true
	}
	return false
}
