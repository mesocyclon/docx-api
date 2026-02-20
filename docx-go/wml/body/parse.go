package body

import (
	"bytes"
	"encoding/xml"
)

// Parse decodes a document.xml byte slice into a CT_Document.
func Parse(data []byte) (*CT_Document, error) {
	doc := &CT_Document{}
	if err := xml.Unmarshal(data, doc); err != nil {
		return nil, err
	}
	return doc, nil
}

// Serialize encodes a CT_Document back to XML bytes, including the XML
// declaration header expected by the OPC package.
func Serialize(doc *CT_Document) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)

	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(doc); err != nil {
		return nil, err
	}
	if err := enc.Flush(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}
