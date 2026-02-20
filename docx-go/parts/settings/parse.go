package settings

import (
	"bytes"
	"encoding/xml"
)

// xmlDeclaration is prepended to serialised output.
const xmlDeclaration = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

// Parse deserialises raw XML bytes into a CT_Settings structure.
func Parse(data []byte) (*CT_Settings, error) {
	var s CT_Settings
	if err := xml.Unmarshal(data, &s); err != nil {
		return nil, err
	}
	return &s, nil
}

// Serialize serialises a CT_Settings structure back to XML bytes,
// preserving original element order and namespace declarations.
func Serialize(s *CT_Settings) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xmlDeclaration)

	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(s); err != nil {
		return nil, err
	}
	if err := enc.Flush(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}
