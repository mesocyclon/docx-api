package styles

import (
	"bytes"
	"encoding/xml"
)

// Parse deserialises raw XML bytes into a CT_Styles structure.
func Parse(data []byte) (*CT_Styles, error) {
	var s CT_Styles
	if err := xml.Unmarshal(data, &s); err != nil {
		return nil, err
	}
	return &s, nil
}

// Serialize serialises a CT_Styles structure back into XML bytes
// with the standard XML declaration header.
func Serialize(s *CT_Styles) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)
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
