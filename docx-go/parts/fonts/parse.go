package fonts

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// Parse decodes the raw bytes of word/fontTable.xml into a CT_FontsList.
// It uses a NormalizingDecoder to accept both Transitional and Strict
// namespace URIs.
func Parse(data []byte) (*CT_FontsList, error) {
	dec := xmltypes.NewNormalizingDecoder(bytes.NewReader(data))

	var fl CT_FontsList

	// Advance to the root <w:fonts> start element.
	for {
		tok, err := dec.Token()
		if err != nil {
			return nil, err
		}
		if start, ok := tok.(xml.StartElement); ok {
			if err := dec.DecodeElement(&fl, &start); err != nil {
				return nil, err
			}
			break
		}
	}

	return &fl, nil
}

// Serialize encodes CT_FontsList back into XML bytes suitable for
// word/fontTable.xml, including the XML declaration.
func Serialize(fl *CT_FontsList) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)

	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(fl); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}
