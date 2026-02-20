package ppr

import (
	"bytes"
	"encoding/xml"
	"strconv"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// encodeChild encodes a child element in the w: namespace.
func encodeChild(e *xml.Encoder, local string, v interface{}) error {
	return e.EncodeElement(v, xml.StartElement{
		Name: xml.Name{Space: xmltypes.NSw, Local: local},
	})
}

// encodeRaw re-encodes a RawXML element preserving its original structure.
func encodeRaw(e *xml.Encoder, raw shared.RawXML) error {
	start := xml.StartElement{Name: raw.XMLName, Attr: raw.Attrs}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if len(raw.Inner) > 0 {
		e.Flush()
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
	return e.EncodeToken(start.End())
}

// intToStr converts int to string.
func intToStr(i int) string {
	return strconv.Itoa(i)
}

// strToInt converts string to int (0 on error).
func strToInt(s string) int {
	v, _ := strconv.Atoi(s)
	return v
}
