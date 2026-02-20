// Package document implements Parse/Serialize for the main document part
// (word/document.xml) of a DOCX package.
//
// Contract: C-20 in contracts.md
// Dependencies: wml/body, opc (opc is listed as a dependency in the contract
// but is not directly used here — it is consumed by the packaging layer that
// calls Parse/Serialize and resolves relationships).
package document

import (
	"bytes"
	"encoding/xml"
	"fmt"

	"github.com/vortex/docx-go/wml/body"
)

// xmlHeader is the standard XML declaration prepended to every serialized part.
const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

// Parse deserializes raw XML bytes (word/document.xml) into a typed
// body.CT_Document structure.
//
// CT_Document implements custom UnmarshalXML (see wml/body) that:
//   - preserves namespace declarations from the root <w:document> element,
//   - delegates body content parsing to CT_Body (paragraphs, tables, sectPr),
//   - captures unknown elements as shared.RawXML for round-trip fidelity.
func Parse(data []byte) (*body.CT_Document, error) {
	doc := &body.CT_Document{}
	if err := xml.NewDecoder(bytes.NewReader(data)).Decode(doc); err != nil {
		return nil, fmt.Errorf("parts/document: parse: %w", err)
	}
	return doc, nil
}

// Serialize marshals a body.CT_Document back into XML bytes suitable for
// writing to word/document.xml inside a DOCX ZIP package.
//
// CT_Document implements custom MarshalXML (see wml/body) that:
//   - emits the root <w:document> with all preserved (or default) namespace
//     declarations,
//   - serializes <w:body> children in document order,
//   - appends Extra (unknown) elements for round-trip preservation.
//
// The returned bytes are prefixed with the standard XML declaration.
func Serialize(doc *body.CT_Document) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString(xmlHeader)

	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(doc); err != nil {
		return nil, fmt.Errorf("parts/document: serialize: %w", err)
	}
	if err := enc.Flush(); err != nil {
		return nil, fmt.Errorf("parts/document: flush: %w", err)
	}

	return buf.Bytes(), nil
}
