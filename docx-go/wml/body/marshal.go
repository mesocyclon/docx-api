package body

import (
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// welem builds an xml.Name in the WML namespace.
// Go's encoder tracks prefix bindings registered via {Space:"xmlns",Local:"w"}
// attrs on ancestor elements.  When it encounters an element name with
// Space == NSw it reuses the "w" prefix → <w:body>, <w:p>, etc.
func welem(local string) xml.Name {
	return xml.Name{Space: xmltypes.NSw, Local: local}
}

// nsattr builds a namespace-declaration attribute that the Go encoder
// recognises and registers in its prefix table: {Space:"xmlns", Local:prefix}.
func nsattr(prefix, uri string) xml.Attr {
	return xml.Attr{Name: xml.Name{Space: "xmlns", Local: prefix}, Value: uri}
}

// ---------------------------------------------------------------------------
// CT_Document — MarshalXML
// ---------------------------------------------------------------------------

// MarshalXML writes the <w:document> root element with preserved namespaces.
func (doc *CT_Document) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = welem("document")

	if len(doc.Namespaces) > 0 {
		start.Attr = doc.Namespaces
	} else {
		start.Attr = defaultDocumentNamespaces()
	}

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Body
	if doc.Body != nil {
		bodyStart := xml.StartElement{Name: welem("body")}
		if err := e.EncodeElement(doc.Body, bodyStart); err != nil {
			return err
		}
	}

	// Extra (unknown children of <w:document> besides <w:body>)
	for i := range doc.Extra {
		if err := e.EncodeElement(&doc.Extra[i], xml.StartElement{Name: doc.Extra[i].XMLName}); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// CT_Body — MarshalXML
// ---------------------------------------------------------------------------

// MarshalXML writes the <w:body> element in document order:
// block-level content first, then the trailing sectPr.
func (b *CT_Body) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Block-level content in order
	for _, el := range b.Content {
		if err := marshalBlockElement(e, el); err != nil {
			return err
		}
	}

	// Trailing section properties
	if b.SectPr != nil {
		sectStart := xml.StartElement{Name: welem("sectPr")}
		if err := e.EncodeElement(b.SectPr, sectStart); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// CT_SdtBlock — MarshalXML
// ---------------------------------------------------------------------------

// MarshalXML writes the <w:sdt> element.
func (sdt *CT_SdtBlock) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if sdt.SdtPr != nil {
		if err := e.EncodeElement(sdt.SdtPr, xml.StartElement{Name: welem("sdtPr")}); err != nil {
			return err
		}
	}
	if sdt.SdtEndPr != nil {
		if err := e.EncodeElement(sdt.SdtEndPr, xml.StartElement{Name: welem("sdtEndPr")}); err != nil {
			return err
		}
	}

	// sdtContent wrapper
	if len(sdt.SdtContent) > 0 {
		contentStart := xml.StartElement{Name: welem("sdtContent")}
		if err := e.EncodeToken(contentStart); err != nil {
			return err
		}
		for _, el := range sdt.SdtContent {
			if err := marshalBlockElement(e, el); err != nil {
				return err
			}
		}
		if err := e.EncodeToken(contentStart.End()); err != nil {
			return err
		}
	}

	return e.EncodeToken(start.End())
}

// ---------------------------------------------------------------------------
// marshalBlockElement dispatches a BlockLevelElement to the encoder.
// ---------------------------------------------------------------------------

func marshalBlockElement(e *xml.Encoder, el interface{}) error {
	switch v := el.(type) {
	case ParagraphElement:
		return e.EncodeElement(v.P, xml.StartElement{Name: welem("p")})
	case TableElement:
		return e.EncodeElement(v.T, xml.StartElement{Name: welem("tbl")})
	case SdtBlockElement:
		return e.EncodeElement(v.Sdt, xml.StartElement{Name: welem("sdt")})
	case RawBlockElement:
		return e.EncodeElement(&v.Raw, xml.StartElement{Name: v.Raw.XMLName})
	}
	return nil
}

// ---------------------------------------------------------------------------
// Default namespaces for new documents
// ---------------------------------------------------------------------------

// defaultDocumentNamespaces returns the standard set of xmlns declarations
// for a new document.xml.  Attributes use {Space:"xmlns", Local:prefix} so
// Go's encoder registers the prefix bindings and all child elements that
// reference the same Space URI reuse the declared prefixes.
func defaultDocumentNamespaces() []xml.Attr {
	return []xml.Attr{
		nsattr("wpc", xmltypes.NSwpc),
		nsattr("mc", xmltypes.NSmc),
		nsattr("o", xmltypes.NSo),
		nsattr("r", xmltypes.NSr),
		nsattr("m", xmltypes.NSm),
		nsattr("v", xmltypes.NSv),
		nsattr("wp14", xmltypes.NSwp14),
		nsattr("wp", xmltypes.NSwp),
		nsattr("w10", xmltypes.NSw10),
		nsattr("w", xmltypes.NSw),
		nsattr("w14", xmltypes.NSw14),
		nsattr("w15", xmltypes.NSw15),
		nsattr("wpg", xmltypes.NSwpg),
		nsattr("wpi", xmltypes.NSwpi),
		nsattr("wne", xmltypes.NSwne),
		nsattr("wps", xmltypes.NSwps),
		// mc:Ignorable is a regular (non-xmlns) attribute in the mc namespace.
		{Name: xml.Name{Space: xmltypes.NSmc, Local: "Ignorable"}, Value: "w14 w15 wp14"},
	}
}
