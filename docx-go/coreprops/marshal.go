package coreprops

import (
	"bytes"
	"encoding/xml"

	"github.com/vortex/docx-go/xmltypes"
)

// marshalCoreXML serializes core properties with correct namespace prefixes.
// Go's encoding/xml does not emit user-friendly prefixes for multiple namespaces,
// so we use a manual encoder approach.
func marshalCoreXML(x *xmlCoreProperties) ([]byte, error) {
	var buf bytes.Buffer
	e := xml.NewEncoder(&buf)
	e.Indent("", "  ")

	start := xml.StartElement{
		Name: xml.Name{Local: "cp:coreProperties"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns:cp"}, Value: xmltypes.NScp},
			{Name: xml.Name{Local: "xmlns:dc"}, Value: xmltypes.NSdc},
			{Name: xml.Name{Local: "xmlns:dcterms"}, Value: xmltypes.NSdcterms},
			{Name: xml.Name{Local: "xmlns:dcmitype"}, Value: xmltypes.NSdcmitype},
			{Name: xml.Name{Local: "xmlns:xsi"}, Value: xmltypes.NSxsi},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return nil, err
	}

	// dc:title
	encodeSimple(e, "dc:title", x.Title)
	// dc:subject
	encodeSimple(e, "dc:subject", x.Subject)
	// dc:creator
	encodeSimple(e, "dc:creator", x.Creator)
	// cp:keywords
	encodeSimple(e, "cp:keywords", x.Keywords)
	// dc:description
	encodeSimple(e, "dc:description", x.Description)
	// cp:lastModifiedBy
	encodeSimple(e, "cp:lastModifiedBy", x.LastModifiedBy)
	// cp:revision
	encodeSimple(e, "cp:revision", x.Revision)

	// dcterms:created with xsi:type
	if x.Created != nil {
		encodeDatetime(e, "dcterms:created", x.Created)
	}
	// dcterms:modified with xsi:type
	if x.Modified != nil {
		encodeDatetime(e, "dcterms:modified", x.Modified)
	}

	// cp:category (optional)
	if x.Category != "" {
		encodeSimple(e, "cp:category", x.Category)
	}
	// cp:contentStatus (optional)
	if x.ContentStatus != "" {
		encodeSimple(e, "cp:contentStatus", x.ContentStatus)
	}

	if err := e.EncodeToken(start.End()); err != nil {
		return nil, err
	}
	if err := e.Flush(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// marshalAppXML serializes app properties with correct namespace prefixes.
func marshalAppXML(x *xmlAppProperties) ([]byte, error) {
	var buf bytes.Buffer
	e := xml.NewEncoder(&buf)
	e.Indent("", "  ")

	start := xml.StartElement{
		Name: xml.Name{Local: "Properties"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns"}, Value: xmltypes.NSep},
			{Name: xml.Name{Local: "xmlns:vt"}, Value: xmltypes.NSvt},
		},
	}
	if err := e.EncodeToken(start); err != nil {
		return nil, err
	}

	encodeSimple(e, "Template", x.Template)
	encodeSimple(e, "TotalTime", x.TotalTime)
	encodeSimple(e, "Pages", x.Pages)
	encodeSimple(e, "Words", x.Words)
	encodeSimple(e, "Characters", x.Characters)
	encodeSimple(e, "Application", x.Application)
	encodeSimple(e, "DocSecurity", x.DocSecurity)
	encodeSimple(e, "Lines", x.Lines)
	encodeSimple(e, "Paragraphs", x.Paragraphs)
	encodeSimple(e, "ScaleCrop", x.ScaleCrop)
	encodeSimple(e, "Company", x.Company)
	encodeSimple(e, "LinksUpToDate", x.LinksUpToDate)
	encodeSimple(e, "CharactersWithSpaces", x.CharactersWithSpaces)
	encodeSimple(e, "SharedDoc", x.SharedDoc)
	encodeSimple(e, "HyperlinksChanged", x.HyperlinksChanged)
	encodeSimple(e, "AppVersion", x.AppVersion)

	if err := e.EncodeToken(start.End()); err != nil {
		return nil, err
	}
	if err := e.Flush(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// encodeSimple writes a simple text element: <name>value</name>.
func encodeSimple(e *xml.Encoder, name, value string) {
	s := xml.StartElement{Name: xml.Name{Local: name}}
	e.EncodeToken(s)
	if value != "" {
		e.EncodeToken(xml.CharData(value))
	}
	e.EncodeToken(s.End())
}

// encodeDatetime writes a dcterms date element with xsi:type attribute.
func encodeDatetime(e *xml.Encoder, name string, dt *xmlW3CDTF) {
	s := xml.StartElement{
		Name: xml.Name{Local: name},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xsi:type"}, Value: dt.Type},
		},
	}
	e.EncodeToken(s)
	if dt.Value != "" {
		e.EncodeToken(xml.CharData(dt.Value))
	}
	e.EncodeToken(s.End())
}
