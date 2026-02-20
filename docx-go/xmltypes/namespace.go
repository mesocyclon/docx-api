// Package xmltypes provides shared XML type definitions and namespace constants
// for the docx-go library. It has no dependencies beyond encoding/xml and io.
package xmltypes

import (
	"encoding/xml"
	"io"
)

// ============================================================
// Namespace URI — Transitional (99.9% of real documents)
// ============================================================

const (
	// WordprocessingML (main)
	NSw = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

	// Relationships
	NSr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

	// DrawingML
	NSa   = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NSwp  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	NSpic = "http://schemas.openxmlformats.org/drawingml/2006/picture"

	// Markup Compatibility
	NSmc = "http://schemas.openxmlformats.org/markup-compatibility/2006"

	// Math
	NSm = "http://schemas.openxmlformats.org/officeDocument/2006/math"

	// VML (legacy)
	NSv   = "urn:schemas-microsoft-com:vml"
	NSo   = "urn:schemas-microsoft-com:office:office"
	NSw10 = "urn:schemas-microsoft-com:office:word"

	// Microsoft Extensions (Word 2010+)
	NSw14   = "http://schemas.microsoft.com/office/word/2010/wordml"
	NSw15   = "http://schemas.microsoft.com/office/word/2012/wordml"
	NSw16se = "http://schemas.microsoft.com/office/word/2015/wordml/symex"

	// Microsoft WordprocessingDrawing / Canvas / Shape extensions
	NSwpc  = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
	NSwp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
	NSwpg  = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
	NSwpi  = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
	NSwne  = "http://schemas.microsoft.com/office/word/2006/wordml"
	NSwps  = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

	// OPC / Package
	NSContentTypes  = "http://schemas.openxmlformats.org/package/2006/content-types"
	NSRelationships = "http://schemas.openxmlformats.org/package/2006/relationships"

	// Document Properties
	NScp       = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	NSdc       = "http://purl.org/dc/elements/1.1/"
	NSdcterms  = "http://purl.org/dc/terms/"
	NSdcmitype = "http://purl.org/dc/dcmitype/"
	NSxsi      = "http://www.w3.org/2001/XMLSchema-instance"

	// Extended properties
	NSvt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
	NSep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	NSsl = "http://schemas.openxmlformats.org/schemaLibrary/2006/main"
)

// ============================================================
// Strict → Transitional namespace mapping
// ============================================================

// strictToTransitional maps ECMA-376 Strict namespace URIs to their
// Transitional equivalents. When unmarshalling, we accept both; when
// marshalling, we always emit Transitional for maximum compatibility.
var strictToTransitional = map[string]string{
	"http://purl.oclc.org/ooxml/wordprocessingml/main":           NSw,
	"http://purl.oclc.org/ooxml/officeDocument/relationships":    NSr,
	"http://purl.oclc.org/ooxml/drawingml/main":                  NSa,
	"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": NSwp,
	"http://purl.oclc.org/ooxml/officeDocument/math":             NSm,
	"http://purl.oclc.org/ooxml/drawingml/picture":               NSpic,
}

// NormalizeNamespace converts a Strict namespace URI to its Transitional
// equivalent. If the URI is not a known Strict namespace, it is returned
// unchanged.
func NormalizeNamespace(ns string) string {
	if mapped, ok := strictToTransitional[ns]; ok {
		return mapped
	}
	return ns
}

// IsWNS reports whether space represents the WordprocessingML namespace.
// It accepts: the Transitional URI (NSw), the Strict URI, an empty string
// (encoding/xml omits the namespace for children of a prefixed parent),
// and the bare prefix "w" (encoding/xml may use the prefix as Space).
func IsWNS(space string) bool {
	return NormalizeNamespace(space) == NSw || space == "" || space == "w"
}

// ============================================================
// NormalizingDecoder — wraps xml.Decoder with auto-normalization
// ============================================================

// NormalizingDecoder wraps an xml.Decoder to automatically convert
// Strict OOXML namespace URIs to Transitional equivalents on every
// token read.
type NormalizingDecoder struct {
	inner *xml.Decoder
}

// NewNormalizingDecoder creates a NormalizingDecoder that reads from r.
func NewNormalizingDecoder(r io.Reader) *NormalizingDecoder {
	return &NormalizingDecoder{inner: xml.NewDecoder(r)}
}

// Token reads the next XML token and normalizes any Strict namespaces.
func (d *NormalizingDecoder) Token() (xml.Token, error) {
	tok, err := d.inner.Token()
	if err != nil {
		return tok, err
	}
	switch t := tok.(type) {
	case xml.StartElement:
		t.Name.Space = NormalizeNamespace(t.Name.Space)
		for i := range t.Attr {
			t.Attr[i].Name.Space = NormalizeNamespace(t.Attr[i].Name.Space)
		}
		return t, nil
	case xml.EndElement:
		t.Name.Space = NormalizeNamespace(t.Name.Space)
		return t, nil
	}
	return tok, nil
}

// DecodeElement works like xml.Decoder.DecodeElement but with namespace
// normalization. It delegates to the inner decoder.
func (d *NormalizingDecoder) DecodeElement(v interface{}, start *xml.StartElement) error {
	return d.inner.DecodeElement(v, start)
}

// Skip skips the current element, delegating to the inner decoder.
func (d *NormalizingDecoder) Skip() error {
	return d.inner.Skip()
}
