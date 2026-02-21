// Package oxml provides low-level XML element manipulation for Office Open XML documents.
package oxml

import (
	"fmt"
	"strings"
)

// Nsmap maps namespace prefixes to their URIs.
var Nsmap = map[string]string{
	"a":       "http://schemas.openxmlformats.org/drawingml/2006/main",
	"c":       "http://schemas.openxmlformats.org/drawingml/2006/chart",
	"cp":      "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
	"dc":      "http://purl.org/dc/elements/1.1/",
	"dcmitype": "http://purl.org/dc/dcmitype/",
	"dcterms": "http://purl.org/dc/terms/",
	"dgm":     "http://schemas.openxmlformats.org/drawingml/2006/diagram",
	"m":       "http://schemas.openxmlformats.org/officeDocument/2006/math",
	"pic":     "http://schemas.openxmlformats.org/drawingml/2006/picture",
	"r":       "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"sl":      "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
	"w":       "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	"w14":     "http://schemas.microsoft.com/office/word/2010/wordml",
	"wp":      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
	"xml":     "http://www.w3.org/XML/1998/namespace",
	"xsi":     "http://www.w3.org/2001/XMLSchema-instance",
}

// Pfxmap is the reverse mapping of URI → prefix.
var Pfxmap map[string]string

func init() {
	Pfxmap = make(map[string]string, len(Nsmap))
	for pfx, uri := range Nsmap {
		Pfxmap[uri] = pfx
	}
}

// Qn converts a namespace-prefixed tag to Clark notation.
// For example, Qn("w:p") returns "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p".
func Qn(tag string) string {
	prefix, local, ok := strings.Cut(tag, ":")
	if !ok {
		return tag
	}
	uri, exists := Nsmap[prefix]
	if !exists {
		panic(fmt.Sprintf("oxml.Qn: unknown namespace prefix %q in tag %q", prefix, tag))
	}
	return "{" + uri + "}" + local
}

// NsDecls returns a namespace declaration string for the given prefixes.
// For example, NsDecls("w", "r") returns `xmlns:w="..." xmlns:r="..."`.
func NsDecls(prefixes ...string) string {
	parts := make([]string, 0, len(prefixes))
	for _, pfx := range prefixes {
		uri, ok := Nsmap[pfx]
		if !ok {
			panic(fmt.Sprintf("oxml.NsDecls: unknown namespace prefix %q", pfx))
		}
		parts = append(parts, fmt.Sprintf(`xmlns:%s="%s"`, pfx, uri))
	}
	return strings.Join(parts, " ")
}

// NsPfxMap returns a subset of Nsmap for the specified prefixes.
func NsPfxMap(prefixes ...string) map[string]string {
	result := make(map[string]string, len(prefixes))
	for _, pfx := range prefixes {
		if uri, ok := Nsmap[pfx]; ok {
			result[pfx] = uri
		}
	}
	return result
}

// NamespacePrefixedTag is a value object that knows the semantics of an XML tag
// with a namespace prefix, such as "w:p".
type NamespacePrefixedTag struct {
	prefix    string
	localPart string
	nsURI     string
}

// NewNSPTag creates a NamespacePrefixedTag from a prefixed tag string like "w:p".
func NewNSPTag(nstag string) NamespacePrefixedTag {
	prefix, local, ok := strings.Cut(nstag, ":")
	if !ok {
		panic(fmt.Sprintf("oxml.NewNSPTag: invalid namespace-prefixed tag %q", nstag))
	}
	uri, exists := Nsmap[prefix]
	if !exists {
		panic(fmt.Sprintf("oxml.NewNSPTag: unknown namespace prefix %q in tag %q", prefix, nstag))
	}
	return NamespacePrefixedTag{
		prefix:    prefix,
		localPart: local,
		nsURI:     uri,
	}
}

// NSPTagFromClark creates a NamespacePrefixedTag from Clark notation like "{http://...}p".
func NSPTagFromClark(clark string) NamespacePrefixedTag {
	if len(clark) == 0 || clark[0] != '{' {
		panic(fmt.Sprintf("oxml.NSPTagFromClark: invalid Clark notation %q", clark))
	}
	closeBrace := strings.Index(clark, "}")
	if closeBrace < 0 {
		panic(fmt.Sprintf("oxml.NSPTagFromClark: invalid Clark notation %q", clark))
	}
	nsURI := clark[1:closeBrace]
	local := clark[closeBrace+1:]

	pfx, ok := Pfxmap[nsURI]
	if !ok {
		panic(fmt.Sprintf("oxml.NSPTagFromClark: unknown namespace URI %q", nsURI))
	}
	return NamespacePrefixedTag{
		prefix:    pfx,
		localPart: local,
		nsURI:     nsURI,
	}
}

// ClarkName returns the Clark notation for this tag, e.g. "{http://...}p".
func (t NamespacePrefixedTag) ClarkName() string {
	return "{" + t.nsURI + "}" + t.localPart
}

// LocalPart returns the local part of the tag, e.g. "p" for "w:p".
func (t NamespacePrefixedTag) LocalPart() string {
	return t.localPart
}

// Prefix returns the namespace prefix, e.g. "w" for "w:p".
func (t NamespacePrefixedTag) Prefix() string {
	return t.prefix
}

// NsURI returns the namespace URI.
func (t NamespacePrefixedTag) NsURI() string {
	return t.nsURI
}

// String returns the prefixed tag string, e.g. "w:p".
func (t NamespacePrefixedTag) String() string {
	return t.prefix + ":" + t.localPart
}

// NsMap returns a single-member map of this tag's prefix to its namespace URI.
func (t NamespacePrefixedTag) NsMap() map[string]string {
	return map[string]string{t.prefix: t.nsURI}
}
