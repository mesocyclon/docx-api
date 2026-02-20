// Package opc implements Open Packaging Conventions (OPC) for OOXML files.
// It handles reading/writing ZIP-based packages with parts, content types,
// and relationships.
package opc

// Relationship Type URI constants.
const (
	RelOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	RelCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	RelExtProperties  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	RelStyles         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	RelSettings       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	RelWebSettings    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	RelFontTable      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	RelNumbering      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
	RelFootnotes      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
	RelEndnotes       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
	RelComments       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	RelHeader         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	RelFooter         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	RelImage          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	RelHyperlink      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	RelTheme          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
)

// Relationship represents an OPC relationship entry from a .rels file.
type Relationship struct {
	ID         string // "rId1"
	Type       string // full URI
	Target     string // "styles.xml" or URL
	TargetMode string // "" (Internal) | "External"
}

// Part represents a single part within an OPC package.
type Part struct {
	Name        string         // "/word/document.xml"
	ContentType string         // MIME content type
	Data        []byte         // raw part data
	Rels        []Relationship // part-level relationships
}

// AddRel adds an internal relationship to the part and returns the generated rId.
func (pt *Part) AddRel(relType, target string) string {
	id := nextRelIDFrom(pt.Rels)
	pt.Rels = append(pt.Rels, Relationship{
		ID:     id,
		Type:   relType,
		Target: target,
	})
	return id
}

// AddExternalRel adds an external relationship to the part and returns the generated rId.
func (pt *Part) AddExternalRel(relType, target string) string {
	id := nextRelIDFrom(pt.Rels)
	pt.Rels = append(pt.Rels, Relationship{
		ID:         id,
		Type:       relType,
		Target:     target,
		TargetMode: "External",
	})
	return id
}

// RelsByType returns all relationships of the given type.
func (pt *Part) RelsByType(relType string) []Relationship {
	var result []Relationship
	for _, r := range pt.Rels {
		if r.Type == relType {
			result = append(result, r)
		}
	}
	return result
}

// RelByID returns the relationship with the given ID.
func (pt *Part) RelByID(id string) (Relationship, bool) {
	for _, r := range pt.Rels {
		if r.ID == id {
			return r, true
		}
	}
	return Relationship{}, false
}

// Package represents an OPC package (a ZIP-based container).
type Package struct {
	parts       map[string]*Part  // normalized name → part
	packageRels []Relationship    // package-level relationships (_rels/.rels)
	defaults    map[string]string // extension → content type (Default entries)
	overrides   map[string]string // part name → content type (Override entries)
}
