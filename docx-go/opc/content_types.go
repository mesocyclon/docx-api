package opc

import (
	"encoding/xml"
	"path"
	"sort"
	"strings"

	"github.com/vortex/docx-go/xmltypes"
)

// xmlTypes represents the root of [Content_Types].xml.
type xmlTypes struct {
	XMLName   xml.Name      `xml:"Types"`
	Defaults  []xmlDefault  `xml:"Default"`
	Overrides []xmlOverride `xml:"Override"`
}

type xmlDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type xmlOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// parseContentTypes parses [Content_Types].xml data.
func parseContentTypes(data []byte) (defaults map[string]string, overrides map[string]string, err error) {
	var ct xmlTypes
	if err = xml.Unmarshal(data, &ct); err != nil {
		return nil, nil, err
	}
	defaults = make(map[string]string, len(ct.Defaults))
	for _, d := range ct.Defaults {
		defaults[strings.ToLower(d.Extension)] = d.ContentType
	}
	overrides = make(map[string]string, len(ct.Overrides))
	for _, o := range ct.Overrides {
		overrides[normalizeName(o.PartName)] = o.ContentType
	}
	return defaults, overrides, nil
}

// buildContentTypes constructs [Content_Types].xml bytes from the package state.
func buildContentTypes(pkg *Package) ([]byte, error) {
	ct := xmlTypes{
		XMLName: xml.Name{Space: xmltypes.NSContentTypes, Local: "Types"},
	}

	// Collect defaults: start from stored defaults, ensure rels + xml present.
	defs := make(map[string]string)
	for ext, contentType := range pkg.defaults {
		defs[ext] = contentType
	}
	if _, ok := defs["rels"]; !ok {
		defs["rels"] = "application/vnd.openxmlformats-package.relationships+xml"
	}
	if _, ok := defs["xml"]; !ok {
		defs["xml"] = "application/xml"
	}

	// Check if any parts have media extensions needing Default entries.
	mediaDefaults := map[string]string{
		"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
		"gif": "image/gif", "bmp": "image/bmp", "tiff": "image/tiff",
		"emf": "image/x-emf", "wmf": "image/x-wmf", "svg": "image/svg+xml",
	}
	for _, pt := range pkg.parts {
		ext := strings.ToLower(strings.TrimPrefix(path.Ext(pt.Name), "."))
		if ct, ok := mediaDefaults[ext]; ok {
			if _, exists := defs[ext]; !exists {
				defs[ext] = ct
			}
		}
	}

	// Sort and emit defaults.
	extKeys := make([]string, 0, len(defs))
	for ext := range defs {
		extKeys = append(extKeys, ext)
	}
	sort.Strings(extKeys)
	for _, ext := range extKeys {
		ct.Defaults = append(ct.Defaults, xmlDefault{
			Extension:   ext,
			ContentType: defs[ext],
		})
	}

	// Build overrides from parts (use stored overrides + current parts).
	overrideMap := make(map[string]string)
	for name, contentType := range pkg.overrides {
		overrideMap[name] = contentType
	}
	for _, pt := range pkg.parts {
		// Parts always get an Override entry (standard OPC behavior).
		overrideMap[pt.Name] = pt.ContentType
	}

	names := make([]string, 0, len(overrideMap))
	for name := range overrideMap {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		ct.Overrides = append(ct.Overrides, xmlOverride{
			PartName:    name,
			ContentType: overrideMap[name],
		})
	}

	return marshalXMLWithHeader(ct)
}

// resolveContentType looks up the content type for a part name.
func resolveContentType(name string, defaults, overrides map[string]string) string {
	name = normalizeName(name)
	if ct, ok := overrides[name]; ok {
		return ct
	}
	ext := strings.ToLower(strings.TrimPrefix(path.Ext(name), "."))
	if ct, ok := defaults[ext]; ok {
		return ct
	}
	return "application/octet-stream"
}

func marshalXMLWithHeader(v interface{}) ([]byte, error) {
	data, err := xml.MarshalIndent(v, "", "  ")
	if err != nil {
		return nil, err
	}
	return append([]byte(xml.Header), data...), nil
}
