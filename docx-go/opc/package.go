package opc

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"sort"
	"strings"
)

// New creates an empty OPC package.
func New() *Package {
	return &Package{
		parts:     make(map[string]*Part),
		defaults:  make(map[string]string),
		overrides: make(map[string]string),
	}
}

// OpenFile opens an OPC package from a file path.
func OpenFile(path string) (*Package, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("opc: open file: %w", err)
	}
	defer f.Close()

	info, err := f.Stat()
	if err != nil {
		return nil, fmt.Errorf("opc: stat file: %w", err)
	}
	return OpenReader(f, info.Size())
}

// OpenReader opens an OPC package from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (*Package, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("opc: open zip: %w", err)
	}

	pkg := New()

	// Read all ZIP entries into a map for lookup.
	entries := make(map[string][]byte, len(zr.File))
	for _, zf := range zr.File {
		data, err := readZipFile(zf)
		if err != nil {
			return nil, fmt.Errorf("opc: read %s: %w", zf.Name, err)
		}
		entries[zf.Name] = data
	}

	// Parse [Content_Types].xml.
	ctData, ok := entries["[Content_Types].xml"]
	if !ok {
		return nil, fmt.Errorf("opc: missing [Content_Types].xml")
	}
	defaults, overrides, err := parseContentTypes(ctData)
	if err != nil {
		return nil, fmt.Errorf("opc: parse content types: %w", err)
	}
	pkg.defaults = defaults
	pkg.overrides = overrides

	// Parse package-level relationships (_rels/.rels).
	if relsData, ok := entries["_rels/.rels"]; ok {
		rels, err := parseRels(relsData)
		if err != nil {
			return nil, fmt.Errorf("opc: parse package rels: %w", err)
		}
		pkg.packageRels = rels
	}

	// Build parts from entries (skip [Content_Types].xml, .rels files).
	for name, data := range entries {
		if name == "[Content_Types].xml" {
			continue
		}
		if isRelsPath(name) {
			continue
		}
		normalized := normalizeName(name)
		ct := resolveContentType(normalized, defaults, overrides)
		part := &Part{
			Name:        normalized,
			ContentType: ct,
			Data:        data,
		}
		pkg.parts[normalized] = part
	}

	// Parse part-level relationships.
	for name, part := range pkg.parts {
		rp := relsZipPath(name)
		if relsData, ok := entries[rp]; ok {
			rels, err := parseRels(relsData)
			if err != nil {
				return nil, fmt.Errorf("opc: parse rels for %s: %w", name, err)
			}
			part.Rels = rels
		}
	}

	return pkg, nil
}

// SaveFile writes the OPC package to a file.
func (p *Package) SaveFile(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("opc: create file: %w", err)
	}
	defer f.Close()

	if err := p.SaveWriter(f); err != nil {
		return err
	}
	return f.Close()
}

// SaveWriter writes the OPC package to a writer as a ZIP archive.
func (p *Package) SaveWriter(w io.Writer) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write [Content_Types].xml.
	ctData, err := buildContentTypes(p)
	if err != nil {
		return fmt.Errorf("opc: build content types: %w", err)
	}
	if err := writeZipEntry(zw, "[Content_Types].xml", ctData); err != nil {
		return err
	}

	// Write _rels/.rels (package-level relationships).
	if len(p.packageRels) > 0 {
		relsData, err := buildRels(p.packageRels)
		if err != nil {
			return fmt.Errorf("opc: build package rels: %w", err)
		}
		if err := writeZipEntry(zw, "_rels/.rels", relsData); err != nil {
			return err
		}
	}

	// Write parts (sorted for deterministic output).
	names := make([]string, 0, len(p.parts))
	for name := range p.parts {
		names = append(names, name)
	}
	sort.Strings(names)

	for _, name := range names {
		part := p.parts[name]
		// Write the part data.
		if err := writeZipEntry(zw, stripLeadingSlash(name), part.Data); err != nil {
			return err
		}
		// Write part-level .rels if present.
		if len(part.Rels) > 0 {
			relsData, err := buildRels(part.Rels)
			if err != nil {
				return fmt.Errorf("opc: build rels for %s: %w", name, err)
			}
			rp := relsZipPath(name)
			if err := writeZipEntry(zw, rp, relsData); err != nil {
				return err
			}
		}
	}

	return zw.Close()
}

// Part returns the part with the given name.
func (p *Package) Part(name string) (*Part, bool) {
	name = normalizeName(name)
	pt, ok := p.parts[name]
	return pt, ok
}

// AddPart adds a new part to the package.
func (p *Package) AddPart(name, contentType string, data []byte) *Part {
	name = normalizeName(name)
	pt := &Part{
		Name:        name,
		ContentType: contentType,
		Data:        data,
	}
	p.parts[name] = pt
	return pt
}

// RemovePart removes the part with the given name. Returns true if it existed.
func (p *Package) RemovePart(name string) bool {
	name = normalizeName(name)
	if _, ok := p.parts[name]; ok {
		delete(p.parts, name)
		return true
	}
	return false
}

// Parts returns all parts in the package (sorted by name).
func (p *Package) Parts() []*Part {
	names := make([]string, 0, len(p.parts))
	for name := range p.parts {
		names = append(names, name)
	}
	sort.Strings(names)

	result := make([]*Part, 0, len(names))
	for _, name := range names {
		result = append(result, p.parts[name])
	}
	return result
}

// PackageRels returns package-level relationships.
func (p *Package) PackageRels() []Relationship {
	result := make([]Relationship, len(p.packageRels))
	copy(result, p.packageRels)
	return result
}

// AddPackageRel adds a package-level relationship and returns the generated rId.
func (p *Package) AddPackageRel(relType, target string) string {
	id := nextRelIDFrom(p.packageRels)
	p.packageRels = append(p.packageRels, Relationship{
		ID:     id,
		Type:   relType,
		Target: target,
	})
	return id
}

// PackageRelsByType returns package-level relationships of the given type.
func (p *Package) PackageRelsByType(relType string) []Relationship {
	var result []Relationship
	for _, r := range p.packageRels {
		if r.Type == relType {
			result = append(result, r)
		}
	}
	return result
}

// --- Internal helpers ---

func isRelsPath(zipName string) bool {
	return strings.Contains(zipName, "_rels/") && strings.HasSuffix(zipName, ".rels")
}

func readZipFile(zf *zip.File) ([]byte, error) {
	rc, err := zf.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	var buf bytes.Buffer
	if _, err := io.Copy(&buf, rc); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func writeZipEntry(zw *zip.Writer, name string, data []byte) error {
	w, err := zw.Create(name)
	if err != nil {
		return fmt.Errorf("opc: create zip entry %s: %w", name, err)
	}
	_, err = w.Write(data)
	if err != nil {
		return fmt.Errorf("opc: write zip entry %s: %w", name, err)
	}
	return nil
}
