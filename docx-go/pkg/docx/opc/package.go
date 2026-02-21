package opc

import (
	"bytes"
	"fmt"
	"io"
	"os"
)

// OpcPackage is the root object representing an OPC package.
type OpcPackage struct {
	rels        *Relationships
	partFactory *PartFactory
	parts       map[PackURI]Part
}

// NewOpcPackage creates an empty OpcPackage.
func NewOpcPackage(factory *PartFactory) *OpcPackage {
	if factory == nil {
		factory = NewPartFactory()
	}
	return &OpcPackage{
		rels:        NewRelationships("/"),
		partFactory: factory,
		parts:       make(map[PackURI]Part),
	}
}

// --------------------------------------------------------------------------
// Open
// --------------------------------------------------------------------------

// Open reads an OPC package from an io.ReaderAt.
func Open(r io.ReaderAt, size int64, factory *PartFactory) (*OpcPackage, error) {
	physReader, err := NewPhysPkgReader(r, size)
	if err != nil {
		return nil, err
	}
	defer physReader.Close()
	return openFromPhysReader(physReader, factory)
}

// OpenFile opens an OPC package from a file path.
func OpenFile(path string, factory *PartFactory) (*OpcPackage, error) {
	physReader, err := NewPhysPkgReaderFromFile(path)
	if err != nil {
		return nil, err
	}
	defer physReader.Close()
	return openFromPhysReader(physReader, factory)
}

// OpenBytes opens an OPC package from in-memory bytes.
func OpenBytes(data []byte, factory *PartFactory) (*OpcPackage, error) {
	physReader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		return nil, err
	}
	defer physReader.Close()
	return openFromPhysReader(physReader, factory)
}

func openFromPhysReader(physReader *PhysPkgReader, factory *PartFactory) (*OpcPackage, error) {
	if factory == nil {
		factory = NewPartFactory()
	}
	pkg := NewOpcPackage(factory)

	reader := &PackageReader{}
	result, err := reader.Read(physReader)
	if err != nil {
		return nil, err
	}

	// Unmarshal: create parts
	parts := make(map[PackURI]Part, len(result.SParts))
	for _, sp := range result.SParts {
		part, err := factory.New(sp.Partname, sp.ContentType, sp.RelType, sp.Blob, pkg)
		if err != nil {
			return nil, fmt.Errorf("opc: creating part %q: %w", sp.Partname, err)
		}
		parts[sp.Partname] = part
	}

	// Wire up package-level relationships
	for _, srel := range result.PkgSRels {
		var target interface{} = srel.TargetRef
		var targetPart Part
		if !srel.IsExternal() {
			pn := srel.TargetPartname()
			p, ok := parts[pn]
			if !ok {
				continue // skip unresolvable rels
			}
			targetPart = p
			target = p
		}
		_ = target
		pkg.rels.Load(srel.RID, srel.RelType, srel.TargetRef, targetPart, srel.IsExternal())
	}

	// Wire up part-level relationships
	for _, sp := range result.SParts {
		part, ok := parts[sp.Partname]
		if !ok {
			continue
		}
		rels := NewRelationships(sp.Partname.BaseURI())
		for _, srel := range sp.SRels {
			var targetPart Part
			if !srel.IsExternal() {
				pn := srel.TargetPartname()
				if p, ok := parts[pn]; ok {
					targetPart = p
				}
			}
			rels.Load(srel.RID, srel.RelType, srel.TargetRef, targetPart, srel.IsExternal())
		}
		part.SetRels(rels)
	}

	pkg.parts = parts

	// Call AfterUnmarshal on all parts
	for _, part := range parts {
		part.AfterUnmarshal()
	}

	return pkg, nil
}

// --------------------------------------------------------------------------
// Save
// --------------------------------------------------------------------------

// Save writes the package to an io.Writer.
func (p *OpcPackage) Save(w io.Writer) error {
	// Call BeforeMarshal on all parts
	for _, part := range p.parts {
		part.BeforeMarshal()
	}

	pw := &PackageWriter{}
	return pw.Write(w, p.rels, p.Parts())
}

// SaveToFile writes the package to a file.
func (p *OpcPackage) SaveToFile(path string) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("opc: creating file %q: %w", path, err)
	}
	defer f.Close()
	return p.Save(f)
}

// SaveToBytes returns the package as a byte slice.
func (p *OpcPackage) SaveToBytes() ([]byte, error) {
	var buf bytes.Buffer
	if err := p.Save(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// --------------------------------------------------------------------------
// Accessors
// --------------------------------------------------------------------------

// Rels returns the package-level relationships.
func (p *OpcPackage) Rels() *Relationships {
	return p.rels
}

// Parts returns all parts in the package (deterministic order by partname).
func (p *OpcPackage) Parts() []Part {
	result := make([]Part, 0, len(p.parts))
	for _, part := range p.parts {
		result = append(result, part)
	}
	return result
}

// PartByName returns a part by its PackURI.
func (p *OpcPackage) PartByName(pn PackURI) (Part, bool) {
	part, ok := p.parts[pn]
	return part, ok
}

// RelatedPart returns the part that the package has a relationship of relType to.
func (p *OpcPackage) RelatedPart(relType string) (Part, error) {
	rel, err := p.rels.GetByRelType(relType)
	if err != nil {
		return nil, err
	}
	if rel.IsExternal || rel.TargetPart == nil {
		return nil, fmt.Errorf("opc: relationship %q is external or unresolved", relType)
	}
	return rel.TargetPart, nil
}

// MainDocumentPart returns the main document part (via RT.OFFICE_DOCUMENT relationship).
func (p *OpcPackage) MainDocumentPart() (Part, error) {
	return p.RelatedPart(RTOfficeDocument)
}

// RelateTo creates or returns an existing package-level relationship to the given part.
func (p *OpcPackage) RelateTo(part Part, relType string) string {
	rel := p.rels.GetOrAdd(relType, part)
	return rel.RID
}

// AddPart adds a part to the package.
func (p *OpcPackage) AddPart(part Part) {
	p.parts[part.PartName()] = part
}

// NextPartname returns the next available partname matching the template (printf-style).
// E.g. NextPartname("/word/header%d.xml") might return "/word/header1.xml".
func (p *OpcPackage) NextPartname(template string) PackURI {
	partnames := make(map[PackURI]bool, len(p.parts))
	for pn := range p.parts {
		partnames[pn] = true
	}
	for n := 1; n <= len(partnames)+2; n++ {
		candidate := PackURI(fmt.Sprintf(template, n))
		if !partnames[candidate] {
			return candidate
		}
	}
	return PackURI(fmt.Sprintf(template, len(partnames)+1))
}

// IterParts generates all parts reachable via the relationship graph.
func (p *OpcPackage) IterParts() []Part {
	var result []Part
	visited := make(map[Part]bool)
	p.walkParts(p.rels, visited, &result)
	return result
}

func (p *OpcPackage) walkParts(rels *Relationships, visited map[Part]bool, result *[]Part) {
	for _, rel := range rels.All() {
		if rel.IsExternal || rel.TargetPart == nil {
			continue
		}
		part := rel.TargetPart
		if visited[part] {
			continue
		}
		visited[part] = true
		*result = append(*result, part)
		p.walkParts(part.Rels(), visited, result)
	}
}
