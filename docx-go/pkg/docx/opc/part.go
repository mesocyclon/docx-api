package opc

import (
	"github.com/beevik/etree"
)

// --------------------------------------------------------------------------
// Part interface
// --------------------------------------------------------------------------

// Part represents an element within an OPC package.
type Part interface {
	PartName() PackURI
	ContentType() string
	Blob() []byte
	Rels() *Relationships
	SetRels(rels *Relationships)
	BeforeMarshal()
	AfterUnmarshal()
}

// --------------------------------------------------------------------------
// BasePart — default implementation of Part
// --------------------------------------------------------------------------

// BasePart is the base implementation of the Part interface for binary parts.
type BasePart struct {
	partName    PackURI
	contentType string
	blob        []byte
	rels        *Relationships
	pkg         *OpcPackage
}

// NewBasePart creates a new BasePart.
func NewBasePart(partName PackURI, contentType string, blob []byte, pkg *OpcPackage) *BasePart {
	return &BasePart{
		partName:    partName,
		contentType: contentType,
		blob:        blob,
		pkg:         pkg,
		rels:        NewRelationships(partName.BaseURI()),
	}
}

func (p *BasePart) PartName() PackURI         { return p.partName }
func (p *BasePart) ContentType() string        { return p.contentType }
func (p *BasePart) Blob() []byte               { return p.blob }
func (p *BasePart) Rels() *Relationships       { return p.rels }
func (p *BasePart) SetRels(rels *Relationships) { p.rels = rels }
func (p *BasePart) Package() *OpcPackage       { return p.pkg }
func (p *BasePart) BeforeMarshal()             {}
func (p *BasePart) AfterUnmarshal()            {}

// SetPartName updates the part name.
func (p *BasePart) SetPartName(pn PackURI) {
	p.partName = pn
}

// SetBlob replaces the blob.
func (p *BasePart) SetBlob(blob []byte) {
	p.blob = blob
}

// --------------------------------------------------------------------------
// XmlPart — Part with parsed XML content
// --------------------------------------------------------------------------

// XmlPart extends BasePart with a parsed etree.Element.
type XmlPart struct {
	BasePart
	element *etree.Element
}

// NewXmlPart creates an XmlPart by parsing the blob as XML.
func NewXmlPart(partName PackURI, contentType string, blob []byte, pkg *OpcPackage) (*XmlPart, error) {
	doc := etree.NewDocument()
	doc.ReadSettings.Permissive = true
	if err := doc.ReadFromBytes(blob); err != nil {
		return nil, err
	}
	root := doc.Root()
	return &XmlPart{
		BasePart: *NewBasePart(partName, contentType, nil, pkg),
		element:  root,
	}, nil
}

// NewXmlPartFromElement creates an XmlPart from an existing element.
func NewXmlPartFromElement(partName PackURI, contentType string, element *etree.Element, pkg *OpcPackage) *XmlPart {
	return &XmlPart{
		BasePart: *NewBasePart(partName, contentType, nil, pkg),
		element:  element,
	}
}

// Element returns the root XML element.
func (p *XmlPart) Element() *etree.Element {
	return p.element
}

// SetElement replaces the root XML element.
func (p *XmlPart) SetElement(el *etree.Element) {
	p.element = el
}

// Blob serializes the XML element to bytes.
func (p *XmlPart) Blob() []byte {
	if p.element == nil {
		return nil
	}
	doc := etree.NewDocument()
	doc.CreateProcInst("xml", `version="1.0" encoding="UTF-8" standalone="yes"`)
	doc.SetRoot(p.element.Copy())
	doc.Indent(0)
	b, _ := doc.WriteToBytes()
	return b
}

// --------------------------------------------------------------------------
// PartConstructor — factory function type
// --------------------------------------------------------------------------

// PartConstructor is a function that creates a Part from serialized data.
type PartConstructor func(partName PackURI, contentType, relType string, blob []byte, pkg *OpcPackage) (Part, error)

// --------------------------------------------------------------------------
// PartFactory — registry of part constructors
// --------------------------------------------------------------------------

// PartFactory maps content types to Part constructors.
type PartFactory struct {
	constructors map[string]PartConstructor
	selector     func(contentType, relType string) PartConstructor
}

// NewPartFactory creates an empty PartFactory.
func NewPartFactory() *PartFactory {
	return &PartFactory{
		constructors: make(map[string]PartConstructor),
	}
}

// Register maps a content type to a constructor.
func (f *PartFactory) Register(contentType string, ctor PartConstructor) {
	f.constructors[contentType] = ctor
}

// SetSelector sets a custom selector function that takes precedence over content type map.
func (f *PartFactory) SetSelector(sel func(contentType, relType string) PartConstructor) {
	f.selector = sel
}

// New creates a Part using the registered constructors.
// Falls back to BasePart if no constructor matches.
func (f *PartFactory) New(partName PackURI, contentType, relType string, blob []byte, pkg *OpcPackage) (Part, error) {
	// Try selector first
	if f.selector != nil {
		if ctor := f.selector(contentType, relType); ctor != nil {
			return ctor(partName, contentType, relType, blob, pkg)
		}
	}
	// Try content type map
	if ctor, ok := f.constructors[contentType]; ok {
		return ctor(partName, contentType, relType, blob, pkg)
	}
	// Default: create a simple BasePart
	return NewBasePart(partName, contentType, blob, pkg), nil
}
