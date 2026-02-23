package opc

import (
	"fmt"

	"github.com/beevik/etree"
)

// --------------------------------------------------------------------------
// Part interface
// --------------------------------------------------------------------------

// Part represents an element within an OPC package.
type Part interface {
	PartName() PackURI
	ContentType() string
	Blob() ([]byte, error)
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

func (p *BasePart) PartName() PackURI           { return p.partName }
func (p *BasePart) ContentType() string         { return p.contentType }
func (p *BasePart) Blob() ([]byte, error)       { return p.blob, nil }
func (p *BasePart) Rels() *Relationships        { return p.rels }
func (p *BasePart) SetRels(rels *Relationships) { p.rels = rels }
func (p *BasePart) Package() *OpcPackage        { return p.pkg }
func (p *BasePart) BeforeMarshal()              {}
func (p *BasePart) AfterUnmarshal()             {}

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

// xmlProcInst is the standard XML declaration for OPC parts.
const xmlProcInst = `version="1.0" encoding="UTF-8" standalone="yes"`

// XmlPart extends BasePart with a parsed XML document.
//
// Internally it stores the owning *etree.Document rather than a bare
// *etree.Element. This lets Blob() serialize the tree directly without
// the deep-copy that would be required if we had to re-parent the element
// into a temporary Document via SetRoot on every call.
type XmlPart struct {
	BasePart
	doc *etree.Document
}

// newXmlDoc creates a Document pre-configured with the standard OPC XML
// processing instruction and compact write settings.
func newXmlDoc() *etree.Document {
	doc := etree.NewDocument()
	doc.CreateProcInst("xml", xmlProcInst)
	doc.WriteSettings.CanonicalEndTags = true
	return doc
}

// ensureProcInst normalizes the XML processing instruction to the standard
// OPC form (version="1.0" encoding="UTF-8" standalone="yes").
// This guarantees output identical to the previous implementation that
// always created a fresh Document in Blob().
func ensureProcInst(doc *etree.Document) {
	for _, tok := range doc.Child {
		if pi, ok := tok.(*etree.ProcInst); ok && pi.Target == "xml" {
			pi.Inst = xmlProcInst
			return
		}
	}
	// No <?xml ...?> found — prepend one.
	pi := &etree.ProcInst{Target: "xml", Inst: xmlProcInst}
	doc.Child = append([]etree.Token{pi}, doc.Child...)
}

// NewXmlPart creates an XmlPart by parsing the blob as XML.
func NewXmlPart(partName PackURI, contentType string, blob []byte, pkg *OpcPackage) (*XmlPart, error) {
	doc := etree.NewDocument()
	doc.ReadSettings.Permissive = true
	doc.WriteSettings.CanonicalEndTags = true
	if err := doc.ReadFromBytes(blob); err != nil {
		return nil, err
	}
	// Normalize the declaration so Blob() output matches the previous
	// implementation that always wrote a fresh standalone="yes" header.
	ensureProcInst(doc)
	return &XmlPart{
		BasePart: *NewBasePart(partName, contentType, nil, pkg),
		doc:      doc,
	}, nil
}

// NewXmlPartFromElement creates an XmlPart from an existing element.
// The element is adopted into a new Document — it will be detached
// from any previous parent.
func NewXmlPartFromElement(partName PackURI, contentType string, element *etree.Element, pkg *OpcPackage) *XmlPart {
	doc := newXmlDoc()
	doc.SetRoot(element)
	return &XmlPart{
		BasePart: *NewBasePart(partName, contentType, nil, pkg),
		doc:      doc,
	}
}

// Element returns the root XML element, or nil if the document is empty.
func (p *XmlPart) Element() *etree.Element {
	if p.doc == nil {
		return nil
	}
	return p.doc.Root()
}

// SetElement replaces the root XML element.
// The element is adopted by the internal Document.
func (p *XmlPart) SetElement(el *etree.Element) {
	if p.doc == nil {
		p.doc = newXmlDoc()
	}
	p.doc.SetRoot(el)
}

// Blob serializes the XML document to bytes.
// Output is compact (no insignificant whitespace), with a standard
// XML declaration — matching Python's serialize_part_xml behavior.
//
// Unlike the previous implementation, no deep-copy of the element tree
// is performed: the Document already owns the root element.
func (p *XmlPart) Blob() ([]byte, error) {
	if p.doc == nil || p.doc.Root() == nil {
		return nil, nil
	}
	b, err := p.doc.WriteToBytes()
	if err != nil {
		return nil, fmt.Errorf("opc: serializing XML part %q: %w", p.partName, err)
	}
	return b, nil
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
