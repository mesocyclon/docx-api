package opc

import (
	"testing"
)

// ---------------------------------------------------------------------------
// BasePart
// ---------------------------------------------------------------------------

func TestBasePart_Accessors(t *testing.T) {
	t.Parallel()

	pkg := NewOpcPackage(nil)
	blob := []byte("binary data")
	part := NewBasePart("/word/document.xml", CTWmlDocumentMain, blob, pkg)

	if part.PartName() != "/word/document.xml" {
		t.Errorf("PartName: got %q", part.PartName())
	}
	if part.ContentType() != CTWmlDocumentMain {
		t.Errorf("ContentType: got %q", part.ContentType())
	}
	gotBlob, err := part.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}
	if string(gotBlob) != "binary data" {
		t.Errorf("Blob: got %q", string(gotBlob))
	}
	if part.Rels() == nil {
		t.Error("Rels should not be nil")
	}
	if part.Package() != pkg {
		t.Error("Package mismatch")
	}

	// SetPartName
	part.SetPartName("/word/newname.xml")
	if part.PartName() != "/word/newname.xml" {
		t.Errorf("after SetPartName: got %q", part.PartName())
	}

	// SetBlob
	part.SetBlob([]byte("new data"))
	gotBlob, _ = part.Blob()
	if string(gotBlob) != "new data" {
		t.Errorf("after SetBlob: got %q", string(gotBlob))
	}

	// SetRels
	newRels := NewRelationships("/word")
	part.SetRels(newRels)
	if part.Rels() != newRels {
		t.Error("SetRels did not update")
	}

	// BeforeMarshal and AfterUnmarshal should be no-ops (no panic)
	part.BeforeMarshal()
	part.AfterUnmarshal()
}

// ---------------------------------------------------------------------------
// XmlPart
// ---------------------------------------------------------------------------

func TestXmlPart_FromValidXml(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0" encoding="UTF-8"?><root><child/></root>`)
	part, err := NewXmlPart("/word/document.xml", CTWmlDocumentMain, xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}
	el := part.Element()
	if el == nil {
		t.Fatal("Element should not be nil")
	}
	if el.Tag != "root" {
		t.Errorf("expected root tag, got %q", el.Tag)
	}
}

func TestXmlPart_FromInvalidXml(t *testing.T) {
	t.Parallel()

	garbage := []byte("this is not XML at all <<<>>>")
	_, err := NewXmlPart("/word/document.xml", CTWmlDocumentMain, garbage, nil)
	if err == nil {
		t.Fatal("expected error for invalid XML, got nil")
	}
}

func TestXmlPart_Blob_RoundTrip(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0" encoding="UTF-8"?><root><child attr="val"></child></root>`)
	part, err := NewXmlPart("/word/document.xml", CTWmlDocumentMain, xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}

	blob, err := part.Blob()
	if err != nil {
		t.Fatalf("Blob: %v", err)
	}
	if len(blob) == 0 {
		t.Fatal("expected non-empty blob")
	}
	// Should contain XML declaration
	if !containsSubstring(string(blob), "<?xml") {
		t.Error("blob should contain <?xml declaration")
	}
	// Should contain our content
	if !containsSubstring(string(blob), "root") {
		t.Error("blob should contain root element")
	}
}

func TestXmlPart_Blob_NilDoc(t *testing.T) {
	t.Parallel()

	part := &XmlPart{
		BasePart: *NewBasePart("/word/document.xml", CTWmlDocumentMain, nil, nil),
		doc:      nil,
	}

	blob, err := part.Blob()
	if err != nil {
		t.Fatalf("Blob with nil doc: %v", err)
	}
	if blob != nil {
		t.Errorf("expected nil blob for nil doc, got %d bytes", len(blob))
	}
}

func TestXmlPart_SetElement(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0"?><old/>`)
	part, err := NewXmlPart("/test.xml", "application/xml", xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}
	if part.Element().Tag != "old" {
		t.Fatalf("expected 'old' tag, got %q", part.Element().Tag)
	}

	newXml := []byte(`<?xml version="1.0"?><new/>`)
	part2, _ := NewXmlPart("/test2.xml", "application/xml", newXml, nil)
	part.SetElement(part2.Element())

	if part.Element().Tag != "new" {
		t.Errorf("after SetElement: expected 'new' tag, got %q", part.Element().Tag)
	}
}

func TestXmlPartFromElement(t *testing.T) {
	t.Parallel()

	xml := []byte(`<?xml version="1.0"?><root/>`)
	original, err := NewXmlPart("/temp.xml", "application/xml", xml, nil)
	if err != nil {
		t.Fatalf("NewXmlPart: %v", err)
	}

	part := NewXmlPartFromElement("/word/document.xml", CTWmlDocumentMain, original.Element(), nil)
	if part.PartName() != "/word/document.xml" {
		t.Errorf("PartName: got %q", part.PartName())
	}
	if part.Element().Tag != "root" {
		t.Errorf("Element tag: got %q", part.Element().Tag)
	}
}

// ---------------------------------------------------------------------------
// PartFactory
// ---------------------------------------------------------------------------

func TestPartFactory_ContentTypeMap(t *testing.T) {
	t.Parallel()

	factory := NewPartFactory()
	factory.Register(CTWmlDocumentMain, func(pn PackURI, ct, rt string, blob []byte, pkg *OpcPackage) (Part, error) {
		return NewXmlPart(pn, ct, blob, pkg)
	})

	xml := []byte(`<?xml version="1.0"?><w:document/>`)
	part, err := factory.New("/word/document.xml", CTWmlDocumentMain, RTOfficeDocument, xml, nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	if _, ok := part.(*XmlPart); !ok {
		t.Errorf("expected *XmlPart, got %T", part)
	}
}

func TestPartFactory_Selector(t *testing.T) {
	t.Parallel()

	factory := NewPartFactory()
	// Register a content-type constructor (should NOT be used)
	factory.Register(CTWmlDocumentMain, func(pn PackURI, ct, rt string, blob []byte, pkg *OpcPackage) (Part, error) {
		return NewBasePart(pn, ct, blob, pkg), nil
	})
	// Register a selector that overrides
	factory.SetSelector(func(ct, rt string) PartConstructor {
		if rt == RTOfficeDocument {
			return func(pn PackURI, ct, rt string, blob []byte, pkg *OpcPackage) (Part, error) {
				return NewXmlPart(pn, ct, blob, pkg)
			}
		}
		return nil
	})

	xml := []byte(`<?xml version="1.0"?><w:document/>`)
	part, err := factory.New("/word/document.xml", CTWmlDocumentMain, RTOfficeDocument, xml, nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	// Selector should have produced XmlPart, not BasePart
	if _, ok := part.(*XmlPart); !ok {
		t.Errorf("expected selector to produce *XmlPart, got %T", part)
	}
}

func TestPartFactory_DefaultFallback(t *testing.T) {
	t.Parallel()

	factory := NewPartFactory()

	blob := []byte("binary data")
	part, err := factory.New("/word/media/image1.png", "image/png", RTImage, blob, nil)
	if err != nil {
		t.Fatalf("New: %v", err)
	}
	if _, ok := part.(*BasePart); !ok {
		t.Errorf("expected *BasePart fallback, got %T", part)
	}
	gotBlob, _ := part.Blob()
	if string(gotBlob) != "binary data" {
		t.Errorf("blob mismatch: got %q", string(gotBlob))
	}
}
