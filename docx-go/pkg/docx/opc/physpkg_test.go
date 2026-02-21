package opc

import (
	"bytes"
	"testing"

	"github.com/user/go-docx/pkg/docx/templates"
)

func loadDefaultDocx(t *testing.T) []byte {
	t.Helper()
	data, err := templates.FS.ReadFile("default.docx")
	if err != nil {
		t.Fatalf("reading default.docx: %v", err)
	}
	return data
}

func TestPhysPkgReader_Open(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	uris := reader.URIs()
	if len(uris) == 0 {
		t.Fatal("expected non-empty URIs list")
	}

	// Check that /word/document.xml is among the URIs
	found := false
	for _, uri := range uris {
		if uri == "/word/document.xml" {
			found = true
			break
		}
	}
	if !found {
		t.Error("expected /word/document.xml in URIs list")
	}
}

func TestPhysPkgReader_BlobFor(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.BlobFor("/word/document.xml")
	if err != nil {
		t.Fatalf("BlobFor: %v", err)
	}
	if len(blob) == 0 {
		t.Error("expected non-empty blob for /word/document.xml")
	}
}

func TestPhysPkgReader_ContentTypesXml(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.ContentTypesXml()
	if err != nil {
		t.Fatalf("ContentTypesXml: %v", err)
	}
	if len(blob) == 0 {
		t.Error("expected non-empty [Content_Types].xml")
	}
	if !bytes.Contains(blob, []byte("ContentType")) {
		t.Error("expected [Content_Types].xml to contain ContentType")
	}
}

func TestPhysPkgReader_RelsXmlFor(t *testing.T) {
	data := loadDefaultDocx(t)
	reader, err := NewPhysPkgReaderFromBytes(data)
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.RelsXmlFor(PackageURI)
	if err != nil {
		t.Fatalf("RelsXmlFor: %v", err)
	}
	if blob == nil {
		t.Error("expected package-level .rels to exist")
	}
	if !bytes.Contains(blob, []byte("Relationship")) {
		t.Error("expected .rels to contain Relationship elements")
	}
}

func TestPhysPkgWriter_RoundTrip(t *testing.T) {
	// Write a simple package and read it back
	var buf bytes.Buffer
	writer := NewPhysPkgWriter(&buf)
	err := writer.Write("/test/data.xml", []byte("<root/>"))
	if err != nil {
		t.Fatalf("Write: %v", err)
	}
	err = writer.Close()
	if err != nil {
		t.Fatalf("Close: %v", err)
	}

	// Read back
	reader, err := NewPhysPkgReaderFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("NewPhysPkgReaderFromBytes: %v", err)
	}
	defer reader.Close()

	blob, err := reader.BlobFor("/test/data.xml")
	if err != nil {
		t.Fatalf("BlobFor: %v", err)
	}
	if string(blob) != "<root/>" {
		t.Errorf("got %q, want %q", string(blob), "<root/>")
	}
}
