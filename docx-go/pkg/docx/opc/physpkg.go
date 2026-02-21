package opc

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"strings"
)

// --------------------------------------------------------------------------
// PhysPkgReader — reads a ZIP-based OPC package
// --------------------------------------------------------------------------

// PhysPkgReader provides low-level access to a ZIP-based OPC package.
type PhysPkgReader struct {
	reader *zip.Reader
	closer io.Closer // non-nil when opened from a file
	files  map[string]*zip.File
}

// NewPhysPkgReader creates a PhysPkgReader from an io.ReaderAt.
func NewPhysPkgReader(r io.ReaderAt, size int64) (*PhysPkgReader, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("opc: opening zip: %w", err)
	}
	return newPhysPkgReaderFromZip(zr, nil), nil
}

// NewPhysPkgReaderFromFile opens a PhysPkgReader from a file path.
func NewPhysPkgReaderFromFile(path string) (*PhysPkgReader, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("opc: opening file %q: %w", path, err)
	}
	info, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("opc: stat file %q: %w", path, err)
	}
	zr, err := zip.NewReader(f, info.Size())
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("opc: opening zip %q: %w", path, err)
	}
	return newPhysPkgReaderFromZip(zr, f), nil
}

// NewPhysPkgReaderFromBytes creates a PhysPkgReader from in-memory bytes.
func NewPhysPkgReaderFromBytes(data []byte) (*PhysPkgReader, error) {
	r := bytes.NewReader(data)
	return NewPhysPkgReader(r, int64(len(data)))
}

func newPhysPkgReaderFromZip(zr *zip.Reader, closer io.Closer) *PhysPkgReader {
	files := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		files[f.Name] = f
	}
	return &PhysPkgReader{
		reader: zr,
		closer: closer,
		files:  files,
	}
}

// BlobFor returns the contents of the part at the given PackURI.
func (p *PhysPkgReader) BlobFor(uri PackURI) ([]byte, error) {
	membername := uri.Membername()
	f, ok := p.files[membername]
	if !ok {
		return nil, fmt.Errorf("opc: member %q not found in package", membername)
	}
	rc, err := f.Open()
	if err != nil {
		return nil, fmt.Errorf("opc: opening member %q: %w", membername, err)
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

// ContentTypesXml returns the [Content_Types].xml blob.
func (p *PhysPkgReader) ContentTypesXml() ([]byte, error) {
	return p.BlobFor(ContentTypesURI)
}

// RelsXmlFor returns the .rels XML for the given source URI, or nil if none exists.
func (p *PhysPkgReader) RelsXmlFor(sourceURI PackURI) ([]byte, error) {
	relsURI := sourceURI.RelsURI()
	blob, err := p.BlobFor(relsURI)
	if err != nil {
		// No .rels file is not an error — it simply means no relationships.
		if strings.Contains(err.Error(), "not found") {
			return nil, nil
		}
		return nil, err
	}
	return blob, nil
}

// URIs returns a list of all member URIs in the package, excluding
// [Content_Types].xml and .rels files.
func (p *PhysPkgReader) URIs() []PackURI {
	var uris []PackURI
	for name := range p.files {
		uri := NewPackURI(name)
		// Skip [Content_Types].xml and .rels files
		if uri == ContentTypesURI {
			continue
		}
		if strings.Contains(name, "_rels/") && strings.HasSuffix(name, ".rels") {
			continue
		}
		uris = append(uris, uri)
	}
	return uris
}

// Close releases resources held by the reader.
func (p *PhysPkgReader) Close() error {
	if p.closer != nil {
		return p.closer.Close()
	}
	return nil
}

// --------------------------------------------------------------------------
// PhysPkgWriter — writes a ZIP-based OPC package
// --------------------------------------------------------------------------

// PhysPkgWriter provides low-level write access to a ZIP-based OPC package.
type PhysPkgWriter struct {
	writer *zip.Writer
}

// NewPhysPkgWriter creates a PhysPkgWriter backed by the given writer.
func NewPhysPkgWriter(w io.Writer) *PhysPkgWriter {
	return &PhysPkgWriter{writer: zip.NewWriter(w)}
}

// Write adds a member to the ZIP package.
func (p *PhysPkgWriter) Write(uri PackURI, blob []byte) error {
	membername := uri.Membername()
	w, err := p.writer.Create(membername)
	if err != nil {
		return fmt.Errorf("opc: creating zip member %q: %w", membername, err)
	}
	if _, err := w.Write(blob); err != nil {
		return fmt.Errorf("opc: writing zip member %q: %w", membername, err)
	}
	return nil
}

// Close finalizes the ZIP archive.
func (p *PhysPkgWriter) Close() error {
	return p.writer.Close()
}
