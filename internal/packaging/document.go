// Package packaging provides a high-level typed view over an OPC package.
// It wraps the lower-level opc layer from docx-go and classifies parts
// (Styles, Headers, Media, etc.) for convenient access by the service layer.
package packaging

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"

	"github.com/user/go-docx/pkg/docx/opc"
)

// --------------------------------------------------------------------------
// Document — high-level typed view over an OPC package
// --------------------------------------------------------------------------

// Document represents an opened .docx with parts classified by type.
type Document struct {
	pkg     *opc.OpcPackage
	docPart opc.Part

	// Core metadata (Dublin Core).
	CoreProps *CoreProperties

	// Extended / application properties.
	AppProps *AppProperties

	// Named XML parts stored as raw blobs (nil when absent).
	Styles    []byte
	Settings  []byte
	Numbering []byte
	Comments  []byte
	Footnotes []byte
	Endnotes  []byte
	Fonts     []byte

	// Single-instance blob parts (empty when absent).
	Theme       []byte
	WebSettings []byte

	// Multi-instance parts.
	Headers [][]byte
	Footers [][]byte

	// Media files keyed by part name (e.g. "/word/media/image1.png").
	Media map[string][]byte

	// Parts that don't match any known relationship type.
	UnknownParts []UnknownPart
}

// CoreProperties holds Dublin Core metadata from core.xml.
type CoreProperties struct {
	Title       string
	Creator     string
	Description string
}

// AppProperties holds extended-property metadata from app.xml.
type AppProperties struct {
	Application string
}

// UnknownPart is a package part with no recognised relationship type.
type UnknownPart struct {
	PartName    string
	ContentType string
	Blob        []byte
}

// --------------------------------------------------------------------------
// Open helpers
// --------------------------------------------------------------------------

// OpenReader opens a .docx from an io.ReaderAt.
func OpenReader(r io.ReaderAt, size int64) (*Document, error) {
	pkg, err := opc.Open(r, size, nil)
	if err != nil {
		return nil, fmt.Errorf("packaging: open: %w", err)
	}
	return classify(pkg)
}

// OpenBytes opens a .docx from in-memory bytes.
func OpenBytes(data []byte) (*Document, error) {
	pkg, err := opc.OpenBytes(data, nil)
	if err != nil {
		return nil, fmt.Errorf("packaging: open bytes: %w", err)
	}
	return classify(pkg)
}

// --------------------------------------------------------------------------
// Save helpers
// --------------------------------------------------------------------------

// SaveWriter writes the document back as a .docx ZIP archive.
func (d *Document) SaveWriter(w io.Writer) error {
	return d.pkg.Save(w)
}

// SaveBytes returns the document as a byte slice.
func (d *Document) SaveBytes() ([]byte, error) {
	var buf bytes.Buffer
	if err := d.SaveWriter(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// --------------------------------------------------------------------------
// classify — walk the OPC relationship graph and fill Document fields
// --------------------------------------------------------------------------

func classify(pkg *opc.OpcPackage) (*Document, error) {
	doc := &Document{
		pkg:   pkg,
		Media: make(map[string][]byte),
	}

	// Package-level relationships.
	for _, rel := range pkg.Rels().All() {
		if rel.IsExternal || rel.TargetPart == nil {
			continue
		}
		switch rel.RelType {
		case opc.RTOfficeDocument:
			doc.docPart = rel.TargetPart
		case opc.RTCoreProperties:
			if b, err := rel.TargetPart.Blob(); err == nil {
				doc.CoreProps = parseCoreProps(b)
			}
		case opc.RTExtendedProperties:
			if b, err := rel.TargetPart.Blob(); err == nil {
				doc.AppProps = parseAppProps(b)
			}
		}
	}

	if doc.docPart == nil {
		return nil, fmt.Errorf("packaging: no main document part found")
	}

	classified := make(map[opc.PackURI]bool)
	classified[doc.docPart.PartName()] = true

	// Document-level relationships (from document.xml.rels).
	if docRels := doc.docPart.Rels(); docRels != nil {
		for _, rel := range docRels.All() {
			if rel.IsExternal || rel.TargetPart == nil {
				continue
			}
			part := rel.TargetPart
			classified[part.PartName()] = true
			blob, err := part.Blob()
			if err != nil {
				return nil, fmt.Errorf("packaging: reading part %q: %w", part.PartName(), err)
			}

			switch rel.RelType {
			case opc.RTStyles:
				doc.Styles = blob
			case opc.RTSettings:
				doc.Settings = blob
			case opc.RTNumbering:
				doc.Numbering = blob
			case opc.RTComments:
				doc.Comments = blob
			case opc.RTFootnotes:
				doc.Footnotes = blob
			case opc.RTEndnotes:
				doc.Endnotes = blob
			case opc.RTFontTable:
				doc.Fonts = blob
			case opc.RTTheme:
				doc.Theme = blob
			case opc.RTWebSettings:
				doc.WebSettings = blob
			case opc.RTHeader:
				doc.Headers = append(doc.Headers, blob)
			case opc.RTFooter:
				doc.Footers = append(doc.Footers, blob)
			case opc.RTImage:
				doc.Media[string(part.PartName())] = blob
			default:
				if isMediaContentType(part.ContentType()) {
					doc.Media[string(part.PartName())] = blob
				}
			}
		}
	}

	// Mark package-level targets as classified too.
	for _, rel := range pkg.Rels().All() {
		if !rel.IsExternal && rel.TargetPart != nil {
			classified[rel.TargetPart.PartName()] = true
		}
	}

	// Remaining parts → UnknownParts.
	for _, part := range pkg.Parts() {
		if classified[part.PartName()] {
			continue
		}
		blob, err := part.Blob()
		if err != nil {
			return nil, fmt.Errorf("packaging: reading unknown part %q: %w", part.PartName(), err)
		}
		doc.UnknownParts = append(doc.UnknownParts, UnknownPart{
			PartName:    string(part.PartName()),
			ContentType: part.ContentType(),
			Blob:        blob,
		})
	}

	return doc, nil
}

func isMediaContentType(ct string) bool {
	return strings.HasPrefix(ct, "image/")
}

// --------------------------------------------------------------------------
// Minimal XML parsing for core / app properties
// --------------------------------------------------------------------------

type xmlCoreProperties struct {
	XMLName     xml.Name `xml:"coreProperties"`
	Title       string   `xml:"title"`
	Creator     string   `xml:"creator"`
	Description string   `xml:"description"`
}

func parseCoreProps(blob []byte) *CoreProperties {
	if len(blob) == 0 {
		return nil
	}
	var props xmlCoreProperties
	if err := xml.Unmarshal(blob, &props); err != nil {
		return &CoreProperties{}
	}
	return &CoreProperties{
		Title:       props.Title,
		Creator:     props.Creator,
		Description: props.Description,
	}
}

type xmlAppProperties struct {
	XMLName     xml.Name `xml:"Properties"`
	Application string   `xml:"Application"`
}

func parseAppProps(blob []byte) *AppProperties {
	if len(blob) == 0 {
		return nil
	}
	var props xmlAppProperties
	if err := xml.Unmarshal(blob, &props); err != nil {
		return &AppProperties{}
	}
	return &AppProperties{
		Application: props.Application,
	}
}
