// Package packaging provides the central document model that ties together
// all OPC parts into a single, navigable [Document]. It handles reading and
// writing .docx files, dispatching each part to the appropriate typed parser
// from parts/*, and preserving unknown parts/relationships for lossless
// round-trip.
//
// Contract: C-30 in contracts.md
package packaging

import (
	"io"
	"sync"

	"github.com/vortex/docx-go/coreprops"
	"github.com/vortex/docx-go/opc"
	"github.com/vortex/docx-go/wml/body"
	"github.com/vortex/docx-go/wml/hdft"

	"github.com/vortex/docx-go/parts/comments"
	"github.com/vortex/docx-go/parts/fonts"
	"github.com/vortex/docx-go/parts/footnotes"
	"github.com/vortex/docx-go/parts/numbering"
	"github.com/vortex/docx-go/parts/settings"
	"github.com/vortex/docx-go/parts/styles"
)

// ──────────────────────────────────────────────
// Relationship-type URIs (from reference-appendix §1.3)
// ──────────────────────────────────────────────

const (
	// Package-level relationship types
	relOfficeDocument = opc.RelOfficeDocument
	relCoreProperties = opc.RelCoreProperties
	relExtProperties  = opc.RelExtProperties

	// Part-level relationship types (from document.xml)
	relStyles      = opc.RelStyles
	relSettings    = opc.RelSettings
	relWebSettings = opc.RelWebSettings
	relFontTable   = opc.RelFontTable
	relNumbering   = opc.RelNumbering
	relFootnotes   = opc.RelFootnotes
	relEndnotes    = opc.RelEndnotes
	relComments    = opc.RelComments
	relHeader      = opc.RelHeader
	relFooter      = opc.RelFooter
	relImage       = opc.RelImage
	relTheme       = opc.RelTheme
)

// ──────────────────────────────────────────────
// Content-type strings (from reference-appendix §1.4)
// ──────────────────────────────────────────────

const (
	ctDocument    = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
	ctStyles      = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
	ctSettings    = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
	ctFontTable   = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
	ctNumbering   = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
	ctFootnotes   = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
	ctEndnotes    = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
	ctComments    = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
	ctHeader      = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
	ctFooter      = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
	ctWebSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
	ctTheme       = "application/vnd.openxmlformats-officedocument.theme+xml"
	ctCore        = "application/vnd.openxmlformats-package.core-properties+xml"
	ctExtended    = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
)

// mediaContentTypes maps file extensions to MIME content types.
var mediaContentTypes = map[string]string{
	".png":  "image/png",
	".jpg":  "image/jpeg",
	".jpeg": "image/jpeg",
	".gif":  "image/gif",
	".bmp":  "image/bmp",
	".tiff": "image/tiff",
	".tif":  "image/tiff",
	".svg":  "image/svg+xml",
	".emf":  "image/x-emf",
	".wmf":  "image/x-wmf",
}

// ──────────────────────────────────────────────
// Document — the central model
// ──────────────────────────────────────────────

// Document holds all typed and raw parts of a .docx file. It is the central
// data structure through which every module accesses, modifies, and persists
// the document.
type Document struct {
	// ── Typed parts ────────────────────────────
	Document  *body.CT_Document
	Styles    *styles.CT_Styles
	Numbering *numbering.CT_Numbering // nil if absent
	Settings  *settings.CT_Settings
	Fonts     *fonts.CT_FontsList
	Comments  *comments.CT_Comments     // nil if absent
	Footnotes *footnotes.CT_Footnotes   // nil if absent
	Endnotes  *footnotes.CT_Footnotes   // nil if absent
	CoreProps *coreprops.CoreProperties
	AppProps  *coreprops.AppProperties

	// ── Headers and footers: rId → parsed ──────
	Headers map[string]*hdft.CT_HdrFtr
	Footers map[string]*hdft.CT_HdrFtr

	// ── Raw / pass-through parts ───────────────
	Theme       []byte
	WebSettings []byte

	// Media stores embedded images/files. Key is the short filename
	// relative to word/media/ (e.g. "image1.png").
	Media map[string][]byte

	// UnknownParts preserves parts whose relationship type is not
	// recognised, keyed by their full part name (e.g. "/word/glossary/document.xml").
	UnknownParts map[string][]byte

	// UnknownRels preserves package-level and document-level
	// relationships whose type is not handled by this module.
	UnknownRels []opc.Relationship

	// ── Internal state ─────────────────────────
	pkg *opc.Package

	mu         sync.Mutex
	nextRelSeq int // monotonically increasing suffix for rId generation
	nextBmkID  int // monotonically increasing bookmark/comment/ins/del ID

	// docPartName is the part name of word/document.xml (resolved from rels).
	docPartName string
}

// ──────────────────────────────────────────────
// Public constructors (implementation in open.go)
// ──────────────────────────────────────────────

// Open reads a .docx file from disk and returns a fully parsed Document.
func Open(path string) (*Document, error) {
	pkg, err := opc.OpenFile(path)
	if err != nil {
		return nil, err
	}
	return load(pkg)
}

// OpenReader reads a .docx from an io.ReaderAt (e.g. bytes.NewReader).
func OpenReader(r io.ReaderAt, size int64) (*Document, error) {
	pkg, err := opc.OpenReader(r, size)
	if err != nil {
		return nil, err
	}
	return load(pkg)
}

// ──────────────────────────────────────────────
// Public save methods (implementation in save.go)
// ──────────────────────────────────────────────

// Save writes the document to disk at the given path.
func (d *Document) Save(path string) error {
	if err := d.buildPackage(); err != nil {
		return err
	}
	return d.pkg.SaveFile(path)
}

// SaveWriter serialises the document as a ZIP stream to w.
func (d *Document) SaveWriter(w io.Writer) error {
	if err := d.buildPackage(); err != nil {
		return err
	}
	return d.pkg.SaveWriter(w)
}
