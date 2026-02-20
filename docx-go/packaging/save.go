package packaging

import (
	"fmt"
	"path"
	"strings"

	"github.com/vortex/docx-go/coreprops"
	"github.com/vortex/docx-go/opc"

	"github.com/vortex/docx-go/parts/comments"
	"github.com/vortex/docx-go/parts/document"
	"github.com/vortex/docx-go/parts/fonts"
	"github.com/vortex/docx-go/parts/footnotes"
	"github.com/vortex/docx-go/parts/headers"
	"github.com/vortex/docx-go/parts/numbering"
	"github.com/vortex/docx-go/parts/settings"
	"github.com/vortex/docx-go/parts/styles"
	"github.com/vortex/docx-go/parts/theme"
	"github.com/vortex/docx-go/parts/websettings"
)

// buildPackage serialises every typed part back into the OPC package,
// rebuilding relationships and content types. Existing parts that were
// not modified keep their original bytes; typed parts are re-serialised.
func (d *Document) buildPackage() error {
	// Start from a fresh OPC package so we don't carry stale parts.
	pkg := opc.New()
	d.pkg = pkg

	// ── 1. Determine document part name ───────────────────────────
	docPartName := d.docPartName
	if docPartName == "" {
		docPartName = "/word/document.xml"
	}
	docDir := path.Dir(docPartName) // "/word"

	// ── 2. Serialise and add document.xml ─────────────────────────
	if d.Document != nil {
		data, err := document.Serialize(d.Document)
		if err != nil {
			return fmt.Errorf("packaging: serialising document.xml: %w", err)
		}
		docPart := pkg.AddPart(docPartName, ctDocument, data)

		// ── 3. Add document-level relationships and their parts ───
		if err := d.addDocParts(pkg, docPart, docDir); err != nil {
			return err
		}
	}

	// ── 4. Add package-level relationships ────────────────────────
	pkg.AddPackageRel(relOfficeDocument, strings.TrimPrefix(docPartName, "/"))

	// Core properties
	if d.CoreProps != nil {
		data, err := coreprops.SerializeCore(d.CoreProps)
		if err != nil {
			return fmt.Errorf("packaging: serialising core.xml: %w", err)
		}
		pkg.AddPart("/docProps/core.xml", ctCore, data)
		pkg.AddPackageRel(relCoreProperties, "docProps/core.xml")
	}

	// Extended (app) properties
	if d.AppProps != nil {
		data, err := coreprops.SerializeApp(d.AppProps)
		if err != nil {
			return fmt.Errorf("packaging: serialising app.xml: %w", err)
		}
		pkg.AddPart("/docProps/app.xml", ctExtended, data)
		pkg.AddPackageRel(relExtProperties, "docProps/app.xml")
	}

	// Unknown package-level rels (e.g. custom properties, thumbnails)
	for _, rel := range d.UnknownRels {
		// Only add package-level unknowns here; doc-level ones were
		// handled in addDocParts.
		if isPackageLevelRel(rel.Type) {
			partName := normalizePartName(rel.Target)
			if data, ok := d.UnknownParts[partName]; ok {
				pkg.AddPart(partName, guessContentType(partName), data)
			}
			pkg.AddPackageRel(rel.Type, rel.Target)
		}
	}

	return nil
}

// addDocParts serialises all document-level parts (styles, settings, etc.)
// and adds them as relationships on docPart.
func (d *Document) addDocParts(pkg *opc.Package, docPart *opc.Part, docDir string) error {
	// Helper: serialise, add part, add rel.
	addPart := func(relType, fileName, ct string, data []byte) {
		partName := docDir + "/" + fileName
		pkg.AddPart(partName, ct, data)
		docPart.AddRel(relType, fileName)
	}

	// ── Styles ────────────────────────────────────────────────────
	if d.Styles != nil {
		data, err := styles.Serialize(d.Styles)
		if err != nil {
			return fmt.Errorf("packaging: serialising styles.xml: %w", err)
		}
		addPart(relStyles, "styles.xml", ctStyles, data)
	}

	// ── Settings ──────────────────────────────────────────────────
	if d.Settings != nil {
		data, err := settings.Serialize(d.Settings)
		if err != nil {
			return fmt.Errorf("packaging: serialising settings.xml: %w", err)
		}
		addPart(relSettings, "settings.xml", ctSettings, data)
	}

	// ── Fonts ─────────────────────────────────────────────────────
	if d.Fonts != nil {
		data, err := fonts.Serialize(d.Fonts)
		if err != nil {
			return fmt.Errorf("packaging: serialising fontTable.xml: %w", err)
		}
		addPart(relFontTable, "fontTable.xml", ctFontTable, data)
	}

	// ── Numbering (optional) ──────────────────────────────────────
	if d.Numbering != nil {
		data, err := numbering.Serialize(d.Numbering)
		if err != nil {
			return fmt.Errorf("packaging: serialising numbering.xml: %w", err)
		}
		addPart(relNumbering, "numbering.xml", ctNumbering, data)
	}

	// ── Footnotes (optional) ──────────────────────────────────────
	if d.Footnotes != nil {
		data, err := footnotes.Serialize(d.Footnotes)
		if err != nil {
			return fmt.Errorf("packaging: serialising footnotes.xml: %w", err)
		}
		addPart(relFootnotes, "footnotes.xml", ctFootnotes, data)
	}

	// ── Endnotes (optional) ───────────────────────────────────────
	if d.Endnotes != nil {
		data, err := footnotes.Serialize(d.Endnotes)
		if err != nil {
			return fmt.Errorf("packaging: serialising endnotes.xml: %w", err)
		}
		addPart(relEndnotes, "endnotes.xml", ctEndnotes, data)
	}

	// ── Comments (optional) ───────────────────────────────────────
	if d.Comments != nil {
		data, err := comments.Serialize(d.Comments)
		if err != nil {
			return fmt.Errorf("packaging: serialising comments.xml: %w", err)
		}
		addPart(relComments, "comments.xml", ctComments, data)
	}

	// ── Web Settings (raw) ────────────────────────────────────────
	if len(d.WebSettings) > 0 {
		ws, err := websettings.Serialize(d.WebSettings)
		if err != nil {
			return fmt.Errorf("packaging: serialising webSettings.xml: %w", err)
		}
		addPart(relWebSettings, "webSettings.xml", ctWebSettings, ws)
	}

	// ── Theme (raw) ───────────────────────────────────────────────
	if len(d.Theme) > 0 {
		th, err := theme.Serialize(d.Theme)
		if err != nil {
			return fmt.Errorf("packaging: serialising theme1.xml: %w", err)
		}
		partName := docDir + "/theme/theme1.xml"
		pkg.AddPart(partName, ctTheme, th)
		docPart.AddRel(relTheme, "theme/theme1.xml")
	}

	// ── Headers ───────────────────────────────────────────────────
	headerIdx := 1
	for rID, hdr := range d.Headers {
		fileName := fmt.Sprintf("header%d.xml", headerIdx)
		headerIdx++
		data, err := headers.Serialize(hdr)
		if err != nil {
			return fmt.Errorf("packaging: serialising %s (rel %s): %w", fileName, rID, err)
		}
		partName := docDir + "/" + fileName
		pkg.AddPart(partName, ctHeader, data)
		docPart.AddRel(relHeader, fileName)
	}

	// ── Footers ───────────────────────────────────────────────────
	footerIdx := 1
	for rID, ftr := range d.Footers {
		fileName := fmt.Sprintf("footer%d.xml", footerIdx)
		footerIdx++
		data, err := headers.SerializeFooter(ftr)
		if err != nil {
			return fmt.Errorf("packaging: serialising %s (rel %s): %w", fileName, rID, err)
		}
		partName := docDir + "/" + fileName
		pkg.AddPart(partName, ctFooter, data)
		docPart.AddRel(relFooter, fileName)
	}

	// ── Media (images) ────────────────────────────────────────────
	for shortName, data := range d.Media {
		partName := docDir + "/media/" + shortName
		ct := guessMediaContentType(shortName)
		pkg.AddPart(partName, ct, data)
		docPart.AddRel(relImage, "media/"+shortName)
	}

	// ── Unknown document-level parts ──────────────────────────────
	for _, rel := range d.UnknownRels {
		if isPackageLevelRel(rel.Type) {
			continue // handled at package level
		}
		if rel.TargetMode == "External" {
			docPart.AddExternalRel(rel.Type, rel.Target)
			continue
		}
		target := rel.Target
		partName := resolveTarget(d.docPartName, target)
		if data, ok := d.UnknownParts[partName]; ok {
			pkg.AddPart(partName, guessContentType(partName), data)
		}
		docPart.AddRel(rel.Type, target)
	}

	return nil
}

// ──────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────

// isPackageLevelRel returns true for relationship types that belong at the
// package level (/_rels/.rels) rather than the document level.
func isPackageLevelRel(relType string) bool {
	switch relType {
	case relCoreProperties, relExtProperties:
		return true
	}
	// Thumbnail and custom-properties are also package-level
	if strings.Contains(relType, "metadata/") {
		return true
	}
	return false
}

// guessMediaContentType returns the MIME content type for a media filename.
func guessMediaContentType(filename string) string {
	ext := strings.ToLower(path.Ext(filename))
	if ct, ok := mediaContentTypes[ext]; ok {
		return ct
	}
	return "application/octet-stream"
}

// guessContentType provides a best-effort content type for unknown parts
// based on file extension.
func guessContentType(partName string) string {
	ext := strings.ToLower(path.Ext(partName))
	switch ext {
	case ".xml":
		return "application/xml"
	case ".rels":
		return "application/vnd.openxmlformats-package.relationships+xml"
	}
	return guessMediaContentType(partName)
}
