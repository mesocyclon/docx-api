package packaging

import (
	"fmt"
	"path"
	"strconv"
	"strings"

	"github.com/vortex/docx-go/coreprops"
	"github.com/vortex/docx-go/opc"
	"github.com/vortex/docx-go/wml/hdft"

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

// load takes an opened OPC package and populates a Document by inspecting
// relationships and parsing each known part.
func load(pkg *opc.Package) (*Document, error) {
	d := &Document{
		pkg:          pkg,
		Headers:      make(map[string]*hdft.CT_HdrFtr),
		Footers:      make(map[string]*hdft.CT_HdrFtr),
		Media:        make(map[string][]byte),
		UnknownParts: make(map[string][]byte),
	}

	// ── 1. Resolve the main document part via package-level rels ──
	docRels := pkg.PackageRelsByType(relOfficeDocument)
	if len(docRels) == 0 {
		return nil, fmt.Errorf("packaging: no officeDocument relationship in package")
	}
	d.docPartName = normalizePartName(docRels[0].Target)

	docPart, ok := pkg.Part(d.docPartName)
	if !ok {
		return nil, fmt.Errorf("packaging: document part %q not found", d.docPartName)
	}

	// ── 2. Parse document.xml ─────────────────────────────────────
	doc, err := document.Parse(docPart.Data)
	if err != nil {
		return nil, fmt.Errorf("packaging: parsing document.xml: %w", err)
	}
	d.Document = doc

	// ── 3. Walk document-level relationships ──────────────────────
	knownDocRels := make(map[string]bool) // track handled rIds

	if err := d.loadByRel(docPart, relStyles, &knownDocRels, func(data []byte) error {
		s, err := styles.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing styles.xml: %w", err)
		}
		d.Styles = s
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relSettings, &knownDocRels, func(data []byte) error {
		s, err := settings.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing settings.xml: %w", err)
		}
		d.Settings = s
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relFontTable, &knownDocRels, func(data []byte) error {
		f, err := fonts.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing fontTable.xml: %w", err)
		}
		d.Fonts = f
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relNumbering, &knownDocRels, func(data []byte) error {
		n, err := numbering.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing numbering.xml: %w", err)
		}
		d.Numbering = n
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relFootnotes, &knownDocRels, func(data []byte) error {
		fn, err := footnotes.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing footnotes.xml: %w", err)
		}
		d.Footnotes = fn
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relEndnotes, &knownDocRels, func(data []byte) error {
		en, err := footnotes.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing endnotes.xml: %w", err)
		}
		d.Endnotes = en
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relComments, &knownDocRels, func(data []byte) error {
		c, err := comments.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing comments.xml: %w", err)
		}
		d.Comments = c
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relWebSettings, &knownDocRels, func(data []byte) error {
		raw, err := websettings.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing webSettings.xml: %w", err)
		}
		d.WebSettings = raw
		return nil
	}); err != nil {
		return nil, err
	}

	if err := d.loadByRel(docPart, relTheme, &knownDocRels, func(data []byte) error {
		raw, err := theme.Parse(data)
		if err != nil {
			return fmt.Errorf("parsing theme.xml: %w", err)
		}
		d.Theme = raw
		return nil
	}); err != nil {
		return nil, err
	}

	// ── 4. Headers & footers (there may be multiple of each) ──────
	for _, rel := range docPart.RelsByType(relHeader) {
		knownDocRels[rel.ID] = true
		partName := resolveTarget(d.docPartName, rel.Target)
		part, ok := pkg.Part(partName)
		if !ok {
			continue
		}
		hdr, err := headers.Parse(part.Data)
		if err != nil {
			return nil, fmt.Errorf("packaging: parsing header %s: %w", partName, err)
		}
		d.Headers[rel.ID] = hdr
	}
	for _, rel := range docPart.RelsByType(relFooter) {
		knownDocRels[rel.ID] = true
		partName := resolveTarget(d.docPartName, rel.Target)
		part, ok := pkg.Part(partName)
		if !ok {
			continue
		}
		ftr, err := headers.Parse(part.Data)
		if err != nil {
			return nil, fmt.Errorf("packaging: parsing footer %s: %w", partName, err)
		}
		d.Footers[rel.ID] = ftr
	}

	// ── 5. Images / media ─────────────────────────────────────────
	for _, rel := range docPart.RelsByType(relImage) {
		knownDocRels[rel.ID] = true
		if rel.TargetMode == "External" {
			continue
		}
		partName := resolveTarget(d.docPartName, rel.Target)
		part, ok := pkg.Part(partName)
		if !ok {
			continue
		}
		shortName := path.Base(partName)
		d.Media[shortName] = part.Data
	}

	// ── 6. Collect unknown document-level relationships ───────────
	for _, rel := range docPart.Rels {
		if knownDocRels[rel.ID] {
			continue
		}
		if rel.TargetMode == "External" {
			// External hyperlinks etc. — preserve as-is
			d.UnknownRels = append(d.UnknownRels, rel)
			continue
		}
		// Unknown internal relationship — preserve part bytes
		partName := resolveTarget(d.docPartName, rel.Target)
		if part, ok := pkg.Part(partName); ok {
			d.UnknownParts[partName] = part.Data
		}
		d.UnknownRels = append(d.UnknownRels, rel)
	}

	// ── 7. Package-level properties (core.xml, app.xml) ───────────
	knownPkgRels := make(map[string]bool)
	knownPkgRels[docRels[0].ID] = true

	for _, rel := range pkg.PackageRelsByType(relCoreProperties) {
		knownPkgRels[rel.ID] = true
		partName := normalizePartName(rel.Target)
		if part, ok := pkg.Part(partName); ok {
			cp, err := coreprops.ParseCore(part.Data)
			if err != nil {
				return nil, fmt.Errorf("packaging: parsing core.xml: %w", err)
			}
			d.CoreProps = cp
		}
	}
	for _, rel := range pkg.PackageRelsByType(relExtProperties) {
		knownPkgRels[rel.ID] = true
		partName := normalizePartName(rel.Target)
		if part, ok := pkg.Part(partName); ok {
			ap, err := coreprops.ParseApp(part.Data)
			if err != nil {
				return nil, fmt.Errorf("packaging: parsing app.xml: %w", err)
			}
			d.AppProps = ap
		}
	}

	// Preserve unknown package-level rels
	for _, rel := range pkg.PackageRels() {
		if !knownPkgRels[rel.ID] {
			d.UnknownRels = append(d.UnknownRels, rel)
		}
	}

	// ── 8. Compute next rId and bookmark ID seeds ─────────────────
	d.seedRelIDCounter(docPart)
	d.seedBookmarkIDCounter()

	return d, nil
}

// ──────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────

// loadByRel looks up the first relationship of the given type on a part,
// resolves the target, reads the data and calls parse. If no relationship
// exists, the function is a no-op (many parts are optional).
func (d *Document) loadByRel(
	docPart *opc.Part,
	relType string,
	known *map[string]bool,
	parse func([]byte) error,
) error {
	rels := docPart.RelsByType(relType)
	if len(rels) == 0 {
		return nil
	}
	rel := rels[0]
	(*known)[rel.ID] = true

	partName := resolveTarget(d.docPartName, rel.Target)
	part, ok := d.pkg.Part(partName)
	if !ok {
		return nil // silently skip missing parts — Word can recover
	}
	if err := parse(part.Data); err != nil {
		return fmt.Errorf("packaging: %w", err)
	}
	return nil
}

// normalizePartName ensures the part name starts with '/'.
func normalizePartName(target string) string {
	if !strings.HasPrefix(target, "/") {
		return "/" + target
	}
	return target
}

// resolveTarget resolves a relative target against the directory of the
// source part. For example, given source "/word/document.xml" and target
// "styles.xml", returns "/word/styles.xml".
func resolveTarget(sourcePart, target string) string {
	if strings.HasPrefix(target, "/") {
		return target
	}
	dir := path.Dir(sourcePart)
	return path.Join(dir, target)
}

// seedRelIDCounter scans all existing relationship IDs on the document part
// and sets nextRelSeq to max+1 so that NextRelID never collides.
func (d *Document) seedRelIDCounter(docPart *opc.Part) {
	maxID := 0
	for _, rel := range docPart.Rels {
		n := parseRelIDNum(rel.ID)
		if n > maxID {
			maxID = n
		}
	}
	// Also scan headers/footers for their own rels if needed
	d.nextRelSeq = maxID + 1
}

// seedBookmarkIDCounter scans the document body (and comments, footnotes)
// for the highest numeric ID used in bookmarks, comments, ins/del, and
// sets nextBmkID accordingly. For a simple implementation we start at a
// safe high value; a full implementation would walk the DOM.
func (d *Document) seedBookmarkIDCounter() {
	// Start at a reasonably high value to avoid collisions with IDs
	// already present in a typical document.
	d.nextBmkID = 100
}

// parseRelIDNum extracts the numeric suffix from an rId string ("rId3" → 3).
func parseRelIDNum(id string) int {
	s := strings.TrimPrefix(id, "rId")
	n, _ := strconv.Atoi(s)
	return n
}
