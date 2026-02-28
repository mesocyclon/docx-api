package docx

import (
	"fmt"
	"io"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Document is the top-level object for a .docx file.
//
// Mirrors Python Document(ElementProxy).
type Document struct {
	element *oxml.CT_Document
	part    *parts.DocumentPart
	wmlPkg  *parts.WmlPackage
	body    *Body // lazy, mirrors Python _body
}

// newDocument creates a Document from its constituent pieces.
func newDocument(docPart *parts.DocumentPart, wmlPkg *parts.WmlPackage) (*Document, error) {
	el := docPart.Element()
	if el == nil {
		return nil, fmt.Errorf("docx: document part element is nil")
	}
	ctDoc := &oxml.CT_Document{Element: oxml.WrapElement(el)}
	return &Document{
		element: ctDoc,
		part:    docPart,
		wmlPkg:  wmlPkg,
	}, nil
}

// --------------------------------------------------------------------------
// Content mutation
// --------------------------------------------------------------------------

// AddHeading appends a heading paragraph to the end of the document.
// Level 0 produces a "Title" style; 1-9 produce "Heading N".
//
// Mirrors Python Document.add_heading.
func (d *Document) AddHeading(text string, level int) (*Paragraph, error) {
	if level < 0 || level > 9 {
		return nil, fmt.Errorf("docx: level must be in range 0-9, got %d", level)
	}
	style := "Title"
	if level > 0 {
		style = fmt.Sprintf("Heading %d", level)
	}
	return d.AddParagraph(text, StyleName(style))
}

// AddPageBreak appends a new paragraph containing only a page break.
//
// Mirrors Python Document.add_page_break.
func (d *Document) AddPageBreak() (*Paragraph, error) {
	para, err := d.AddParagraph("")
	if err != nil {
		return nil, err
	}
	run, err := para.AddRun("")
	if err != nil {
		return nil, err
	}
	if err := run.AddBreak(enum.WdBreakTypePage); err != nil {
		return nil, err
	}
	return para, nil
}

// AddParagraph appends a new paragraph to the end of the document body.
// text may contain tab (\t) and newline (\n, \r) characters. style may
// be a StyleName or a *BaseStyle. Omit to apply no explicit style.
//
// Mirrors Python Document.add_paragraph → self._body.add_paragraph(text, style).
func (d *Document) AddParagraph(text string, style ...StyleRef) (*Paragraph, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, err
	}
	return b.AddParagraph(text, style...)
}

// AddPicture adds an inline picture in its own paragraph at the end of the
// document. The image is read from r. width and height are optional EMU
// values; pass nil for native/proportional sizing.
//
// Mirrors Python Document.add_picture → add_paragraph().add_run().add_picture().
func (d *Document) AddPicture(r io.ReadSeeker, width, height *int64) (*InlineShape, error) {
	para, err := d.AddParagraph("")
	if err != nil {
		return nil, fmt.Errorf("docx: add picture paragraph: %w", err)
	}
	run, err := para.AddRun("")
	if err != nil {
		return nil, fmt.Errorf("docx: add picture run: %w", err)
	}
	return run.AddPicture(r, width, height)
}

// AddSection adds a new section break at the end of the document and returns
// the new Section. startType defaults to WdSectionStartNewPage.
//
// Mirrors Python Document.add_section.
func (d *Document) AddSection(startType enum.WdSectionStart) (*Section, error) {
	body := d.element.Body()
	if body == nil {
		return nil, fmt.Errorf("docx: document has no body")
	}
	newSectPr := body.AddSectionBreak()
	if err := newSectPr.SetStartType(startType); err != nil {
		return nil, fmt.Errorf("docx: setting section start type: %w", err)
	}
	return newSection(newSectPr, d.part), nil
}

// AddTable appends a new table with the given row and column counts.
// style may be a StyleName or *BaseStyle. Omit to apply no explicit style.
//
// Mirrors Python Document.add_table.
func (d *Document) AddTable(rows, cols int, style ...StyleRef) (*Table, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, err
	}
	bw, err := d.blockWidth()
	if err != nil {
		return nil, err
	}
	table, err := b.AddTable(rows, cols, bw)
	if err != nil {
		return nil, err
	}
	if raw := resolveStyleRef(style); raw != nil {
		if err := table.setStyleRaw(raw); err != nil {
			return nil, fmt.Errorf("docx: setting table style: %w", err)
		}
	}
	return table, nil
}

// AddComment adds a comment anchored to the specified runs.
// runs must contain at least one Run; the first and last are used to
// delimit the comment range. text, author, and initials populate the
// comment metadata.
//
// Mirrors Python Document.add_comment.
func (d *Document) AddComment(runs []*Run, text, author string, initials *string) (*Comment, error) {
	if len(runs) == 0 {
		return nil, fmt.Errorf("docx: at least one run required for comment")
	}
	firstRun := runs[0]
	lastRun := runs[len(runs)-1]

	comments, err := d.Comments()
	if err != nil {
		return nil, err
	}
	comment, err := comments.AddComment(text, author, initials)
	if err != nil {
		return nil, err
	}
	commentID, err := comment.CommentID()
	if err != nil {
		return nil, fmt.Errorf("docx: getting comment ID: %w", err)
	}
	if err := firstRun.MarkCommentRange(lastRun, commentID); err != nil {
		return nil, fmt.Errorf("docx: marking comment range: %w", err)
	}
	return comment, nil
}

// ReplaceText replaces all occurrences of old with new throughout the entire
// document: body, headers, footers, and comments of all sections.
//
// Headers/footers without their own definition (linked to previous) are
// skipped. Additionally, already-processed StoryParts are tracked by pointer
// to avoid double replacement when multiple sections share the same
// HeaderPart/FooterPart.
//
// Returns the total number of replacements performed.
func (d *Document) ReplaceText(old, new string) (int, error) {
	if old == "" {
		return 0, nil
	}

	// 1. Document body.
	b, err := d.getBody()
	if err != nil {
		return 0, err
	}
	count := b.ReplaceText(old, new)

	// 2. Headers/footers of all sections, with deduplication.
	seen := map[*parts.StoryPart]bool{}
	for _, sect := range d.Sections().Iter() {
		hfs := []*baseHeaderFooter{
			&sect.Header().baseHeaderFooter,
			&sect.Footer().baseHeaderFooter,
			&sect.EvenPageHeader().baseHeaderFooter,
			&sect.EvenPageFooter().baseHeaderFooter,
			&sect.FirstPageHeader().baseHeaderFooter,
			&sect.FirstPageFooter().baseHeaderFooter,
		}
		for _, hf := range hfs {
			n, err := hf.replaceTextDedup(old, new, seen)
			if err != nil {
				return count, fmt.Errorf("docx: replacing text in %s: %w", hf.ops.kind(), err)
			}
			count += n
		}
	}

	// 3. Comments.
	n, err := d.replaceTextInComments(old, new)
	if err != nil {
		return count, err
	}
	count += n

	return count, nil
}

// replaceTextInComments replaces text in all comments. Returns 0 if
// no comments part exists (avoids creating one as a side effect).
func (d *Document) replaceTextInComments(old, new string) (int, error) {
	if !d.part.HasCommentsPart() {
		return 0, nil
	}
	comments, err := d.Comments()
	if err != nil {
		return 0, fmt.Errorf("docx: replacing text in comments: %w", err)
	}
	return comments.ReplaceText(old, new), nil
}

// --------------------------------------------------------------------------
// Properties
// --------------------------------------------------------------------------

// Comments returns the Comments collection for this document.
//
// Mirrors Python Document.comments → self._part.comments.
func (d *Document) Comments() (*Comments, error) {
	cp, err := d.part.CommentsPart()
	if err != nil {
		return nil, fmt.Errorf("docx: getting comments part: %w", err)
	}
	elm, err := cp.CommentsElement()
	if err != nil {
		return nil, fmt.Errorf("docx: getting comments element: %w", err)
	}
	return newComments(elm, cp), nil
}

// CoreProperties returns the CoreProperties for this document.
//
// Mirrors Python Document.core_properties → self._part.core_properties.
func (d *Document) CoreProperties() (*CoreProperties, error) {
	cpp, err := d.part.CoreProperties()
	if err != nil {
		return nil, fmt.Errorf("docx: getting core properties: %w", err)
	}
	elm, err := cpp.CT()
	if err != nil {
		return nil, fmt.Errorf("docx: getting core properties element: %w", err)
	}
	return newCoreProperties(elm), nil
}

// InlineShapes returns the InlineShapes collection for this document.
//
// Mirrors Python Document.inline_shapes → self._part.inline_shapes.
func (d *Document) InlineShapes() (*InlineShapes, error) {
	body := d.element.Body()
	if body == nil || body.RawElement() == nil {
		return nil, fmt.Errorf("docx: document has no body element")
	}
	return newInlineShapes(body.RawElement()), nil
}

// IterInnerContent returns all paragraphs and tables in document order.
//
// Mirrors Python Document.iter_inner_content → self._body.iter_inner_content().
func (d *Document) IterInnerContent() ([]*InnerContentItem, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, fmt.Errorf("docx: getting body: %w", err)
	}
	return b.IterInnerContent(), nil
}

// Paragraphs returns all top-level paragraphs in document order.
//
// Mirrors Python Document.paragraphs → self._body.paragraphs.
func (d *Document) Paragraphs() ([]*Paragraph, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, fmt.Errorf("docx: getting body: %w", err)
	}
	return b.Paragraphs(), nil
}

// Part returns the DocumentPart for this document.
//
// Mirrors Python Document.part.
func (d *Document) Part() *parts.DocumentPart {
	return d.part
}

// Sections returns the Sections collection for this document.
//
// Mirrors Python Document.sections → Sections(self._element, self._part).
func (d *Document) Sections() *Sections {
	return newSections(d.element, d.part)
}

// Settings returns the Settings proxy for this document.
//
// Mirrors Python Document.settings → self._part.settings.
func (d *Document) Settings() (*Settings, error) {
	elm, err := d.part.Settings()
	if err != nil {
		return nil, fmt.Errorf("docx: getting settings: %w", err)
	}
	return newSettings(elm), nil
}

// Styles returns the Styles proxy for this document.
//
// Mirrors Python Document.styles → self._part.styles.
func (d *Document) Styles() (*Styles, error) {
	elm, err := d.part.Styles()
	if err != nil {
		return nil, fmt.Errorf("docx: getting styles: %w", err)
	}
	return newStyles(elm), nil
}

// Tables returns all top-level tables in document order.
//
// Mirrors Python Document.tables → self._body.tables.
func (d *Document) Tables() ([]*Table, error) {
	b, err := d.getBody()
	if err != nil {
		return nil, fmt.Errorf("docx: getting body: %w", err)
	}
	return b.Tables(), nil
}

// --------------------------------------------------------------------------
// Save
// --------------------------------------------------------------------------

// Save writes this document to w.
//
// Mirrors Python Document.save(stream).
func (d *Document) Save(w io.Writer) error {
	return d.wmlPkg.Save(w)
}

// SaveFile writes this document to a file.
//
// Mirrors Python Document.save(path).
func (d *Document) SaveFile(path string) error {
	return d.wmlPkg.SaveToFile(path)
}

// --------------------------------------------------------------------------
// Internal
// --------------------------------------------------------------------------

// blockWidth returns the available width between margins of the last section,
// in twips. Used for table column width calculation.
//
// Mirrors Python Document._block_width (but in twips, not EMU, since Go
// Section methods return twips).
func (d *Document) blockWidth() (int, error) {
	sections := d.Sections()
	if sections.Len() == 0 {
		return Inches(6.5).Twips(), nil
	}
	last, err := sections.Get(sections.Len() - 1)
	if err != nil {
		return 0, fmt.Errorf("docx: getting last section: %w", err)
	}

	pageWidth := Inches(8.5).Twips()
	if pw, err := last.PageWidth(); err == nil && pw != nil {
		pageWidth = *pw
	}
	leftMargin := Inches(1).Twips()
	if lm, err := last.LeftMargin(); err == nil && lm != nil {
		leftMargin = *lm
	}
	rightMargin := Inches(1).Twips()
	if rm, err := last.RightMargin(); err == nil && rm != nil {
		rightMargin = *rm
	}

	return pageWidth - leftMargin - rightMargin, nil
}

// getBody returns the cached Body, creating it on first call.
//
// Mirrors Python Document._body (lazy property).
func (d *Document) getBody() (*Body, error) {
	if d.body != nil {
		return d.body, nil
	}
	bodyElm := d.element.Body()
	if bodyElm == nil {
		return nil, fmt.Errorf("docx: document has no body element")
	}
	d.body = newBody(bodyElm, &d.part.StoryPart)
	return d.body, nil
}
