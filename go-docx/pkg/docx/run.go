package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
	"github.com/vortex/go-docx/pkg/docx/parts"
)

// Run is a proxy object wrapping a <w:r> element.
//
// Mirrors Python Run(StoryChild).
type Run struct {
	r    *oxml.CT_R
	part *parts.StoryPart
}

// NewRun creates a new Run proxy.
func NewRun(r *oxml.CT_R, part *parts.StoryPart) *Run {
	return &Run{r: r, part: part}
}

// AddBreak adds a break element of the given type to this run.
//
// Mirrors Python Run.add_break. Maps break_type to (type_, clear) pairs.
func (run *Run) AddBreak(breakType enum.WdBreakType) error {
	type_, clear := breakTypeToAttrs(breakType)
	br := run.r.AddBr()
	if type_ != "" {
		if err := br.SetType(type_); err != nil {
			return err
		}
	}
	if clear != "" {
		if err := br.SetClear(clear); err != nil {
			return err
		}
	}
	return nil
}

// breakTypeToAttrs maps WdBreakType to (type, clear) attribute values.
func breakTypeToAttrs(bt enum.WdBreakType) (string, string) {
	switch bt {
	case enum.WdBreakTypeLine:
		return "", ""
	case enum.WdBreakTypePage:
		return "page", ""
	case enum.WdBreakTypeColumn:
		return "column", ""
	case enum.WdBreakTypeLineClearLeft:
		return "textWrapping", "left"
	case enum.WdBreakTypeLineClearRight:
		return "textWrapping", "right"
	case enum.WdBreakTypeLineClearAll:
		return "textWrapping", "all"
	default:
		return "", ""
	}
}

// AddPicture adds an inline picture to this run and returns the InlineShape.
//
// Mirrors Python Run.add_picture.
func (run *Run) AddPicture(imgPart *parts.ImagePart, width, height *int64) (*InlineShape, error) {
	inline, err := run.part.NewPicInline(imgPart, width, height)
	if err != nil {
		return nil, fmt.Errorf("docx: creating pic inline: %w", err)
	}
	run.r.AddDrawingWithInline(inline)
	return NewInlineShape(inline), nil
}

// AddTab adds a <w:tab/> element at the end of the run.
//
// Mirrors Python Run.add_tab.
func (run *Run) AddTab() {
	run.r.AddTab()
}

// AddText appends a <w:t> element with the given text to the run.
//
// Mirrors Python Run.add_text.
func (run *Run) AddText(text string) {
	run.r.AddTWithText(text)
}

// Bold returns the tri-state bold value (delegates to Font).
//
// Mirrors Python Run.bold (getter).
func (run *Run) Bold() *bool {
	return run.Font().Bold()
}

// SetBold sets the tri-state bold value (delegates to Font).
//
// Mirrors Python Run.bold (setter).
func (run *Run) SetBold(v *bool) error {
	return run.Font().SetBold(v)
}

// Clear removes all content from this run, preserving formatting.
//
// Mirrors Python Run.clear.
func (run *Run) Clear() {
	run.r.ClearContent()
}

// ContainsPageBreak returns true when rendered page-breaks occur in this run.
//
// Mirrors Python Run.contains_page_break.
func (run *Run) ContainsPageBreak() bool {
	return len(run.r.LastRenderedPageBreaks()) > 0
}

// Font returns the Font providing access to character formatting properties.
//
// Mirrors Python Run.font.
func (run *Run) Font() *Font {
	return NewFont(run.r)
}

// Italic returns the tri-state italic value (delegates to Font).
//
// Mirrors Python Run.italic (getter).
func (run *Run) Italic() *bool {
	return run.Font().Italic()
}

// SetItalic sets the tri-state italic value (delegates to Font).
//
// Mirrors Python Run.italic (setter).
func (run *Run) SetItalic(v *bool) error {
	return run.Font().SetItalic(v)
}

// MarkCommentRange marks the range of runs from this run to lastRun as
// belonging to the comment identified by commentID.
//
// Mirrors Python Run.mark_comment_range.
func (run *Run) MarkCommentRange(lastRun *Run, commentID int) error {
	run.r.InsertCommentRangeStartAbove(commentID)
	lastRun.r.InsertCommentRangeEndAndReferenceBelow(commentID)
	return nil
}

// Style returns the character style applied to this run.
//
// Mirrors Python Run.style (getter).
func (run *Run) Style() (*oxml.CT_Style, error) {
	styleID, err := run.r.Style()
	if err != nil {
		return nil, err
	}
	return run.part.GetStyle(styleID, enum.WdStyleTypeCharacter)
}

// SetStyle sets the character style. style can be a string name or nil.
//
// Mirrors Python Run.style (setter).
func (run *Run) SetStyle(style interface{}) error {
	styleID, err := run.part.GetStyleID(style, enum.WdStyleTypeCharacter)
	if err != nil {
		return err
	}
	return run.r.SetStyle(styleID)
}

// Text returns the textual content of this run.
//
// Mirrors Python Run.text (getter).
func (run *Run) Text() string {
	return run.r.RunText()
}

// SetText replaces all run content with elements representing the given text.
//
// Mirrors Python Run.text (setter).
func (run *Run) SetText(text string) {
	run.r.SetRunText(text)
}

// Underline returns the underline value (delegates to Font).
//
// Mirrors Python Run.underline (getter).
func (run *Run) Underline() interface{} {
	return run.Font().Underline()
}

// SetUnderline sets the underline value (delegates to Font).
//
// Mirrors Python Run.underline (setter).
func (run *Run) SetUnderline(v interface{}) error {
	return run.Font().SetUnderline(v)
}

// CT_R returns the underlying oxml element.
func (run *Run) CT_R() *oxml.CT_R { return run.r }
