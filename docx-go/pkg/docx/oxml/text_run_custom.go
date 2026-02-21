package oxml

import (
	"strings"

	"github.com/beevik/etree"
)

// --- CT_R custom methods ---

// AddTWithText adds a <w:t> element containing the given text.
// Sets xml:space="preserve" if the text has leading or trailing whitespace.
func (r *CT_R) AddTWithText(text string) *CT_Text {
	t := r.addT()
	t.SetText(text)
	if len(strings.TrimSpace(text)) < len(text) {
		t.E.CreateAttr("xml:space", "preserve")
	}
	return t
}

// AddDrawingWithInline adds a <w:drawing> element containing the given inline element.
func (r *CT_R) AddDrawingWithInline(inline *CT_Inline) *CT_Drawing {
	drawing := r.addDrawing()
	drawing.E.AddChild(inline.E)
	return drawing
}

// ClearContent removes all child elements except <w:rPr>.
func (r *CT_R) ClearContent() {
	var toRemove []*etree.Element
	for _, child := range r.E.ChildElements() {
		if !(child.Space == "w" && child.Tag == "rPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		r.E.RemoveChild(child)
	}
}

// Style returns the styleId of the run, or nil if not set.
func (r *CT_R) Style() *string {
	rPr := r.RPr()
	if rPr == nil {
		return nil
	}
	return rPr.StyleVal()
}

// SetStyle sets the run style. Passing nil removes the rStyle element.
func (r *CT_R) SetStyle(styleID *string) {
	rPr := r.GetOrAddRPr()
	rPr.SetStyleVal(styleID)
}

// RunText returns the textual content of this run by concatenating text equivalents
// of all inner-content elements (w:t, w:br, w:cr, w:tab, w:noBreakHyphen, w:ptab).
func (r *CT_R) RunText() string {
	var sb strings.Builder
	for _, child := range r.E.ChildElements() {
		if child.Space != "w" {
			continue
		}
		switch child.Tag {
		case "t":
			sb.WriteString(child.Text())
		case "br":
			br := &CT_Br{Element{E: child}}
			sb.WriteString(br.TextEquivalent())
		case "cr":
			sb.WriteString("\n")
		case "tab":
			sb.WriteString("\t")
		case "noBreakHyphen":
			sb.WriteString("-")
		case "ptab":
			sb.WriteString("\t")
		}
	}
	return sb.String()
}

// SetRunText replaces all run content with elements representing the given text.
// Tab characters become <w:tab/>, newlines/carriage-returns become <w:br/>,
// and regular characters are grouped into <w:t> elements.
func (r *CT_R) SetRunText(text string) {
	r.ClearContent()
	appendRunContentFromText(r, text)
}

// LastRenderedPageBreaks returns all <w:lastRenderedPageBreak> descendants of this run.
func (r *CT_R) LastRenderedPageBreaks() []*CT_LastRenderedPageBreak {
	var result []*CT_LastRenderedPageBreak
	for _, child := range r.E.ChildElements() {
		if child.Space == "w" && child.Tag == "lastRenderedPageBreak" {
			result = append(result, &CT_LastRenderedPageBreak{Element{E: child}})
		}
	}
	return result
}

// --- CT_Br custom methods ---

// TextEquivalent returns the text equivalent of this break element.
// Line breaks produce "\n"; column and page breaks produce "".
func (br *CT_Br) TextEquivalent() string {
	if br.Type() == "textWrapping" {
		return "\n"
	}
	return ""
}

// --- CT_Cr custom methods ---

// TextEquivalent returns the text equivalent of a carriage return element: "\n".
func (cr *CT_Cr) TextEquivalent() string {
	return "\n"
}

// --- CT_NoBreakHyphen custom methods ---

// TextEquivalent returns the text equivalent of a non-breaking hyphen: "-".
func (nbh *CT_NoBreakHyphen) TextEquivalent() string {
	return "-"
}

// --- CT_PTab custom methods ---

// TextEquivalent returns the text equivalent of an absolute-position tab: "\t".
func (pt *CT_PTab) TextEquivalent() string {
	return "\t"
}

// --- CT_Text custom methods ---

// ContentText returns the text content of this <w:t> element, or empty string if none.
func (t *CT_Text) ContentText() string {
	return t.E.Text()
}

// SetPreserveSpace sets xml:space="preserve" on this <w:t> element.
func (t *CT_Text) SetPreserveSpace() {
	t.E.CreateAttr("xml:space", "preserve")
}

// --- Run content appender utility ---

// appendRunContentFromText translates a string into run content elements.
// Tabs → <w:tab/>, newlines → <w:br/>, regular chars → <w:t>.
func appendRunContentFromText(r *CT_R, text string) {
	var buf strings.Builder
	flush := func() {
		if buf.Len() > 0 {
			r.AddTWithText(buf.String())
			buf.Reset()
		}
	}
	for _, ch := range text {
		switch ch {
		case '\t':
			flush()
			r.AddTab()
		case '\n', '\r':
			flush()
			r.AddBr()
		default:
			buf.WriteRune(ch)
		}
	}
	flush()
}
