package oxml

import (
	"testing"

	"github.com/user/go-docx/pkg/docx/enum"
)

// --- CT_P tests ---

func TestCT_P_ParagraphText(t *testing.T) {
	// Build <w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	r1 := p.AddR()
	r1.AddTWithText("Hello ")

	r2 := p.AddR()
	r2.AddTWithText("World")

	got := p.ParagraphText()
	if got != "Hello World" {
		t.Errorf("CT_P.ParagraphText() = %q, want %q", got, "Hello World")
	}
}

func TestCT_P_ParagraphTextWithHyperlink(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	r1 := p.AddR()
	r1.AddTWithText("Click ")

	h := p.AddHyperlink()
	hr := h.AddR()
	hr.AddTWithText("here")

	r2 := p.AddR()
	r2.AddTWithText(" now")

	got := p.ParagraphText()
	if got != "Click here now" {
		t.Errorf("CT_P.ParagraphText() = %q, want %q", got, "Click here now")
	}
}

func TestCT_P_Alignment_RoundTrip(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	// Initially nil
	if p.Alignment() != nil {
		t.Error("expected nil alignment for new paragraph")
	}

	// Set center
	center := enum.WdParagraphAlignmentCenter
	p.SetAlignment(&center)
	got := p.Alignment()
	if got == nil || *got != enum.WdParagraphAlignmentCenter {
		t.Errorf("expected center alignment, got %v", got)
	}

	// Set nil removes
	p.SetAlignment(nil)
	if p.Alignment() != nil {
		t.Error("expected nil alignment after setting nil")
	}
}

func TestCT_P_Style_RoundTrip(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	if p.Style() != nil {
		t.Error("expected nil style for new paragraph")
	}

	s := "Heading1"
	p.SetStyle(&s)
	got := p.Style()
	if got == nil || *got != "Heading1" {
		t.Errorf("expected Heading1 style, got %v", got)
	}

	p.SetStyle(nil)
	if p.Style() != nil {
		t.Error("expected nil style after removing")
	}
}

func TestCT_P_ClearContent(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	p.GetOrAddPPr() // adds pPr
	p.AddR()
	p.AddR()
	p.AddHyperlink()

	p.ClearContent()

	// pPr should remain
	if p.PPr() == nil {
		t.Error("pPr should be preserved after ClearContent")
	}
	// r and hyperlink should be gone
	if len(p.RList()) != 0 {
		t.Error("runs should be removed after ClearContent")
	}
	if len(p.HyperlinkList()) != 0 {
		t.Error("hyperlinks should be removed after ClearContent")
	}
}

func TestCT_P_AddPBefore(t *testing.T) {
	// Create a parent body with one paragraph
	body := OxmlElement("w:body")
	pEl := OxmlElement("w:p")
	body.AddChild(pEl)
	p := &CT_P{Element{E: pEl}}

	newP := p.AddPBefore()
	if newP == nil {
		t.Fatal("AddPBefore returned nil")
	}

	// The new paragraph should be before the original
	children := body.ChildElements()
	if len(children) != 2 {
		t.Fatalf("expected 2 children, got %d", len(children))
	}
	if children[0] != newP.E {
		t.Error("new paragraph should be first child")
	}
	if children[1] != p.E {
		t.Error("original paragraph should be second child")
	}
}

func TestCT_P_InnerContentElements(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	p.GetOrAddPPr()
	p.AddR()
	p.AddHyperlink()
	p.AddR()

	elems := p.InnerContentElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 inner content elements, got %d", len(elems))
	}
	// First should be CT_R, second CT_Hyperlink, third CT_R
	if _, ok := elems[0].(*CT_R); !ok {
		t.Error("first element should be *CT_R")
	}
	if _, ok := elems[1].(*CT_Hyperlink); !ok {
		t.Error("second element should be *CT_Hyperlink")
	}
	if _, ok := elems[2].(*CT_R); !ok {
		t.Error("third element should be *CT_R")
	}
}

// --- CT_R tests ---

func TestCT_R_AddTWithText_PreservesSpace(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{E: rEl}}

	t1 := r.AddTWithText(" hello ")
	// Check xml:space="preserve" is set
	val := t1.E.SelectAttrValue("xml:space", "")
	if val != "preserve" {
		t.Errorf("expected xml:space=preserve for text with spaces, got %q", val)
	}

	r2El := OxmlElement("w:r")
	r2 := &CT_R{Element{E: r2El}}
	t2 := r2.AddTWithText("hello")
	val2 := t2.E.SelectAttrValue("xml:space", "")
	if val2 != "" {
		t.Errorf("expected no xml:space for trimmed text, got %q", val2)
	}
}

func TestCT_R_RunText(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{E: rEl}}

	r.AddTWithText("Hello")
	r.AddTab()
	r.AddTWithText("World")

	got := r.RunText()
	if got != "Hello\tWorld" {
		t.Errorf("RunText() = %q, want %q", got, "Hello\tWorld")
	}
}

func TestCT_R_RunTextWithBr(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{E: rEl}}

	r.AddTWithText("Line1")
	r.AddBr() // default type = textWrapping → "\n"
	r.AddTWithText("Line2")

	got := r.RunText()
	if got != "Line1\nLine2" {
		t.Errorf("RunText() = %q, want %q", got, "Line1\nLine2")
	}
}

func TestCT_R_ClearContent(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{E: rEl}}

	r.GetOrAddRPr()
	r.AddTWithText("text")
	r.AddBr()

	r.ClearContent()

	if r.RPr() == nil {
		t.Error("rPr should be preserved after ClearContent")
	}
	if len(r.TList()) != 0 || len(r.BrList()) != 0 {
		t.Error("content should be removed after ClearContent")
	}
}

func TestCT_R_Style_RoundTrip(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{E: rEl}}

	if r.Style() != nil {
		t.Error("expected nil style for new run")
	}

	s := "Emphasis"
	r.SetStyle(&s)
	got := r.Style()
	if got == nil || *got != "Emphasis" {
		t.Errorf("expected Emphasis style, got %v", got)
	}

	r.SetStyle(nil)
	if r.Style() != nil {
		t.Error("expected nil style after removing")
	}
}

func TestCT_R_SetRunText(t *testing.T) {
	rEl := OxmlElement("w:r")
	r := &CT_R{Element{E: rEl}}

	r.GetOrAddRPr() // should be preserved
	r.SetRunText("Hello\tWorld\nNew")

	// Check rPr still exists
	if r.RPr() == nil {
		t.Error("rPr should be preserved after SetRunText")
	}

	got := r.RunText()
	if got != "Hello\tWorld\nNew" {
		t.Errorf("after SetRunText, RunText() = %q, want %q", got, "Hello\tWorld\nNew")
	}
}

// --- CT_Br tests ---

func TestCT_Br_TextEquivalent(t *testing.T) {
	// Default (textWrapping)
	br1 := &CT_Br{Element{E: OxmlElement("w:br")}}
	if br1.TextEquivalent() != "\n" {
		t.Error("expected newline for default break type")
	}

	// Page break
	br2 := &CT_Br{Element{E: OxmlElement("w:br")}}
	br2.SetType("page")
	if br2.TextEquivalent() != "" {
		t.Error("expected empty string for page break")
	}
}

// --- CT_RPr tests ---

func TestCT_RPr_BoldVal_TriState(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	// Initially nil (not set)
	if rPr.BoldVal() != nil {
		t.Error("expected nil bold for new rPr")
	}

	// Set true → <w:b/> (no val attr)
	bTrue := true
	rPr.SetBoldVal(&bTrue)
	got := rPr.BoldVal()
	if got == nil || !*got {
		t.Error("expected *true after SetBoldVal(true)")
	}

	// Set false → <w:b w:val="false"/>
	bFalse := false
	rPr.SetBoldVal(&bFalse)
	got = rPr.BoldVal()
	if got == nil || *got {
		t.Error("expected *false after SetBoldVal(false)")
	}

	// Set nil → remove element
	rPr.SetBoldVal(nil)
	if rPr.BoldVal() != nil {
		t.Error("expected nil after SetBoldVal(nil)")
	}
}

func TestCT_RPr_ItalicVal_TriState(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	v := true
	rPr.SetItalicVal(&v)
	got := rPr.ItalicVal()
	if got == nil || !*got {
		t.Error("expected *true for italic")
	}
}

func TestCT_RPr_ColorVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if rPr.ColorVal() != nil {
		t.Error("expected nil color for new rPr")
	}

	c := "FF0000"
	rPr.SetColorVal(&c)
	got := rPr.ColorVal()
	if got == nil || *got != "FF0000" {
		t.Errorf("expected FF0000, got %v", got)
	}

	rPr.SetColorVal(nil)
	if rPr.ColorVal() != nil {
		t.Error("expected nil after removing color")
	}
}

func TestCT_RPr_SzVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if rPr.SzVal() != nil {
		t.Error("expected nil sz for new rPr")
	}

	var sz int64 = 24 // 12pt in half-points
	rPr.SetSzVal(&sz)
	got := rPr.SzVal()
	if got == nil || *got != 24 {
		t.Errorf("expected 24, got %v", got)
	}

	rPr.SetSzVal(nil)
	if rPr.SzVal() != nil {
		t.Error("expected nil after removing sz")
	}
}

func TestCT_RPr_RFontsAscii(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if rPr.RFontsAscii() != nil {
		t.Error("expected nil font for new rPr")
	}

	f := "Arial"
	rPr.SetRFontsAscii(&f)
	got := rPr.RFontsAscii()
	if got == nil || *got != "Arial" {
		t.Errorf("expected Arial, got %v", got)
	}

	rPr.SetRFontsAscii(nil)
	if rPr.RFontsAscii() != nil {
		t.Error("expected nil after removing font")
	}
}

func TestCT_RPr_StyleVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	s := "CommentReference"
	rPr.SetStyleVal(&s)
	got := rPr.StyleVal()
	if got == nil || *got != "CommentReference" {
		t.Errorf("expected CommentReference, got %v", got)
	}

	rPr.SetStyleVal(nil)
	if rPr.StyleVal() != nil {
		t.Error("expected nil after removing style")
	}
}

func TestCT_RPr_UVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if rPr.UVal() != nil {
		t.Error("expected nil underline for new rPr")
	}

	u := "single"
	rPr.SetUVal(&u)
	got := rPr.UVal()
	if got == nil || *got != "single" {
		t.Errorf("expected single, got %v", got)
	}

	rPr.SetUVal(nil)
	if rPr.UVal() != nil {
		t.Error("expected nil after removing underline")
	}
}

func TestCT_RPr_Subscript(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if rPr.Subscript() != nil {
		t.Error("expected nil subscript for new rPr")
	}

	bTrue := true
	rPr.SetSubscript(&bTrue)
	got := rPr.Subscript()
	if got == nil || !*got {
		t.Error("expected *true for subscript")
	}

	bFalse := false
	rPr.SetSubscript(&bFalse)
	// Should remove since current is subscript and setting to false
	if rPr.Subscript() != nil {
		t.Error("expected nil after setting subscript to false (was subscript)")
	}
}

// --- CT_PPr tests ---

func TestCT_PPr_SpacingBefore_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	if pPr.SpacingBefore() != nil {
		t.Error("expected nil spacing before for new pPr")
	}

	v := 240 // 240 twips
	pPr.SetSpacingBefore(&v)
	got := pPr.SpacingBefore()
	if got == nil || *got != 240 {
		t.Errorf("expected 240, got %v", got)
	}
}

func TestCT_PPr_SpacingAfter_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	v := 120
	pPr.SetSpacingAfter(&v)
	got := pPr.SpacingAfter()
	if got == nil || *got != 120 {
		t.Errorf("expected 120, got %v", got)
	}
}

func TestCT_PPr_SpacingLineRule(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	// Set line without lineRule → default to MULTIPLE
	line := 480
	pPr.SetSpacingLine(&line)
	got := pPr.SpacingLineRule()
	if got == nil || *got != enum.WdLineSpacingMultiple {
		t.Errorf("expected MULTIPLE default, got %v", got)
	}
}

func TestCT_PPr_IndLeft_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	if pPr.IndLeft() != nil {
		t.Error("expected nil indent for new pPr")
	}

	v := 720 // 720 twips = 0.5 inch
	pPr.SetIndLeft(&v)
	got := pPr.IndLeft()
	if got == nil || *got != 720 {
		t.Errorf("expected 720, got %v", got)
	}
}

func TestCT_PPr_FirstLineIndent(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	// Positive first-line indent
	v := 360
	pPr.SetFirstLineIndent(&v)
	got := pPr.FirstLineIndent()
	if got == nil || *got != 360 {
		t.Errorf("expected 360, got %v", got)
	}

	// Negative (hanging) indent
	neg := -720
	pPr.SetFirstLineIndent(&neg)
	got = pPr.FirstLineIndent()
	if got == nil || *got != -720 {
		t.Errorf("expected -720 (hanging), got %v", got)
	}

	// Nil clears both
	pPr.SetFirstLineIndent(nil)
	got = pPr.FirstLineIndent()
	if got != nil {
		t.Errorf("expected nil after clearing, got %v", got)
	}
}

func TestCT_PPr_KeepLines_TriState(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	if pPr.KeepLinesVal() != nil {
		t.Error("expected nil keepLines for new pPr")
	}

	v := true
	pPr.SetKeepLinesVal(&v)
	got := pPr.KeepLinesVal()
	if got == nil || !*got {
		t.Error("expected *true for keepLines")
	}

	pPr.SetKeepLinesVal(nil)
	if pPr.KeepLinesVal() != nil {
		t.Error("expected nil after removing keepLines")
	}
}

func TestCT_PPr_PageBreakBefore(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	v := true
	pPr.SetPageBreakBeforeVal(&v)
	got := pPr.PageBreakBeforeVal()
	if got == nil || !*got {
		t.Error("expected *true for pageBreakBefore")
	}
}

// --- CT_Hyperlink tests ---

func TestCT_Hyperlink_Text(t *testing.T) {
	hEl := OxmlElement("w:hyperlink")
	h := &CT_Hyperlink{Element{E: hEl}}

	r := h.AddR()
	r.AddTWithText("Click here")

	got := h.HyperlinkText()
	if got != "Click here" {
		t.Errorf("HyperlinkText() = %q, want %q", got, "Click here")
	}
}

// --- CT_LastRenderedPageBreak tests ---

func TestCT_LastRenderedPageBreak_PrecedesAllContent(t *testing.T) {
	// Build: <w:p><w:r><w:lastRenderedPageBreak/><w:t>text</w:t></w:r></w:p>
	pEl := OxmlElement("w:p")
	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)
	tEl := OxmlElement("w:t")
	tEl.SetText("text")
	rEl.AddChild(tEl)

	lrpb := &CT_LastRenderedPageBreak{Element{E: lrpbEl}}

	if !lrpb.PrecedesAllContent() {
		t.Error("expected PrecedesAllContent to be true when lrpb is first in first run")
	}
}

func TestCT_LastRenderedPageBreak_FollowsAllContent(t *testing.T) {
	// Build: <w:p><w:r><w:t>text</w:t><w:lastRenderedPageBreak/></w:r></w:p>
	pEl := OxmlElement("w:p")
	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)
	tEl := OxmlElement("w:t")
	tEl.SetText("text")
	rEl.AddChild(tEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	lrpb := &CT_LastRenderedPageBreak{Element{E: lrpbEl}}

	if !lrpb.FollowsAllContent() {
		t.Error("expected FollowsAllContent to be true when lrpb is last in last run")
	}
}

func TestCT_LastRenderedPageBreak_IsInHyperlink(t *testing.T) {
	// Build: <w:p><w:hyperlink><w:r><w:lastRenderedPageBreak/></w:r></w:hyperlink></w:p>
	pEl := OxmlElement("w:p")
	hEl := OxmlElement("w:hyperlink")
	pEl.AddChild(hEl)
	rEl := OxmlElement("w:r")
	hEl.AddChild(rEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	lrpb := &CT_LastRenderedPageBreak{Element{E: lrpbEl}}

	if !lrpb.IsInHyperlink() {
		t.Error("expected IsInHyperlink to be true")
	}
}

func TestCT_LastRenderedPageBreak_EnclosingP(t *testing.T) {
	pEl := OxmlElement("w:p")
	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	lrpb := &CT_LastRenderedPageBreak{Element{E: lrpbEl}}
	p := lrpb.EnclosingP()
	if p == nil || p.E != pEl {
		t.Error("EnclosingP should return the parent w:p")
	}
}

// --- CT_TabStops tests ---

func TestCT_TabStops_InsertTabInOrder(t *testing.T) {
	tabsEl := OxmlElement("w:tabs")
	tabs := &CT_TabStops{Element{E: tabsEl}}

	tabs.InsertTabInOrder(2880, enum.WdTabAlignmentCenter, enum.WdTabLeaderDots)
	tabs.InsertTabInOrder(720, enum.WdTabAlignmentLeft, enum.WdTabLeaderSpaces)
	tabs.InsertTabInOrder(5760, enum.WdTabAlignmentRight, enum.WdTabLeaderDashes)

	list := tabs.TabList()
	if len(list) != 3 {
		t.Fatalf("expected 3 tabs, got %d", len(list))
	}

	// Verify order
	pos0, _ := list[0].Pos()
	pos1, _ := list[1].Pos()
	pos2, _ := list[2].Pos()
	if pos0 != 720 || pos1 != 2880 || pos2 != 5760 {
		t.Errorf("tabs not in order: %d, %d, %d", pos0, pos1, pos2)
	}
}
