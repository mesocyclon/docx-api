package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
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
	if a, err := p.Alignment(); err != nil {
		t.Fatalf("Alignment: %v", err)
	} else if a != nil {
		t.Error("expected nil alignment for new paragraph")
	}

	// Set center
	center := enum.WdParagraphAlignmentCenter
	if err := p.SetAlignment(&center); err != nil {
		t.Fatalf("SetAlignment: %v", err)
	}
	got, err := p.Alignment()
	if err != nil {
		t.Fatalf("Alignment: %v", err)
	}
	if got == nil || *got != enum.WdParagraphAlignmentCenter {
		t.Errorf("expected center alignment, got %v", got)
	}

	// Set nil removes
	if err := p.SetAlignment(nil); err != nil {
		t.Fatalf("SetAlignment(nil): %v", err)
	}
	if a, err := p.Alignment(); err != nil {
		t.Fatalf("Alignment: %v", err)
	} else if a != nil {
		t.Error("expected nil alignment after setting nil")
	}
}

func TestCT_P_Style_RoundTrip(t *testing.T) {
	pEl := OxmlElement("w:p")
	p := &CT_P{Element{E: pEl}}

	if s, err := p.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style for new paragraph")
	}

	s := "Heading1"
	if err := p.SetStyle(&s); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	got, err := p.Style()
	if err != nil {
		t.Fatalf("Style: %v", err)
	}
	if got == nil || *got != "Heading1" {
		t.Errorf("expected Heading1 style, got %v", got)
	}

	if err := p.SetStyle(nil); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	if s, err := p.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
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

	if s, err := r.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
		t.Error("expected nil style for new run")
	}

	s := "Emphasis"
	if err := r.SetStyle(&s); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	got, err := r.Style()
	if err != nil {
		t.Fatalf("Style: %v", err)
	}
	if got == nil || *got != "Emphasis" {
		t.Errorf("expected Emphasis style, got %v", got)
	}

	if err := r.SetStyle(nil); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	if s, err := r.Style(); err != nil {
		t.Fatalf("Style: %v", err)
	} else if s != nil {
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
	if err := br2.SetType("page"); err != nil {
		t.Fatalf("SetType: %v", err)
	}
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
	if err := rPr.SetBoldVal(&bTrue); err != nil {
		t.Fatalf("SetBoldVal: %v", err)
	}
	got := rPr.BoldVal()
	if got == nil || !*got {
		t.Error("expected *true after SetBoldVal(true)")
	}

	// Set false → <w:b w:val="false"/>
	bFalse := false
	if err := rPr.SetBoldVal(&bFalse); err != nil {
		t.Fatalf("SetBoldVal: %v", err)
	}
	got = rPr.BoldVal()
	if got == nil || *got {
		t.Error("expected *false after SetBoldVal(false)")
	}

	// Set nil → remove element
	if err := rPr.SetBoldVal(nil); err != nil {
		t.Fatalf("SetBoldVal: %v", err)
	}
	if rPr.BoldVal() != nil {
		t.Error("expected nil after SetBoldVal(nil)")
	}
}

func TestCT_RPr_ItalicVal_TriState(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	v := true
	if err := rPr.SetItalicVal(&v); err != nil {
		t.Fatalf("SetItalicVal: %v", err)
	}
	got := rPr.ItalicVal()
	if got == nil || !*got {
		t.Error("expected *true for italic")
	}
}

func TestCT_RPr_ColorVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if cv, err := rPr.ColorVal(); err != nil {
		t.Fatalf("ColorVal: %v", err)
	} else if cv != nil {
		t.Error("expected nil color for new rPr")
	}

	c := "FF0000"
	if err := rPr.SetColorVal(&c); err != nil {
		t.Fatalf("SetColorVal: %v", err)
	}
	got, err := rPr.ColorVal()
	if err != nil {
		t.Fatalf("ColorVal: %v", err)
	}
	if got == nil || *got != "FF0000" {
		t.Errorf("expected FF0000, got %v", got)
	}

	if err := rPr.SetColorVal(nil); err != nil {
		t.Fatalf("SetColorVal: %v", err)
	}
	if cv, err := rPr.ColorVal(); err != nil {
		t.Fatalf("ColorVal: %v", err)
	} else if cv != nil {
		t.Error("expected nil after removing color")
	}
}

func TestCT_RPr_SzVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if sv, err := rPr.SzVal(); err != nil {
		t.Fatalf("SzVal: %v", err)
	} else if sv != nil {
		t.Error("expected nil sz for new rPr")
	}

	var sz int64 = 24 // 12pt in half-points
	if err := rPr.SetSzVal(&sz); err != nil {
		t.Fatalf("SetSzVal: %v", err)
	}
	got, err := rPr.SzVal()
	if err != nil {
		t.Fatalf("SzVal: %v", err)
	}
	if got == nil || *got != 24 {
		t.Errorf("expected 24, got %v", got)
	}

	if err := rPr.SetSzVal(nil); err != nil {
		t.Fatalf("SetSzVal: %v", err)
	}
	if sv, err := rPr.SzVal(); err != nil {
		t.Fatalf("SzVal: %v", err)
	} else if sv != nil {
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
	if err := rPr.SetRFontsAscii(&f); err != nil {
		t.Fatalf("SetRFontsAscii: %v", err)
	}
	got := rPr.RFontsAscii()
	if got == nil || *got != "Arial" {
		t.Errorf("expected Arial, got %v", got)
	}

	if err := rPr.SetRFontsAscii(nil); err != nil {
		t.Fatalf("SetRFontsAscii: %v", err)
	}
	if rPr.RFontsAscii() != nil {
		t.Error("expected nil after removing font")
	}
}

func TestCT_RPr_StyleVal(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	s := "CommentReference"
	if err := rPr.SetStyleVal(&s); err != nil {
		t.Fatalf("SetStyleVal: %v", err)
	}
	got, err := rPr.StyleVal()
	if err != nil {
		t.Fatalf("StyleVal: %v", err)
	}
	if got == nil || *got != "CommentReference" {
		t.Errorf("expected CommentReference, got %v", got)
	}

	if err := rPr.SetStyleVal(nil); err != nil {
		t.Fatalf("SetStyleVal: %v", err)
	}
	if sv, err := rPr.StyleVal(); err != nil {
		t.Fatalf("StyleVal: %v", err)
	} else if sv != nil {
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
	if err := rPr.SetUVal(&u); err != nil {
		t.Fatalf("SetUVal: %v", err)
	}
	got := rPr.UVal()
	if got == nil || *got != "single" {
		t.Errorf("expected single, got %v", got)
	}

	if err := rPr.SetUVal(nil); err != nil {
		t.Fatalf("SetUVal: %v", err)
	}
	if rPr.UVal() != nil {
		t.Error("expected nil after removing underline")
	}
}

func TestCT_RPr_Subscript(t *testing.T) {
	rPrEl := OxmlElement("w:rPr")
	rPr := &CT_RPr{Element{E: rPrEl}}

	if sub, err := rPr.Subscript(); err != nil {
		t.Fatalf("Subscript: %v", err)
	} else if sub != nil {
		t.Error("expected nil subscript for new rPr")
	}

	bTrue := true
	if err := rPr.SetSubscript(&bTrue); err != nil {
		t.Fatalf("SetSubscript: %v", err)
	}
	got, err := rPr.Subscript()
	if err != nil {
		t.Fatalf("Subscript: %v", err)
	}
	if got == nil || !*got {
		t.Error("expected *true for subscript")
	}

	bFalse := false
	if err := rPr.SetSubscript(&bFalse); err != nil {
		t.Fatalf("SetSubscript: %v", err)
	}
	// Should remove since current is subscript and setting to false
	if sub, err := rPr.Subscript(); err != nil {
		t.Fatalf("Subscript: %v", err)
	} else if sub != nil {
		t.Error("expected nil after setting subscript to false (was subscript)")
	}
}

// --- CT_PPr tests ---

func TestCT_PPr_SpacingBefore_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	if sb, err := pPr.SpacingBefore(); err != nil {
		t.Fatalf("SpacingBefore: %v", err)
	} else if sb != nil {
		t.Error("expected nil spacing before for new pPr")
	}

	v := 240 // 240 twips
	if err := pPr.SetSpacingBefore(&v); err != nil {
		t.Fatalf("SetSpacingBefore: %v", err)
	}
	got, err := pPr.SpacingBefore()
	if err != nil {
		t.Fatalf("SpacingBefore: %v", err)
	}
	if got == nil || *got != 240 {
		t.Errorf("expected 240, got %v", got)
	}
}

func TestCT_PPr_SpacingAfter_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	v := 120
	if err := pPr.SetSpacingAfter(&v); err != nil {
		t.Fatalf("SetSpacingAfter: %v", err)
	}
	got, err := pPr.SpacingAfter()
	if err != nil {
		t.Fatalf("SpacingAfter: %v", err)
	}
	if got == nil || *got != 120 {
		t.Errorf("expected 120, got %v", got)
	}
}

func TestCT_PPr_SpacingLineRule(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	// Set line without lineRule → default to MULTIPLE
	line := 480
	if err := pPr.SetSpacingLine(&line); err != nil {
		t.Fatalf("SetSpacingLine: %v", err)
	}
	got, err := pPr.SpacingLineRule()
	if err != nil {
		t.Fatalf("SpacingLineRule: %v", err)
	}
	if got == nil || *got != enum.WdLineSpacingMultiple {
		t.Errorf("expected MULTIPLE default, got %v", got)
	}
}

func TestCT_PPr_IndLeft_RoundTrip(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	if il, err := pPr.IndLeft(); err != nil {
		t.Fatalf("IndLeft: %v", err)
	} else if il != nil {
		t.Error("expected nil indent for new pPr")
	}

	v := 720 // 720 twips = 0.5 inch
	if err := pPr.SetIndLeft(&v); err != nil {
		t.Fatalf("SetIndLeft: %v", err)
	}
	got, err := pPr.IndLeft()
	if err != nil {
		t.Fatalf("IndLeft: %v", err)
	}
	if got == nil || *got != 720 {
		t.Errorf("expected 720, got %v", got)
	}
}

func TestCT_PPr_FirstLineIndent(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	// Positive first-line indent
	v := 360
	if err := pPr.SetFirstLineIndent(&v); err != nil {
		t.Fatalf("SetFirstLineIndent: %v", err)
	}
	got, err := pPr.FirstLineIndent()
	if err != nil {
		t.Fatalf("FirstLineIndent: %v", err)
	}
	if got == nil || *got != 360 {
		t.Errorf("expected 360, got %v", got)
	}

	// Negative (hanging) indent
	neg := -720
	if err := pPr.SetFirstLineIndent(&neg); err != nil {
		t.Fatalf("SetFirstLineIndent: %v", err)
	}
	got, err = pPr.FirstLineIndent()
	if err != nil {
		t.Fatalf("FirstLineIndent: %v", err)
	}
	if got == nil || *got != -720 {
		t.Errorf("expected -720 (hanging), got %v", got)
	}

	// Nil clears both
	if err := pPr.SetFirstLineIndent(nil); err != nil {
		t.Fatalf("SetFirstLineIndent: %v", err)
	}
	got, err = pPr.FirstLineIndent()
	if err != nil {
		t.Fatalf("FirstLineIndent: %v", err)
	}
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
	if err := pPr.SetKeepLinesVal(&v); err != nil {
		t.Fatalf("SetKeepLinesVal: %v", err)
	}
	got := pPr.KeepLinesVal()
	if got == nil || !*got {
		t.Error("expected *true for keepLines")
	}

	if err := pPr.SetKeepLinesVal(nil); err != nil {
		t.Fatalf("SetKeepLinesVal: %v", err)
	}
	if pPr.KeepLinesVal() != nil {
		t.Error("expected nil after removing keepLines")
	}
}

func TestCT_PPr_PageBreakBefore(t *testing.T) {
	pPrEl := OxmlElement("w:pPr")
	pPr := &CT_PPr{Element{E: pPrEl}}

	v := true
	if err := pPr.SetPageBreakBeforeVal(&v); err != nil {
		t.Fatalf("SetPageBreakBeforeVal: %v", err)
	}
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

	if _, err := tabs.InsertTabInOrder(2880, enum.WdTabAlignmentCenter, enum.WdTabLeaderDots); err != nil {
		t.Fatalf("InsertTabInOrder: %v", err)
	}
	if _, err := tabs.InsertTabInOrder(720, enum.WdTabAlignmentLeft, enum.WdTabLeaderSpaces); err != nil {
		t.Fatalf("InsertTabInOrder: %v", err)
	}
	if _, err := tabs.InsertTabInOrder(5760, enum.WdTabAlignmentRight, enum.WdTabLeaderDashes); err != nil {
		t.Fatalf("InsertTabInOrder: %v", err)
	}

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

// ===========================================================================
// CT_RPr tri-state boolean properties — table-driven test
// ===========================================================================

// triStateProp describes a getter/setter pair for testing.
type triStateProp struct {
	name   string
	get    func(rPr *CT_RPr) *bool
	set    func(rPr *CT_RPr, v *bool) error
}

func TestCT_RPr_TriStateBooleans(t *testing.T) {
	t.Parallel()

	props := []triStateProp{
		{"Caps", (*CT_RPr).CapsVal, (*CT_RPr).SetCapsVal},
		{"SmallCaps", (*CT_RPr).SmallCapsVal, (*CT_RPr).SetSmallCapsVal},
		{"Strike", (*CT_RPr).StrikeVal, (*CT_RPr).SetStrikeVal},
		{"Dstrike", (*CT_RPr).DstrikeVal, (*CT_RPr).SetDstrikeVal},
		{"Outline", (*CT_RPr).OutlineVal, (*CT_RPr).SetOutlineVal},
		{"Shadow", (*CT_RPr).ShadowVal, (*CT_RPr).SetShadowVal},
		{"Emboss", (*CT_RPr).EmbossVal, (*CT_RPr).SetEmbossVal},
		{"Imprint", (*CT_RPr).ImprintVal, (*CT_RPr).SetImprintVal},
		{"NoProof", (*CT_RPr).NoProofVal, (*CT_RPr).SetNoProofVal},
		{"SnapToGrid", (*CT_RPr).SnapToGridVal, (*CT_RPr).SetSnapToGridVal},
		{"Vanish", (*CT_RPr).VanishVal, (*CT_RPr).SetVanishVal},
		{"WebHidden", (*CT_RPr).WebHiddenVal, (*CT_RPr).SetWebHiddenVal},
		{"SpecVanish", (*CT_RPr).SpecVanishVal, (*CT_RPr).SetSpecVanishVal},
		{"OMath", (*CT_RPr).OMathVal, (*CT_RPr).SetOMathVal},
	}

	for _, p := range props {
		p := p
		t.Run(p.name, func(t *testing.T) {
			t.Parallel()
			rPr := &CT_RPr{Element{E: OxmlElement("w:rPr")}}

			// Initially nil
			if p.get(rPr) != nil {
				t.Errorf("%s: expected nil initially", p.name)
			}

			// Set true
			bTrue := true
			if err := p.set(rPr, &bTrue); err != nil {
				t.Fatalf("%s: set true: %v", p.name, err)
			}
			got := p.get(rPr)
			if got == nil || !*got {
				t.Errorf("%s: expected *true, got %v", p.name, got)
			}

			// Set false
			bFalse := false
			if err := p.set(rPr, &bFalse); err != nil {
				t.Fatalf("%s: set false: %v", p.name, err)
			}
			got = p.get(rPr)
			if got == nil || *got {
				t.Errorf("%s: expected *false, got %v", p.name, got)
			}

			// Set nil (remove)
			if err := p.set(rPr, nil); err != nil {
				t.Fatalf("%s: set nil: %v", p.name, err)
			}
			if p.get(rPr) != nil {
				t.Errorf("%s: expected nil after removal", p.name)
			}
		})
	}
}

// ===========================================================================
// CT_RPr additional property tests
// ===========================================================================

func TestCT_RPr_ColorTheme_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{E: OxmlElement("w:rPr")}}

	// Initially nil
	if ct, err := rPr.ColorTheme(); err != nil {
		t.Fatalf("ColorTheme: %v", err)
	} else if ct != nil {
		t.Error("expected nil color theme initially")
	}

	// Set theme
	tc := enum.MsoThemeColorIndexAccent1
	if err := rPr.SetColorTheme(&tc); err != nil {
		t.Fatalf("SetColorTheme: %v", err)
	}
	got, err := rPr.ColorTheme()
	if err != nil {
		t.Fatalf("ColorTheme: %v", err)
	}
	if got == nil || *got != enum.MsoThemeColorIndexAccent1 {
		t.Errorf("expected Accent1, got %v", got)
	}

	// Remove
	if err := rPr.SetColorTheme(nil); err != nil {
		t.Fatalf("SetColorTheme nil: %v", err)
	}
	got, _ = rPr.ColorTheme()
	if got != nil {
		t.Error("expected nil after removing theme color")
	}
}

func TestCT_RPr_HighlightVal_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{E: OxmlElement("w:rPr")}}

	if hv, err := rPr.HighlightVal(); err != nil {
		t.Fatalf("HighlightVal: %v", err)
	} else if hv != nil {
		t.Error("expected nil highlight initially")
	}

	h := "yellow"
	if err := rPr.SetHighlightVal(&h); err != nil {
		t.Fatalf("SetHighlightVal: %v", err)
	}
	got, err := rPr.HighlightVal()
	if err != nil {
		t.Fatalf("HighlightVal: %v", err)
	}
	if got == nil || *got != "yellow" {
		t.Errorf("expected yellow, got %v", got)
	}

	if err := rPr.SetHighlightVal(nil); err != nil {
		t.Fatalf("SetHighlightVal nil: %v", err)
	}
	if hv, _ := rPr.HighlightVal(); hv != nil {
		t.Error("expected nil after removing highlight")
	}
}

func TestCT_RPr_Superscript_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{E: OxmlElement("w:rPr")}}

	// Initially nil
	if sup, err := rPr.Superscript(); err != nil {
		t.Fatalf("Superscript: %v", err)
	} else if sup != nil {
		t.Error("expected nil initially")
	}

	// Set true
	bTrue := true
	if err := rPr.SetSuperscript(&bTrue); err != nil {
		t.Fatalf("SetSuperscript true: %v", err)
	}
	got, err := rPr.Superscript()
	if err != nil {
		t.Fatalf("Superscript: %v", err)
	}
	if got == nil || !*got {
		t.Error("expected *true for superscript")
	}

	// Set false clears only if currently superscript
	bFalse := false
	if err := rPr.SetSuperscript(&bFalse); err != nil {
		t.Fatalf("SetSuperscript false: %v", err)
	}
	if sup, _ := rPr.Superscript(); sup != nil {
		t.Error("expected nil after false (was superscript)")
	}
}

func TestCT_RPr_RFontsHAnsi_RoundTrip(t *testing.T) {
	t.Parallel()

	rPr := &CT_RPr{Element{E: OxmlElement("w:rPr")}}

	if rPr.RFontsHAnsi() != nil {
		t.Error("expected nil hAnsi initially")
	}

	f := "Times New Roman"
	if err := rPr.SetRFontsHAnsi(&f); err != nil {
		t.Fatalf("SetRFontsHAnsi: %v", err)
	}
	got := rPr.RFontsHAnsi()
	if got == nil || *got != "Times New Roman" {
		t.Errorf("expected Times New Roman, got %v", got)
	}

	// Set nil
	if err := rPr.SetRFontsHAnsi(nil); err != nil {
		t.Fatalf("SetRFontsHAnsi nil: %v", err)
	}
}

// ===========================================================================
// CT_LastRenderedPageBreak fragmentation tests
// ===========================================================================

// buildParagraphWithBreakInRun builds:
// <w:p><w:pPr/><w:r><w:t>before</w:t><w:lastRenderedPageBreak/><w:t>after</w:t></w:r></w:p>
func buildParagraphWithBreakInRun() (*CT_P, *CT_LastRenderedPageBreak) {
	pEl := OxmlElement("w:p")
	pPrEl := OxmlElement("w:pPr")
	pEl.AddChild(pPrEl)

	rEl := OxmlElement("w:r")
	pEl.AddChild(rEl)

	t1 := OxmlElement("w:t")
	t1.SetText("before")
	rEl.AddChild(t1)

	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	rEl.AddChild(lrpbEl)

	t2 := OxmlElement("w:t")
	t2.SetText("after")
	rEl.AddChild(t2)

	return &CT_P{Element{E: pEl}}, &CT_LastRenderedPageBreak{Element{E: lrpbEl}}
}

func TestCT_LastRenderedPageBreak_PrecedingFragment_InRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInRun()

	frag, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP: %v", err)
	}

	// The preceding fragment should contain "before" but not "after"
	text := frag.ParagraphText()
	if text != "before" {
		t.Errorf("PrecedingFragmentP text: got %q, want %q", text, "before")
	}
}

func TestCT_LastRenderedPageBreak_FollowingFragment_InRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInRun()

	frag, err := lrpb.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP: %v", err)
	}

	text := frag.ParagraphText()
	if text != "after" {
		t.Errorf("FollowingFragmentP text: got %q, want %q", text, "after")
	}
}

// buildParagraphWithBreakInHyperlink builds:
// <w:p><w:pPr/><w:r><w:t>pre</w:t></w:r>
//   <w:hyperlink><w:r><w:lastRenderedPageBreak/><w:t>link</w:t></w:r></w:hyperlink>
//   <w:r><w:t>post</w:t></w:r></w:p>
func buildParagraphWithBreakInHyperlink() (*CT_P, *CT_LastRenderedPageBreak) {
	pEl := OxmlElement("w:p")
	pPrEl := OxmlElement("w:pPr")
	pEl.AddChild(pPrEl)

	r1 := OxmlElement("w:r")
	t1 := OxmlElement("w:t")
	t1.SetText("pre")
	r1.AddChild(t1)
	pEl.AddChild(r1)

	hEl := OxmlElement("w:hyperlink")
	hr := OxmlElement("w:r")
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	hr.AddChild(lrpbEl)
	tLink := OxmlElement("w:t")
	tLink.SetText("link")
	hr.AddChild(tLink)
	hEl.AddChild(hr)
	pEl.AddChild(hEl)

	r2 := OxmlElement("w:r")
	t2 := OxmlElement("w:t")
	t2.SetText("post")
	r2.AddChild(t2)
	pEl.AddChild(r2)

	return &CT_P{Element{E: pEl}}, &CT_LastRenderedPageBreak{Element{E: lrpbEl}}
}

func TestCT_LastRenderedPageBreak_PrecedingFragment_InHyperlink(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInHyperlink()

	frag, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP (hyperlink): %v", err)
	}

	// Preceding should include pre-run and the hyperlink (without lrpb),
	// but not the post-run
	text := frag.ParagraphText()
	if text != "prelink" {
		t.Errorf("PrecedingFragmentP text: got %q, want %q", text, "prelink")
	}
}

func TestCT_LastRenderedPageBreak_FollowingFragment_InHyperlink(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInHyperlink()

	frag, err := lrpb.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP (hyperlink): %v", err)
	}

	// Following should include content after the hyperlink
	text := frag.ParagraphText()
	if text != "post" {
		t.Errorf("FollowingFragmentP text: got %q, want %q", text, "post")
	}
}

func TestCT_LastRenderedPageBreak_Fragment_PreservesProperties(t *testing.T) {
	t.Parallel()

	_, lrpb := buildParagraphWithBreakInRun()

	// pPr should survive in both fragments
	preceding, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP: %v", err)
	}
	if preceding.E.FindElement("w:pPr") == nil {
		t.Error("pPr should be preserved in preceding fragment")
	}

	_, lrpb2 := buildParagraphWithBreakInRun()
	following, err := lrpb2.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP: %v", err)
	}
	if following.E.FindElement("w:pPr") == nil {
		t.Error("pPr should be preserved in following fragment")
	}
}

// buildMultiRunParagraphWithBreak builds:
// <w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t><w:lastRenderedPageBreak/><w:t>C</w:t></w:r><w:r><w:t>D</w:t></w:r></w:p>
func buildMultiRunParagraphWithBreak() (*CT_P, *CT_LastRenderedPageBreak) {
	pEl := OxmlElement("w:p")

	r1 := OxmlElement("w:r")
	t1 := OxmlElement("w:t")
	t1.SetText("A")
	r1.AddChild(t1)
	pEl.AddChild(r1)

	r2 := OxmlElement("w:r")
	t2 := OxmlElement("w:t")
	t2.SetText("B")
	r2.AddChild(t2)
	lrpbEl := OxmlElement("w:lastRenderedPageBreak")
	r2.AddChild(lrpbEl)
	t3 := OxmlElement("w:t")
	t3.SetText("C")
	r2.AddChild(t3)
	pEl.AddChild(r2)

	r3 := OxmlElement("w:r")
	t4 := OxmlElement("w:t")
	t4.SetText("D")
	r3.AddChild(t4)
	pEl.AddChild(r3)

	return &CT_P{Element{E: pEl}}, &CT_LastRenderedPageBreak{Element{E: lrpbEl}}
}

func TestCT_LastRenderedPageBreak_PrecedingFragment_MultiRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildMultiRunParagraphWithBreak()

	frag, err := lrpb.PrecedingFragmentP()
	if err != nil {
		t.Fatalf("PrecedingFragmentP: %v", err)
	}
	text := frag.ParagraphText()
	if text != "AB" {
		t.Errorf("PrecedingFragmentP (multi-run): got %q, want %q", text, "AB")
	}
}

func TestCT_LastRenderedPageBreak_FollowingFragment_MultiRun(t *testing.T) {
	t.Parallel()

	_, lrpb := buildMultiRunParagraphWithBreak()

	frag, err := lrpb.FollowingFragmentP()
	if err != nil {
		t.Fatalf("FollowingFragmentP: %v", err)
	}
	text := frag.ParagraphText()
	if text != "CD" {
		t.Errorf("FollowingFragmentP (multi-run): got %q, want %q", text, "CD")
	}
}
