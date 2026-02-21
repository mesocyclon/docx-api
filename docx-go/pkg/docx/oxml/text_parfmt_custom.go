package oxml

import (
	"github.com/user/go-docx/pkg/docx/enum"
)

// --- CT_PPr custom methods ---

// JcVal returns the paragraph justification value, or nil if not set.
func (pPr *CT_PPr) JcVal() *enum.WdParagraphAlignment {
	jc := pPr.Jc()
	if jc == nil {
		return nil
	}
	v, err := jc.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetJcVal sets the justification value. Passing nil removes the jc element.
func (pPr *CT_PPr) SetJcVal(v *enum.WdParagraphAlignment) {
	if v == nil {
		pPr.RemoveJc()
		return
	}
	pPr.GetOrAddJc().SetVal(*v)
}

// --- Style ---

// StyleVal returns the paragraph style string, or nil if not set.
func (pPr *CT_PPr) StyleVal() *string {
	pStyle := pPr.PStyle()
	if pStyle == nil {
		return nil
	}
	v, err := pStyle.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetStyleVal sets the paragraph style. Passing nil removes pStyle.
func (pPr *CT_PPr) SetStyleVal(v *string) {
	if v == nil {
		pPr.RemovePStyle()
		return
	}
	pPr.GetOrAddPStyle().SetVal(*v)
}

// --- Spacing properties ---

// SpacingBefore returns the value of w:spacing/@w:before in twips, or nil if not present.
func (pPr *CT_PPr) SpacingBefore() *int {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil
	}
	v, ok := spacing.GetAttr("w:before")
	if !ok {
		return nil
	}
	i := parseIntAttr(v)
	return &i
}

// SetSpacingBefore sets the w:spacing/@w:before value in twips.
// Passing nil removes the attribute (creates spacing element if needed for other attrs).
func (pPr *CT_PPr) SetSpacingBefore(v *int) {
	if v == nil && pPr.Spacing() == nil {
		return
	}
	spacing := pPr.GetOrAddSpacing()
	if v == nil {
		spacing.SetBefore(0) // removes attr via generated code
	} else {
		spacing.SetBefore(*v)
	}
}

// SpacingAfter returns the value of w:spacing/@w:after in twips, or nil if not present.
func (pPr *CT_PPr) SpacingAfter() *int {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil
	}
	v, ok := spacing.GetAttr("w:after")
	if !ok {
		return nil
	}
	i := parseIntAttr(v)
	return &i
}

// SetSpacingAfter sets the w:spacing/@w:after value in twips.
func (pPr *CT_PPr) SetSpacingAfter(v *int) {
	if v == nil && pPr.Spacing() == nil {
		return
	}
	spacing := pPr.GetOrAddSpacing()
	if v == nil {
		spacing.SetAfter(0)
	} else {
		spacing.SetAfter(*v)
	}
}

// SpacingLine returns the value of w:spacing/@w:line in twips, or nil if not present.
func (pPr *CT_PPr) SpacingLine() *int {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil
	}
	v, ok := spacing.GetAttr("w:line")
	if !ok {
		return nil
	}
	i := parseIntAttr(v)
	return &i
}

// SetSpacingLine sets the w:spacing/@w:line value.
func (pPr *CT_PPr) SetSpacingLine(v *int) {
	if v == nil && pPr.Spacing() == nil {
		return
	}
	spacing := pPr.GetOrAddSpacing()
	if v == nil {
		spacing.SetLine(0)
	} else {
		spacing.SetLine(*v)
	}
}

// SpacingLineRule returns the line spacing rule, or nil if not present.
// Defaults to WdLineSpacingMultiple if spacing/@w:line is present but lineRule is absent.
func (pPr *CT_PPr) SpacingLineRule() *enum.WdLineSpacing {
	spacing := pPr.Spacing()
	if spacing == nil {
		return nil
	}
	lr := spacing.LineRule()
	if lr == "" {
		// Check if line is present; if so, default to MULTIPLE
		_, hasLine := spacing.GetAttr("w:line")
		if hasLine {
			v := enum.WdLineSpacingMultiple
			return &v
		}
		return nil
	}
	v, err := enum.WdLineSpacingFromXml(lr)
	if err != nil {
		return nil
	}
	return &v
}

// SetSpacingLineRule sets the line spacing rule.
func (pPr *CT_PPr) SetSpacingLineRule(v *enum.WdLineSpacing) {
	if v == nil && pPr.Spacing() == nil {
		return
	}
	spacing := pPr.GetOrAddSpacing()
	if v == nil {
		spacing.SetLineRule("")
	} else {
		xml, err := v.ToXml()
		if err == nil {
			spacing.SetLineRule(xml)
		}
	}
}

// --- Indentation properties ---

// IndLeft returns the value of w:ind/@w:left in twips, or nil if not present.
func (pPr *CT_PPr) IndLeft() *int {
	ind := pPr.Ind()
	if ind == nil {
		return nil
	}
	_, ok := ind.GetAttr("w:left")
	if !ok {
		return nil
	}
	v := ind.Left()
	return &v
}

// SetIndLeft sets the w:ind/@w:left in twips.
func (pPr *CT_PPr) SetIndLeft(v *int) {
	if v == nil && pPr.Ind() == nil {
		return
	}
	ind := pPr.GetOrAddInd()
	if v == nil {
		ind.SetLeft(0)
	} else {
		ind.SetLeft(*v)
	}
}

// IndRight returns the value of w:ind/@w:right in twips, or nil if not present.
func (pPr *CT_PPr) IndRight() *int {
	ind := pPr.Ind()
	if ind == nil {
		return nil
	}
	_, ok := ind.GetAttr("w:right")
	if !ok {
		return nil
	}
	v := ind.Right()
	return &v
}

// SetIndRight sets the w:ind/@w:right in twips.
func (pPr *CT_PPr) SetIndRight(v *int) {
	if v == nil && pPr.Ind() == nil {
		return
	}
	ind := pPr.GetOrAddInd()
	if v == nil {
		ind.SetRight(0)
	} else {
		ind.SetRight(*v)
	}
}

// FirstLineIndent returns a calculated indentation from w:ind/@w:firstLine and
// w:ind/@w:hanging. A hanging indent is returned as negative.
// Returns nil if no w:ind element.
func (pPr *CT_PPr) FirstLineIndent() *int {
	ind := pPr.Ind()
	if ind == nil {
		return nil
	}
	_, hasHanging := ind.GetAttr("w:hanging")
	if hasHanging {
		v := -ind.Hanging()
		return &v
	}
	_, hasFirstLine := ind.GetAttr("w:firstLine")
	if !hasFirstLine {
		return nil
	}
	v := ind.FirstLine()
	return &v
}

// SetFirstLineIndent sets the first-line indent. Negative values become hanging indents.
// nil clears both firstLine and hanging.
func (pPr *CT_PPr) SetFirstLineIndent(v *int) {
	if pPr.Ind() == nil && v == nil {
		return
	}
	ind := pPr.GetOrAddInd()
	ind.SetFirstLine(0)
	ind.SetHanging(0)
	if v == nil {
		return
	}
	if *v < 0 {
		ind.SetHanging(-*v)
	} else {
		ind.SetFirstLine(*v)
	}
}

// --- Paragraph formatting booleans (keepLines, keepNext, pageBreakBefore, widowControl) ---

// pPrBoolVal reads a tri-state from a CT_OnOff child by tag.
func (pPr *CT_PPr) pPrBoolVal(tag string) *bool {
	child := pPr.FindChild(tag)
	if child == nil {
		return nil
	}
	onOff := &CT_OnOff{Element{E: child}}
	v := onOff.Val()
	return &v
}

// KeepLinesVal returns the tri-state keepLines value.
func (pPr *CT_PPr) KeepLinesVal() *bool {
	return pPr.pPrBoolVal("w:keepLines")
}

// SetKeepLinesVal sets keepLines. nil removes the element.
func (pPr *CT_PPr) SetKeepLinesVal(v *bool) {
	if v == nil {
		pPr.RemoveKeepLines()
	} else {
		pPr.GetOrAddKeepLines().SetVal(*v)
	}
}

// KeepNextVal returns the tri-state keepNext value.
func (pPr *CT_PPr) KeepNextVal() *bool {
	return pPr.pPrBoolVal("w:keepNext")
}

// SetKeepNextVal sets keepNext. nil removes the element.
func (pPr *CT_PPr) SetKeepNextVal(v *bool) {
	if v == nil {
		pPr.RemoveKeepNext()
	} else {
		pPr.GetOrAddKeepNext().SetVal(*v)
	}
}

// PageBreakBeforeVal returns the tri-state pageBreakBefore value.
func (pPr *CT_PPr) PageBreakBeforeVal() *bool {
	return pPr.pPrBoolVal("w:pageBreakBefore")
}

// SetPageBreakBeforeVal sets pageBreakBefore. nil removes the element.
func (pPr *CT_PPr) SetPageBreakBeforeVal(v *bool) {
	if v == nil {
		pPr.RemovePageBreakBefore()
	} else {
		pPr.GetOrAddPageBreakBefore().SetVal(*v)
	}
}

// WidowControlVal returns the tri-state widowControl value.
func (pPr *CT_PPr) WidowControlVal() *bool {
	return pPr.pPrBoolVal("w:widowControl")
}

// SetWidowControlVal sets widowControl. nil removes the element.
func (pPr *CT_PPr) SetWidowControlVal(v *bool) {
	if v == nil {
		pPr.RemoveWidowControl()
	} else {
		pPr.GetOrAddWidowControl().SetVal(*v)
	}
}

// --- CT_TabStops custom methods ---

// InsertTabInOrder inserts a new <w:tab> child element in position order.
func (tabs *CT_TabStops) InsertTabInOrder(pos int, align enum.WdTabAlignment, leader enum.WdTabLeader) *CT_TabStop {
	newTab := tabs.newTab()
	newTab.SetPos(pos)
	newTab.SetVal(align)
	newTab.SetLeader(leader)

	for _, tab := range tabs.TabList() {
		tabPos, err := tab.Pos()
		if err == nil && pos < tabPos {
			insertBefore(tabs.E, newTab.E, tab.E)
			return newTab
		}
	}
	tabs.E.AddChild(newTab.E)
	return newTab
}
