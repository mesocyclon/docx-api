package oxml

import (
	"fmt"

	"github.com/beevik/etree"
	"github.com/user/go-docx/pkg/docx/enum"
)

// ===========================================================================
// CT_Tbl — custom methods
// ===========================================================================

// NewTbl creates a new <w:tbl> element with the given number of rows and columns.
// Width (in twips) is distributed evenly among columns.
func NewTbl(rows, cols int, widthTwips int) *CT_Tbl {
	tblE := OxmlElement("w:tbl")
	tbl := &CT_Tbl{Element{E: tblE}}

	// tblPr
	tblPrE := tblE.CreateElement("tblPr")
	tblPrE.Space = "w"
	tblW := tblPrE.CreateElement("tblW")
	tblW.Space = "w"
	tblW.CreateAttr("w:type", "auto")
	tblW.CreateAttr("w:w", "0")
	tblLook := tblPrE.CreateElement("tblLook")
	tblLook.Space = "w"
	tblLook.CreateAttr("w:firstColumn", "1")
	tblLook.CreateAttr("w:firstRow", "1")
	tblLook.CreateAttr("w:lastColumn", "0")
	tblLook.CreateAttr("w:lastRow", "0")
	tblLook.CreateAttr("w:noHBand", "0")
	tblLook.CreateAttr("w:noVBand", "1")
	tblLook.CreateAttr("w:val", "04A0")

	// tblGrid
	colWidth := 0
	if cols > 0 {
		colWidth = widthTwips / cols
	}
	tblGridE := tblE.CreateElement("tblGrid")
	tblGridE.Space = "w"
	for i := 0; i < cols; i++ {
		gc := tblGridE.CreateElement("gridCol")
		gc.Space = "w"
		gc.CreateAttr("w:w", formatIntAttr(colWidth))
	}

	// rows
	for r := 0; r < rows; r++ {
		trE := tblE.CreateElement("tr")
		trE.Space = "w"
		for c := 0; c < cols; c++ {
			tcE := trE.CreateElement("tc")
			tcE.Space = "w"
			tcPrE := tcE.CreateElement("tcPr")
			tcPrE.Space = "w"
			tcW := tcPrE.CreateElement("tcW")
			tcW.Space = "w"
			tcW.CreateAttr("w:type", "dxa")
			tcW.CreateAttr("w:w", formatIntAttr(colWidth))
			pE := tcE.CreateElement("p")
			pE.Space = "w"
		}
	}

	return tbl
}

// TblStyleVal returns the value of tblPr/tblStyle/@w:val or "" if not present.
func (t *CT_Tbl) TblStyleVal() string {
	tblPr := t.TblPr()
	ts := tblPr.TblStyle()
	if ts == nil {
		return ""
	}
	v, err := ts.Val()
	if err != nil {
		return ""
	}
	return v
}

// SetTblStyleVal sets tblPr/tblStyle/@w:val. Passing "" removes tblStyle.
func (t *CT_Tbl) SetTblStyleVal(styleID string) {
	tblPr := t.TblPr()
	tblPr.RemoveTblStyle()
	if styleID == "" {
		return
	}
	tblPr.GetOrAddTblStyle().SetVal(styleID)
}

// AlignmentVal returns the table alignment from tblPr/jc, or nil if not set.
func (t *CT_Tbl) AlignmentVal() *enum.WdTableAlignment {
	tblPr := t.TblPr()
	jc := tblPr.Jc()
	if jc == nil {
		return nil
	}
	val, ok := jc.GetAttr("w:val")
	if !ok {
		return nil
	}
	v, err := enum.WdTableAlignmentFromXml(val)
	if err != nil {
		return nil
	}
	return &v
}

// SetAlignmentVal sets the table alignment. Passing nil removes jc.
func (t *CT_Tbl) SetAlignmentVal(v *enum.WdTableAlignment) {
	tblPr := t.TblPr()
	tblPr.RemoveJc()
	if v == nil {
		return
	}
	jc := tblPr.GetOrAddJc()
	jc.SetAttr("w:val", v.ToXml())
}

// BidiVisualVal returns the value of tblPr/bidiVisual, or nil if not present.
func (t *CT_Tbl) BidiVisualVal() *bool {
	tblPr := t.TblPr()
	bidi := tblPr.BidiVisual()
	if bidi == nil {
		return nil
	}
	v := bidi.Val()
	return &v
}

// SetBidiVisualVal sets tblPr/bidiVisual. Passing nil removes it.
func (t *CT_Tbl) SetBidiVisualVal(v *bool) {
	tblPr := t.TblPr()
	if v == nil {
		tblPr.RemoveBidiVisual()
		return
	}
	tblPr.GetOrAddBidiVisual().SetVal(*v)
}

// Autofit returns false when there is a tblLayout with type="fixed", true otherwise.
func (t *CT_Tbl) Autofit() bool {
	tblPr := t.TblPr()
	layout := tblPr.TblLayout()
	if layout == nil {
		return true
	}
	return layout.Type() != "fixed"
}

// SetAutofit sets the table layout to "autofit" or "fixed".
func (t *CT_Tbl) SetAutofit(v bool) {
	tblPr := t.TblPr()
	layout := tblPr.GetOrAddTblLayout()
	if v {
		layout.SetType("autofit")
	} else {
		layout.SetType("fixed")
	}
}

// ColCount returns the number of grid columns defined in tblGrid.
func (t *CT_Tbl) ColCount() int {
	return len(t.TblGrid().GridColList())
}

// IterTcs generates each w:tc element in this table, left to right, top to bottom.
func (t *CT_Tbl) IterTcs() []*CT_Tc {
	var result []*CT_Tc
	for _, tr := range t.TrList() {
		result = append(result, tr.TcList()...)
	}
	return result
}

// ColWidths returns the widths (in twips) of each grid column.
func (t *CT_Tbl) ColWidths() []int {
	cols := t.TblGrid().GridColList()
	result := make([]int, len(cols))
	for i, col := range cols {
		result[i] = col.W()
	}
	return result
}

// ===========================================================================
// CT_TblPr — custom methods
// ===========================================================================

// AlignmentVal returns the table alignment, or nil.
func (pr *CT_TblPr) AlignmentVal() *enum.WdTableAlignment {
	jc := pr.Jc()
	if jc == nil {
		return nil
	}
	val, ok := jc.GetAttr("w:val")
	if !ok {
		return nil
	}
	v, err := enum.WdTableAlignmentFromXml(val)
	if err != nil {
		return nil
	}
	return &v
}

// SetAlignmentVal sets the table alignment. Passing nil removes jc.
func (pr *CT_TblPr) SetAlignmentVal(v *enum.WdTableAlignment) {
	pr.RemoveJc()
	if v == nil {
		return
	}
	jc := pr.GetOrAddJc()
	jc.SetAttr("w:val", v.ToXml())
}

// AutofitVal returns false when tblLayout type="fixed", true otherwise.
func (pr *CT_TblPr) AutofitVal() bool {
	layout := pr.TblLayout()
	if layout == nil {
		return true
	}
	return layout.Type() != "fixed"
}

// SetAutofitVal sets the autofit property.
func (pr *CT_TblPr) SetAutofitVal(v bool) {
	layout := pr.GetOrAddTblLayout()
	if v {
		layout.SetType("autofit")
	} else {
		layout.SetType("fixed")
	}
}

// StyleVal returns the value of tblStyle/@w:val or "" if absent.
func (pr *CT_TblPr) StyleVal() string {
	ts := pr.TblStyle()
	if ts == nil {
		return ""
	}
	v, err := ts.Val()
	if err != nil {
		return ""
	}
	return v
}

// SetStyleVal sets the table style. Passing "" removes tblStyle.
func (pr *CT_TblPr) SetStyleVal(v string) {
	pr.RemoveTblStyle()
	if v == "" {
		return
	}
	pr.GetOrAddTblStyle().SetVal(v)
}

// ===========================================================================
// CT_Row — custom methods
// ===========================================================================

// TrIdx returns the index of this w:tr within its parent w:tbl.
// Returns -1 if parent is not found.
func (r *CT_Row) TrIdx() int {
	parent := r.E.Parent()
	if parent == nil {
		return -1
	}
	idx := 0
	for _, child := range parent.ChildElements() {
		if child.Space == "w" && child.Tag == "tr" {
			if child == r.E {
				return idx
			}
			idx++
		}
	}
	return -1
}

// GridBeforeVal returns the number of unpopulated grid cells at the start of this row.
func (r *CT_Row) GridBeforeVal() int {
	trPr := r.TrPr()
	if trPr == nil {
		return 0
	}
	return trPr.GridBeforeVal()
}

// GridAfterVal returns the number of unpopulated grid cells at the end of this row.
func (r *CT_Row) GridAfterVal() int {
	trPr := r.TrPr()
	if trPr == nil {
		return 0
	}
	return trPr.GridAfterVal()
}

// TcAtGridOffset returns the w:tc at the given grid column offset.
// Returns error if no tc at that exact offset.
func (r *CT_Row) TcAtGridOffset(gridOffset int) (*CT_Tc, error) {
	remaining := gridOffset - r.GridBeforeVal()
	for _, tc := range r.TcList() {
		if remaining < 0 {
			break
		}
		if remaining == 0 {
			return tc, nil
		}
		remaining -= tc.GridSpanVal()
	}
	return nil, fmt.Errorf("no tc element at grid_offset=%d", gridOffset)
}

// TrHeightVal returns the value of trPr/trHeight/@w:val (twips), or nil.
func (r *CT_Row) TrHeightVal() *int {
	trPr := r.TrPr()
	if trPr == nil {
		return nil
	}
	return trPr.TrHeightValTwips()
}

// SetTrHeightVal sets the row height. Passing nil removes it.
func (r *CT_Row) SetTrHeightVal(twips *int) {
	if twips == nil {
		trPr := r.TrPr()
		if trPr != nil {
			trPr.RemoveTrHeight()
		}
		return
	}
	trPr := r.GetOrAddTrPr()
	h := trPr.GetOrAddTrHeight()
	h.SetVal(*twips)
}

// TrHeightHRule returns the row height rule, or nil.
func (r *CT_Row) TrHeightHRule() *enum.WdRowHeightRule {
	trPr := r.TrPr()
	if trPr == nil {
		return nil
	}
	return trPr.TrHeightHRuleVal()
}

// SetTrHeightHRule sets the row height rule. Passing nil removes it.
func (r *CT_Row) SetTrHeightHRule(rule *enum.WdRowHeightRule) {
	if rule == nil {
		trPr := r.TrPr()
		if trPr != nil {
			h := trPr.TrHeight()
			if h != nil {
				h.SetHRule(enum.WdRowHeightRule(0))
			}
		}
		return
	}
	trPr := r.GetOrAddTrPr()
	h := trPr.GetOrAddTrHeight()
	h.SetHRule(*rule)
}

// ===========================================================================
// CT_TrPr — custom methods
// ===========================================================================

// GridBeforeVal returns the value of gridBefore/@w:val or 0 if not present.
func (pr *CT_TrPr) GridBeforeVal() int {
	gb := pr.GridBefore()
	if gb == nil {
		return 0
	}
	v, err := gb.Val()
	if err != nil {
		return 0
	}
	return v
}

// GridAfterVal returns the value of gridAfter/@w:val or 0 if not present.
func (pr *CT_TrPr) GridAfterVal() int {
	ga := pr.GridAfter()
	if ga == nil {
		return 0
	}
	v, err := ga.Val()
	if err != nil {
		return 0
	}
	return v
}

// TrHeightValTwips returns the value of trHeight/@w:val in twips, or nil.
func (pr *CT_TrPr) TrHeightValTwips() *int {
	h := pr.TrHeight()
	if h == nil {
		return nil
	}
	v := h.Val()
	if v == 0 {
		return nil
	}
	return &v
}

// SetTrHeightValTwips sets the trHeight value. Passing nil removes trHeight.
func (pr *CT_TrPr) SetTrHeightValTwips(twips *int) {
	if twips == nil {
		pr.RemoveTrHeight()
		return
	}
	h := pr.GetOrAddTrHeight()
	h.SetVal(*twips)
}

// TrHeightHRuleVal returns the height rule, or nil.
func (pr *CT_TrPr) TrHeightHRuleVal() *enum.WdRowHeightRule {
	h := pr.TrHeight()
	if h == nil {
		return nil
	}
	v := h.HRule()
	if v == enum.WdRowHeightRule(0) {
		return nil
	}
	return &v
}

// SetTrHeightHRuleVal sets the height rule. Passing nil removes it.
func (pr *CT_TrPr) SetTrHeightHRuleVal(rule *enum.WdRowHeightRule) {
	if rule == nil {
		h := pr.TrHeight()
		if h != nil {
			h.SetHRule(enum.WdRowHeightRule(0))
		}
		return
	}
	h := pr.GetOrAddTrHeight()
	h.SetHRule(*rule)
}

// ===========================================================================
// CT_Tc — custom methods
// ===========================================================================

// NewTc creates a new <w:tc> element with a single empty <w:p>.
func NewTc() *CT_Tc {
	tcE := OxmlElement("w:tc")
	pE := tcE.CreateElement("p")
	pE.Space = "w"
	return &CT_Tc{Element{E: tcE}}
}

// GridSpanVal returns the number of grid columns this cell spans (default 1).
func (tc *CT_Tc) GridSpanVal() int {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return 1
	}
	return tcPr.GridSpanVal()
}

// SetGridSpanVal sets the grid span. Values ≤ 1 remove the gridSpan element.
func (tc *CT_Tc) SetGridSpanVal(v int) {
	tcPr := tc.GetOrAddTcPr()
	tcPr.SetGridSpanVal(v)
}

// VMergeVal returns the value of tcPr/vMerge/@w:val, or nil if vMerge is not present.
// When vMerge is present without @val, returns "continue".
func (tc *CT_Tc) VMergeVal() *string {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return nil
	}
	return tcPr.VMergeValStr()
}

// SetVMergeVal sets the vMerge value. Passing nil removes vMerge.
func (tc *CT_Tc) SetVMergeVal(v *string) {
	tcPr := tc.GetOrAddTcPr()
	tcPr.SetVMergeValStr(v)
}

// WidthTwips returns the cell width in twips from tcPr/tcW, or nil if not present
// or not dxa type.
func (tc *CT_Tc) WidthTwips() *int {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return nil
	}
	return tcPr.WidthTwips()
}

// SetWidthTwips sets the cell width in twips.
func (tc *CT_Tc) SetWidthTwips(twips int) {
	tcPr := tc.GetOrAddTcPr()
	tcPr.SetWidthTwips(twips)
}

// VAlignVal returns the vertical alignment of this cell, or nil.
func (tc *CT_Tc) VAlignVal() *enum.WdCellVerticalAlignment {
	tcPr := tc.TcPr()
	if tcPr == nil {
		return nil
	}
	return tcPr.VAlignValEnum()
}

// SetVAlignVal sets the vertical alignment. Passing nil removes vAlign.
func (tc *CT_Tc) SetVAlignVal(v *enum.WdCellVerticalAlignment) {
	tcPr := tc.GetOrAddTcPr()
	tcPr.SetVAlignValEnum(v)
}

// InnerContentElements returns all w:p and w:tbl direct children in document order.
func (tc *CT_Tc) InnerContentElements() []interface{} {
	var result []interface{}
	for _, child := range tc.E.ChildElements() {
		if child.Space == "w" && child.Tag == "p" {
			result = append(result, &CT_P{Element{E: child}})
		} else if child.Space == "w" && child.Tag == "tbl" {
			result = append(result, &CT_Tbl{Element{E: child}})
		}
	}
	return result
}

// IterBlockItems generates all block-level content elements: w:p, w:tbl, w:sdt.
func (tc *CT_Tc) IterBlockItems() []*etree.Element {
	var result []*etree.Element
	for _, child := range tc.E.ChildElements() {
		if child.Space == "w" {
			switch child.Tag {
			case "p", "tbl", "sdt":
				result = append(result, child)
			}
		}
	}
	return result
}

// ClearContent removes all children except w:tcPr.
// NOTE: This leaves the cell in an invalid state (missing required w:p).
// Caller must add a w:p afterwards.
func (tc *CT_Tc) ClearContent() {
	var toRemove []*etree.Element
	for _, child := range tc.E.ChildElements() {
		if !(child.Space == "w" && child.Tag == "tcPr") {
			toRemove = append(toRemove, child)
		}
	}
	for _, child := range toRemove {
		tc.E.RemoveChild(child)
	}
}

// GridOffset returns the starting offset of this cell in the grid columns.
func (tc *CT_Tc) GridOffset() int {
	tr := tc.parentTr()
	if tr == nil {
		return 0
	}
	offset := tr.GridBeforeVal()
	for _, child := range tr.E.ChildElements() {
		if child.Space == "w" && child.Tag == "tc" {
			if child == tc.E {
				return offset
			}
			sibling := &CT_Tc{Element{E: child}}
			offset += sibling.GridSpanVal()
		}
	}
	return offset
}

// Left is an alias for GridOffset.
func (tc *CT_Tc) Left() int {
	return tc.GridOffset()
}

// Right returns the grid column just past the right edge of this cell.
func (tc *CT_Tc) Right() int {
	return tc.GridOffset() + tc.GridSpanVal()
}

// Top returns the top-most row index in the vertical span of this cell.
func (tc *CT_Tc) Top() int {
	vm := tc.VMergeVal()
	if vm == nil || *vm == "restart" {
		return tc.trIdx()
	}
	above := tc.tcAbove()
	if above != nil {
		return above.Top()
	}
	return tc.trIdx()
}

// Bottom returns the row index just past the bottom of the vertical span.
func (tc *CT_Tc) Bottom() int {
	vm := tc.VMergeVal()
	if vm != nil {
		below := tc.tcBelow()
		if below != nil {
			bvm := below.VMergeVal()
			if bvm != nil && *bvm == "continue" {
				return below.Bottom()
			}
		}
	}
	return tc.trIdx() + 1
}

// IsEmpty returns true if this cell contains only a single empty w:p.
func (tc *CT_Tc) IsEmpty() bool {
	blocks := tc.IterBlockItems()
	if len(blocks) != 1 {
		return false
	}
	b := blocks[0]
	if !(b.Space == "w" && b.Tag == "p") {
		return false
	}
	p := &CT_P{Element{E: b}}
	return len(p.RList()) == 0
}

// NextTc returns the w:tc element immediately following this one in the row, or nil.
func (tc *CT_Tc) NextTc() *CT_Tc {
	found := false
	tr := tc.parentTr()
	if tr == nil {
		return nil
	}
	for _, child := range tr.E.ChildElements() {
		if found && child.Space == "w" && child.Tag == "tc" {
			return &CT_Tc{Element{E: child}}
		}
		if child == tc.E {
			found = true
		}
	}
	return nil
}

// AddWidthOf adds the width of other to this cell. Does nothing if either has no width.
func (tc *CT_Tc) AddWidthOf(other *CT_Tc) {
	w1 := tc.WidthTwips()
	w2 := other.WidthTwips()
	if w1 != nil && w2 != nil {
		sum := *w1 + *w2
		tc.SetWidthTwips(sum)
	}
}

// MoveContentTo appends the block-level content of this cell to other.
// Leaves this cell with a single empty w:p.
func (tc *CT_Tc) MoveContentTo(other *CT_Tc) {
	if other.E == tc.E {
		return
	}
	if tc.IsEmpty() {
		return
	}
	// Remove trailing empty p from other
	other.removeTrailingEmptyP()
	// Move all block items
	for _, block := range tc.IterBlockItems() {
		tc.E.RemoveChild(block)
		other.E.AddChild(block)
	}
	// Restore minimum required p
	pE := tc.E.CreateElement("p")
	pE.Space = "w"
}

// RemoveElement removes this tc from its parent row.
func (tc *CT_Tc) RemoveElement() {
	parent := tc.E.Parent()
	if parent != nil {
		parent.RemoveChild(tc.E)
	}
}

// Merge merges the rectangular region defined by this tc and other as diagonal corners.
// Returns the top-left tc element of the new span.
func (tc *CT_Tc) Merge(other *CT_Tc) (*CT_Tc, error) {
	top, left, height, width, err := tc.spanDimensions(other)
	if err != nil {
		return nil, err
	}
	tbl := tc.parentTbl()
	if tbl == nil {
		return nil, fmt.Errorf("tc has no parent tbl")
	}
	trs := tbl.TrList()
	if top >= len(trs) {
		return nil, fmt.Errorf("top row %d out of range", top)
	}
	topTc, err := trs[top].TcAtGridOffset(left)
	if err != nil {
		return nil, err
	}
	err = topTc.growTo(width, height, topTc)
	if err != nil {
		return nil, err
	}
	return topTc, nil
}

// --- private helpers ---

func (tc *CT_Tc) parentTr() *CT_Row {
	p := tc.E.Parent()
	if p == nil || !(p.Space == "w" && p.Tag == "tr") {
		return nil
	}
	return &CT_Row{Element{E: p}}
}

func (tc *CT_Tc) parentTbl() *CT_Tbl {
	tr := tc.parentTr()
	if tr == nil {
		return nil
	}
	p := tr.E.Parent()
	if p == nil || !(p.Space == "w" && p.Tag == "tbl") {
		return nil
	}
	return &CT_Tbl{Element{E: p}}
}

func (tc *CT_Tc) trIdx() int {
	tr := tc.parentTr()
	if tr == nil {
		return 0
	}
	return tr.TrIdx()
}

func (tc *CT_Tc) tcAbove() *CT_Tc {
	tbl := tc.parentTbl()
	if tbl == nil {
		return nil
	}
	trs := tbl.TrList()
	idx := tc.trIdx()
	if idx <= 0 {
		return nil
	}
	above, err := trs[idx-1].TcAtGridOffset(tc.GridOffset())
	if err != nil {
		return nil
	}
	return above
}

func (tc *CT_Tc) tcBelow() *CT_Tc {
	tbl := tc.parentTbl()
	if tbl == nil {
		return nil
	}
	trs := tbl.TrList()
	idx := tc.trIdx()
	if idx >= len(trs)-1 {
		return nil
	}
	below, err := trs[idx+1].TcAtGridOffset(tc.GridOffset())
	if err != nil {
		return nil
	}
	return below
}

func (tc *CT_Tc) removeTrailingEmptyP() {
	blocks := tc.IterBlockItems()
	if len(blocks) == 0 {
		return
	}
	last := blocks[len(blocks)-1]
	if !(last.Space == "w" && last.Tag == "p") {
		return
	}
	p := &CT_P{Element{E: last}}
	if len(p.RList()) > 0 {
		return
	}
	tc.E.RemoveChild(last)
}

func (tc *CT_Tc) spanDimensions(other *CT_Tc) (top, left, height, width int, err error) {
	aTop, aLeft := tc.Top(), tc.Left()
	aBottom, aRight := tc.Bottom(), tc.Right()
	bTop, bLeft := other.Top(), other.Left()
	bBottom, bRight := other.Bottom(), other.Right()

	// Check inverted-L
	if aTop == bTop && aBottom != bBottom {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (inverted-L)")
	}
	if aLeft == bLeft && aRight != bRight {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (inverted-L)")
	}
	// Check tee-shaped
	topMost, otherTc := tc, other
	if otherTc.Top() < topMost.Top() {
		topMost, otherTc = otherTc, topMost
	}
	if topMost.Top() < otherTc.Top() && topMost.Bottom() > otherTc.Bottom() {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (tee)")
	}
	leftMost, otherTc2 := tc, other
	if otherTc2.Left() < leftMost.Left() {
		leftMost, otherTc2 = otherTc2, leftMost
	}
	if leftMost.Left() < otherTc2.Left() && leftMost.Right() > otherTc2.Right() {
		return 0, 0, 0, 0, fmt.Errorf("requested span not rectangular (tee)")
	}

	if aTop < bTop {
		top = aTop
	} else {
		top = bTop
	}
	if aLeft < bLeft {
		left = aLeft
	} else {
		left = bLeft
	}
	bottom := aBottom
	if bBottom > bottom {
		bottom = bBottom
	}
	right := aRight
	if bRight > right {
		right = bRight
	}
	return top, left, bottom - top, right - left, nil
}

func (tc *CT_Tc) growTo(width, height int, topTc *CT_Tc) error {
	vMerge := ""
	if topTc.E != tc.E {
		vMerge = "continue"
	} else if height > 1 {
		vMerge = "restart"
	}

	tc.MoveContentTo(topTc)
	// Span to width
	for tc.GridSpanVal() < width {
		next := tc.NextTc()
		if next == nil {
			return fmt.Errorf("not enough grid columns")
		}
		if tc.GridSpanVal()+next.GridSpanVal() > width {
			return fmt.Errorf("span is not rectangular")
		}
		next.MoveContentTo(topTc)
		tc.AddWidthOf(next)
		tc.SetGridSpanVal(tc.GridSpanVal() + next.GridSpanVal())
		next.RemoveElement()
	}

	if vMerge == "" {
		// Remove vMerge entirely
		tc.SetVMergeVal(nil)
	} else {
		tc.SetVMergeVal(&vMerge)
	}

	if height > 1 {
		below := tc.tcBelow()
		if below == nil {
			return fmt.Errorf("not enough rows for vertical span")
		}
		return below.growTo(width, height-1, topTc)
	}
	return nil
}

// ===========================================================================
// CT_TcPr — custom methods
// ===========================================================================

// GridSpanVal returns the grid span value (default 1).
func (pr *CT_TcPr) GridSpanVal() int {
	gs := pr.GridSpan()
	if gs == nil {
		return 1
	}
	v, err := gs.Val()
	if err != nil {
		return 1
	}
	return v
}

// SetGridSpanVal sets the grid span. Values ≤ 1 remove the gridSpan element.
func (pr *CT_TcPr) SetGridSpanVal(v int) {
	pr.RemoveGridSpan()
	if v > 1 {
		pr.GetOrAddGridSpan().SetVal(v)
	}
}

// VMergeValStr returns the vMerge value as a string pointer.
// nil means vMerge is not present. "continue" or "restart".
func (pr *CT_TcPr) VMergeValStr() *string {
	vm := pr.VMerge()
	if vm == nil {
		return nil
	}
	v := vm.Val()
	return &v
}

// SetVMergeValStr sets the vMerge value. nil removes vMerge.
func (pr *CT_TcPr) SetVMergeValStr(v *string) {
	pr.RemoveVMerge()
	if v == nil {
		return
	}
	vm := pr.GetOrAddVMerge()
	vm.SetVal(*v)
}

// VAlignValEnum returns the vertical alignment enum, or nil.
func (pr *CT_TcPr) VAlignValEnum() *enum.WdCellVerticalAlignment {
	va := pr.VAlign()
	if va == nil {
		return nil
	}
	v, err := va.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetVAlignValEnum sets the vertical alignment. nil removes vAlign.
func (pr *CT_TcPr) SetVAlignValEnum(v *enum.WdCellVerticalAlignment) {
	if v == nil {
		pr.RemoveVAlign()
		return
	}
	pr.GetOrAddVAlign().SetVal(*v)
}

// WidthTwips returns the cell width in twips from tcW, or nil if not dxa or absent.
func (pr *CT_TcPr) WidthTwips() *int {
	tcW := pr.TcW()
	if tcW == nil {
		return nil
	}
	t, err := tcW.Type()
	if err != nil || t != "dxa" {
		return nil
	}
	w, err := tcW.W()
	if err != nil {
		return nil
	}
	return &w
}

// SetWidthTwips sets the cell width to dxa type with the given twips value.
func (pr *CT_TcPr) SetWidthTwips(twips int) {
	tcW := pr.GetOrAddTcW()
	tcW.SetType("dxa")
	tcW.SetW(twips)
}

// ===========================================================================
// CT_TblWidth — custom methods
// ===========================================================================

// WidthTwips returns the width in twips if type is "dxa", otherwise nil.
func (tw *CT_TblWidth) WidthTwips() *int {
	t, err := tw.Type()
	if err != nil || t != "dxa" {
		return nil
	}
	w, err := tw.W()
	if err != nil {
		return nil
	}
	return &w
}

// SetWidthDxa sets the width in dxa (twips) and type to "dxa".
func (tw *CT_TblWidth) SetWidthDxa(twips int) {
	tw.SetType("dxa")
	tw.SetW(twips)
}

// ===========================================================================
// CT_TblGridCol — custom methods
// ===========================================================================

// GridColIdx returns the index of this gridCol among its siblings.
func (gc *CT_TblGridCol) GridColIdx() int {
	parent := gc.E.Parent()
	if parent == nil {
		return -1
	}
	idx := 0
	for _, child := range parent.ChildElements() {
		if child.Space == "w" && child.Tag == "gridCol" {
			if child == gc.E {
				return idx
			}
			idx++
		}
	}
	return -1
}
