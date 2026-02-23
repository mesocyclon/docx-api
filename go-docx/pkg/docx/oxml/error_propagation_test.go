package oxml

import (
	"errors"
	"strings"
	"testing"

	"github.com/beevik/etree"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// ===========================================================================
// Phase 4 — Error propagation tests
//
// Verify that malformed XML attributes surface errors instead of being
// silently swallowed. Each test injects garbage into an attribute and
// checks the full error chain: generated getter → custom wrapper → caller.
// ===========================================================================

// mustParseAttrErr asserts err wraps a *ParseAttrError with expected fields.
func mustParseAttrErr(t *testing.T, err error, wantElemSuffix, wantAttr string) {
	t.Helper()
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	var pe *ParseAttrError
	if !errors.As(err, &pe) {
		t.Fatalf("error %T does not wrap *ParseAttrError: %v", err, err)
	}
	if !strings.HasSuffix(pe.Element, wantElemSuffix) {
		t.Errorf("ParseAttrError.Element = %q, want suffix %q", pe.Element, wantElemSuffix)
	}
	if pe.Attr != wantAttr {
		t.Errorf("ParseAttrError.Attr = %q, want %q", pe.Attr, wantAttr)
	}
}

// inject sets an attribute to a garbage value on a raw etree element.
func inject(el *etree.Element, attr, garbage string) {
	el.CreateAttr(attr, garbage)
}

// ---------------------------------------------------------------------------
// Group 1 — Generated int getters (ParseAttrError on non-numeric)
// ---------------------------------------------------------------------------

func TestGen_CT_PageMar_Top_Garbage(t *testing.T) {
	el := OxmlElement("w:pgMar")
	inject(el, "w:top", "not-a-number")
	_, err := (&CT_PageMar{Element{E: el}}).Top()
	mustParseAttrErr(t, err, "pgMar", "w:top")
}

func TestGen_CT_PageMar_Bottom_Garbage(t *testing.T) {
	el := OxmlElement("w:pgMar")
	inject(el, "w:bottom", "abc")
	_, err := (&CT_PageMar{Element{E: el}}).Bottom()
	mustParseAttrErr(t, err, "pgMar", "w:bottom")
}

func TestGen_CT_PageSz_W_Garbage(t *testing.T) {
	el := OxmlElement("w:pgSz")
	inject(el, "w:w", "twelve")
	_, err := (&CT_PageSz{Element{E: el}}).W()
	mustParseAttrErr(t, err, "pgSz", "w:w")
}

func TestGen_CT_DecimalNumber_Val_Garbage(t *testing.T) {
	el := OxmlElement("w:gridSpan")
	inject(el, "w:val", "NaN")
	_, err := (&CT_DecimalNumber{Element{E: el}}).Val()
	mustParseAttrErr(t, err, "gridSpan", "w:val")
}

func TestGen_CT_Height_Val_Garbage(t *testing.T) {
	el := OxmlElement("w:trHeight")
	inject(el, "w:val", "tall")
	_, err := (&CT_Height{Element{E: el}}).Val()
	mustParseAttrErr(t, err, "trHeight", "w:val")
}

func TestGen_CT_Ind_Left_Garbage(t *testing.T) {
	el := OxmlElement("w:ind")
	inject(el, "w:left", "xxx")
	_, err := (&CT_Ind{Element{E: el}}).Left()
	mustParseAttrErr(t, err, "ind", "w:left")
}

func TestGen_CT_TblGridCol_W_Garbage(t *testing.T) {
	el := OxmlElement("w:gridCol")
	inject(el, "w:w", "wide")
	_, err := (&CT_TblGridCol{Element{E: el}}).W()
	mustParseAttrErr(t, err, "gridCol", "w:w")
}

func TestGen_CT_TblWidth_W_Garbage(t *testing.T) {
	el := OxmlElement("w:tblW")
	inject(el, "w:w", "bogus")
	inject(el, "w:type", "dxa")
	_, err := (&CT_TblWidth{Element{E: el}}).W()
	mustParseAttrErr(t, err, "tblW", "w:w")
}

// ---------------------------------------------------------------------------
// Group 2 — Generated enum getters (ParseAttrError on bad enum string)
// ---------------------------------------------------------------------------

func TestGen_CT_Jc_Val_BadEnum(t *testing.T) {
	el := OxmlElement("w:jc")
	inject(el, "w:val", "not-an-alignment")
	_, err := (&CT_Jc{Element{E: el}}).Val()
	mustParseAttrErr(t, err, "jc", "w:val")
}

func TestGen_CT_VerticalJc_Val_BadEnum(t *testing.T) {
	el := OxmlElement("w:vAlign")
	inject(el, "w:val", "diagonal")
	_, err := (&CT_VerticalJc{Element{E: el}}).Val()
	mustParseAttrErr(t, err, "vAlign", "w:val")
}

func TestGen_CT_PageSz_Orient_BadEnum(t *testing.T) {
	el := OxmlElement("w:pgSz")
	inject(el, "w:orient", "upside-down")
	_, err := (&CT_PageSz{Element{E: el}}).Orient()
	mustParseAttrErr(t, err, "pgSz", "w:orient")
}

func TestGen_CT_Height_HRule_BadEnum(t *testing.T) {
	el := OxmlElement("w:trHeight")
	inject(el, "w:hRule", "maybe")
	_, err := (&CT_Height{Element{E: el}}).HRule()
	mustParseAttrErr(t, err, "trHeight", "w:hRule")
}

// ---------------------------------------------------------------------------
// Group 3 — Custom method propagation (errors bubble up through wrappers)
// ---------------------------------------------------------------------------

func TestCustom_SectPr_TopMargin_Propagates(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	pgMar := sp.GetOrAddPgMar()
	inject(pgMar.E, "w:top", "garbage")
	_, err := sp.TopMargin()
	mustParseAttrErr(t, err, "pgMar", "w:top")
}

func TestCustom_SectPr_PageWidth_Propagates(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	pgSz := sp.GetOrAddPgSz()
	inject(pgSz.E, "w:w", "bad")
	_, err := sp.PageWidth()
	mustParseAttrErr(t, err, "pgSz", "w:w")
}

func TestCustom_SectPr_Orientation_Propagates(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	pgSz := sp.GetOrAddPgSz()
	inject(pgSz.E, "w:orient", "sideways")
	_, err := sp.Orientation()
	mustParseAttrErr(t, err, "pgSz", "w:orient")
}

func TestCustom_Row_TrHeightVal_Propagates(t *testing.T) {
	tbl := NewTbl(1, 1, 1000)
	tr := tbl.TrList()[0]
	h := tr.GetOrAddTrPr().GetOrAddTrHeight()
	inject(h.E, "w:val", "high")
	_, err := tr.TrHeightVal()
	mustParseAttrErr(t, err, "trHeight", "w:val")
}

func TestCustom_Tc_GridSpanVal_Propagates(t *testing.T) {
	tc := NewTc()
	gs := tc.GetOrAddTcPr().GetOrAddGridSpan()
	inject(gs.E, "w:val", "many")
	_, err := tc.GridSpanVal()
	mustParseAttrErr(t, err, "gridSpan", "w:val")
}

func TestCustom_Tc_VAlignVal_Propagates(t *testing.T) {
	tc := NewTc()
	va := tc.GetOrAddTcPr().GetOrAddVAlign()
	inject(va.E, "w:val", "diagonal")
	_, err := tc.VAlignVal()
	mustParseAttrErr(t, err, "vAlign", "w:val")
}

func TestCustom_PPr_JcVal_Propagates(t *testing.T) {
	pPr := &CT_PPr{Element{E: OxmlElement("w:pPr")}}
	jc := pPr.GetOrAddJc()
	inject(jc.E, "w:val", "crooked")
	_, err := pPr.JcVal()
	mustParseAttrErr(t, err, "jc", "w:val")
}

func TestCustom_PPr_IndLeft_Propagates(t *testing.T) {
	pPr := &CT_PPr{Element{E: OxmlElement("w:pPr")}}
	ind := pPr.GetOrAddInd()
	inject(ind.E, "w:left", "far")
	_, err := pPr.IndLeft()
	mustParseAttrErr(t, err, "ind", "w:left")
}

func TestCustom_CT_P_Alignment_Propagates(t *testing.T) {
	p := &CT_P{Element{E: OxmlElement("w:p")}}
	pPr := p.GetOrAddPPr()
	jc := pPr.GetOrAddJc()
	inject(jc.E, "w:val", "crooked")
	_, err := p.Alignment()
	mustParseAttrErr(t, err, "jc", "w:val")
}

// ---------------------------------------------------------------------------
// Group 4 — ColWidths propagates through collection
// ---------------------------------------------------------------------------

func TestCustom_Tbl_ColWidths_Propagates(t *testing.T) {
	tbl := NewTbl(1, 3, 9000)
	grid, err := tbl.TblGrid()
	if err != nil {
		t.Fatal(err)
	}
	cols := grid.GridColList()
	inject(cols[1].E, "w:w", "oops")

	_, err = tbl.ColWidths()
	if err == nil {
		t.Fatal("expected error from ColWidths with corrupt gridCol")
	}
	if !strings.Contains(err.Error(), "grid col 1") {
		t.Errorf("error should mention col index: %v", err)
	}
	var pe *ParseAttrError
	if !errors.As(err, &pe) {
		t.Errorf("error should wrap ParseAttrError: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Group 5 — Cascading: GridOffset walks siblings, hits corrupt span
// ---------------------------------------------------------------------------

func TestCustom_Tc_GridOffset_PropagatesSpanError(t *testing.T) {
	tbl := NewTbl(1, 3, 3000)
	tcs := tbl.TrList()[0].TcList()
	gs := tcs[0].GetOrAddTcPr().GetOrAddGridSpan()
	inject(gs.E, "w:val", "bad")

	_, err := tcs[1].GridOffset()
	if err == nil {
		t.Fatal("expected error from GridOffset when sibling gridSpan is corrupt")
	}
	var pe *ParseAttrError
	if !errors.As(err, &pe) {
		t.Errorf("error should wrap ParseAttrError: %v", err)
	}
}

func TestCustom_Tc_Right_PropagatesSpanError(t *testing.T) {
	tc := NewTc()
	gs := tc.GetOrAddTcPr().GetOrAddGridSpan()
	inject(gs.E, "w:val", "bad")
	_, err := tc.Right()
	if err == nil {
		t.Fatal("expected error from Right when gridSpan is corrupt")
	}
}

// ---------------------------------------------------------------------------
// Group 6 — Happy path sanity (refactor didn't break normal flow)
// ---------------------------------------------------------------------------

func TestGen_CT_PageMar_Top_HappyPath(t *testing.T) {
	m := &CT_PageMar{Element{E: OxmlElement("w:pgMar")}}
	v, err := m.Top()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v != nil {
		t.Errorf("expected nil for absent attr, got %d", *v)
	}
	val := 1440
	if err := m.SetTop(&val); err != nil {
		t.Fatal(err)
	}
	v, err = m.Top()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v == nil || *v != 1440 {
		t.Errorf("expected 1440, got %v", v)
	}
}

func TestGen_CT_Jc_Val_HappyPath(t *testing.T) {
	jc := &CT_Jc{Element{E: OxmlElement("w:jc")}}
	if err := jc.SetVal(enum.WdParagraphAlignmentCenter); err != nil {
		t.Fatal(err)
	}
	v, err := jc.Val()
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if v != enum.WdParagraphAlignmentCenter {
		t.Errorf("expected Center, got %v", v)
	}
}

func TestCustom_SectPr_Margins_HappyRoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}
	v := 1440
	if err := sp.SetTopMargin(&v); err != nil {
		t.Fatal(err)
	}
	got, err := sp.TopMargin()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || *got != 1440 {
		t.Errorf("expected 1440, got %v", got)
	}
}

// ---------------------------------------------------------------------------
// Group 7 — Zero vs absent (nullable semantics)
// ---------------------------------------------------------------------------

func TestGen_CT_PageMar_Top_ZeroIsNotAbsent(t *testing.T) {
	el := OxmlElement("w:pgMar")
	m := &CT_PageMar{Element{E: el}}

	// Absent → nil
	v, err := m.Top()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("absent Top should be nil")
	}

	// Set 0 → &0, NOT nil
	zero := 0
	if err := m.SetTop(&zero); err != nil {
		t.Fatal(err)
	}
	v, err = m.Top()
	if err != nil {
		t.Fatal(err)
	}
	if v == nil {
		t.Fatal("Top=0 should NOT be nil")
	}
	if *v != 0 {
		t.Errorf("expected 0, got %d", *v)
	}

	// Verify XML attr actually present
	if _, ok := m.GetAttr("w:top"); !ok {
		t.Error("w:top attr should be present after SetTop(0)")
	}

	// Set nil → remove attr
	if err := m.SetTop(nil); err != nil {
		t.Fatal(err)
	}
	v, err = m.Top()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("should be nil after SetTop(nil)")
	}
	if _, ok := m.GetAttr("w:top"); ok {
		t.Error("w:top attr should be removed after SetTop(nil)")
	}
}

func TestGen_CT_Ind_Left_ZeroIsNotAbsent(t *testing.T) {
	el := OxmlElement("w:ind")
	ind := &CT_Ind{Element{E: el}}

	// Absent → nil
	v, err := ind.Left()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("absent Left should be nil")
	}

	// Set 0 → &0
	zero := 0
	if err := ind.SetLeft(&zero); err != nil {
		t.Fatal(err)
	}
	v, err = ind.Left()
	if err != nil {
		t.Fatal(err)
	}
	if v == nil {
		t.Fatal("Left=0 should NOT be nil")
	}
	if *v != 0 {
		t.Errorf("expected 0, got %d", *v)
	}
}

func TestGen_CT_Spacing_Before_ZeroIsNotAbsent(t *testing.T) {
	el := OxmlElement("w:spacing")
	sp := &CT_Spacing{Element{E: el}}

	// Absent → nil
	v, err := sp.Before()
	if err != nil {
		t.Fatal(err)
	}
	if v != nil {
		t.Error("absent Before should be nil")
	}

	// Set 0 → &0
	zero := 0
	if err := sp.SetBefore(&zero); err != nil {
		t.Fatal(err)
	}
	v, err = sp.Before()
	if err != nil {
		t.Fatal(err)
	}
	if v == nil {
		t.Fatal("Before=0 should NOT be nil")
	}
	if *v != 0 {
		t.Errorf("expected 0, got %d", *v)
	}
}

func TestCustom_SectPr_TopMargin_ZeroRoundTrip(t *testing.T) {
	sp := &CT_SectPr{Element{E: OxmlElement("w:sectPr")}}

	// Set zero margin
	zero := 0
	if err := sp.SetTopMargin(&zero); err != nil {
		t.Fatal(err)
	}
	got, err := sp.TopMargin()
	if err != nil {
		t.Fatal(err)
	}
	if got == nil {
		t.Fatal("TopMargin=0 should NOT be nil")
	}
	if *got != 0 {
		t.Errorf("expected 0, got %d", *got)
	}

	// Set nil → remove
	if err := sp.SetTopMargin(nil); err != nil {
		t.Fatal(err)
	}
	got, err = sp.TopMargin()
	if err != nil {
		t.Fatal(err)
	}
	if got != nil {
		t.Error("expected nil after SetTopMargin(nil)")
	}
}
