package ppr

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/xmltypes"
)

const nsW = xmltypes.NSw

// TestPPrBaseRoundTrip tests unmarshal → marshal → unmarshal with known and unknown elements.
func TestPPrBaseRoundTrip(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `" xmlns:w14="` + xmltypes.NSw14 + `">` +
		`<w:pStyle w:val="Heading1"/>` +
		`<w:keepNext/>` +
		`<w:keepLines/>` +
		`<w:spacing w:before="240" w:after="0"/>` +
		`<w:jc w:val="center"/>` +
		`<w:outlineLvl w:val="0"/>` +
		`<w14:someExtension w14:val="test"/>` +
		`</w:pPr>`

	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	// Verify known fields
	if ppr.PStyle == nil || ppr.PStyle.Val != "Heading1" {
		t.Error("PStyle not parsed correctly")
	}
	if !ppr.KeepNext.Bool(false) {
		t.Error("KeepNext should be true")
	}
	if !ppr.KeepLines.Bool(false) {
		t.Error("KeepLines should be true")
	}
	if ppr.Spacing == nil {
		t.Fatal("Spacing not parsed")
	}
	if ppr.Spacing.Before == nil || *ppr.Spacing.Before != 240 {
		t.Error("Spacing.Before not 240")
	}
	if ppr.Spacing.After == nil || *ppr.Spacing.After != 0 {
		t.Error("Spacing.After not 0")
	}
	if ppr.Jc == nil || ppr.Jc.Val != "center" {
		t.Error("Jc not parsed correctly")
	}
	if ppr.OutlineLvl == nil || ppr.OutlineLvl.Val != 0 {
		t.Error("OutlineLvl not parsed correctly")
	}
	if len(ppr.Extra) != 1 {
		t.Fatalf("Expected 1 Extra element, got %d", len(ppr.Extra))
	}
	if ppr.Extra[0].XMLName.Local != "someExtension" {
		t.Errorf("Extra[0] local = %q, want someExtension", ppr.Extra[0].XMLName.Local)
	}
	if ppr.PageBreakBefore != nil {
		t.Error("PageBreakBefore should be nil")
	}

	// Marshal
	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}

	// Verify element ordering
	s := string(output)
	order := []string{"pStyle", "keepNext", "keepLines", "spacing", "jc", "outlineLvl"}
	for i := 0; i < len(order)-1; i++ {
		a, b := order[i], order[i+1]
		if strings.Index(s, a) >= strings.Index(s, b) {
			t.Errorf("%q should come before %q in output", a, b)
		}
	}

	// Re-unmarshal and compare
	var ppr2 CT_PPrBase
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if ppr2.PStyle == nil || ppr2.PStyle.Val != ppr.PStyle.Val {
		t.Error("round-trip lost PStyle")
	}
	if !ppr2.KeepNext.Bool(false) {
		t.Error("round-trip lost KeepNext")
	}
	if len(ppr2.Extra) != 1 || ppr2.Extra[0].XMLName.Local != "someExtension" {
		t.Error("round-trip lost Extra")
	}
}

// TestPPrFullRoundTrip tests CT_PPr with rPr and numPr.
func TestPPrFullRoundTrip(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `">` +
		`<w:pStyle w:val="ListParagraph"/>` +
		`<w:numPr>` +
		`<w:ilvl w:val="0"/>` +
		`<w:numId w:val="1"/>` +
		`</w:numPr>` +
		`<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>` +
		`<w:ind w:start="720" w:hanging="360"/>` +
		`<w:rPr/>` +
		`</w:pPr>`

	var ppr CT_PPr
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if ppr.Base.PStyle == nil || ppr.Base.PStyle.Val != "ListParagraph" {
		t.Error("PStyle not parsed")
	}
	if ppr.Base.NumPr == nil {
		t.Fatal("NumPr not parsed")
	}
	if ppr.Base.NumPr.Ilvl == nil || ppr.Base.NumPr.Ilvl.Val != 0 {
		t.Error("NumPr.Ilvl not 0")
	}
	if ppr.Base.NumPr.NumId == nil || ppr.Base.NumPr.NumId.Val != 1 {
		t.Error("NumPr.NumId not 1")
	}
	if ppr.Base.Ind == nil || ppr.Base.Ind.Start == nil || *ppr.Base.Ind.Start != 720 {
		t.Error("Ind.Start not 720")
	}
	if ppr.Base.Ind.Hanging == nil || *ppr.Base.Ind.Hanging != 360 {
		t.Error("Ind.Hanging not 360")
	}
	if ppr.RPr == nil {
		t.Error("rPr not parsed")
	}

	// Marshal + re-unmarshal
	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}
	var ppr2 CT_PPr
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if ppr2.Base.PStyle == nil || ppr2.Base.PStyle.Val != "ListParagraph" {
		t.Error("round-trip lost PStyle")
	}
	if ppr2.Base.NumPr == nil || ppr2.Base.NumPr.NumId == nil {
		t.Error("round-trip lost NumPr")
	}
}

// TestPPrChangeRoundTrip tests CT_PPrChange parsing and round-trip.
func TestPPrChangeRoundTrip(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `">` +
		`<w:jc w:val="center"/>` +
		`<w:pPrChange w:id="5" w:author="John" w:date="2025-01-15T10:00:00Z">` +
		`<w:pPr>` +
		`<w:jc w:val="left"/>` +
		`</w:pPr>` +
		`</w:pPrChange>` +
		`</w:pPr>`

	var ppr CT_PPr
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}

	if ppr.Base.Jc == nil || ppr.Base.Jc.Val != "center" {
		t.Error("Jc not parsed")
	}
	if ppr.PPrChange == nil {
		t.Fatal("PPrChange not parsed")
	}
	if ppr.PPrChange.ID != 5 {
		t.Errorf("PPrChange.ID = %d, want 5", ppr.PPrChange.ID)
	}
	if ppr.PPrChange.Author != "John" {
		t.Errorf("PPrChange.Author = %q, want John", ppr.PPrChange.Author)
	}
	if ppr.PPrChange.PPr == nil || ppr.PPrChange.PPr.Jc == nil {
		t.Error("PPrChange inner PPr not parsed")
	}

	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}
	var ppr2 CT_PPr
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if ppr2.PPrChange == nil || ppr2.PPrChange.ID != 5 {
		t.Error("round-trip lost PPrChange")
	}
}
