package ppr

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/xmltypes"
)

// TestPPrBaseTabs tests tab stop parsing and round-trip.
func TestPPrBaseTabs(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `">` +
		`<w:tabs>` +
		`<w:tab w:val="center" w:pos="4680"/>` +
		`<w:tab w:val="end" w:pos="9360" w:leader="dot"/>` +
		`</w:tabs>` +
		`</w:pPr>`

	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}
	if ppr.Tabs == nil || len(ppr.Tabs.Tab) != 2 {
		t.Fatalf("Expected 2 tabs, got %v", ppr.Tabs)
	}
	if ppr.Tabs.Tab[0].Val != "center" || ppr.Tabs.Tab[0].Pos != 4680 {
		t.Error("Tab[0] wrong")
	}
	if ppr.Tabs.Tab[1].Leader == nil || *ppr.Tabs.Tab[1].Leader != "dot" {
		t.Error("Tab[1].Leader not dot")
	}

	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}
	var ppr2 CT_PPrBase
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if ppr2.Tabs == nil || len(ppr2.Tabs.Tab) != 2 {
		t.Error("round-trip lost tabs")
	}
}

// TestPPrBaseBorders tests paragraph border parsing and round-trip.
func TestPPrBaseBorders(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `">` +
		`<w:pBdr>` +
		`<w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>` +
		`<w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>` +
		`</w:pBdr>` +
		`</w:pPr>`

	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal failed: %v", err)
	}
	if ppr.PBdr == nil || ppr.PBdr.Top == nil || ppr.PBdr.Top.Val != "single" {
		t.Error("PBdr.Top not parsed")
	}
	if ppr.PBdr.Bottom == nil || ppr.PBdr.Bottom.Val != "single" {
		t.Error("PBdr.Bottom not parsed")
	}
	if ppr.PBdr.Left != nil {
		t.Error("PBdr.Left should be nil")
	}

	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal failed: %v", err)
	}
	var ppr2 CT_PPrBase
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal failed: %v", err)
	}
	if ppr2.PBdr == nil || ppr2.PBdr.Top == nil {
		t.Error("round-trip lost borders")
	}
}

// TestPPrBaseOnOffVariants tests CT_OnOff tri-state semantics.
func TestPPrBaseOnOffVariants(t *testing.T) {
	tests := []struct {
		name   string
		xml    string
		expect bool
	}{
		{"bare", `<w:pPr xmlns:w="` + nsW + `"><w:keepNext/></w:pPr>`, true},
		{"true", `<w:pPr xmlns:w="` + nsW + `"><w:keepNext w:val="true"/></w:pPr>`, true},
		{"1", `<w:pPr xmlns:w="` + nsW + `"><w:keepNext w:val="1"/></w:pPr>`, true},
		{"false", `<w:pPr xmlns:w="` + nsW + `"><w:keepNext w:val="false"/></w:pPr>`, false},
		{"0", `<w:pPr xmlns:w="` + nsW + `"><w:keepNext w:val="0"/></w:pPr>`, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var ppr CT_PPrBase
			if err := xml.Unmarshal([]byte(tt.xml), &ppr); err != nil {
				t.Fatalf("Unmarshal: %v", err)
			}
			if got := ppr.KeepNext.Bool(false); got != tt.expect {
				t.Errorf("KeepNext.Bool = %v, want %v", got, tt.expect)
			}
		})
	}
}

// TestPPrBaseElementOrder verifies marshalled output follows xsd:sequence.
func TestPPrBaseElementOrder(t *testing.T) {
	before240 := 240
	after0 := 0
	left720 := 720
	ppr := CT_PPrBase{
		PStyle:            &xmltypes.CT_String{Val: "Normal"},
		KeepNext:          xmltypes.NewOnOff(true),
		Spacing:           &CT_Spacing{Before: &before240, After: &after0},
		Ind:               &CT_Ind{Start: &left720},
		ContextualSpacing: xmltypes.NewOnOff(true),
		Jc:                &CT_Jc{Val: "both"},
		OutlineLvl:        &xmltypes.CT_DecimalNumber{Val: 2},
	}

	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	s := string(output)

	order := []string{"pStyle", "keepNext", "spacing", "ind", "contextualSpacing", "jc", "outlineLvl"}
	for i := 0; i < len(order)-1; i++ {
		a, b := order[i], order[i+1]
		if strings.Index(s, a) >= strings.Index(s, b) {
			t.Errorf("%q should precede %q in:\n%s", a, b, s)
		}
	}
}

// TestRealWorldHeading tests Heading 1 from reference-appendix 3.1.
func TestRealWorldHeading(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `">` +
		`<w:pStyle w:val="Heading1"/>` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>`

	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if ppr.PStyle == nil || ppr.PStyle.Val != "Heading1" {
		t.Error("PStyle wrong")
	}
	if ppr.Jc == nil || ppr.Jc.Val != "center" {
		t.Error("Jc wrong")
	}
}

// TestRealWorldFootnoteSpacing tests footnote pPr from reference-appendix 3.8.
func TestRealWorldFootnoteSpacing(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `">` +
		`<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
		`</w:pPr>`

	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if ppr.Spacing == nil || ppr.Spacing.After == nil || *ppr.Spacing.After != 0 {
		t.Error("Spacing.After wrong")
	}
	if ppr.Spacing.Line == nil || *ppr.Spacing.Line != 240 {
		t.Error("Spacing.Line wrong")
	}
	if ppr.Spacing.LineRule == nil || *ppr.Spacing.LineRule != "auto" {
		t.Error("Spacing.LineRule wrong")
	}
}

// TestMultipleUnknownElements tests multiple unknown extension elements round-trip.
func TestMultipleUnknownElements(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `" xmlns:w14="` + xmltypes.NSw14 +
		`" xmlns:w15="` + xmltypes.NSw15 + `">` +
		`<w:pStyle w:val="Normal"/>` +
		`<w14:ext1 w14:val="a"/>` +
		`<w15:ext2 w15:val="b"/>` +
		`</w:pPr>`

	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if len(ppr.Extra) != 2 {
		t.Fatalf("Expected 2 Extra, got %d", len(ppr.Extra))
	}
	if ppr.Extra[0].XMLName.Local != "ext1" {
		t.Errorf("Extra[0] = %q", ppr.Extra[0].XMLName.Local)
	}
	if ppr.Extra[1].XMLName.Local != "ext2" {
		t.Errorf("Extra[1] = %q", ppr.Extra[1].XMLName.Local)
	}

	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var ppr2 CT_PPrBase
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal: %v", err)
	}
	if len(ppr2.Extra) != 2 {
		t.Errorf("round-trip lost Extra: %d", len(ppr2.Extra))
	}
}

// TestEmptyPPr tests that an empty pPr round-trips correctly.
func TestEmptyPPr(t *testing.T) {
	input := `<w:pPr xmlns:w="` + nsW + `"></w:pPr>`
	var ppr CT_PPrBase
	if err := xml.Unmarshal([]byte(input), &ppr); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	output, err := xml.Marshal(&ppr)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	var ppr2 CT_PPrBase
	if err := xml.Unmarshal(output, &ppr2); err != nil {
		t.Fatalf("Re-unmarshal: %v", err)
	}
}
