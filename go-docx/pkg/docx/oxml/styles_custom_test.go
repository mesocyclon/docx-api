package oxml

import (
	"testing"

	"github.com/vortex/go-docx/pkg/docx/enum"
)

// These tests cover CT_Style and CT_LatentStyles methods NOT already tested
// in table_section_styles_custom_test.go.

func TestCT_Style_UnhideWhenUsedVal_RoundTrip(t *testing.T) {
	t.Parallel()

	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	s, err := styles.AddStyleOfType("Custom", enum.WdStyleTypeParagraph, false)
	if err != nil {
		t.Fatal(err)
	}

	if s.UnhideWhenUsedVal() {
		t.Error("expected false by default")
	}

	if err := s.SetUnhideWhenUsedVal(true); err != nil {
		t.Fatalf("SetUnhideWhenUsedVal: %v", err)
	}
	if !s.UnhideWhenUsedVal() {
		t.Error("expected true")
	}

	if err := s.SetUnhideWhenUsedVal(false); err != nil {
		t.Fatalf("SetUnhideWhenUsedVal: %v", err)
	}
	if s.UnhideWhenUsedVal() {
		t.Error("expected false after removing")
	}
}

func TestCT_LatentStyles_BoolProp(t *testing.T) {
	t.Parallel()

	styles := &CT_Styles{Element{E: OxmlElement("w:styles")}}
	ls := styles.GetOrAddLatentStyles()

	// Initially false (not set)
	if ls.BoolProp("w:defSemiHidden") {
		t.Error("expected false for unset property")
	}

	if err := ls.SetBoolProp("w:defSemiHidden", true); err != nil {
		t.Fatalf("SetBoolProp: %v", err)
	}
	if !ls.BoolProp("w:defSemiHidden") {
		t.Error("expected true after setting")
	}

	if err := ls.SetBoolProp("w:defSemiHidden", false); err != nil {
		t.Fatalf("SetBoolProp: %v", err)
	}
	if ls.BoolProp("w:defSemiHidden") {
		t.Error("expected false after clearing")
	}
}
