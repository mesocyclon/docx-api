package xmltypes

import (
	"encoding/xml"
	"testing"
)

// ============================================================
// Simple wrapper types — table-driven round-trip tests
// ============================================================

func TestSimpleTypes_RoundTrip(t *testing.T) {
	t.Parallel()

	t.Run("CT_String", func(t *testing.T) {
		input := CT_String{Val: "Heading1"}
		var got CT_String
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != "Heading1" {
			t.Errorf("Val = %q, want %q", got.Val, "Heading1")
		}
	})
	t.Run("CT_DecimalNumber", func(t *testing.T) {
		input := CT_DecimalNumber{Val: 42}
		var got CT_DecimalNumber
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != 42 {
			t.Errorf("Val = %d, want %d", got.Val, 42)
		}
	})
	t.Run("CT_UnsignedDecimalNumber", func(t *testing.T) {
		input := CT_UnsignedDecimalNumber{Val: 100}
		var got CT_UnsignedDecimalNumber
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != 100 {
			t.Errorf("Val = %d, want %d", got.Val, 100)
		}
	})
	t.Run("CT_TwipsMeasure", func(t *testing.T) {
		input := CT_TwipsMeasure{Val: 1440}
		var got CT_TwipsMeasure
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != 1440 {
			t.Errorf("Val = %d, want %d", got.Val, 1440)
		}
	})
	t.Run("CT_HpsMeasure", func(t *testing.T) {
		input := CT_HpsMeasure{Val: 24}
		var got CT_HpsMeasure
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != 24 {
			t.Errorf("Val = %d, want %d", got.Val, 24)
		}
	})
	t.Run("CT_TextScale", func(t *testing.T) {
		input := CT_TextScale{Val: 150}
		var got CT_TextScale
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != 150 {
			t.Errorf("Val = %d, want %d", got.Val, 150)
		}
	})
	t.Run("CT_LongHexNumber", func(t *testing.T) {
		input := CT_LongHexNumber{Val: "00A4B3C2"}
		var got CT_LongHexNumber
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
		if got.Val != "00A4B3C2" {
			t.Errorf("Val = %q, want %q", got.Val, "00A4B3C2")
		}
	})
	t.Run("CT_Empty", func(t *testing.T) {
		input := CT_Empty{}
		var got CT_Empty
		if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
			t.Fatal(err)
		}
	})
}

// ============================================================
// Complex types round-trip tests
// ============================================================

func TestCTColor_RoundTrip(t *testing.T) {
	t.Parallel()
	input := CT_Color{Val: "FF0000", ThemeColor: strPtr("accent1"), ThemeTint: strPtr("BF")}
	var got CT_Color
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Val != "FF0000" {
		t.Errorf("Val = %q, want %q", got.Val, "FF0000")
	}
	if got.ThemeColor == nil || *got.ThemeColor != "accent1" {
		t.Error("ThemeColor lost")
	}
	if got.ThemeTint == nil || *got.ThemeTint != "BF" {
		t.Error("ThemeTint lost")
	}
	if got.ThemeShade != nil {
		t.Error("ThemeShade should be nil")
	}
}

func TestCTBorder_RoundTrip(t *testing.T) {
	t.Parallel()
	sz, space := 4, 1
	color := "auto"
	input := CT_Border{Val: "single", Sz: &sz, Space: &space, Color: &color}
	var got CT_Border
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Val != "single" {
		t.Errorf("Val = %q, want %q", got.Val, "single")
	}
	if got.Sz == nil || *got.Sz != 4 {
		t.Error("Sz lost")
	}
	if got.Space == nil || *got.Space != 1 {
		t.Error("Space lost")
	}
	if got.Color == nil || *got.Color != "auto" {
		t.Error("Color lost")
	}
}

func TestCTShd_RoundTrip(t *testing.T) {
	t.Parallel()
	fill, color := "FFFF00", "auto"
	input := CT_Shd{Val: "clear", Color: &color, Fill: &fill}
	var got CT_Shd
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Val != "clear" {
		t.Errorf("Val = %q, want %q", got.Val, "clear")
	}
	if got.Fill == nil || *got.Fill != "FFFF00" {
		t.Error("Fill lost")
	}
}

func TestCTFonts_RoundTrip(t *testing.T) {
	t.Parallel()
	input := CT_Fonts{Ascii: strPtr("Arial"), HAnsi: strPtr("Arial"), CS: strPtr("Arial"), EastAsia: strPtr("SimSun")}
	var got CT_Fonts
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Ascii == nil || *got.Ascii != "Arial" {
		t.Error("Ascii lost")
	}
	if got.EastAsia == nil || *got.EastAsia != "SimSun" {
		t.Error("EastAsia lost")
	}
	if got.Hint != nil {
		t.Error("Hint should be nil")
	}
}

func TestCTLanguage_RoundTrip(t *testing.T) {
	t.Parallel()
	input := CT_Language{Val: strPtr("en-US"), EastAsia: strPtr("zh-CN"), Bidi: strPtr("ar-SA")}
	var got CT_Language
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Val == nil || *got.Val != "en-US" {
		t.Error("Val lost")
	}
	if got.EastAsia == nil || *got.EastAsia != "zh-CN" {
		t.Error("EastAsia lost")
	}
	if got.Bidi == nil || *got.Bidi != "ar-SA" {
		t.Error("Bidi lost")
	}
}

func TestCTUnderline_RoundTrip(t *testing.T) {
	t.Parallel()
	input := CT_Underline{Val: strPtr("single"), Color: strPtr("FF0000")}
	var got CT_Underline
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Val == nil || *got.Val != "single" {
		t.Error("Val lost")
	}
	if got.Color == nil || *got.Color != "FF0000" {
		t.Error("Color lost")
	}
}

func TestCTHighlight_RoundTrip(t *testing.T) {
	t.Parallel()
	input := CT_Highlight{Val: "yellow"}
	var got CT_Highlight
	if err := xml.Unmarshal(roundTrip(t, &input), &got); err != nil {
		t.Fatal(err)
	}
	if got.Val != "yellow" {
		t.Errorf("Val = %q, want %q", got.Val, "yellow")
	}
}

// ============================================================
// Helpers
// ============================================================

func strPtr(s string) *string {
	return &s
}

func roundTrip(t *testing.T, v interface{}) []byte {
	t.Helper()
	out, err := xml.Marshal(v)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	return out
}
