package units

import (
	"math"
	"testing"
)

// ============================================================
// Helper
// ============================================================

const epsilon = 1e-9

func almostEqual(a, b, eps float64) bool {
	return math.Abs(a-b) < eps
}

// ============================================================
// DXA conversions
// ============================================================

func TestInchToDXA(t *testing.T) {
	tests := []struct {
		in   float64
		want DXA
	}{
		{1.0, 1440},
		{8.5, 12240},  // Letter width
		{11.0, 15840}, // Letter height
		{0.0, 0},
		{0.5, 720},
	}
	for _, tc := range tests {
		got := InchToDXA(tc.in)
		if got != tc.want {
			t.Errorf("InchToDXA(%v) = %d, want %d", tc.in, got, tc.want)
		}
	}
}

func TestCmToDXA(t *testing.T) {
	// 2.54 cm = 1 inch = 1440 DXA
	got := CmToDXA(2.54)
	if got != 1440 {
		t.Errorf("CmToDXA(2.54) = %d, want 1440", got)
	}

	// 21.0 cm ≈ 11906 DXA (A4 width)
	gotA4 := CmToDXA(21.0)
	if gotA4 != A4W {
		t.Errorf("CmToDXA(21.0) = %d, want %d (A4W)", gotA4, A4W)
	}
}

func TestMmToDXA(t *testing.T) {
	// 25.4 mm = 1 inch = 1440 DXA
	got := MmToDXA(25.4)
	if got != 1440 {
		t.Errorf("MmToDXA(25.4) = %d, want 1440", got)
	}

	// 210 mm = A4W
	gotA4 := MmToDXA(210.0)
	if gotA4 != A4W {
		t.Errorf("MmToDXA(210.0) = %d, want %d (A4W)", gotA4, A4W)
	}
}

func TestPtToDXA(t *testing.T) {
	tests := []struct {
		pt   float64
		want DXA
	}{
		{1.0, 20},
		{12.0, 240},
		{72.0, 1440}, // 72pt = 1 inch
		{0.0, 0},
	}
	for _, tc := range tests {
		got := PtToDXA(tc.pt)
		if got != tc.want {
			t.Errorf("PtToDXA(%v) = %d, want %d", tc.pt, got, tc.want)
		}
	}
}

// ============================================================
// EMU conversions
// ============================================================

func TestInchToEMU(t *testing.T) {
	got := InchToEMU(1.0)
	if got != 914400 {
		t.Errorf("InchToEMU(1.0) = %d, want 914400", got)
	}

	got2 := InchToEMU(2.0)
	if got2 != 1828800 {
		t.Errorf("InchToEMU(2.0) = %d, want 1828800", got2)
	}
}

func TestCmToEMU(t *testing.T) {
	got := CmToEMU(1.0)
	if got != 360000 {
		t.Errorf("CmToEMU(1.0) = %d, want 360000", got)
	}
}

func TestMmToEMU(t *testing.T) {
	got := MmToEMU(1.0)
	if got != 36000 {
		t.Errorf("MmToEMU(1.0) = %d, want 36000", got)
	}
}

func TestPtToEMU(t *testing.T) {
	got := PtToEMU(1.0)
	if got != 12700 {
		t.Errorf("PtToEMU(1.0) = %d, want 12700", got)
	}

	got72 := PtToEMU(72.0)
	if got72 != 914400 {
		t.Errorf("PtToEMU(72.0) = %d, want 914400", got72)
	}
}

func TestDXAToEMU(t *testing.T) {
	got := DXAToEMU(1440)
	if got != 914400 {
		t.Errorf("DXAToEMU(1440) = %d, want 914400", got)
	}

	got20 := DXAToEMU(20)
	if got20 != 12700 {
		t.Errorf("DXAToEMU(20) = %d, want 12700", got20)
	}
}

// ============================================================
// Reverse conversions
// ============================================================

func TestEMUToDXA(t *testing.T) {
	got := EMUToDXA(914400)
	if got != 1440 {
		t.Errorf("EMUToDXA(914400) = %d, want 1440", got)
	}

	got2 := EMUToDXA(12700)
	if got2 != 20 {
		t.Errorf("EMUToDXA(12700) = %d, want 20", got2)
	}
}

func TestDXAInches(t *testing.T) {
	d := DXA(1440)
	if !almostEqual(d.Inches(), 1.0, epsilon) {
		t.Errorf("DXA(1440).Inches() = %f, want 1.0", d.Inches())
	}

	d2 := DXA(12240)
	if !almostEqual(d2.Inches(), 8.5, epsilon) {
		t.Errorf("DXA(12240).Inches() = %f, want 8.5", d2.Inches())
	}
}

func TestDXACm(t *testing.T) {
	d := DXA(1440) // 1 inch
	if !almostEqual(d.Cm(), 2.54, epsilon) {
		t.Errorf("DXA(1440).Cm() = %f, want 2.54", d.Cm())
	}
}

func TestDXAMm(t *testing.T) {
	d := DXA(1440) // 1 inch
	if !almostEqual(d.Mm(), 25.4, epsilon) {
		t.Errorf("DXA(1440).Mm() = %f, want 25.4", d.Mm())
	}
}

func TestDXAPt(t *testing.T) {
	d := DXA(20)
	if !almostEqual(d.Pt(), 1.0, epsilon) {
		t.Errorf("DXA(20).Pt() = %f, want 1.0", d.Pt())
	}

	d2 := DXA(1440)
	if !almostEqual(d2.Pt(), 72.0, epsilon) {
		t.Errorf("DXA(1440).Pt() = %f, want 72.0", d2.Pt())
	}
}

func TestEMUInches(t *testing.T) {
	e := EMU(914400)
	if !almostEqual(e.Inches(), 1.0, epsilon) {
		t.Errorf("EMU(914400).Inches() = %f, want 1.0", e.Inches())
	}
}

func TestEMUCm(t *testing.T) {
	e := EMU(360000)
	if !almostEqual(e.Cm(), 1.0, epsilon) {
		t.Errorf("EMU(360000).Cm() = %f, want 1.0", e.Cm())
	}
}

// ============================================================
// Font size conversions
// ============================================================

func TestPtToHalfPoint(t *testing.T) {
	tests := []struct {
		pt   float64
		want HalfPoint
	}{
		{12.0, 24},
		{10.5, 21},
		{8.0, 16},
		{72.0, 144},
	}
	for _, tc := range tests {
		got := PtToHalfPoint(tc.pt)
		if got != tc.want {
			t.Errorf("PtToHalfPoint(%v) = %d, want %d", tc.pt, got, tc.want)
		}
	}
}

func TestHalfPointToPt(t *testing.T) {
	tests := []struct {
		hp   HalfPoint
		want float64
	}{
		{24, 12.0},
		{21, 10.5},
		{16, 8.0},
	}
	for _, tc := range tests {
		got := HalfPointToPt(tc.hp)
		if !almostEqual(got, tc.want, epsilon) {
			t.Errorf("HalfPointToPt(%d) = %f, want %f", tc.hp, got, tc.want)
		}
	}
}

func TestPtToEighthPoint(t *testing.T) {
	tests := []struct {
		pt   float64
		want EighthPoint
	}{
		{0.5, 4}, // common border width
		{1.0, 8},
		{0.25, 2},
		{72.0, 576}, // 1 inch
	}
	for _, tc := range tests {
		got := PtToEighthPoint(tc.pt)
		if got != tc.want {
			t.Errorf("PtToEighthPoint(%v) = %d, want %d", tc.pt, got, tc.want)
		}
	}
}

// ============================================================
// ParseUniversalMeasure
// ============================================================

func TestParseUniversalMeasure(t *testing.T) {
	tests := []struct {
		input string
		want  DXA
	}{
		{"2.54cm", 1440}, // 1 inch
		{"1in", 1440},
		{"25.4mm", 1440},
		{"72pt", 1440}, // 72pt = 1 inch
		{"1pt", 20},
		{"1pc", 240}, // 1 pica = 12pt = 240 DXA
		{"0.5in", 720},
	}
	for _, tc := range tests {
		got, err := ParseUniversalMeasure(tc.input)
		if err != nil {
			t.Errorf("ParseUniversalMeasure(%q) error: %v", tc.input, err)
			continue
		}
		if got != tc.want {
			t.Errorf("ParseUniversalMeasure(%q) = %d, want %d", tc.input, got, tc.want)
		}
	}
}

func TestParseUniversalMeasureErrors(t *testing.T) {
	badInputs := []string{
		"",
		"cm",
		"abc",
		"12xyz",
		"12",
	}
	for _, s := range badInputs {
		_, err := ParseUniversalMeasure(s)
		if err == nil {
			t.Errorf("ParseUniversalMeasure(%q) expected error, got nil", s)
		}
	}
}

// ============================================================
// Constants
// ============================================================

func TestPageConstants(t *testing.T) {
	// Letter: 8.5 x 11 inches
	if LetterW != 12240 {
		t.Errorf("LetterW = %d, want 12240", LetterW)
	}
	if LetterH != 15840 {
		t.Errorf("LetterH = %d, want 15840", LetterH)
	}

	// A4: 210 x 297 mm
	if A4W != 11906 {
		t.Errorf("A4W = %d, want 11906", A4W)
	}
	if A4H != 16838 {
		t.Errorf("A4H = %d, want 16838", A4H)
	}
}

// ============================================================
// Round-trip: convert → reverse → compare
// ============================================================

func TestDXARoundTrip(t *testing.T) {
	// inch → DXA → inch
	original := 8.5
	dxa := InchToDXA(original)
	back := dxa.Inches()
	if !almostEqual(back, original, 0.001) {
		t.Errorf("inch round-trip: %f → %d → %f", original, dxa, back)
	}

	// cm → DXA → cm
	originalCm := 21.0
	dxaCm := CmToDXA(originalCm)
	backCm := dxaCm.Cm()
	if !almostEqual(backCm, originalCm, 0.01) {
		t.Errorf("cm round-trip: %f → %d → %f", originalCm, dxaCm, backCm)
	}

	// mm → DXA → mm
	originalMm := 297.0
	dxaMm := MmToDXA(originalMm)
	backMm := dxaMm.Mm()
	if !almostEqual(backMm, originalMm, 0.1) {
		t.Errorf("mm round-trip: %f → %d → %f", originalMm, dxaMm, backMm)
	}

	// pt → DXA → pt
	originalPt := 12.0
	dxaPt := PtToDXA(originalPt)
	backPt := dxaPt.Pt()
	if !almostEqual(backPt, originalPt, epsilon) {
		t.Errorf("pt round-trip: %f → %d → %f", originalPt, dxaPt, backPt)
	}
}

func TestEMURoundTrip(t *testing.T) {
	// DXA → EMU → DXA
	original := DXA(1440)
	emu := DXAToEMU(original)
	back := EMUToDXA(emu)
	if back != original {
		t.Errorf("DXA→EMU→DXA: %d → %d → %d", original, emu, back)
	}

	// inch → EMU → inch
	originalIn := 2.5
	emuIn := InchToEMU(originalIn)
	backIn := emuIn.Inches()
	if !almostEqual(backIn, originalIn, epsilon) {
		t.Errorf("inch→EMU→inch: %f → %d → %f", originalIn, emuIn, backIn)
	}
}

func TestFontSizeRoundTrip(t *testing.T) {
	// pt → half-point → pt
	originalPt := 10.5
	hp := PtToHalfPoint(originalPt)
	backPt := HalfPointToPt(hp)
	if !almostEqual(backPt, originalPt, epsilon) {
		t.Errorf("pt→hp→pt: %f → %d → %f", originalPt, hp, backPt)
	}
}
