// Package units provides OOXML measurement types and conversion functions.
//
// OOXML uses several measurement systems:
//   - DXA (twentieths of a point, aka "twips") — page size, margins, indents
//   - EMU (English Metric Units) — images, drawings
//   - HalfPoint — font sizes (w:sz val="24" = 12pt)
//   - EighthPoint — border widths
//
// Conversion cheat sheet:
//
//	1 inch = 1440 DXA = 914400 EMU = 72 pt = 144 half-pt = 576 eighth-pt
//	1 cm   = 567  DXA = 360000 EMU ≈ 28.35 pt
//	1 mm   = 56.7 DXA = 36000  EMU ≈ 2.835 pt
//	1 pt   = 20   DXA = 12700  EMU
package units

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// ============================================================
// Types
// ============================================================

// EMU represents English Metric Units (1 inch = 914400 EMU).
// Used for image/drawing dimensions.
type EMU int64

// DXA represents twentieths of a point (1 inch = 1440 DXA).
// Used for page size, margins, indents, spacing.
type DXA int

// HalfPoint represents half-points (1 pt = 2 half-points).
// Used for font sizes: w:sz val="24" means 12pt.
type HalfPoint int

// EighthPoint represents eighth-points (1 pt = 8 eighth-points).
// Used for border widths: w:val="4" means 0.5pt.
type EighthPoint int

// ============================================================
// Internal constants
// ============================================================

const (
	dxaPerInch = 1440
	emuPerInch = 914400
	emuPerCm   = 360000
	emuPerMm   = 36000
	emuPerPt   = 12700
	emuPerDXA  = 635 // 914400 / 1440
	dxaPerPt   = 20
	cmPerInch  = 2.54
	mmPerInch  = 25.4
	ptPerInch  = 72.0
)

// ============================================================
// Page size constants (DXA)
// ============================================================

const (
	LetterW DXA = 12240 // 8.5 inch
	LetterH DXA = 15840 // 11 inch
	A4W     DXA = 11906 // 210 mm
	A4H     DXA = 16838 // 297 mm
)

// ============================================================
// Conversions → DXA
// ============================================================

// InchToDXA converts inches to DXA.
func InchToDXA(in float64) DXA {
	return DXA(math.Round(in * dxaPerInch))
}

// CmToDXA converts centimeters to DXA.
func CmToDXA(cm float64) DXA {
	return DXA(math.Round(cm * dxaPerInch / cmPerInch))
}

// MmToDXA converts millimeters to DXA.
func MmToDXA(mm float64) DXA {
	return DXA(math.Round(mm * dxaPerInch / mmPerInch))
}

// PtToDXA converts points to DXA.
func PtToDXA(pt float64) DXA {
	return DXA(math.Round(pt * dxaPerPt))
}

// ============================================================
// Conversions → EMU
// ============================================================

// InchToEMU converts inches to EMU.
func InchToEMU(in float64) EMU {
	return EMU(math.Round(in * emuPerInch))
}

// CmToEMU converts centimeters to EMU.
func CmToEMU(cm float64) EMU {
	return EMU(math.Round(cm * emuPerCm))
}

// MmToEMU converts millimeters to EMU.
func MmToEMU(mm float64) EMU {
	return EMU(math.Round(mm * emuPerMm))
}

// PtToEMU converts points to EMU.
func PtToEMU(pt float64) EMU {
	return EMU(math.Round(pt * emuPerPt))
}

// DXAToEMU converts DXA to EMU.
func DXAToEMU(d DXA) EMU {
	return EMU(d) * emuPerDXA
}

// ============================================================
// Reverse conversions
// ============================================================

// EMUToDXA converts EMU to DXA.
func EMUToDXA(e EMU) DXA {
	return DXA(math.Round(float64(e) / emuPerDXA))
}

// Inches returns the DXA value in inches.
func (d DXA) Inches() float64 {
	return float64(d) / dxaPerInch
}

// Cm returns the DXA value in centimeters.
func (d DXA) Cm() float64 {
	return float64(d) / dxaPerInch * cmPerInch
}

// Mm returns the DXA value in millimeters.
func (d DXA) Mm() float64 {
	return float64(d) / dxaPerInch * mmPerInch
}

// Pt returns the DXA value in points.
func (d DXA) Pt() float64 {
	return float64(d) / dxaPerPt
}

// Inches returns the EMU value in inches.
func (e EMU) Inches() float64 {
	return float64(e) / emuPerInch
}

// Cm returns the EMU value in centimeters.
func (e EMU) Cm() float64 {
	return float64(e) / emuPerCm
}

// ============================================================
// Font size conversions
// ============================================================

// PtToHalfPoint converts points to half-points (for w:sz).
func PtToHalfPoint(pt float64) HalfPoint {
	return HalfPoint(math.Round(pt * 2))
}

// HalfPointToPt converts half-points back to points.
func HalfPointToPt(hp HalfPoint) float64 {
	return float64(hp) / 2.0
}

// PtToEighthPoint converts points to eighth-points (for border widths).
func PtToEighthPoint(pt float64) EighthPoint {
	return EighthPoint(math.Round(pt * 8))
}

// ============================================================
// Parsing XML string values
// ============================================================

// ParseUniversalMeasure parses an xsd:ST_UniversalMeasure string
// (e.g. "2.54cm", "1in", "72pt", "25.4mm") and returns the value in DXA.
func ParseUniversalMeasure(s string) (DXA, error) {
	s = strings.TrimSpace(s)
	if len(s) == 0 {
		return 0, fmt.Errorf("units: empty measure string")
	}

	// Find where the numeric part ends and the unit suffix begins.
	i := len(s)
	for i > 0 && !isDigitOrDot(s[i-1]) {
		i--
	}
	if i == 0 {
		return 0, fmt.Errorf("units: no numeric value in %q", s)
	}

	numStr := s[:i]
	unit := strings.ToLower(s[i:])

	val, err := strconv.ParseFloat(numStr, 64)
	if err != nil {
		return 0, fmt.Errorf("units: invalid number in %q: %w", s, err)
	}

	switch unit {
	case "in":
		return InchToDXA(val), nil
	case "cm":
		return CmToDXA(val), nil
	case "mm":
		return MmToDXA(val), nil
	case "pt":
		return PtToDXA(val), nil
	case "pc": // pica = 12pt
		return PtToDXA(val * 12), nil
	case "pi": // alternative pica notation
		return PtToDXA(val * 12), nil
	default:
		return 0, fmt.Errorf("units: unknown unit %q in %q", unit, s)
	}
}

func isDigitOrDot(c byte) bool {
	return (c >= '0' && c <= '9') || c == '.' || c == '-' || c == '+'
}
