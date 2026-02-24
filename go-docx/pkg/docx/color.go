package docx

import (
	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// ColorFormat provides access to color settings such as RGB and theme color.
//
// Mirrors Python ColorFormat(ElementProxy).
type ColorFormat struct {
	r *oxml.CT_R // rPr parent (CT_R or CT_Style wrapping rPr)
}

// NewColorFormat creates a new ColorFormat proxy.
func NewColorFormat(r *oxml.CT_R) *ColorFormat {
	return &ColorFormat{r: r}
}

// RGB returns the RGB color value, or nil if not set or auto.
//
// Mirrors Python ColorFormat.rgb (getter).
func (cf *ColorFormat) RGB() *RGBColor {
	rPr := cf.r.RPr()
	if rPr == nil {
		return nil
	}
	val, err := rPr.ColorVal()
	if err != nil {
		return nil
	}
	if val == nil || *val == "auto" {
		return nil
	}
	c, err := RGBColorFromString(*val)
	if err != nil {
		return nil
	}
	return &c
}

// SetRGB sets the RGB color. Passing nil removes the color.
//
// Mirrors Python ColorFormat.rgb (setter).
func (cf *ColorFormat) SetRGB(v *RGBColor) error {
	if v == nil {
		rPr := cf.r.RPr()
		if rPr == nil {
			return nil
		}
		return rPr.SetColorVal(nil)
	}
	rPr := cf.r.GetOrAddRPr()
	hex := v.String()
	return rPr.SetColorVal(&hex)
}

// ThemeColor returns the theme color index, or nil if not set.
//
// Mirrors Python ColorFormat.theme_color (getter).
func (cf *ColorFormat) ThemeColor() (*enum.MsoThemeColorIndex, error) {
	rPr := cf.r.RPr()
	if rPr == nil {
		return nil, nil
	}
	return rPr.ColorTheme()
}

// SetThemeColor sets the theme color. Passing nil removes the color entirely.
//
// Mirrors Python ColorFormat.theme_color (setter).
func (cf *ColorFormat) SetThemeColor(v *enum.MsoThemeColorIndex) error {
	if v == nil {
		rPr := cf.r.RPr()
		if rPr == nil {
			return nil
		}
		return rPr.SetColorVal(nil)
	}
	rPr := cf.r.GetOrAddRPr()
	return rPr.SetColorTheme(v)
}

// Type returns the color type: RGB, THEME, AUTO, or nil if no color is applied.
//
// Mirrors Python ColorFormat.type (getter).
func (cf *ColorFormat) Type() *enum.MsoColorType {
	rPr := cf.r.RPr()
	if rPr == nil {
		return nil
	}
	theme, _ := rPr.ColorTheme()
	if theme != nil {
		ct := enum.MsoColorTypeTheme
		return &ct
	}
	val, err := rPr.ColorVal()
	if err != nil || val == nil {
		return nil
	}
	if *val == "auto" {
		ct := enum.MsoColorTypeAuto
		return &ct
	}
	ct := enum.MsoColorTypeRGB
	return &ct
}
