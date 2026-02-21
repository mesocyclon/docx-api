package oxml

import (
	"github.com/user/go-docx/pkg/docx/enum"
)

// --- CT_RPr custom methods ---

// getBoolVal reads a tri-state boolean value from a CT_OnOff child element.
//   - nil: element not present (inherit)
//   - *true: element present with val=true or val absent (e.g. <w:b/>)
//   - *false: element present with val=false (e.g. <w:b w:val="false"/>)
func (rPr *CT_RPr) getBoolVal(tag string) *bool {
	child := rPr.FindChild(tag)
	if child == nil {
		return nil
	}
	onOff := &CT_OnOff{Element{E: child}}
	v := onOff.Val()
	return &v
}

// setBoolVal sets a tri-state boolean for a CT_OnOff child element.
//   - nil: remove the element
//   - *true: add element without val attr (e.g. <w:b/>)
//   - *false: add element with val="false" (e.g. <w:b w:val="false"/>)
//
// getOrAdd and remove are function params to handle the correct child tag.
func (rPr *CT_RPr) setBoolValWith(val *bool, getOrAdd func() *CT_OnOff, remove func()) {
	if val == nil {
		remove()
		return
	}
	el := getOrAdd()
	el.SetVal(*val)
}

// --- Bold ---

// BoldVal returns the tri-state bold value: nil (inherit), *true, or *false.
func (rPr *CT_RPr) BoldVal() *bool {
	return rPr.getBoolVal("w:b")
}

// SetBoldVal sets the bold tri-state.
func (rPr *CT_RPr) SetBoldVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddB, rPr.RemoveB)
}

// --- Italic ---

// ItalicVal returns the tri-state italic value.
func (rPr *CT_RPr) ItalicVal() *bool {
	return rPr.getBoolVal("w:i")
}

// SetItalicVal sets the italic tri-state.
func (rPr *CT_RPr) SetItalicVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddI, rPr.RemoveI)
}

// --- Caps ---

// CapsVal returns the tri-state all-caps value.
func (rPr *CT_RPr) CapsVal() *bool {
	return rPr.getBoolVal("w:caps")
}

// SetCapsVal sets the all-caps tri-state.
func (rPr *CT_RPr) SetCapsVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddCaps, rPr.RemoveCaps)
}

// --- SmallCaps ---

// SmallCapsVal returns the tri-state small-caps value.
func (rPr *CT_RPr) SmallCapsVal() *bool {
	return rPr.getBoolVal("w:smallCaps")
}

// SetSmallCapsVal sets the small-caps tri-state.
func (rPr *CT_RPr) SetSmallCapsVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddSmallCaps, rPr.RemoveSmallCaps)
}

// --- Strike ---

// StrikeVal returns the tri-state strikethrough value.
func (rPr *CT_RPr) StrikeVal() *bool {
	return rPr.getBoolVal("w:strike")
}

// SetStrikeVal sets the strikethrough tri-state.
func (rPr *CT_RPr) SetStrikeVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddStrike, rPr.RemoveStrike)
}

// --- Dstrike (double strikethrough) ---

// DstrikeVal returns the tri-state double-strikethrough value.
func (rPr *CT_RPr) DstrikeVal() *bool {
	return rPr.getBoolVal("w:dstrike")
}

// SetDstrikeVal sets the double-strikethrough tri-state.
func (rPr *CT_RPr) SetDstrikeVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddDstrike, rPr.RemoveDstrike)
}

// --- Outline ---

// OutlineVal returns the tri-state outline value.
func (rPr *CT_RPr) OutlineVal() *bool {
	return rPr.getBoolVal("w:outline")
}

// SetOutlineVal sets the outline tri-state.
func (rPr *CT_RPr) SetOutlineVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddOutline, rPr.RemoveOutline)
}

// --- Shadow ---

// ShadowVal returns the tri-state shadow value.
func (rPr *CT_RPr) ShadowVal() *bool {
	return rPr.getBoolVal("w:shadow")
}

// SetShadowVal sets the shadow tri-state.
func (rPr *CT_RPr) SetShadowVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddShadow, rPr.RemoveShadow)
}

// --- Emboss ---

// EmbossVal returns the tri-state emboss value.
func (rPr *CT_RPr) EmbossVal() *bool {
	return rPr.getBoolVal("w:emboss")
}

// SetEmbossVal sets the emboss tri-state.
func (rPr *CT_RPr) SetEmbossVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddEmboss, rPr.RemoveEmboss)
}

// --- Imprint ---

// ImprintVal returns the tri-state imprint value.
func (rPr *CT_RPr) ImprintVal() *bool {
	return rPr.getBoolVal("w:imprint")
}

// SetImprintVal sets the imprint tri-state.
func (rPr *CT_RPr) SetImprintVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddImprint, rPr.RemoveImprint)
}

// --- NoProof ---

// NoProofVal returns the tri-state noProof value.
func (rPr *CT_RPr) NoProofVal() *bool {
	return rPr.getBoolVal("w:noProof")
}

// SetNoProofVal sets the noProof tri-state.
func (rPr *CT_RPr) SetNoProofVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddNoProof, rPr.RemoveNoProof)
}

// --- SnapToGrid ---

// SnapToGridVal returns the tri-state snapToGrid value.
func (rPr *CT_RPr) SnapToGridVal() *bool {
	return rPr.getBoolVal("w:snapToGrid")
}

// SetSnapToGridVal sets the snapToGrid tri-state.
func (rPr *CT_RPr) SetSnapToGridVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddSnapToGrid, rPr.RemoveSnapToGrid)
}

// --- Vanish ---

// VanishVal returns the tri-state vanish value.
func (rPr *CT_RPr) VanishVal() *bool {
	return rPr.getBoolVal("w:vanish")
}

// SetVanishVal sets the vanish tri-state.
func (rPr *CT_RPr) SetVanishVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddVanish, rPr.RemoveVanish)
}

// --- WebHidden ---

// WebHiddenVal returns the tri-state webHidden value.
func (rPr *CT_RPr) WebHiddenVal() *bool {
	return rPr.getBoolVal("w:webHidden")
}

// SetWebHiddenVal sets the webHidden tri-state.
func (rPr *CT_RPr) SetWebHiddenVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddWebHidden, rPr.RemoveWebHidden)
}

// --- SpecVanish ---

// SpecVanishVal returns the tri-state specVanish value.
func (rPr *CT_RPr) SpecVanishVal() *bool {
	return rPr.getBoolVal("w:specVanish")
}

// SetSpecVanishVal sets the specVanish tri-state.
func (rPr *CT_RPr) SetSpecVanishVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddSpecVanish, rPr.RemoveSpecVanish)
}

// --- OMath ---

// OMathVal returns the tri-state oMath value.
func (rPr *CT_RPr) OMathVal() *bool {
	return rPr.getBoolVal("w:oMath")
}

// SetOMathVal sets the oMath tri-state.
func (rPr *CT_RPr) SetOMathVal(v *bool) {
	rPr.setBoolValWith(v, rPr.GetOrAddOMath, rPr.RemoveOMath)
}

// --- Color ---

// ColorVal returns the hex color string from w:color/@w:val, or nil if not present.
func (rPr *CT_RPr) ColorVal() *string {
	color := rPr.Color()
	if color == nil {
		return nil
	}
	val, err := color.Val()
	if err != nil {
		return nil
	}
	return &val
}

// SetColorVal sets the color hex value. Passing nil removes the color element.
func (rPr *CT_RPr) SetColorVal(v *string) {
	if v == nil {
		rPr.RemoveColor()
		return
	}
	color := rPr.GetOrAddColor()
	color.SetVal(*v)
}

// ColorTheme returns the theme color from w:color/@w:themeColor, or nil if not present.
func (rPr *CT_RPr) ColorTheme() *enum.MsoThemeColorIndex {
	color := rPr.Color()
	if color == nil {
		return nil
	}
	tc := color.ThemeColor()
	if tc == "" {
		return nil
	}
	v, err := enum.MsoThemeColorIndexFromXml(tc)
	if err != nil {
		return nil
	}
	return &v
}

// SetColorTheme sets the theme color. Passing nil removes the themeColor attribute.
func (rPr *CT_RPr) SetColorTheme(v *enum.MsoThemeColorIndex) {
	if v == nil {
		color := rPr.Color()
		if color != nil {
			color.SetThemeColor("")
		}
		return
	}
	color := rPr.GetOrAddColor()
	xml, err := v.ToXml()
	if err == nil {
		color.SetThemeColor(xml)
	}
}

// --- Size ---

// SzVal returns the font size from w:sz/@w:val as half-points, or nil if not present.
func (rPr *CT_RPr) SzVal() *int64 {
	sz := rPr.Sz()
	if sz == nil {
		return nil
	}
	val, err := sz.Val()
	if err != nil {
		return nil
	}
	return &val
}

// SetSzVal sets the font size in half-points. Passing nil removes the sz element.
func (rPr *CT_RPr) SetSzVal(v *int64) {
	if v == nil {
		rPr.RemoveSz()
		return
	}
	sz := rPr.GetOrAddSz()
	sz.SetVal(*v)
}

// --- Fonts ---

// RFontsAscii returns the ascii font name, or nil if not present.
func (rPr *CT_RPr) RFontsAscii() *string {
	rFonts := rPr.RFonts()
	if rFonts == nil {
		return nil
	}
	v := rFonts.Ascii()
	if v == "" {
		return nil
	}
	return &v
}

// SetRFontsAscii sets the ascii font name. Passing nil removes the rFonts element.
func (rPr *CT_RPr) SetRFontsAscii(v *string) {
	if v == nil {
		rPr.RemoveRFonts()
		return
	}
	rFonts := rPr.GetOrAddRFonts()
	rFonts.SetAscii(*v)
}

// RFontsHAnsi returns the hAnsi font name, or nil if not present.
func (rPr *CT_RPr) RFontsHAnsi() *string {
	rFonts := rPr.RFonts()
	if rFonts == nil {
		return nil
	}
	v := rFonts.HAnsi()
	if v == "" {
		return nil
	}
	return &v
}

// SetRFontsHAnsi sets the hAnsi font name. Passing nil leaves rFonts alone.
func (rPr *CT_RPr) SetRFontsHAnsi(v *string) {
	if v == nil && rPr.RFonts() == nil {
		return
	}
	rFonts := rPr.GetOrAddRFonts()
	if v == nil {
		rFonts.SetHAnsi("")
	} else {
		rFonts.SetHAnsi(*v)
	}
}

// --- Underline ---

// UVal returns the underline style from w:u/@w:val, or nil if not present.
func (rPr *CT_RPr) UVal() *string {
	u := rPr.U()
	if u == nil {
		return nil
	}
	v := u.Val()
	if v == "" {
		return nil
	}
	return &v
}

// SetUVal sets the underline style. Passing nil removes the u element.
func (rPr *CT_RPr) SetUVal(v *string) {
	rPr.RemoveU()
	if v != nil {
		u := rPr.addU()
		u.SetVal(*v)
	}
}

// --- Highlight ---

// HighlightVal returns the highlight color string, or nil if not present.
func (rPr *CT_RPr) HighlightVal() *string {
	h := rPr.Highlight()
	if h == nil {
		return nil
	}
	v, err := h.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetHighlightVal sets the highlight color. Passing nil removes the highlight element.
func (rPr *CT_RPr) SetHighlightVal(v *string) {
	if v == nil {
		rPr.RemoveHighlight()
		return
	}
	h := rPr.GetOrAddHighlight()
	h.SetVal(*v)
}

// --- Style ---

// StyleVal returns the run style string from w:rStyle/@w:val, or nil if not present.
func (rPr *CT_RPr) StyleVal() *string {
	rStyle := rPr.RStyle()
	if rStyle == nil {
		return nil
	}
	v, err := rStyle.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetStyleVal sets the run style. Passing nil removes the rStyle element.
func (rPr *CT_RPr) SetStyleVal(v *string) {
	if v == nil {
		rPr.RemoveRStyle()
		return
	}
	rStyle := rPr.RStyle()
	if rStyle == nil {
		s := rPr.addRStyle()
		s.SetVal(*v)
	} else {
		rStyle.SetVal(*v)
	}
}

// --- Subscript / Superscript ---

// Subscript returns true if vertAlign is "subscript", false if it's something else,
// nil if vertAlign is not present.
func (rPr *CT_RPr) Subscript() *bool {
	va := rPr.VertAlign()
	if va == nil {
		return nil
	}
	v, err := va.Val()
	if err != nil {
		return nil
	}
	result := v == "subscript"
	return &result
}

// SetSubscript sets the subscript state. nil removes vertAlign,
// true sets it to "subscript", false clears only if currently "subscript".
func (rPr *CT_RPr) SetSubscript(v *bool) {
	if v == nil {
		rPr.RemoveVertAlign()
	} else if *v {
		rPr.GetOrAddVertAlign().SetVal("subscript")
	} else {
		va := rPr.VertAlign()
		if va != nil {
			val, _ := va.Val()
			if val == "subscript" {
				rPr.RemoveVertAlign()
			}
		}
	}
}

// Superscript returns true if vertAlign is "superscript", false if it's something else,
// nil if vertAlign is not present.
func (rPr *CT_RPr) Superscript() *bool {
	va := rPr.VertAlign()
	if va == nil {
		return nil
	}
	v, err := va.Val()
	if err != nil {
		return nil
	}
	result := v == "superscript"
	return &result
}

// SetSuperscript sets the superscript state.
func (rPr *CT_RPr) SetSuperscript(v *bool) {
	if v == nil {
		rPr.RemoveVertAlign()
	} else if *v {
		rPr.GetOrAddVertAlign().SetVal("superscript")
	} else {
		va := rPr.VertAlign()
		if va != nil {
			val, _ := va.Val()
			if val == "superscript" {
				rPr.RemoveVertAlign()
			}
		}
	}
}
