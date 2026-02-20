package xmltypes

// ============================================================
// Simple wrapper types — each wraps a single val attribute
// ============================================================

// CT_String wraps a single string attribute.
//
//	XML: <w:pStyle w:val="Heading1"/>
type CT_String struct {
	Val string `xml:"val,attr"`
}

// CT_DecimalNumber wraps a signed integer attribute.
//
//	XML: <w:outlineLvl w:val="0"/>
type CT_DecimalNumber struct {
	Val int `xml:"val,attr"`
}

// CT_UnsignedDecimalNumber wraps an unsigned integer attribute.
type CT_UnsignedDecimalNumber struct {
	Val uint `xml:"val,attr"`
}

// CT_TwipsMeasure wraps a DXA (twips) measurement.
//
//	XML: <w:before w:val="240"/>
type CT_TwipsMeasure struct {
	Val int `xml:"val,attr"`
}

// CT_SignedTwipsMeasure wraps a signed DXA measurement (can be negative).
type CT_SignedTwipsMeasure struct {
	Val int `xml:"val,attr"`
}

// CT_HpsMeasure wraps a half-point measurement (font size).
//
//	XML: <w:sz w:val="24"/>  → 12pt
type CT_HpsMeasure struct {
	Val int `xml:"val,attr"`
}

// CT_SignedHpsMeasure wraps a signed half-point measurement.
type CT_SignedHpsMeasure struct {
	Val int `xml:"val,attr"`
}

// CT_TextScale wraps a text scaling percentage.
//
//	XML: <w:w w:val="100"/>  → 100%
type CT_TextScale struct {
	Val int `xml:"val,attr"`
}

// CT_LongHexNumber wraps an 8-character hex string.
//
//	XML: <w:rsidR w:val="00A4B3C2"/>
type CT_LongHexNumber struct {
	Val string `xml:"val,attr"`
}

// CT_ShortHexNumber wraps a 4-character hex string.
type CT_ShortHexNumber struct {
	Val string `xml:"val,attr"`
}

// CT_Guid wraps a GUID string.
//
//	XML: <w:docId w:val="{XXXXXXXX-XXXX-...}"/>
type CT_Guid struct {
	Val string `xml:"val,attr"`
}

// CT_Lang wraps a language tag.
//
//	XML: <w:lang w:val="en-US"/>
type CT_Lang struct {
	Val string `xml:"val,attr"`
}

// CT_Empty represents a self-closing element with no attributes or content.
//
//	XML: <w:noProof/>
type CT_Empty struct{}

// ============================================================
// Complex types — multiple attributes
// ============================================================

// CT_Color represents a color specification with optional theme references.
//
//	XML: <w:color w:val="FF0000" w:themeColor="accent1" w:themeTint="BF"/>
type CT_Color struct {
	Val        string  `xml:"val,attr"`                  // hex "FF0000" or "auto"
	ThemeColor *string `xml:"themeColor,attr,omitempty"`  // "accent1", "dark1", etc.
	ThemeTint  *string `xml:"themeTint,attr,omitempty"`   // hex "BF"
	ThemeShade *string `xml:"themeShade,attr,omitempty"`  // hex "80"
}

// CT_Underline represents an underline style with optional color.
//
//	XML: <w:u w:val="single" w:color="FF0000"/>
type CT_Underline struct {
	Val        *string `xml:"val,attr,omitempty"`         // "single", "double", etc.
	Color      *string `xml:"color,attr,omitempty"`       // hex color
	ThemeColor *string `xml:"themeColor,attr,omitempty"`
}

// CT_Highlight represents a text highlight color.
//
//	XML: <w:highlight w:val="yellow"/>
type CT_Highlight struct {
	Val string `xml:"val,attr"` // "yellow", "green", etc.
}

// CT_Border represents a single border line.
//
//	XML: <w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>
type CT_Border struct {
	Val        string  `xml:"val,attr"`                   // "single", "double", "none", etc.
	Sz         *int    `xml:"sz,attr,omitempty"`           // eighth-points
	Space      *int    `xml:"space,attr,omitempty"`        // points
	Color      *string `xml:"color,attr,omitempty"`        // hex color
	ThemeColor *string `xml:"themeColor,attr,omitempty"`
	ThemeTint  *string `xml:"themeTint,attr,omitempty"`
	ThemeShade *string `xml:"themeShade,attr,omitempty"`
	Shadow     *bool   `xml:"shadow,attr,omitempty"`
	Frame      *bool   `xml:"frame,attr,omitempty"`
}

// CT_Shd represents paragraph/cell shading.
//
//	XML: <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
type CT_Shd struct {
	Val            string  `xml:"val,attr"`                    // "clear", "solid", "pct10", etc.
	Color          *string `xml:"color,attr,omitempty"`
	Fill           *string `xml:"fill,attr,omitempty"`
	ThemeColor     *string `xml:"themeColor,attr,omitempty"`
	ThemeFill      *string `xml:"themeFill,attr,omitempty"`
	ThemeFillTint  *string `xml:"themeFillTint,attr,omitempty"`
	ThemeFillShade *string `xml:"themeFillShade,attr,omitempty"`
}

// CT_Fonts specifies font families for different Unicode ranges.
//
//	XML: <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
type CT_Fonts struct {
	Ascii         *string `xml:"ascii,attr,omitempty"`
	HAnsi         *string `xml:"hAnsi,attr,omitempty"`
	EastAsia      *string `xml:"eastAsia,attr,omitempty"`
	CS            *string `xml:"cs,attr,omitempty"`
	AsciiTheme    *string `xml:"asciiTheme,attr,omitempty"`
	HAnsiTheme    *string `xml:"hAnsiTheme,attr,omitempty"`
	EastAsiaTheme *string `xml:"eastAsiaTheme,attr,omitempty"`
	CSTheme       *string `xml:"cstheme,attr,omitempty"`
	Hint          *string `xml:"hint,attr,omitempty"`
}

// CT_Language specifies language settings for proofing.
//
//	XML: <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
type CT_Language struct {
	Val      *string `xml:"val,attr,omitempty"`
	EastAsia *string `xml:"eastAsia,attr,omitempty"`
	Bidi     *string `xml:"bidi,attr,omitempty"`
}
