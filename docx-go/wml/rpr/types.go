// Package rpr implements WML run properties (CT_RPr, CT_RPrBase, CT_ParaRPr).
// Contract: C-10 in contracts.md.
// Dependencies: xmltypes, wml/shared.
package rpr

// CT_TextEffect represents the w:effect element.
type CT_TextEffect struct {
	Val string `xml:"val,attr"` // "blinkBackground", "lights", etc.
}

// CT_FitText represents the w:fitText element.
type CT_FitText struct {
	Val int  `xml:"val,attr"`            // DXA
	ID  *int `xml:"id,attr,omitempty"`
}

// CT_VerticalAlignRun represents the w:vertAlign element.
type CT_VerticalAlignRun struct {
	Val string `xml:"val,attr"` // "baseline", "superscript", "subscript"
}

// CT_Em represents the w:em element (East Asian emphasis marks).
type CT_Em struct {
	Val string `xml:"val,attr"` // "none", "dot", "comma", "circle", "underDot"
}

// CT_EastAsianLayout represents the w:eastAsianLayout element.
type CT_EastAsianLayout struct {
	ID              *int    `xml:"id,attr,omitempty"`
	Combine         *bool   `xml:"combine,attr,omitempty"`
	CombineBrackets *string `xml:"combineBrackets,attr,omitempty"`
	Vert            *bool   `xml:"vert,attr,omitempty"`
	VertCompress    *bool   `xml:"vertCompress,attr,omitempty"`
}

// CT_TrackChangeRef represents an ins/del tracking reference within CT_ParaRPr.
type CT_TrackChangeRef struct {
	ID     int    `xml:"id,attr"`
	Author string `xml:"author,attr"`
	Date   string `xml:"date,attr,omitempty"`
}
