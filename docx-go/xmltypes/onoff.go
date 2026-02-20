package xmltypes

import "encoding/xml"

// CT_OnOff represents the WML three-state boolean logic:
//   - nil pointer:      not set (inherit from style)
//   - Val == nil:        element present without val attribute → true (e.g., <w:b/>)
//   - Val == "0"|"false"|"off": explicitly false
//   - Val == "1"|"true"|"on":   explicitly true
type CT_OnOff struct {
	Val *string // nil → element present means true
}

// NewOnOff creates a CT_OnOff value.
//
//	true  → &CT_OnOff{Val: nil}   — most compact form: <w:b/>
//	false → &CT_OnOff{Val: &"0"}  — explicit off: <w:b w:val="0"/>
func NewOnOff(v bool) *CT_OnOff {
	if v {
		return &CT_OnOff{Val: nil}
	}
	s := "0"
	return &CT_OnOff{Val: &s}
}

// Bool returns the boolean interpretation of CT_OnOff.
// If the receiver is nil, defaultVal is returned (style inheritance).
func (o *CT_OnOff) Bool(defaultVal bool) bool {
	if o == nil {
		return defaultVal
	}
	if o.Val == nil {
		return true // <w:b/> without val = true
	}
	switch *o.Val {
	case "true", "1", "on":
		return true
	case "false", "0", "off":
		return false
	default:
		return true // unknown value → true (Word behavior)
	}
}

// IsExplicitlySet returns true if the CT_OnOff is non-nil,
// meaning the element was present in the XML (regardless of value).
func (o *CT_OnOff) IsExplicitlySet() bool {
	return o != nil
}

// MarshalXML implements xml.Marshaler. Emits a self-closing element,
// optionally with a w:val attribute.
func (o *CT_OnOff) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if o.Val != nil {
		start.Attr = append(start.Attr, xml.Attr{
			Name:  xml.Name{Space: NSw, Local: "val"},
			Value: *o.Val,
		})
	}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML implements xml.Unmarshaler. Reads the val attribute
// if present and skips any child content.
func (o *CT_OnOff) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			s := attr.Value
			o.Val = &s
		}
	}
	// Skip any child content (usually empty for CT_OnOff)
	return d.Skip()
}
