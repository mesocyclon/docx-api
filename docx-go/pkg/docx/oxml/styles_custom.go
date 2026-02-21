package oxml

import (
	"strings"

	"github.com/user/go-docx/pkg/docx/enum"
)

// ===========================================================================
// styleIdFromName utility
// ===========================================================================

// StyleIdFromName returns the style ID corresponding to a style name.
// Special-case names like "Heading 1" map to "Heading1", etc.
// Default behaviour: remove spaces.
func StyleIdFromName(name string) string {
	special := map[string]string{
		"caption":   "Caption",
		"heading 1": "Heading1",
		"heading 2": "Heading2",
		"heading 3": "Heading3",
		"heading 4": "Heading4",
		"heading 5": "Heading5",
		"heading 6": "Heading6",
		"heading 7": "Heading7",
		"heading 8": "Heading8",
		"heading 9": "Heading9",
	}
	lower := strings.ToLower(name)
	if v, ok := special[lower]; ok {
		return v
	}
	return strings.ReplaceAll(name, " ", "")
}

// ===========================================================================
// CT_Styles — custom methods
// ===========================================================================

// GetByID returns the w:style element whose @w:styleId matches styleID, or nil.
func (ss *CT_Styles) GetByID(styleID string) *CT_Style {
	for _, s := range ss.StyleList() {
		if s.StyleId() == styleID {
			return s
		}
	}
	return nil
}

// GetByName returns the w:style element whose w:name/@w:val matches name, or nil.
func (ss *CT_Styles) GetByName(name string) *CT_Style {
	for _, s := range ss.StyleList() {
		if s.NameVal() == name {
			return s
		}
	}
	return nil
}

// DefaultFor returns the default style for the given type, or nil.
// If multiple defaults exist, returns the last one (per OOXML spec).
func (ss *CT_Styles) DefaultFor(styleType enum.WdStyleType) *CT_Style {
	xmlType := styleType.ToXml()
	var last *CT_Style
	for _, s := range ss.StyleList() {
		if s.Type() == xmlType && s.Default() {
			last = s
		}
	}
	return last
}

// AddStyleOfType creates and adds a new w:style element with the given name, type,
// and builtin flag. Returns the new style element.
func (ss *CT_Styles) AddStyleOfType(name string, styleType enum.WdStyleType, builtin bool) *CT_Style {
	style := ss.AddStyle()
	style.SetType(styleType.ToXml())
	if !builtin {
		style.SetCustomStyle(true)
	}
	style.SetStyleId(StyleIdFromName(name))
	style.SetNameVal(name)
	return style
}

// ===========================================================================
// CT_Style — custom methods
// ===========================================================================

// NameVal returns the value of w:name/@w:val, or "" if not present.
func (s *CT_Style) NameVal() string {
	n := s.Name()
	if n == nil {
		return ""
	}
	v, err := n.Val()
	if err != nil {
		return ""
	}
	return v
}

// SetNameVal sets the w:name/@w:val. Passing "" removes the name element.
func (s *CT_Style) SetNameVal(name string) {
	s.RemoveName()
	if name == "" {
		return
	}
	s.GetOrAddName().SetVal(name)
}

// BasedOnVal returns the value of w:basedOn/@w:val, or "" if not present.
func (s *CT_Style) BasedOnVal() string {
	b := s.BasedOn()
	if b == nil {
		return ""
	}
	v, err := b.Val()
	if err != nil {
		return ""
	}
	return v
}

// SetBasedOnVal sets the basedOn value. Passing "" removes the element.
func (s *CT_Style) SetBasedOnVal(v string) {
	s.RemoveBasedOn()
	if v == "" {
		return
	}
	s.GetOrAddBasedOn().SetVal(v)
}

// NextVal returns the value of w:next/@w:val, or "" if not present.
func (s *CT_Style) NextVal() string {
	n := s.Next()
	if n == nil {
		return ""
	}
	v, err := n.Val()
	if err != nil {
		return ""
	}
	return v
}

// SetNextVal sets the next style ID. Passing "" removes the element.
func (s *CT_Style) SetNextVal(v string) {
	s.RemoveNext()
	if v == "" {
		return
	}
	s.GetOrAddNext().SetVal(v)
}

// LockedVal returns the value of w:locked, or false if not present.
func (s *CT_Style) LockedVal() bool {
	l := s.Locked()
	if l == nil {
		return false
	}
	return l.Val()
}

// SetLockedVal sets the locked flag. Passing false removes the element.
func (s *CT_Style) SetLockedVal(v bool) {
	s.RemoveLocked()
	if v {
		s.GetOrAddLocked().SetVal(true)
	}
}

// SemiHiddenVal returns the value of w:semiHidden, or false if not present.
func (s *CT_Style) SemiHiddenVal() bool {
	sh := s.SemiHidden()
	if sh == nil {
		return false
	}
	return sh.Val()
}

// SetSemiHiddenVal sets the semiHidden flag.
func (s *CT_Style) SetSemiHiddenVal(v bool) {
	s.RemoveSemiHidden()
	if v {
		s.GetOrAddSemiHidden().SetVal(true)
	}
}

// UnhideWhenUsedVal returns the value of w:unhideWhenUsed, or false.
func (s *CT_Style) UnhideWhenUsedVal() bool {
	u := s.UnhideWhenUsed()
	if u == nil {
		return false
	}
	return u.Val()
}

// SetUnhideWhenUsedVal sets the unhideWhenUsed flag.
func (s *CT_Style) SetUnhideWhenUsedVal(v bool) {
	s.RemoveUnhideWhenUsed()
	if v {
		s.GetOrAddUnhideWhenUsed().SetVal(true)
	}
}

// QFormatVal returns the value of w:qFormat, or false.
func (s *CT_Style) QFormatVal() bool {
	q := s.QFormat()
	if q == nil {
		return false
	}
	return q.Val()
}

// SetQFormatVal sets the qFormat flag.
func (s *CT_Style) SetQFormatVal(v bool) {
	s.RemoveQFormat()
	if v {
		s.GetOrAddQFormat()
	}
}

// UiPriorityVal returns the value of w:uiPriority/@w:val, or nil.
func (s *CT_Style) UiPriorityVal() *int {
	u := s.UiPriority()
	if u == nil {
		return nil
	}
	v, err := u.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetUiPriorityVal sets the uiPriority. Passing nil removes the element.
func (s *CT_Style) SetUiPriorityVal(v *int) {
	s.RemoveUiPriority()
	if v == nil {
		return
	}
	s.GetOrAddUiPriority().SetVal(*v)
}

// BaseStyle returns the sibling CT_Style that this style is based on, or nil.
func (s *CT_Style) BaseStyle() *CT_Style {
	basedOn := s.BasedOnVal()
	if basedOn == "" {
		return nil
	}
	parent := s.E.Parent()
	if parent == nil {
		return nil
	}
	styles := &CT_Styles{Element{E: parent}}
	return styles.GetByID(basedOn)
}

// NextStyle returns the sibling CT_Style identified by w:next, or nil.
func (s *CT_Style) NextStyle() *CT_Style {
	nextVal := s.NextVal()
	if nextVal == "" {
		return nil
	}
	parent := s.E.Parent()
	if parent == nil {
		return nil
	}
	styles := &CT_Styles{Element{E: parent}}
	return styles.GetByID(nextVal)
}

// Delete removes this w:style element from its parent w:styles.
func (s *CT_Style) Delete() {
	parent := s.E.Parent()
	if parent != nil {
		parent.RemoveChild(s.E)
	}
}

// IsBuiltin returns true if this is a built-in style (customStyle is not set or false).
func (s *CT_Style) IsBuiltin() bool {
	return !s.CustomStyle()
}

// ===========================================================================
// CT_LatentStyles — custom methods
// ===========================================================================

// GetByName returns the lsdException child with the given name, or nil.
func (ls *CT_LatentStyles) GetByName(name string) *CT_LsdException {
	for _, exc := range ls.LsdExceptionList() {
		n, err := exc.Name()
		if err == nil && n == name {
			return exc
		}
	}
	return nil
}

// BoolProp returns the boolean value of the named attribute, or false if absent.
func (ls *CT_LatentStyles) BoolProp(attrName string) bool {
	val, ok := ls.GetAttr(attrName)
	if !ok {
		return false
	}
	return parseBoolAttr(val)
}

// SetBoolProp sets the named on/off attribute.
func (ls *CT_LatentStyles) SetBoolProp(attrName string, val bool) {
	ls.SetAttr(attrName, formatBoolAttr(val))
}

// ===========================================================================
// CT_LsdException — custom methods
// ===========================================================================

// Delete removes this lsdException element from its parent.
func (exc *CT_LsdException) Delete() {
	parent := exc.E.Parent()
	if parent != nil {
		parent.RemoveChild(exc.E)
	}
}

// OnOffProp returns the boolean value of the named attribute, or nil if absent.
func (exc *CT_LsdException) OnOffProp(attrName string) *bool {
	val, ok := exc.GetAttr(attrName)
	if !ok {
		return nil
	}
	v := parseBoolAttr(val)
	return &v
}

// SetOnOffProp sets the named on/off attribute. Passing nil removes the attribute.
func (exc *CT_LsdException) SetOnOffProp(attrName string, val *bool) {
	if val == nil {
		exc.RemoveAttr(attrName)
		return
	}
	exc.SetAttr(attrName, formatBoolAttr(*val))
}
