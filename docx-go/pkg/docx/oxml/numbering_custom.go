package oxml

import (
	"sort"
)

// ===========================================================================
// CT_Numbering — custom methods
// ===========================================================================

// AddNumWithAbstractNumId adds a new <w:num> referencing the given abstract
// numbering definition id. The new num is assigned the next available numId.
// Returns the newly created CT_Num.
func (n *CT_Numbering) AddNumWithAbstractNumId(abstractNumId int) *CT_Num {
	nextNumId := n.NextNumId()
	num := NewNum(nextNumId, abstractNumId)
	n.insertNum(num)
	return num
}

// NumHavingNumId returns the <w:num> child with the given numId attribute,
// or nil if not found.
func (n *CT_Numbering) NumHavingNumId(numId int) *CT_Num {
	for _, num := range n.NumList() {
		id, err := num.NumId()
		if err == nil && id == numId {
			return num
		}
	}
	return nil
}

// NextNumId returns the first numId not used by any <w:num> element,
// starting at 1 and filling gaps.
func (n *CT_Numbering) NextNumId() int {
	var numIds []int
	for _, num := range n.NumList() {
		id, err := num.NumId()
		if err == nil {
			numIds = append(numIds, id)
		}
	}
	sort.Ints(numIds)
	idSet := make(map[int]bool, len(numIds))
	for _, id := range numIds {
		idSet[id] = true
	}
	for i := 1; i <= len(numIds)+1; i++ {
		if !idSet[i] {
			return i
		}
	}
	return len(numIds) + 1
}

// ===========================================================================
// CT_Num — custom methods
// ===========================================================================

// NewNum creates a new <w:num> element with the given numId and a child
// <w:abstractNumId> referencing abstractNumId.
func NewNum(numId, abstractNumId int) *CT_Num {
	el := OxmlElement("w:num")
	num := &CT_Num{Element{E: el}}
	num.SetNumId(numId)

	// Create <w:abstractNumId w:val="N"/>
	absEl := OxmlElement("w:abstractNumId")
	absNum := &CT_DecimalNumber{Element{E: absEl}}
	absNum.SetVal(abstractNumId)
	el.AddChild(absEl)

	return num
}

// AddLvlOverrideWithIlvl adds a new <w:lvlOverride> child with the given ilvl attribute.
func (n *CT_Num) AddLvlOverrideWithIlvl(ilvl int) *CT_NumLvl {
	lvl := n.AddLvlOverride()
	lvl.SetIlvl(ilvl)
	return lvl
}

// ===========================================================================
// CT_NumLvl — custom methods
// ===========================================================================

// AddStartOverrideWithVal adds a <w:startOverride> child element with the given val.
func (nl *CT_NumLvl) AddStartOverrideWithVal(val int) *CT_DecimalNumber {
	so := nl.GetOrAddStartOverride()
	so.SetVal(val)
	return so
}

// ===========================================================================
// CT_NumPr — custom methods
// ===========================================================================

// NumIdVal returns the value of the w:numId/w:val attribute, or nil if not present.
func (np *CT_NumPr) NumIdVal() *int {
	numId := np.NumId()
	if numId == nil {
		return nil
	}
	v, err := numId.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetNumIdVal sets the w:numId/w:val attribute, creating the element if needed.
func (np *CT_NumPr) SetNumIdVal(val int) {
	np.GetOrAddNumId().SetVal(val)
}

// IlvlVal returns the value of the w:ilvl/w:val attribute, or nil if not present.
func (np *CT_NumPr) IlvlVal() *int {
	ilvl := np.Ilvl()
	if ilvl == nil {
		return nil
	}
	v, err := ilvl.Val()
	if err != nil {
		return nil
	}
	return &v
}

// SetIlvlVal sets the w:ilvl/w:val attribute, creating the element if needed.
func (np *CT_NumPr) SetIlvlVal(val int) {
	np.GetOrAddIlvl().SetVal(val)
}

// ===========================================================================
// CT_DecimalNumber — additional factory method
// ===========================================================================

// NewDecimalNumber creates a new element with the given namespace-prefixed tagname
// and val attribute set. Mirrors CT_DecimalNumber.new() from Python.
func NewDecimalNumber(nspTagname string, val int) *CT_DecimalNumber {
	el := OxmlElement(nspTagname)
	dn := &CT_DecimalNumber{Element{E: el}}
	dn.SetVal(val)
	return dn
}

// ===========================================================================
// CT_String — additional factory method
// ===========================================================================

// NewCtString creates a new element with the given namespace-prefixed tagname
// and val attribute set. Mirrors CT_String.new() from Python.
func NewCtString(nspTagname, val string) *CT_String {
	el := OxmlElement(nspTagname)
	s := &CT_String{Element{E: el}}
	s.SetVal(val)
	return s
}

