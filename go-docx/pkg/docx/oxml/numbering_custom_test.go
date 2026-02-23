package oxml

import (
	"testing"
)

func TestCT_Numbering_NextNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{E: el}}

	if got := n.NextNumId(); got != 1 {
		t.Errorf("expected next numId=1 on empty, got %d", got)
	}
}

func TestCT_Numbering_AddNumWithAbstractNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{E: el}}

	num, err := n.AddNumWithAbstractNumId(0)
	if err != nil {
		t.Fatalf("AddNumWithAbstractNumId: %v", err)
	}
	if num == nil {
		t.Fatal("expected num, got nil")
	}
	numId, err := num.NumId()
	if err != nil {
		t.Fatalf("numId error: %v", err)
	}
	if numId != 1 {
		t.Errorf("expected numId=1, got %d", numId)
	}

	// Check abstractNumId
	absNum, err := num.AbstractNumId()
	if err != nil {
		t.Fatalf("AbstractNumId error: %v", err)
	}
	absVal, err := absNum.Val()
	if err != nil {
		t.Fatalf("abstractNumId val error: %v", err)
	}
	if absVal != 0 {
		t.Errorf("expected abstractNumId=0, got %d", absVal)
	}

	// Add another
	num2, err := n.AddNumWithAbstractNumId(1)
	if err != nil {
		t.Fatalf("AddNumWithAbstractNumId: %v", err)
	}
	numId2, _ := num2.NumId()
	if numId2 != 2 {
		t.Errorf("expected numId=2, got %d", numId2)
	}
}

func TestCT_Numbering_NumHavingNumId(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{E: el}}

	num := n.NumHavingNumId(7)
	if num == nil {
		t.Fatal("expected num with numId=7, got nil")
	}

	if n.NumHavingNumId(999) != nil {
		t.Error("expected nil for nonexistent numId")
	}
}

func TestCT_Numbering_NextNumId_GapFilling(t *testing.T) {
	xml := `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
		`<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>` +
		`</w:numbering>`
	el, _ := ParseXml([]byte(xml))
	n := &CT_Numbering{Element{E: el}}

	// Should find gap at 2
	if got := n.NextNumId(); got != 2 {
		t.Errorf("expected next numId=2 (gap), got %d", got)
	}
}

func TestNewNum(t *testing.T) {
	num, err := NewNum(5, 3)
	if err != nil {
		t.Fatalf("NewNum: %v", err)
	}
	numId, err := num.NumId()
	if err != nil {
		t.Fatalf("numId error: %v", err)
	}
	if numId != 5 {
		t.Errorf("expected numId=5, got %d", numId)
	}
	absNumId, err := num.AbstractNumId()
	if err != nil {
		t.Fatalf("AbstractNumId error: %v", err)
	}
	absVal, err := absNumId.Val()
	if err != nil {
		t.Fatalf("abstractNumId error: %v", err)
	}
	if absVal != 3 {
		t.Errorf("expected abstractNumId=3, got %d", absVal)
	}
}

func TestCT_NumPr_ValAccessors(t *testing.T) {
	xml := `<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:ilvl w:val="2"/>` +
		`<w:numId w:val="5"/>` +
		`</w:numPr>`
	el, _ := ParseXml([]byte(xml))
	np := &CT_NumPr{Element{E: el}}

	ilvl, err := np.IlvlVal()
	if err != nil {
		t.Fatalf("IlvlVal: %v", err)
	}
	if ilvl == nil || *ilvl != 2 {
		t.Errorf("expected ilvl=2, got %v", ilvl)
	}
	numId, err := np.NumIdVal()
	if err != nil {
		t.Fatalf("NumIdVal: %v", err)
	}
	if numId == nil || *numId != 5 {
		t.Errorf("expected numId=5, got %v", numId)
	}
}

func TestCT_NumPr_ValAccessors_Empty(t *testing.T) {
	xml := `<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := ParseXml([]byte(xml))
	np := &CT_NumPr{Element{E: el}}

	if iv, err := np.IlvlVal(); err != nil {
		t.Fatalf("IlvlVal: %v", err)
	} else if iv != nil {
		t.Error("expected nil ilvl on empty numPr")
	}
	if nid, err := np.NumIdVal(); err != nil {
		t.Fatalf("NumIdVal: %v", err)
	} else if nid != nil {
		t.Error("expected nil numId on empty numPr")
	}

	// Set and verify
	if err := np.SetIlvlVal(3); err != nil {
		t.Fatalf("SetIlvlVal: %v", err)
	}
	if err := np.SetNumIdVal(7); err != nil {
		t.Fatalf("SetNumIdVal: %v", err)
	}
	ilvl, err := np.IlvlVal()
	if err != nil {
		t.Fatalf("IlvlVal: %v", err)
	}
	if ilvl == nil || *ilvl != 3 {
		t.Errorf("expected ilvl=3, got %v", ilvl)
	}
	numId, err := np.NumIdVal()
	if err != nil {
		t.Fatalf("NumIdVal: %v", err)
	}
	if numId == nil || *numId != 7 {
		t.Errorf("expected numId=7, got %v", numId)
	}
}
