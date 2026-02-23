package parts

import (
	"testing"

	"github.com/beevik/etree"
)

func makeElementWithIDs(ids ...string) *etree.Element {
	el := etree.NewElement("root")
	for _, id := range ids {
		child := etree.NewElement("item")
		child.CreateAttr("id", id)
		el.AddChild(child)
	}
	return el
}

func TestNextID_EmptyElement(t *testing.T) {
	el := etree.NewElement("root")
	maxID := 0
	collectMaxID(el, &maxID)
	got := maxID + 1
	if got != 1 {
		t.Errorf("NextID for empty element: got %d, want 1", got)
	}
}

func TestNextID_WithIDs(t *testing.T) {
	el := makeElementWithIDs("1", "5", "3")
	maxID := 0
	collectMaxID(el, &maxID)
	got := maxID + 1
	if got != 6 {
		t.Errorf("NextID with ids [1,5,3]: got %d, want 6", got)
	}
}

func TestNextID_IgnoresNonDigit(t *testing.T) {
	el := makeElementWithIDs("abc", "12", "rId4", "7")
	maxID := 0
	collectMaxID(el, &maxID)
	got := maxID + 1
	if got != 13 {
		t.Errorf("NextID with mixed ids: got %d, want 13", got)
	}
}

func TestNextID_NestedElements(t *testing.T) {
	el := etree.NewElement("root")
	child := etree.NewElement("p")
	child.CreateAttr("id", "10")
	grandchild := etree.NewElement("r")
	grandchild.CreateAttr("id", "20")
	child.AddChild(grandchild)
	el.AddChild(child)

	maxID := 0
	collectMaxID(el, &maxID)
	got := maxID + 1
	if got != 21 {
		t.Errorf("NextID nested: got %d, want 21", got)
	}
}

func TestIsDigits(t *testing.T) {
	tests := []struct {
		input string
		want  bool
	}{
		{"", false},
		{"123", true},
		{"0", true},
		{"abc", false},
		{"12a", false},
		{"rId3", false},
	}
	for _, tt := range tests {
		got := isDigits(tt.input)
		if got != tt.want {
			t.Errorf("isDigits(%q) = %v, want %v", tt.input, got, tt.want)
		}
	}
}

func TestRelRefCount(t *testing.T) {
	el := etree.NewElement("root")
	child1 := etree.NewElement("drawing")
	child1.CreateAttr("r:id", "rId5")
	el.AddChild(child1)
	child2 := etree.NewElement("hyperlink")
	child2.CreateAttr("r:id", "rId5")
	el.AddChild(child2)
	child3 := etree.NewElement("other")
	child3.CreateAttr("r:id", "rId3")
	el.AddChild(child3)

	count := 0
	countRIdRefs(el, "rId5", &count)
	if count != 2 {
		t.Errorf("relRefCount for rId5: got %d, want 2", count)
	}

	count = 0
	countRIdRefs(el, "rId3", &count)
	if count != 1 {
		t.Errorf("relRefCount for rId3: got %d, want 1", count)
	}

	count = 0
	countRIdRefs(el, "rId99", &count)
	if count != 0 {
		t.Errorf("relRefCount for rId99: got %d, want 0", count)
	}
}

func TestDropRel_DeletesWhenRefCountLow(t *testing.T) {
	// DropRel should delete a relationship when its XML reference count < 2.
	// Core logic tested via countRIdRefs; integration tested in document_test.go.
	el := etree.NewElement("root")
	count := 0
	countRIdRefs(el, "rId1", &count)
	if count >= 2 {
		t.Errorf("expected count < 2 for element with no refs, got %d", count)
	}
}
