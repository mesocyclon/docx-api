package shared

import (
	"encoding/xml"
	"testing"
)

// dummyBlock is a minimal BlockLevelElement for testing.
type dummyBlock struct{ tag string }

func (dummyBlock) blockLevelElement() {}

// dummyPara is a minimal ParagraphContent for testing.
type dummyPara struct{ tag string }

func (dummyPara) paragraphContent() {}

// dummyRun is a minimal RunContent for testing.
type dummyRun struct{ tag string }

func (dummyRun) runContent() {}

// ---------------------------------------------------------------------------
// BlockLevelFactory
// ---------------------------------------------------------------------------

func TestBlockFactory(t *testing.T) {
	t.Parallel()

	// Reset global state for this test by working through the public API.
	// Since there is no Reset func, we just register and test that the
	// factory chain works. Other tests running in parallel may also have
	// registered factories, so we test a specific name that only our
	// factory handles.
	RegisterBlockFactory(func(name xml.Name) BlockLevelElement {
		if name.Local == "testBlock" {
			return &dummyBlock{tag: "testBlock"}
		}
		return nil
	})

	el := CreateBlockElement(xml.Name{Local: "testBlock"})
	if el == nil {
		t.Fatal("expected non-nil element for 'testBlock'")
	}
	db, ok := el.(*dummyBlock)
	if !ok {
		t.Fatalf("expected *dummyBlock, got %T", el)
	}
	if db.tag != "testBlock" {
		t.Errorf("tag = %q, want %q", db.tag, "testBlock")
	}
}

func TestBlockFactoryUnknownReturnsNil(t *testing.T) {
	t.Parallel()
	el := CreateBlockElement(xml.Name{Local: "completelyUnknown12345"})
	if el != nil {
		t.Errorf("expected nil for unknown element, got %T", el)
	}
}

// ---------------------------------------------------------------------------
// ParagraphContentFactory
// ---------------------------------------------------------------------------

func TestParagraphContentFactory(t *testing.T) {
	t.Parallel()

	RegisterParagraphContentFactory(func(name xml.Name) ParagraphContent {
		if name.Local == "testPara" {
			return &dummyPara{tag: "testPara"}
		}
		return nil
	})

	el := CreateParagraphContent(xml.Name{Local: "testPara"})
	if el == nil {
		t.Fatal("expected non-nil element for 'testPara'")
	}
	dp, ok := el.(*dummyPara)
	if !ok {
		t.Fatalf("expected *dummyPara, got %T", el)
	}
	if dp.tag != "testPara" {
		t.Errorf("tag = %q, want %q", dp.tag, "testPara")
	}
}

func TestParagraphContentFactoryUnknownReturnsNil(t *testing.T) {
	t.Parallel()
	el := CreateParagraphContent(xml.Name{Local: "completelyUnknownPara12345"})
	if el != nil {
		t.Errorf("expected nil for unknown element, got %T", el)
	}
}

// ---------------------------------------------------------------------------
// RunContentFactory
// ---------------------------------------------------------------------------

func TestRunContentFactory(t *testing.T) {
	t.Parallel()

	RegisterRunContentFactory(func(name xml.Name) RunContent {
		if name.Local == "testRun" {
			return &dummyRun{tag: "testRun"}
		}
		return nil
	})

	el := CreateRunContent(xml.Name{Local: "testRun"})
	if el == nil {
		t.Fatal("expected non-nil element for 'testRun'")
	}
	dr, ok := el.(*dummyRun)
	if !ok {
		t.Fatalf("expected *dummyRun, got %T", el)
	}
	if dr.tag != "testRun" {
		t.Errorf("tag = %q, want %q", dr.tag, "testRun")
	}
}

func TestRunContentFactoryUnknownReturnsNil(t *testing.T) {
	t.Parallel()
	el := CreateRunContent(xml.Name{Local: "completelyUnknownRun12345"})
	if el != nil {
		t.Errorf("expected nil for unknown element, got %T", el)
	}
}

// ---------------------------------------------------------------------------
// Multiple factories — first non-nil wins
// ---------------------------------------------------------------------------

func TestMultipleFactoriesFirstWins(t *testing.T) {
	t.Parallel()

	RegisterBlockFactory(func(name xml.Name) BlockLevelElement {
		if name.Local == "sharedName" {
			return &dummyBlock{tag: "factory-A"}
		}
		return nil
	})
	RegisterBlockFactory(func(name xml.Name) BlockLevelElement {
		if name.Local == "sharedName" {
			return &dummyBlock{tag: "factory-B"}
		}
		return nil
	})

	el := CreateBlockElement(xml.Name{Local: "sharedName"})
	if el == nil {
		t.Fatal("expected non-nil element")
	}
	db := el.(*dummyBlock)
	// The first registered factory that returns non-nil wins.
	if db.tag != "factory-A" {
		t.Errorf("expected first factory to win, got tag = %q", db.tag)
	}
}
