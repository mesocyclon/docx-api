package comments

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// referenceXML is the comments.xml sample from reference-appendix.md § 3.5.
const referenceXML = `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
	`<w:comment w:id="0" w:author="Reviewer" w:date="2025-01-20T14:00:00Z" w:initials="R">` +
	`<w:p>` +
	`<w:pPr>` +
	`<w:pStyle w:val="CommentText"/>` +
	`</w:pPr>` +
	`<w:r>` +
	`<w:rPr>` +
	`<w:rStyle w:val="CommentReference"/>` +
	`</w:rPr>` +
	`<w:annotationRef/>` +
	`</w:r>` +
	`<w:r>` +
	`<w:t>Please verify this statement.</w:t>` +
	`</w:r>` +
	`</w:p>` +
	`</w:comment>` +
	`</w:comments>`

// TestRoundTrip verifies that unmarshal → marshal → unmarshal produces
// identical structures and preserves all content (including unknown elements
// stored as RawXML).
func TestRoundTrip(t *testing.T) {
	// --- Phase 1: Unmarshal ---
	comments1, err := Parse([]byte(referenceXML))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// Basic structural checks.
	if len(comments1.Comment) != 1 {
		t.Fatalf("expected 1 comment, got %d", len(comments1.Comment))
	}
	cm := comments1.Comment[0]
	if cm.ID != 0 {
		t.Errorf("expected ID=0, got %d", cm.ID)
	}
	if cm.Author != "Reviewer" {
		t.Errorf("expected Author=Reviewer, got %q", cm.Author)
	}
	if cm.Date != "2025-01-20T14:00:00Z" {
		t.Errorf("expected Date=2025-01-20T14:00:00Z, got %q", cm.Date)
	}
	if cm.Initials != "R" {
		t.Errorf("expected Initials=R, got %q", cm.Initials)
	}
	// Content should have 1 block-level element (the <w:p>), stored as RawXML
	// since no block factory is registered in this test.
	if len(cm.Content) != 1 {
		t.Fatalf("expected 1 content element, got %d", len(cm.Content))
	}
	rawP, ok := cm.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("expected RawXML content, got %T", cm.Content[0])
	}
	if rawP.XMLName.Local != "p" {
		t.Errorf("expected RawXML element 'p', got %q", rawP.XMLName.Local)
	}

	// --- Phase 2: Marshal ---
	serialized, err := Serialize(comments1)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	// --- Phase 3: Re-unmarshal ---
	comments2, err := Parse(serialized)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}

	// --- Phase 4: Compare ---
	if len(comments2.Comment) != len(comments1.Comment) {
		t.Fatalf("round-trip: comment count mismatch: %d vs %d",
			len(comments2.Comment), len(comments1.Comment))
	}

	cm2 := comments2.Comment[0]
	if cm2.ID != cm.ID {
		t.Errorf("round-trip lost ID: %d vs %d", cm2.ID, cm.ID)
	}
	if cm2.Author != cm.Author {
		t.Errorf("round-trip lost Author: %q vs %q", cm2.Author, cm.Author)
	}
	if cm2.Date != cm.Date {
		t.Errorf("round-trip lost Date: %q vs %q", cm2.Date, cm.Date)
	}
	if cm2.Initials != cm.Initials {
		t.Errorf("round-trip lost Initials: %q vs %q", cm2.Initials, cm.Initials)
	}
	if len(cm2.Content) != len(cm.Content) {
		t.Fatalf("round-trip lost content elements: %d vs %d",
			len(cm2.Content), len(cm.Content))
	}

	// Verify the RawXML element name survived.
	rawP2, ok := cm2.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("round-trip: expected RawXML, got %T", cm2.Content[0])
	}
	if rawP2.XMLName.Local != rawP.XMLName.Local {
		t.Errorf("round-trip lost element name: %q vs %q",
			rawP2.XMLName.Local, rawP.XMLName.Local)
	}
}

// TestMultipleComments verifies parsing of multiple comments.
func TestMultipleComments(t *testing.T) {
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:comment w:id="0" w:author="Alice" w:date="2025-01-20T10:00:00Z" w:initials="A">` +
		`<w:p><w:r><w:t>First comment</w:t></w:r></w:p>` +
		`</w:comment>` +
		`<w:comment w:id="1" w:author="Bob" w:initials="B">` +
		`<w:p><w:r><w:t>Second comment</w:t></w:r></w:p>` +
		`</w:comment>` +
		`</w:comments>`

	c, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(c.Comment) != 2 {
		t.Fatalf("expected 2 comments, got %d", len(c.Comment))
	}

	if c.Comment[0].Author != "Alice" {
		t.Errorf("comment 0 author: got %q, want Alice", c.Comment[0].Author)
	}
	if c.Comment[1].Author != "Bob" {
		t.Errorf("comment 1 author: got %q, want Bob", c.Comment[1].Author)
	}
	// Bob's comment has no date — should be empty.
	if c.Comment[1].Date != "" {
		t.Errorf("comment 1 date: expected empty, got %q", c.Comment[1].Date)
	}

	// Round-trip.
	serialized, err := Serialize(c)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	c2, err := Parse(serialized)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}
	if len(c2.Comment) != 2 {
		t.Fatalf("round-trip: expected 2 comments, got %d", len(c2.Comment))
	}
	if c2.Comment[0].ID != 0 || c2.Comment[1].ID != 1 {
		t.Errorf("round-trip lost IDs: %d, %d", c2.Comment[0].ID, c2.Comment[1].ID)
	}
}

// TestEmptyComments verifies an empty <w:comments/> element.
func TestEmptyComments(t *testing.T) {
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>`

	c, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}
	if len(c.Comment) != 0 {
		t.Errorf("expected 0 comments, got %d", len(c.Comment))
	}

	serialized, err := Serialize(c)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	if !strings.Contains(string(serialized), "w:comments") {
		t.Error("serialized output missing w:comments element")
	}
}

// TestNamespacePreservation verifies that extra xmlns declarations survive round-trip.
func TestNamespacePreservation(t *testing.T) {
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
		`<w:comment w:id="0" w:author="Test">` +
		`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>` +
		`</w:comment>` +
		`</w:comments>`

	c, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	// Check that namespace declarations are preserved.
	foundW14 := false
	for _, attr := range c.Namespaces {
		if attr.Value == "http://schemas.microsoft.com/office/word/2010/wordml" {
			foundW14 = true
			break
		}
	}
	if !foundW14 {
		t.Error("lost w14 namespace declaration")
	}

	// Round-trip and verify.
	serialized, err := Serialize(c)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	c2, err := Parse(serialized)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}
	foundW14 = false
	for _, attr := range c2.Namespaces {
		if attr.Value == "http://schemas.microsoft.com/office/word/2010/wordml" {
			foundW14 = true
			break
		}
	}
	if !foundW14 {
		t.Error("round-trip lost w14 namespace declaration")
	}
}

// TestRawXMLContentPreservation verifies that inner XML of block-level
// elements is preserved exactly through round-trip.
func TestRawXMLContentPreservation(t *testing.T) {
	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:comment w:id="5" w:author="QA">` +
		`<w:p>` +
		`<w:pPr><w:pStyle w:val="Normal"/></w:pPr>` +
		`<w:r><w:t>test content</w:t></w:r>` +
		`</w:p>` +
		`<w:tbl>` +
		`<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>` +
		`</w:tbl>` +
		`</w:comment>` +
		`</w:comments>`

	c, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	cm := c.Comment[0]
	if len(cm.Content) != 2 {
		t.Fatalf("expected 2 content elements (p + tbl), got %d", len(cm.Content))
	}

	// Both should be RawXML (no factories registered).
	raw0, ok := cm.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("content[0]: expected RawXML, got %T", cm.Content[0])
	}
	if raw0.XMLName.Local != "p" {
		t.Errorf("content[0]: expected 'p', got %q", raw0.XMLName.Local)
	}

	raw1, ok := cm.Content[1].(shared.RawXML)
	if !ok {
		t.Fatalf("content[1]: expected RawXML, got %T", cm.Content[1])
	}
	if raw1.XMLName.Local != "tbl" {
		t.Errorf("content[1]: expected 'tbl', got %q", raw1.XMLName.Local)
	}

	// Round-trip.
	serialized, err := Serialize(c)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}
	c2, err := Parse(serialized)
	if err != nil {
		t.Fatalf("re-Parse failed: %v", err)
	}

	cm2 := c2.Comment[0]
	if len(cm2.Content) != 2 {
		t.Fatalf("round-trip: expected 2 content elements, got %d", len(cm2.Content))
	}

	r0, ok := cm2.Content[0].(shared.RawXML)
	if !ok || r0.XMLName.Local != "p" {
		t.Error("round-trip lost <w:p> element")
	}
	r1, ok := cm2.Content[1].(shared.RawXML)
	if !ok || r1.XMLName.Local != "tbl" {
		t.Error("round-trip lost <w:tbl> element")
	}

	// Verify inner content survived by checking for key substrings.
	if !strings.Contains(string(r0.Inner), "pStyle") {
		t.Error("round-trip lost pStyle inside <w:p>")
	}
	if !strings.Contains(string(r0.Inner), "test content") {
		t.Error("round-trip lost text inside <w:p>")
	}
	if !strings.Contains(string(r1.Inner), "cell") {
		t.Error("round-trip lost text inside <w:tbl>")
	}
}

// TestWithRegisteredFactory verifies that when a BlockLevelFactory is
// registered, known elements are created via the factory.
func TestWithRegisteredFactory(t *testing.T) {
	// Register a dummy factory that recognizes <w:p>.
	type dummyPara struct {
		shared.RawXML
	}
	shared.RegisterBlockFactory(func(name xml.Name) shared.BlockLevelElement {
		if name.Local == "p" {
			return &dummyPara{}
		}
		return nil
	})

	// Clean up by re-registering without the factory is not easy,
	// so we accept this side-effect in testing.

	input := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:comment w:id="0" w:author="Test">` +
		`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>` +
		`</w:comment>` +
		`</w:comments>`

	c, err := Parse([]byte(input))
	if err != nil {
		t.Fatalf("Parse failed: %v", err)
	}

	cm := c.Comment[0]
	if len(cm.Content) != 1 {
		t.Fatalf("expected 1 content element, got %d", len(cm.Content))
	}

	// The element should have been created by the factory.
	if _, ok := cm.Content[0].(*dummyPara); !ok {
		t.Errorf("expected *dummyPara, got %T", cm.Content[0])
	}
}

// TestSerializeNewDocument verifies that a programmatically constructed
// CT_Comments can be serialized correctly.
func TestSerializeNewDocument(t *testing.T) {
	c := &CT_Comments{
		Comment: []CT_Comment{
			{
				ID:       0,
				Author:   "Author1",
				Date:     "2025-06-01T12:00:00Z",
				Initials: "A",
				Content: []shared.BlockLevelElement{
					shared.RawXML{
						XMLName: xml.Name{Space: xmltypes.NSw, Local: "p"},
						Inner:   []byte(`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>New comment</w:t></w:r>`),
					},
				},
			},
		},
	}

	serialized, err := Serialize(c)
	if err != nil {
		t.Fatalf("Serialize failed: %v", err)
	}

	output := string(serialized)
	if !strings.Contains(output, "Author1") {
		t.Error("missing Author1 in output")
	}
	if !strings.Contains(output, "New comment") {
		t.Error("missing 'New comment' text in output")
	}
	if !strings.Contains(output, `w:id="0"`) && !strings.Contains(output, `id="0"`) {
		t.Error("missing id attribute in output")
	}
}
