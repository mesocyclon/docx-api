package tracking

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// nsHeader is the namespace declaration preamble reused by tests.
const nsHeader = `xmlns:w="` + xmltypes.NSw + `" xmlns:w14="` + xmltypes.NSw14 + `"`

// ---------------------------------------------------------------------------
// CT_RunTrackChange round-trip
// ---------------------------------------------------------------------------

func TestCT_RunTrackChange_Ins_RoundTrip(t *testing.T) {
	// XML from reference-appendix §3.4 — <w:ins> with a run child.
	// Because no ParagraphContentFactory is registered in this test,
	// the child <w:r> will be stored as shared.RawXML.
	input := `<w:ins ` + nsHeader + ` w:id="2" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">` +
		`<w:r w:rsidR="00F12AB3">` +
		`<w:rPr><w:b/></w:rPr>` +
		`<w:t>60</w:t>` +
		`</w:r>` +
		`</w:ins>`

	var tc CT_RunTrackChange
	if err := xml.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// Verify attributes.
	if tc.ID != 2 {
		t.Errorf("ID = %d, want 2", tc.ID)
	}
	if tc.Author != "Jane Smith" {
		t.Errorf("Author = %q, want %q", tc.Author, "Jane Smith")
	}
	if tc.Date == nil || *tc.Date != "2025-01-15T10:30:00Z" {
		t.Errorf("Date = %v, want %q", tc.Date, "2025-01-15T10:30:00Z")
	}
	if len(tc.Content) != 1 {
		t.Fatalf("Content length = %d, want 1", len(tc.Content))
	}

	// The child should be a RawXML (no factory registered).
	raw, ok := tc.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("Content[0] type = %T, want shared.RawXML", tc.Content[0])
	}
	if raw.XMLName.Local != "r" {
		t.Errorf("RawXML Local = %q, want %q", raw.XMLName.Local, "r")
	}

	// Marshal back.
	output, err := xml.Marshal(&tc)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// Re-unmarshal and compare.
	var tc2 CT_RunTrackChange
	if err := xml.Unmarshal(output, &tc2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if tc2.ID != tc.ID {
		t.Errorf("round-trip lost ID: got %d, want %d", tc2.ID, tc.ID)
	}
	if tc2.Author != tc.Author {
		t.Errorf("round-trip lost Author: got %q, want %q", tc2.Author, tc.Author)
	}
	if (tc2.Date == nil) != (tc.Date == nil) {
		t.Errorf("round-trip lost Date presence")
	} else if tc2.Date != nil && *tc2.Date != *tc.Date {
		t.Errorf("round-trip changed Date: got %q, want %q", *tc2.Date, *tc.Date)
	}
	if len(tc2.Content) != len(tc.Content) {
		t.Errorf("round-trip changed Content length: got %d, want %d", len(tc2.Content), len(tc.Content))
	}
}

func TestCT_RunTrackChange_Del_RoundTrip(t *testing.T) {
	input := `<w:del ` + nsHeader + ` w:id="1" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">` +
		`<w:r w:rsidDel="00F12AB3">` +
		`<w:rPr><w:b/></w:rPr>` +
		`<w:delText>30</w:delText>` +
		`</w:r>` +
		`</w:del>`

	var tc CT_RunTrackChange
	if err := xml.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if tc.ID != 1 {
		t.Errorf("ID = %d, want 1", tc.ID)
	}
	if tc.Author != "Jane Smith" {
		t.Errorf("Author = %q, want %q", tc.Author, "Jane Smith")
	}
	if len(tc.Content) != 1 {
		t.Fatalf("Content length = %d, want 1", len(tc.Content))
	}

	output, err := xml.Marshal(&tc)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var tc2 CT_RunTrackChange
	if err := xml.Unmarshal(output, &tc2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if tc2.ID != 1 || tc2.Author != "Jane Smith" || len(tc2.Content) != 1 {
		t.Error("round-trip lost data in del element")
	}
}

func TestCT_RunTrackChange_NoDate(t *testing.T) {
	input := `<w:ins ` + nsHeader + ` w:id="5" w:author="Bot"></w:ins>`

	var tc CT_RunTrackChange
	if err := xml.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if tc.ID != 5 {
		t.Errorf("ID = %d, want 5", tc.ID)
	}
	if tc.Date != nil {
		t.Errorf("Date = %v, want nil", *tc.Date)
	}
	if len(tc.Content) != 0 {
		t.Errorf("Content length = %d, want 0", len(tc.Content))
	}

	output, err := xml.Marshal(&tc)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}
	// Date should not appear in output.
	if strings.Contains(string(output), "date") {
		t.Errorf("output contains date when Date is nil: %s", output)
	}
}

func TestCT_RunTrackChange_UnknownChildPreserved(t *testing.T) {
	input := `<w:ins ` + nsHeader + ` w:id="10" w:author="A">` +
		`<w14:customElem w14:val="test"><w14:inner>data</w14:inner></w14:customElem>` +
		`</w:ins>`

	var tc CT_RunTrackChange
	if err := xml.Unmarshal([]byte(input), &tc); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if len(tc.Content) != 1 {
		t.Fatalf("Content length = %d, want 1", len(tc.Content))
	}
	raw, ok := tc.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("Content[0] type = %T, want shared.RawXML", tc.Content[0])
	}
	if raw.XMLName.Local != "customElem" {
		t.Errorf("RawXML Local = %q, want %q", raw.XMLName.Local, "customElem")
	}

	output, err := xml.Marshal(&tc)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// Re-parse and verify the custom element survived.
	var tc2 CT_RunTrackChange
	if err := xml.Unmarshal(output, &tc2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if len(tc2.Content) != 1 {
		t.Fatalf("round-trip Content length = %d, want 1", len(tc2.Content))
	}
	raw2, ok := tc2.Content[0].(shared.RawXML)
	if !ok {
		t.Fatalf("round-trip Content[0] type = %T, want shared.RawXML", tc2.Content[0])
	}
	if raw2.XMLName.Local != "customElem" {
		t.Errorf("round-trip lost element name: got %q, want %q", raw2.XMLName.Local, "customElem")
	}
}

// ---------------------------------------------------------------------------
// CT_Markup round-trip
// ---------------------------------------------------------------------------

func TestCT_Markup_RoundTrip(t *testing.T) {
	input := `<w:commentRangeStart ` + nsHeader + ` w:id="0"/>`

	var m CT_Markup
	if err := xml.Unmarshal([]byte(input), &m); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if m.ID != 0 {
		t.Errorf("ID = %d, want 0", m.ID)
	}

	output, err := xml.Marshal(&m)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var m2 CT_Markup
	if err := xml.Unmarshal(output, &m2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if m2.ID != m.ID {
		t.Errorf("round-trip lost ID: got %d, want %d", m2.ID, m.ID)
	}
}

func TestCT_Markup_CommentReference(t *testing.T) {
	input := `<w:commentReference ` + nsHeader + ` w:id="7"/>`
	var m CT_Markup
	if err := xml.Unmarshal([]byte(input), &m); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if m.ID != 7 {
		t.Errorf("ID = %d, want 7", m.ID)
	}
}

// ---------------------------------------------------------------------------
// CT_Bookmark round-trip
// ---------------------------------------------------------------------------

func TestCT_Bookmark_RoundTrip(t *testing.T) {
	input := `<w:bookmarkStart ` + nsHeader + ` w:id="0" w:name="_GoBack"/>`

	var b CT_Bookmark
	if err := xml.Unmarshal([]byte(input), &b); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if b.ID != 0 {
		t.Errorf("ID = %d, want 0", b.ID)
	}
	if b.Name != "_GoBack" {
		t.Errorf("Name = %q, want %q", b.Name, "_GoBack")
	}
	if b.ColFirst != nil {
		t.Errorf("ColFirst = %v, want nil", b.ColFirst)
	}

	output, err := xml.Marshal(&b)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var b2 CT_Bookmark
	if err := xml.Unmarshal(output, &b2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if b2.ID != b.ID || b2.Name != b.Name {
		t.Errorf("round-trip lost data: got id=%d name=%q, want id=%d name=%q",
			b2.ID, b2.Name, b.ID, b.Name)
	}
}

func TestCT_Bookmark_WithColumns(t *testing.T) {
	input := `<w:bookmarkStart ` + nsHeader + ` w:id="3" w:name="TableBookmark" w:colFirst="0" w:colLast="2"/>`

	var b CT_Bookmark
	if err := xml.Unmarshal([]byte(input), &b); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if b.ID != 3 {
		t.Errorf("ID = %d, want 3", b.ID)
	}
	if b.Name != "TableBookmark" {
		t.Errorf("Name = %q, want %q", b.Name, "TableBookmark")
	}
	if b.ColFirst == nil || *b.ColFirst != 0 {
		t.Errorf("ColFirst = %v, want 0", b.ColFirst)
	}
	if b.ColLast == nil || *b.ColLast != 2 {
		t.Errorf("ColLast = %v, want 2", b.ColLast)
	}

	output, err := xml.Marshal(&b)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var b2 CT_Bookmark
	if err := xml.Unmarshal(output, &b2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if b2.ColFirst == nil || *b2.ColFirst != 0 {
		t.Errorf("round-trip lost ColFirst")
	}
	if b2.ColLast == nil || *b2.ColLast != 2 {
		t.Errorf("round-trip lost ColLast")
	}
}

// ---------------------------------------------------------------------------
// CT_MarkupRange round-trip
// ---------------------------------------------------------------------------

func TestCT_MarkupRange_RoundTrip(t *testing.T) {
	input := `<w:bookmarkEnd ` + nsHeader + ` w:id="0"/>`

	var mr CT_MarkupRange
	if err := xml.Unmarshal([]byte(input), &mr); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if mr.ID != 0 {
		t.Errorf("ID = %d, want 0", mr.ID)
	}

	output, err := xml.Marshal(&mr)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var mr2 CT_MarkupRange
	if err := xml.Unmarshal(output, &mr2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if mr2.ID != mr.ID {
		t.Errorf("round-trip lost ID: got %d, want %d", mr2.ID, mr.ID)
	}
}

// ---------------------------------------------------------------------------
// CT_MoveBookmark round-trip
// ---------------------------------------------------------------------------

func TestCT_MoveBookmark_RoundTrip(t *testing.T) {
	input := `<w:moveFromRangeStart ` + nsHeader + ` w:id="3" w:author="Editor" w:name="move1"/>`

	var mb CT_MoveBookmark
	if err := xml.Unmarshal([]byte(input), &mb); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if mb.ID != 3 {
		t.Errorf("ID = %d, want 3", mb.ID)
	}
	if mb.Author != "Editor" {
		t.Errorf("Author = %q, want %q", mb.Author, "Editor")
	}
	if mb.Name != "move1" {
		t.Errorf("Name = %q, want %q", mb.Name, "move1")
	}

	output, err := xml.Marshal(&mb)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var mb2 CT_MoveBookmark
	if err := xml.Unmarshal(output, &mb2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if mb2.ID != mb.ID || mb2.Author != mb.Author || mb2.Name != mb.Name {
		t.Errorf("round-trip lost data: got id=%d author=%q name=%q, want id=%d author=%q name=%q",
			mb2.ID, mb2.Author, mb2.Name, mb.ID, mb.Author, mb.Name)
	}
}

// ---------------------------------------------------------------------------
// Integration: full track-change paragraph snippet
// ---------------------------------------------------------------------------

func TestFullTrackChangeSnippet(t *testing.T) {
	// A simplified but complete paragraph with del + ins, based on
	// reference-appendix §3.4.  We test CT_RunTrackChange only here;
	// the paragraph wrapper is tested in wml/para.

	del := `<w:del ` + nsHeader + ` w:id="1" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">` +
		`<w:r w:rsidDel="00F12AB3"><w:rPr><w:b/></w:rPr><w:delText>30</w:delText></w:r>` +
		`</w:del>`
	ins := `<w:ins ` + nsHeader + ` w:id="2" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">` +
		`<w:r w:rsidR="00F12AB3"><w:rPr><w:b/></w:rPr><w:t>60</w:t></w:r>` +
		`</w:ins>`

	for _, tt := range []struct {
		name   string
		input  string
		wantID int
	}{
		{"del", del, 1},
		{"ins", ins, 2},
	} {
		t.Run(tt.name, func(t *testing.T) {
			var tc CT_RunTrackChange
			if err := xml.Unmarshal([]byte(tt.input), &tc); err != nil {
				t.Fatalf("unmarshal: %v", err)
			}
			if tc.ID != tt.wantID {
				t.Errorf("ID = %d, want %d", tc.ID, tt.wantID)
			}

			// Round-trip.
			out, err := xml.Marshal(&tc)
			if err != nil {
				t.Fatalf("marshal: %v", err)
			}
			var tc2 CT_RunTrackChange
			if err := xml.Unmarshal(out, &tc2); err != nil {
				t.Fatalf("re-unmarshal: %v", err)
			}
			if tc2.ID != tc.ID || tc2.Author != tc.Author || len(tc2.Content) != len(tc.Content) {
				t.Errorf("round-trip mismatch")
			}
		})
	}
}

// ---------------------------------------------------------------------------
// Bookmark + CommentRange combined snippet
// ---------------------------------------------------------------------------

func TestBookmarkAndCommentMarkersRoundTrip(t *testing.T) {
	// bookmarkStart → bookmarkEnd pair
	bsInput := `<w:bookmarkStart ` + nsHeader + ` w:id="5" w:name="TestBM"/>`
	beInput := `<w:bookmarkEnd ` + nsHeader + ` w:id="5"/>`

	var bs CT_Bookmark
	if err := xml.Unmarshal([]byte(bsInput), &bs); err != nil {
		t.Fatalf("bookmarkStart unmarshal: %v", err)
	}
	var be CT_MarkupRange
	if err := xml.Unmarshal([]byte(beInput), &be); err != nil {
		t.Fatalf("bookmarkEnd unmarshal: %v", err)
	}
	if bs.ID != be.ID {
		t.Errorf("bookmark ID mismatch: start=%d end=%d", bs.ID, be.ID)
	}

	// commentRangeStart → commentRangeEnd pair
	csInput := `<w:commentRangeStart ` + nsHeader + ` w:id="0"/>`
	ceInput := `<w:commentRangeEnd ` + nsHeader + ` w:id="0"/>`

	var cs CT_Markup
	if err := xml.Unmarshal([]byte(csInput), &cs); err != nil {
		t.Fatalf("commentRangeStart unmarshal: %v", err)
	}
	var ce CT_Markup
	if err := xml.Unmarshal([]byte(ceInput), &ce); err != nil {
		t.Fatalf("commentRangeEnd unmarshal: %v", err)
	}
	if cs.ID != ce.ID {
		t.Errorf("comment range ID mismatch: start=%d end=%d", cs.ID, ce.ID)
	}
}
