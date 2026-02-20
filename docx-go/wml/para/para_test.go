package para

import (
	"encoding/xml"
	"strings"
	"testing"

	"github.com/vortex/docx-go/wml/shared"
	"github.com/vortex/docx-go/xmltypes"
)

// Compile-time interface checks.
var (
	_ shared.BlockLevelElement = (*CT_P)(nil)
	_ shared.ParagraphContent  = RunItem{}
	_ shared.ParagraphContent  = HyperlinkItem{}
	_ shared.ParagraphContent  = SimpleFieldItem{}
	_ shared.ParagraphContent  = InsItem{}
	_ shared.ParagraphContent  = DelItem{}
	_ shared.ParagraphContent  = BookmarkStartItem{}
	_ shared.ParagraphContent  = BookmarkEndItem{}
	_ shared.ParagraphContent  = CommentRangeStartItem{}
	_ shared.ParagraphContent  = CommentRangeEndItem{}
	_ shared.ParagraphContent  = SdtRunItem{}
	_ shared.ParagraphContent  = RawParagraphContent{}
)

// ---------------------------------------------------------------------------
// Test 1: Simple paragraph with heading style and a single run
// Reference: appendix §3.1
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_HeadingWithRun(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `" ` +
		`xmlns:w14="` + xmltypes.NSw14 + `" ` +
		`w:rsidR="00A77B3E" w:rsidRDefault="00A77B3E" w:rsidP="00A77B3E">` +
		`<w:pPr>` +
		`<w:pStyle w:val="Heading1"/>` +
		`<w:jc w:val="center"/>` +
		`</w:pPr>` +
		`<w:r>` +
		`<w:rPr><w:b/></w:rPr>` +
		`<w:t>Document Title</w:t>` +
		`</w:r>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// Verify attributes.
	assertStringPtr(t, "RsidR", p.RsidR, "00A77B3E")
	assertStringPtr(t, "RsidRDefault", p.RsidRDefault, "00A77B3E")
	assertStringPtr(t, "RsidP", p.RsidP, "00A77B3E")

	// Verify pPr.
	if p.PPr == nil {
		t.Fatal("expected pPr to be present")
	}
	if p.PPr.Base.PStyle == nil || p.PPr.Base.PStyle.Val != "Heading1" {
		t.Error("expected pStyle=Heading1")
	}
	if p.PPr.Base.Jc == nil || p.PPr.Base.Jc.Val != "center" {
		t.Error("expected jc=center")
	}

	// Verify content: should have 1 run.
	if len(p.Content) != 1 {
		t.Fatalf("expected 1 content item, got %d", len(p.Content))
	}
	ri, ok := p.Content[0].(RunItem)
	if !ok {
		t.Fatalf("expected RunItem, got %T", p.Content[0])
	}
	if ri.R == nil {
		t.Fatal("RunItem.R is nil")
	}

	// Marshal back.
	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	// Re-unmarshal.
	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	assertStringPtr(t, "round-trip RsidR", p2.RsidR, "00A77B3E")
	if p2.PPr == nil {
		t.Fatal("round-trip lost pPr")
	}
	if p2.PPr.Base.PStyle == nil || p2.PPr.Base.PStyle.Val != "Heading1" {
		t.Error("round-trip lost pStyle")
	}
	if len(p2.Content) != 1 {
		t.Errorf("round-trip: expected 1 content item, got %d", len(p2.Content))
	}
}

// ---------------------------------------------------------------------------
// Test 2: Multiple runs (bold + italic example from appendix §3.1)
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_MultipleRuns(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `" w:rsidR="00B22C47" w:rsidRDefault="00B22C47">` +
		`<w:r><w:t xml:space="preserve">This is </w:t></w:r>` +
		`<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>bold</w:t></w:r>` +
		`<w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
		`<w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t>italic</w:t></w:r>` +
		`<w:r><w:t xml:space="preserve"> text.</w:t></w:r>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if len(p.Content) != 5 {
		t.Fatalf("expected 5 runs, got %d", len(p.Content))
	}

	// All items should be RunItem.
	for i, item := range p.Content {
		if _, ok := item.(RunItem); !ok {
			t.Errorf("content[%d]: expected RunItem, got %T", i, item)
		}
	}

	// Round-trip.
	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	if len(p2.Content) != 5 {
		t.Errorf("round-trip: expected 5 runs, got %d", len(p2.Content))
	}
}

// ---------------------------------------------------------------------------
// Test 3: Track changes (ins + del) from appendix §3.4
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_TrackChanges(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `">` +
		`<w:r><w:t xml:space="preserve">The contract term is </w:t></w:r>` +
		`<w:del w:id="1" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">` +
		`<w:r><w:rPr><w:b/></w:rPr><w:t>30</w:t></w:r>` +
		`</w:del>` +
		`<w:ins w:id="2" w:author="Jane Smith" w:date="2025-01-15T10:30:00Z">` +
		`<w:r><w:rPr><w:b/></w:rPr><w:t>60</w:t></w:r>` +
		`</w:ins>` +
		`<w:r><w:t xml:space="preserve"> days.</w:t></w:r>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// Expected content: run, del, ins, run (4 items).
	if len(p.Content) != 4 {
		t.Fatalf("expected 4 content items, got %d", len(p.Content))
	}

	// Check types.
	if _, ok := p.Content[0].(RunItem); !ok {
		t.Errorf("content[0]: expected RunItem, got %T", p.Content[0])
	}
	if di, ok := p.Content[1].(DelItem); !ok {
		t.Errorf("content[1]: expected DelItem, got %T", p.Content[1])
	} else {
		if di.Del == nil {
			t.Error("DelItem.Del is nil")
		} else if di.Del.ID != 1 {
			t.Errorf("del ID = %d, want 1", di.Del.ID)
		}
	}
	if ii, ok := p.Content[2].(InsItem); !ok {
		t.Errorf("content[2]: expected InsItem, got %T", p.Content[2])
	} else {
		if ii.Ins == nil {
			t.Error("InsItem.Ins is nil")
		} else if ii.Ins.ID != 2 {
			t.Errorf("ins ID = %d, want 2", ii.Ins.ID)
		}
	}

	// Round-trip.
	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	if len(p2.Content) != 4 {
		t.Errorf("round-trip: expected 4 items, got %d", len(p2.Content))
	}
}

// ---------------------------------------------------------------------------
// Test 4: Bookmarks and comments from appendix §3.5
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_BookmarksAndComments(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `">` +
		`<w:bookmarkStart w:id="0" w:name="_Ref123"/>` +
		`<w:commentRangeStart w:id="0"/>` +
		`<w:r><w:t>This text has a comment.</w:t></w:r>` +
		`<w:commentRangeEnd w:id="0"/>` +
		`<w:bookmarkEnd w:id="0"/>` +
		`<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
		`<w:commentReference w:id="0"/></w:r>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// Expected: bookmarkStart, commentRangeStart, run, commentRangeEnd,
	//           bookmarkEnd, run (6 items).
	if len(p.Content) != 6 {
		t.Fatalf("expected 6 content items, got %d", len(p.Content))
	}

	if bs, ok := p.Content[0].(BookmarkStartItem); !ok {
		t.Errorf("content[0]: expected BookmarkStartItem, got %T", p.Content[0])
	} else if bs.B.Name != "_Ref123" {
		t.Errorf("bookmark name = %q, want %q", bs.B.Name, "_Ref123")
	}

	if _, ok := p.Content[1].(CommentRangeStartItem); !ok {
		t.Errorf("content[1]: expected CommentRangeStartItem, got %T", p.Content[1])
	}

	if _, ok := p.Content[3].(CommentRangeEndItem); !ok {
		t.Errorf("content[3]: expected CommentRangeEndItem, got %T", p.Content[3])
	}

	if _, ok := p.Content[4].(BookmarkEndItem); !ok {
		t.Errorf("content[4]: expected BookmarkEndItem, got %T", p.Content[4])
	}

	// Round-trip.
	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if len(p2.Content) != 6 {
		t.Errorf("round-trip: expected 6 items, got %d", len(p2.Content))
	}
}

// ---------------------------------------------------------------------------
// Test 5: Numbered list paragraph (appendix §3.3)
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_NumberedList(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `">` +
		`<w:pPr>` +
		`<w:pStyle w:val="ListParagraph"/>` +
		`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` +
		`</w:pPr>` +
		`<w:r><w:t>First bullet item</w:t></w:r>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if p.PPr == nil {
		t.Fatal("expected pPr")
	}
	if p.PPr.Base.PStyle == nil || p.PPr.Base.PStyle.Val != "ListParagraph" {
		t.Error("expected pStyle=ListParagraph")
	}
	if p.PPr.Base.NumPr == nil {
		t.Error("expected numPr")
	}

	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if p2.PPr == nil || p2.PPr.Base.PStyle == nil || p2.PPr.Base.PStyle.Val != "ListParagraph" {
		t.Error("round-trip lost pStyle")
	}
}

// ---------------------------------------------------------------------------
// Test 6: w14 attributes (paraId, textId)
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_W14Attributes(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `" ` +
		`xmlns:w14="` + xmltypes.NSw14 + `" ` +
		`w14:paraId="00000001" w14:textId="77777777">` +
		`<w:r><w:t>Hello</w:t></w:r>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	assertStringPtr(t, "ParaId", p.ParaId, "00000001")
	assertStringPtr(t, "TextId", p.TextId, "77777777")

	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	assertStringPtr(t, "round-trip ParaId", p2.ParaId, "00000001")
	assertStringPtr(t, "round-trip TextId", p2.TextId, "77777777")
}

// ---------------------------------------------------------------------------
// Test 7: Unknown extension elements preserved as RawParagraphContent
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_UnknownExtension(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `" ` +
		`xmlns:w14="` + xmltypes.NSw14 + `">` +
		`<w:r><w:t>text</w:t></w:r>` +
		`<w14:customExtension w14:val="test"><w14:child>data</w14:child></w14:customExtension>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if len(p.Content) != 2 {
		t.Fatalf("expected 2 content items, got %d", len(p.Content))
	}

	raw, ok := p.Content[1].(RawParagraphContent)
	if !ok {
		t.Fatalf("content[1]: expected RawParagraphContent, got %T", p.Content[1])
	}
	if raw.Raw.XMLName.Local != "customExtension" {
		t.Errorf("raw element name = %q, want %q", raw.Raw.XMLName.Local, "customExtension")
	}

	// Round-trip.
	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if len(p2.Content) != 2 {
		t.Fatalf("round-trip: expected 2 items, got %d", len(p2.Content))
	}

	raw2, ok := p2.Content[1].(RawParagraphContent)
	if !ok {
		t.Fatalf("round-trip content[1]: expected RawParagraphContent, got %T", p2.Content[1])
	}
	if raw2.Raw.XMLName.Local != "customExtension" {
		t.Errorf("round-trip lost extension element name: got %q", raw2.Raw.XMLName.Local)
	}
	// Verify child data survived.
	if !strings.Contains(string(raw2.Raw.Inner), "data") {
		t.Errorf("round-trip lost inner content of extension element")
	}
}

// ---------------------------------------------------------------------------
// Test 8: Empty paragraph (no pPr, no content)
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_Empty(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `"></w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if p.PPr != nil {
		t.Error("expected nil pPr for empty paragraph")
	}
	if len(p.Content) != 0 {
		t.Errorf("expected 0 content items, got %d", len(p.Content))
	}

	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if p2.PPr != nil {
		t.Error("round-trip: expected nil pPr")
	}
}

// ---------------------------------------------------------------------------
// Test 9: Hyperlink with r:id
// ---------------------------------------------------------------------------

func TestCT_Hyperlink_RoundTrip(t *testing.T) {
	input := `<w:hyperlink xmlns:w="` + xmltypes.NSw + `" ` +
		`xmlns:r="` + xmltypes.NSr + `" ` +
		`r:id="rId5">` +
		`<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>` +
		`<w:t>Click here</w:t></w:r>` +
		`</w:hyperlink>`

	var h CT_Hyperlink
	if err := xml.Unmarshal([]byte(input), &h); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	assertStringPtr(t, "RID", h.RID, "rId5")
	if len(h.Content) != 1 {
		t.Fatalf("expected 1 content item, got %d", len(h.Content))
	}

	out, err := xml.Marshal(&h)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var h2 CT_Hyperlink
	if err := xml.Unmarshal(out, &h2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	assertStringPtr(t, "round-trip RID", h2.RID, "rId5")
	if len(h2.Content) != 1 {
		t.Errorf("round-trip: expected 1 item, got %d", len(h2.Content))
	}
}

// ---------------------------------------------------------------------------
// Test 10: Hyperlink with anchor (internal bookmark)
// ---------------------------------------------------------------------------

func TestCT_Hyperlink_Anchor(t *testing.T) {
	input := `<w:hyperlink xmlns:w="` + xmltypes.NSw + `" w:anchor="Section1">` +
		`<w:r><w:t>See Section 1</w:t></w:r>` +
		`</w:hyperlink>`

	var h CT_Hyperlink
	if err := xml.Unmarshal([]byte(input), &h); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	assertStringPtr(t, "Anchor", h.Anchor, "Section1")

	out, err := xml.Marshal(&h)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var h2 CT_Hyperlink
	if err := xml.Unmarshal(out, &h2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	assertStringPtr(t, "round-trip Anchor", h2.Anchor, "Section1")
}

// ---------------------------------------------------------------------------
// Test 11: Simple field
// ---------------------------------------------------------------------------

func TestCT_SimpleField_RoundTrip(t *testing.T) {
	input := `<w:fldSimple xmlns:w="` + xmltypes.NSw + `" w:instr=" PAGE ">` +
		`<w:r><w:t>3</w:t></w:r>` +
		`</w:fldSimple>`

	var f CT_SimpleField
	if err := xml.Unmarshal([]byte(input), &f); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if f.Instr != " PAGE " {
		t.Errorf("Instr = %q, want %q", f.Instr, " PAGE ")
	}

	out, err := xml.Marshal(&f)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var f2 CT_SimpleField
	if err := xml.Unmarshal(out, &f2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if f2.Instr != " PAGE " {
		t.Errorf("round-trip Instr = %q, want %q", f2.Instr, " PAGE ")
	}
}

// ---------------------------------------------------------------------------
// Test 12: BlockLevelFactory registration
// ---------------------------------------------------------------------------

func TestBlockFactory_CreatesP(t *testing.T) {
	el := shared.CreateBlockElement(xml.Name{Space: xmltypes.NSw, Local: "p"})
	if el == nil {
		t.Fatal("factory returned nil for <w:p>")
	}
	if _, ok := el.(*CT_P); !ok {
		t.Errorf("expected *CT_P, got %T", el)
	}
}

func TestBlockFactory_EmptyNamespace(t *testing.T) {
	// encoding/xml sometimes resolves to "" when namespace is inherited.
	el := shared.CreateBlockElement(xml.Name{Space: "", Local: "p"})
	if el == nil {
		t.Fatal("factory returned nil for <p> with empty namespace")
	}
}

func TestBlockFactory_UnknownReturnsNil(t *testing.T) {
	el := shared.CreateBlockElement(xml.Name{Space: xmltypes.NSw, Local: "completelyUnknown"})
	if el != nil {
		t.Errorf("expected nil for unknown element, got %T", el)
	}
}

// ---------------------------------------------------------------------------
// Test 13: Full complex paragraph (comprehensive integration)
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_Complex(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `" ` +
		`xmlns:w14="` + xmltypes.NSw14 + `" ` +
		`xmlns:r="` + xmltypes.NSr + `" ` +
		`w:rsidR="00A77B3E" w:rsidRDefault="00B22C47" ` +
		`w14:paraId="AABB0011" w14:textId="DEADBEEF">` +
		`<w:pPr>` +
		`<w:pStyle w:val="Normal"/>` +
		`<w:spacing w:before="240" w:after="120"/>` +
		`</w:pPr>` +
		`<w:bookmarkStart w:id="1" w:name="intro"/>` +
		`<w:r><w:t>Hello </w:t></w:r>` +
		`<w:hyperlink r:id="rId7">` +
		`<w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>world</w:t></w:r>` +
		`</w:hyperlink>` +
		`<w:r><w:t>!</w:t></w:r>` +
		`<w:bookmarkEnd w:id="1"/>` +
		`<w14:unknownExt w14:data="xyz"/>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	// Attributes.
	assertStringPtr(t, "RsidR", p.RsidR, "00A77B3E")
	assertStringPtr(t, "RsidRDefault", p.RsidRDefault, "00B22C47")
	assertStringPtr(t, "ParaId", p.ParaId, "AABB0011")
	assertStringPtr(t, "TextId", p.TextId, "DEADBEEF")

	// pPr.
	if p.PPr == nil {
		t.Fatal("pPr is nil")
	}
	if p.PPr.Base.PStyle == nil || p.PPr.Base.PStyle.Val != "Normal" {
		t.Error("expected pStyle=Normal")
	}

	// Content: bookmarkStart, run, hyperlink, run, bookmarkEnd, unknownExt
	if len(p.Content) != 6 {
		t.Fatalf("expected 6 content items, got %d (types: %v)", len(p.Content), contentTypes(p.Content))
	}

	if _, ok := p.Content[0].(BookmarkStartItem); !ok {
		t.Errorf("content[0]: expected BookmarkStartItem, got %T", p.Content[0])
	}
	if _, ok := p.Content[1].(RunItem); !ok {
		t.Errorf("content[1]: expected RunItem, got %T", p.Content[1])
	}
	if hi, ok := p.Content[2].(HyperlinkItem); !ok {
		t.Errorf("content[2]: expected HyperlinkItem, got %T", p.Content[2])
	} else {
		assertStringPtr(t, "hyperlink RID", hi.H.RID, "rId7")
	}
	if _, ok := p.Content[3].(RunItem); !ok {
		t.Errorf("content[3]: expected RunItem, got %T", p.Content[3])
	}
	if _, ok := p.Content[4].(BookmarkEndItem); !ok {
		t.Errorf("content[4]: expected BookmarkEndItem, got %T", p.Content[4])
	}
	if raw, ok := p.Content[5].(RawParagraphContent); !ok {
		t.Errorf("content[5]: expected RawParagraphContent, got %T", p.Content[5])
	} else if raw.Raw.XMLName.Local != "unknownExt" {
		t.Errorf("raw name = %q, want %q", raw.Raw.XMLName.Local, "unknownExt")
	}

	// Marshal and re-unmarshal.
	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}

	assertStringPtr(t, "rt RsidR", p2.RsidR, "00A77B3E")
	assertStringPtr(t, "rt ParaId", p2.ParaId, "AABB0011")
	if len(p2.Content) != 6 {
		t.Errorf("round-trip: expected 6 items, got %d (types: %v)", len(p2.Content), contentTypes(p2.Content))
	}
}

// ---------------------------------------------------------------------------
// Test 14: Paragraph with pPr only (no runs)
// ---------------------------------------------------------------------------

func TestCT_P_RoundTrip_PPrOnly(t *testing.T) {
	input := `<w:p xmlns:w="` + xmltypes.NSw + `">` +
		`<w:pPr><w:jc w:val="both"/></w:pPr>` +
		`</w:p>`

	var p CT_P
	if err := xml.Unmarshal([]byte(input), &p); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}

	if p.PPr == nil || p.PPr.Base.Jc == nil || p.PPr.Base.Jc.Val != "both" {
		t.Error("expected jc=both")
	}
	if len(p.Content) != 0 {
		t.Errorf("expected 0 content items, got %d", len(p.Content))
	}

	out, err := xml.Marshal(&p)
	if err != nil {
		t.Fatalf("marshal: %v", err)
	}

	var p2 CT_P
	if err := xml.Unmarshal(out, &p2); err != nil {
		t.Fatalf("re-unmarshal: %v", err)
	}
	if p2.PPr == nil || p2.PPr.Base.Jc == nil || p2.PPr.Base.Jc.Val != "both" {
		t.Error("round-trip lost jc")
	}
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

func assertStringPtr(t *testing.T, name string, got *string, want string) {
	t.Helper()
	if got == nil {
		t.Errorf("%s: expected %q, got nil", name, want)
		return
	}
	if *got != want {
		t.Errorf("%s: expected %q, got %q", name, want, *got)
	}
}

func contentTypes(items []shared.ParagraphContent) []string {
	types := make([]string, len(items))
	for i, item := range items {
		switch item.(type) {
		case RunItem:
			types[i] = "RunItem"
		case HyperlinkItem:
			types[i] = "HyperlinkItem"
		case SimpleFieldItem:
			types[i] = "SimpleFieldItem"
		case InsItem:
			types[i] = "InsItem"
		case DelItem:
			types[i] = "DelItem"
		case BookmarkStartItem:
			types[i] = "BookmarkStartItem"
		case BookmarkEndItem:
			types[i] = "BookmarkEndItem"
		case CommentRangeStartItem:
			types[i] = "CommentRangeStartItem"
		case CommentRangeEndItem:
			types[i] = "CommentRangeEndItem"
		case SdtRunItem:
			types[i] = "SdtRunItem"
		case RawParagraphContent:
			types[i] = "RawParagraphContent"
		default:
			types[i] = "unknown"
		}
	}
	return types
}
