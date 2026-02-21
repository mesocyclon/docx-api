package oxml

import (
	"testing"
	"time"
)

// ===========================================================================
// Shape tests
// ===========================================================================

func TestNewPicInline_Structure(t *testing.T) {
	inline := NewPicInline(1, "rId5", "image1.png", 914400, 457200)

	// Check extent dimensions
	cx := inline.ExtentCx()
	if cx != 914400 {
		t.Errorf("expected cx=914400, got %d", cx)
	}
	cy := inline.ExtentCy()
	if cy != 457200 {
		t.Errorf("expected cy=457200, got %d", cy)
	}

	// Check docPr
	docPr := inline.DocPr()
	id, err := docPr.Id()
	if err != nil {
		t.Fatalf("docPr.Id() error: %v", err)
	}
	if id != 1 {
		t.Errorf("expected docPr id=1, got %d", id)
	}
	name, err := docPr.Name()
	if err != nil {
		t.Fatalf("docPr.Name() error: %v", err)
	}
	if name != "Picture 1" {
		t.Errorf("expected docPr name='Picture 1', got %q", name)
	}

	// Check graphic data URI
	gd := inline.Graphic().GraphicData()
	uri, err := gd.Uri()
	if err != nil {
		t.Fatalf("graphicData.Uri() error: %v", err)
	}
	if uri != "http://schemas.openxmlformats.org/drawingml/2006/picture" {
		t.Errorf("unexpected graphicData uri: %q", uri)
	}

	// Check pic element exists in graphicData
	pic := gd.Pic()
	if pic == nil {
		t.Fatal("expected pic:pic inside graphicData, got nil")
	}

	// Check blipFill has the right rId
	embed := pic.BlipFill().Blip().Embed()
	if embed != "rId5" {
		t.Errorf("expected blip embed='rId5', got %q", embed)
	}
}

func TestCT_Inline_SetExtent(t *testing.T) {
	inline := NewPicInline(1, "rId1", "test.png", 100, 200)
	inline.SetExtentCx(300)
	inline.SetExtentCy(400)
	if inline.ExtentCx() != 300 {
		t.Errorf("expected cx=300, got %d", inline.ExtentCx())
	}
	if inline.ExtentCy() != 400 {
		t.Errorf("expected cy=400, got %d", inline.ExtentCy())
	}
}

func TestCT_ShapeProperties_CxCy(t *testing.T) {
	// Parse a pic:spPr with xfrm/ext
	xml := `<pic:spPr xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>`
	el, err := ParseXml([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	spPr := &CT_ShapeProperties{Element{E: el}}

	cx := spPr.Cx()
	if cx == nil || *cx != 914400 {
		t.Errorf("expected cx=914400, got %v", cx)
	}
	cy := spPr.Cy()
	if cy == nil || *cy != 457200 {
		t.Errorf("expected cy=457200, got %v", cy)
	}

	// Set new values
	spPr.SetCx(1234)
	spPr.SetCy(5678)
	cx = spPr.Cx()
	if cx == nil || *cx != 1234 {
		t.Errorf("after set, expected cx=1234, got %v", cx)
	}
	cy = spPr.Cy()
	if cy == nil || *cy != 5678 {
		t.Errorf("after set, expected cy=5678, got %v", cy)
	}
}

func TestCT_ShapeProperties_CxCy_NoXfrm(t *testing.T) {
	xml := `<pic:spPr xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></pic:spPr>`
	el, _ := ParseXml([]byte(xml))
	spPr := &CT_ShapeProperties{Element{E: el}}

	if cx := spPr.Cx(); cx != nil {
		t.Errorf("expected nil cx on empty spPr, got %v", cx)
	}
	if cy := spPr.Cy(); cy != nil {
		t.Errorf("expected nil cy on empty spPr, got %v", cy)
	}
}

// ===========================================================================
// Comments tests
// ===========================================================================

func TestCT_Comments_AddCommentFull(t *testing.T) {
	xml := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>`
	el, _ := ParseXml([]byte(xml))
	cs := &CT_Comments{Element{E: el}}

	c := cs.AddCommentFull()
	if c == nil {
		t.Fatal("expected comment, got nil")
	}
	id, err := c.Id()
	if err != nil {
		t.Fatalf("comment id error: %v", err)
	}
	if id != 0 {
		t.Errorf("expected first comment id=0, got %d", id)
	}

	// Add another
	c2 := cs.AddCommentFull()
	id2, _ := c2.Id()
	if id2 != 1 {
		t.Errorf("expected second comment id=1, got %d", id2)
	}

	// Check list
	if len(cs.CommentList()) != 2 {
		t.Errorf("expected 2 comments, got %d", len(cs.CommentList()))
	}
}

func TestCT_Comments_GetCommentByID(t *testing.T) {
	xml := `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:comment w:id="5" w:author="Alice"/>` +
		`<w:comment w:id="10" w:author="Bob"/>` +
		`</w:comments>`
	el, _ := ParseXml([]byte(xml))
	cs := &CT_Comments{Element{E: el}}

	c := cs.GetCommentByID(10)
	if c == nil {
		t.Fatal("expected comment with id=10, got nil")
	}
	author, _ := c.Author()
	if author != "Bob" {
		t.Errorf("expected author 'Bob', got %q", author)
	}

	if cs.GetCommentByID(999) != nil {
		t.Error("expected nil for nonexistent comment id")
	}
}

func TestCT_Comment_InnerContentElements(t *testing.T) {
	xml := `<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="0" w:author="">` +
		`<w:p/><w:tbl/><w:p/>` +
		`</w:comment>`
	el, _ := ParseXml([]byte(xml))
	c := &CT_Comment{Element{E: el}}

	elems := c.InnerContentElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elems))
	}
	if _, ok := elems[0].(*CT_P); !ok {
		t.Error("expected first element to be CT_P")
	}
	if _, ok := elems[1].(*CT_Tbl); !ok {
		t.Error("expected second element to be CT_Tbl")
	}
	if _, ok := elems[2].(*CT_P); !ok {
		t.Error("expected third element to be CT_P")
	}
}

// ===========================================================================
// CoreProperties tests
// ===========================================================================

func TestNewCoreProperties(t *testing.T) {
	cp := NewCoreProperties()
	if cp == nil {
		t.Fatal("expected coreProperties, got nil")
	}
	// Check that it has the cp namespace
	_, ok := HasNsDecl(cp.E, "cp")
	if !ok {
		t.Error("expected xmlns:cp declaration")
	}
}

func TestCT_CoreProperties_TextProperties(t *testing.T) {
	cp := NewCoreProperties()

	// All text properties should start empty
	if got := cp.TitleText(); got != "" {
		t.Errorf("expected empty title, got %q", got)
	}

	// Set and get title
	if err := cp.SetTitleText("My Document"); err != nil {
		t.Fatal(err)
	}
	if got := cp.TitleText(); got != "My Document" {
		t.Errorf("expected 'My Document', got %q", got)
	}

	// Set and get author
	if err := cp.SetAuthorText("Alice"); err != nil {
		t.Fatal(err)
	}
	if got := cp.AuthorText(); got != "Alice" {
		t.Errorf("expected 'Alice', got %q", got)
	}

	// Set and get subject
	if err := cp.SetSubjectText("Testing"); err != nil {
		t.Fatal(err)
	}
	if got := cp.SubjectText(); got != "Testing" {
		t.Errorf("expected 'Testing', got %q", got)
	}

	// Set and get category
	if err := cp.SetCategoryText("Test Category"); err != nil {
		t.Fatal(err)
	}
	if got := cp.CategoryText(); got != "Test Category" {
		t.Errorf("expected 'Test Category', got %q", got)
	}

	// Set and get keywords
	if err := cp.SetKeywordsText("go, docx"); err != nil {
		t.Fatal(err)
	}
	if got := cp.KeywordsText(); got != "go, docx" {
		t.Errorf("expected 'go, docx', got %q", got)
	}

	// Set and get comments
	if err := cp.SetCommentsText("A test"); err != nil {
		t.Fatal(err)
	}
	if got := cp.CommentsText(); got != "A test" {
		t.Errorf("expected 'A test', got %q", got)
	}

	// Set and get lastModifiedBy
	if err := cp.SetLastModifiedByText("Bob"); err != nil {
		t.Fatal(err)
	}
	if got := cp.LastModifiedByText(); got != "Bob" {
		t.Errorf("expected 'Bob', got %q", got)
	}

	// Set and get contentStatus
	if err := cp.SetContentStatusText("Draft"); err != nil {
		t.Fatal(err)
	}
	if got := cp.ContentStatusText(); got != "Draft" {
		t.Errorf("expected 'Draft', got %q", got)
	}

	// Set and get language
	if err := cp.SetLanguageText("en-US"); err != nil {
		t.Fatal(err)
	}
	if got := cp.LanguageText(); got != "en-US" {
		t.Errorf("expected 'en-US', got %q", got)
	}

	// Set and get version
	if err := cp.SetVersionText("1.0"); err != nil {
		t.Fatal(err)
	}
	if got := cp.VersionText(); got != "1.0" {
		t.Errorf("expected '1.0', got %q", got)
	}
}

func TestCT_CoreProperties_TextProperty_255Limit(t *testing.T) {
	cp := NewCoreProperties()
	longStr := ""
	for i := 0; i < 256; i++ {
		longStr += "x"
	}
	err := cp.SetTitleText(longStr)
	if err == nil {
		t.Error("expected error for string > 255 chars")
	}
}

func TestCT_CoreProperties_DatetimeProperties(t *testing.T) {
	cp := NewCoreProperties()

	// Created should be nil initially
	if got := cp.CreatedDatetime(); got != nil {
		t.Errorf("expected nil created, got %v", got)
	}

	// Set created
	created := time.Date(2024, 1, 15, 10, 30, 0, 0, time.UTC)
	cp.SetCreatedDatetime(created)
	got := cp.CreatedDatetime()
	if got == nil {
		t.Fatal("expected non-nil created")
	}
	if !got.Equal(created) {
		t.Errorf("expected %v, got %v", created, *got)
	}

	// Set modified
	modified := time.Date(2024, 6, 20, 14, 45, 0, 0, time.UTC)
	cp.SetModifiedDatetime(modified)
	got = cp.ModifiedDatetime()
	if got == nil {
		t.Fatal("expected non-nil modified")
	}
	if !got.Equal(modified) {
		t.Errorf("expected %v, got %v", modified, *got)
	}

	// Set lastPrinted
	lastPrinted := time.Date(2024, 3, 1, 8, 0, 0, 0, time.UTC)
	cp.SetLastPrintedDatetime(lastPrinted)
	got = cp.LastPrintedDatetime()
	if got == nil {
		t.Fatal("expected non-nil lastPrinted")
	}
	if !got.Equal(lastPrinted) {
		t.Errorf("expected %v, got %v", lastPrinted, *got)
	}
}

func TestCT_CoreProperties_RevisionNumber(t *testing.T) {
	cp := NewCoreProperties()

	// Default should be 0
	if got := cp.RevisionNumber(); got != 0 {
		t.Errorf("expected revision 0, got %d", got)
	}

	// Set valid revision
	if err := cp.SetRevisionNumber(5); err != nil {
		t.Fatal(err)
	}
	if got := cp.RevisionNumber(); got != 5 {
		t.Errorf("expected revision 5, got %d", got)
	}

	// Set invalid (< 1) should error
	if err := cp.SetRevisionNumber(0); err == nil {
		t.Error("expected error for revision 0")
	}
	if err := cp.SetRevisionNumber(-1); err == nil {
		t.Error("expected error for negative revision")
	}
}

func TestCT_CoreProperties_RevisionFromXml(t *testing.T) {
	// Test parsing a revision from existing XML
	xml := `<cp:coreProperties ` +
		`xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
		`xmlns:dcterms="http://purl.org/dc/terms/">` +
		`<cp:revision>42</cp:revision>` +
		`</cp:coreProperties>`
	el, _ := ParseXml([]byte(xml))
	cp := &CT_CoreProperties{Element{E: el}}

	if got := cp.RevisionNumber(); got != 42 {
		t.Errorf("expected revision 42, got %d", got)
	}

	// Non-integer revision should return 0
	xml2 := `<cp:coreProperties ` +
		`xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" ` +
		`xmlns:dcterms="http://purl.org/dc/terms/">` +
		`<cp:revision>abc</cp:revision>` +
		`</cp:coreProperties>`
	el2, _ := ParseXml([]byte(xml2))
	cp2 := &CT_CoreProperties{Element{E: el2}}
	if got := cp2.RevisionNumber(); got != 0 {
		t.Errorf("expected 0 for non-integer revision, got %d", got)
	}
}

func TestParseW3CDTF(t *testing.T) {
	tests := []struct {
		input   string
		wantErr bool
		year    int
		month   time.Month
		day     int
	}{
		{"2024-01-15T10:30:00Z", false, 2024, time.January, 15},
		{"2024-01-15T10:30:00", false, 2024, time.January, 15},
		{"2024-01-15", false, 2024, time.January, 15},
		{"2024-01", false, 2024, time.January, 1},
		{"2024", false, 2024, time.January, 1},
		{"", true, 0, 0, 0},
		{"not-a-date", true, 0, 0, 0},
	}
	for _, tt := range tests {
		dt, err := parseW3CDTF(tt.input)
		if tt.wantErr {
			if err == nil {
				t.Errorf("parseW3CDTF(%q): expected error, got nil", tt.input)
			}
			continue
		}
		if err != nil {
			t.Errorf("parseW3CDTF(%q): unexpected error: %v", tt.input, err)
			continue
		}
		if dt.Year() != tt.year || dt.Month() != tt.month || dt.Day() != tt.day {
			t.Errorf("parseW3CDTF(%q): expected %d-%d-%d, got %d-%d-%d",
				tt.input, tt.year, tt.month, tt.day, dt.Year(), dt.Month(), dt.Day())
		}
	}
}

func TestParseW3CDTF_WithOffset(t *testing.T) {
	dt, err := parseW3CDTF("2024-01-15T10:30:00-05:00")
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	// -05:00 means UTC+5 hours when reversed (add +5 to get UTC)
	if dt.Hour() != 15 || dt.Minute() != 30 {
		t.Errorf("expected 15:30 UTC, got %d:%d", dt.Hour(), dt.Minute())
	}
}

// ===========================================================================
// Numbering tests
// ===========================================================================

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

	num := n.AddNumWithAbstractNumId(0)
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
	absNum := num.AbstractNumId()
	absVal, err := absNum.Val()
	if err != nil {
		t.Fatalf("abstractNumId val error: %v", err)
	}
	if absVal != 0 {
		t.Errorf("expected abstractNumId=0, got %d", absVal)
	}

	// Add another
	num2 := n.AddNumWithAbstractNumId(1)
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
	num := NewNum(5, 3)
	numId, err := num.NumId()
	if err != nil {
		t.Fatalf("numId error: %v", err)
	}
	if numId != 5 {
		t.Errorf("expected numId=5, got %d", numId)
	}
	absVal, err := num.AbstractNumId().Val()
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

	ilvl := np.IlvlVal()
	if ilvl == nil || *ilvl != 2 {
		t.Errorf("expected ilvl=2, got %v", ilvl)
	}
	numId := np.NumIdVal()
	if numId == nil || *numId != 5 {
		t.Errorf("expected numId=5, got %v", numId)
	}
}

func TestCT_NumPr_ValAccessors_Empty(t *testing.T) {
	xml := `<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := ParseXml([]byte(xml))
	np := &CT_NumPr{Element{E: el}}

	if np.IlvlVal() != nil {
		t.Error("expected nil ilvl on empty numPr")
	}
	if np.NumIdVal() != nil {
		t.Error("expected nil numId on empty numPr")
	}

	// Set and verify
	np.SetIlvlVal(3)
	np.SetNumIdVal(7)
	ilvl := np.IlvlVal()
	if ilvl == nil || *ilvl != 3 {
		t.Errorf("expected ilvl=3, got %v", ilvl)
	}
	numId := np.NumIdVal()
	if numId == nil || *numId != 7 {
		t.Errorf("expected numId=7, got %v", numId)
	}
}

// ===========================================================================
// Settings tests
// ===========================================================================

func TestCT_Settings_EvenAndOddHeadersVal(t *testing.T) {
	xml := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`
	el, _ := ParseXml([]byte(xml))
	s := &CT_Settings{Element{E: el}}

	// Default should be false
	if s.EvenAndOddHeadersVal() {
		t.Error("expected false by default")
	}

	// Set to true
	boolTrue := true
	s.SetEvenAndOddHeadersVal(&boolTrue)
	if !s.EvenAndOddHeadersVal() {
		t.Error("expected true after setting")
	}

	// Set to false (should remove)
	boolFalse := false
	s.SetEvenAndOddHeadersVal(&boolFalse)
	if s.EvenAndOddHeadersVal() {
		t.Error("expected false after unsetting")
	}

	// Set to true again then nil (should remove)
	s.SetEvenAndOddHeadersVal(&boolTrue)
	s.SetEvenAndOddHeadersVal(nil)
	if s.EvenAndOddHeadersVal() {
		t.Error("expected false after setting nil")
	}
}

// ===========================================================================
// Document tests
// ===========================================================================

func TestCT_Body_InnerContentElements(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/><w:tbl/><w:p/><w:sectPr/>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{E: el}}

	elems := body.InnerContentElements()
	if len(elems) != 3 {
		t.Fatalf("expected 3 elements (sectPr excluded), got %d", len(elems))
	}
	if _, ok := elems[0].(*CT_P); !ok {
		t.Error("expected first element to be CT_P")
	}
	if _, ok := elems[1].(*CT_Tbl); !ok {
		t.Error("expected second element to be CT_Tbl")
	}
	if _, ok := elems[2].(*CT_P); !ok {
		t.Error("expected third element to be CT_P")
	}
}

func TestCT_Body_ClearContent(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/><w:tbl/><w:p/><w:sectPr/>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{E: el}}

	body.ClearContent()

	// Only sectPr should remain
	children := body.E.ChildElements()
	if len(children) != 1 {
		t.Fatalf("expected 1 child (sectPr), got %d", len(children))
	}
	if children[0].Tag != "sectPr" {
		t.Errorf("expected remaining child to be sectPr, got %s", children[0].Tag)
	}
}

func TestCT_Body_ClearContent_NoSectPr(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/><w:tbl/>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{E: el}}

	body.ClearContent()
	if len(body.E.ChildElements()) != 0 {
		t.Errorf("expected 0 children, got %d", len(body.E.ChildElements()))
	}
}

func TestCT_Document_SectPrList(t *testing.T) {
	xml := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:body>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` +
		`<w:p/>` +
		`<w:p><w:pPr><w:sectPr/></w:pPr></w:p>` +
		`<w:sectPr/>` +
		`</w:body>` +
		`</w:document>`
	el, _ := ParseXml([]byte(xml))
	doc := &CT_Document{Element{E: el}}

	sectPrs := doc.SectPrList()
	if len(sectPrs) != 3 {
		t.Errorf("expected 3 sectPr elements, got %d", len(sectPrs))
	}
}

func TestCT_Body_AddSectionBreak(t *testing.T) {
	xml := `<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
		`<w:p/>` +
		`<w:sectPr><w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>` +
		`</w:body>`
	el, _ := ParseXml([]byte(xml))
	body := &CT_Body{Element{E: el}}

	sentinelSectPr := body.AddSectionBreak()
	if sentinelSectPr == nil {
		t.Fatal("expected sentinel sectPr, got nil")
	}

	// The sentinel should have no headerReference now (removed)
	refs := sentinelSectPr.HeaderReferenceList()
	if len(refs) != 0 {
		t.Errorf("expected 0 headerReferences on sentinel, got %d", len(refs))
	}

	// There should be a new paragraph with sectPr (the clone)
	pList := body.PList()
	// Should be at least 2 paragraphs now (original + new one with sectPr)
	if len(pList) < 2 {
		t.Fatalf("expected at least 2 paragraphs, got %d", len(pList))
	}
}

// ===========================================================================
// Factory method tests
// ===========================================================================

func TestNewDecimalNumber(t *testing.T) {
	dn := NewDecimalNumber("w:abstractNumId", 42)
	v, err := dn.Val()
	if err != nil {
		t.Fatal(err)
	}
	if v != 42 {
		t.Errorf("expected val=42, got %d", v)
	}
}

func TestNewCtString(t *testing.T) {
	s := NewCtString("w:pStyle", "Heading1")
	v, err := s.Val()
	if err != nil {
		t.Fatal(err)
	}
	if v != "Heading1" {
		t.Errorf("expected val='Heading1', got %q", v)
	}
}
