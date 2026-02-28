// replace-txt generates a pair of .docx files to visually verify ReplaceText:
//
//	01_before_replace.docx — original document with highlighted placeholders and a spec for each test
//	02_after_replace.docx  — same document reopened, replacements applied, re-saved
//
// Visual verification:
//
//	BEFORE:  Yellow-highlighted text = placeholders that WILL be replaced.
//	         Green-highlighted text  = the expected replacement value (shown in the spec line).
//	AFTER:   Yellow-highlighted text now contains the replacement values.
//	         Compare yellow (actual) vs green (expected) in each section — they must match.
//
// The "before" file is created from scratch, saved to disk, then reopened via
// docx.OpenBytes (full serialization roundtrip) before applying replacements.
// This ensures the test exercises the real read→modify→write pipeline.
//
// Run:
//
//	go run ./visual-regtest/replace-txt --output ./visual-regtest/replace-txt/out
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

func boolPtr(v bool) *bool { return &v }

// FileResult captures the outcome of one generation.
type FileResult struct {
	Name    string `json:"name"`
	OK      bool   `json:"ok"`
	Error   string `json:"error,omitempty"`
	Elapsed string `json:"elapsed"`
}

func main() {
	outputDir := flag.String("output", "", "directory for generated .docx files")
	flag.Parse()

	if *outputDir == "" {
		log.Fatal("--output is required")
	}
	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	var results []FileResult

	// ---- Step 1: build and save the "before" document ----
	start := time.Now()
	beforeDoc, err := buildBeforeDocument()
	if err != nil {
		log.Fatalf("building before document: %v", err)
	}
	beforePath := filepath.Join(*outputDir, "01_before_replace.docx")
	if err := beforeDoc.SaveFile(beforePath); err != nil {
		log.Fatalf("saving before document: %v", err)
	}
	results = append(results, FileResult{
		Name: "01_before_replace.docx", OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("OK   01_before_replace.docx (%s)", time.Since(start))

	// ---- Step 2: reopen saved file and apply replacements ----
	start = time.Now()
	beforeBytes, err := os.ReadFile(beforePath)
	if err != nil {
		log.Fatalf("reading before file: %v", err)
	}
	afterDoc, err := docx.OpenBytes(beforeBytes)
	if err != nil {
		log.Fatalf("opening before file: %v", err)
	}

	totalCount := 0
	for _, r := range allReplacements() {
		n, err := afterDoc.ReplaceText(r.old, r.new)
		if err != nil {
			log.Fatalf("ReplaceText(%q, %q): %v", r.old, r.new, err)
		}
		log.Printf("  %-40s → %-30s  %d hits", truncQuote(r.old), truncQuote(r.new), n)
		totalCount += n
	}
	log.Printf("  total replacements: %d", totalCount)

	afterPath := filepath.Join(*outputDir, "02_after_replace.docx")
	if err := afterDoc.SaveFile(afterPath); err != nil {
		log.Fatalf("saving after document: %v", err)
	}
	results = append(results, FileResult{
		Name: "02_after_replace.docx", OK: true,
		Elapsed: time.Since(start).String(),
	})
	log.Printf("OK   02_after_replace.docx (%s)", time.Since(start))

	// ---- Manifest ----
	data, _ := json.MarshalIndent(results, "", "  ")
	if err := os.WriteFile(filepath.Join(*outputDir, "manifest.json"), data, 0o644); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}
	log.Printf("done: %d files, %d total replacements", len(results), totalCount)
}

// truncQuote quotes a string, truncating if longer than 30 chars.
func truncQuote(s string) string {
	if len(s) > 30 {
		return fmt.Sprintf("%q...", s[:27])
	}
	return fmt.Sprintf("%q", s)
}

// ============================================================================
// Replacement table — every pair is applied to the "before" document
// ============================================================================

type repl struct{ old, new string }

func allReplacements() []repl {
	return []repl{
		// §1  simple placeholders
		{"{{NAME}}", "Иван Петров"},
		{"{{DATE}}", "January 15, 2025"},
		{"{{COMPANY}}", "Acme Corp"},

		// §2  cross-run (text split across differently-formatted runs)
		{"CROSSRUN_REPLACE", "DONE"},

		// §3  table cells
		{"CELL_OLD", "CELL_NEW"},

		// §4  nested table
		{"NESTED_OLD", "NESTED_NEW"},

		// §5  merged cells
		{"MERGED_CELL_TEXT", "MERGED_REPLACED"},

		// §6  header (primary)
		{"HEADER_PLACEHOLDER", "Real Header Title"},

		// §7  footer (primary)
		{"FOOTER_PLACEHOLDER", "Page 1 — Confidential"},

		// §8  first-page header/footer
		{"FIRST_HDR", "First Page Header — Replaced"},
		{"FIRST_FTR", "First Page Footer — Replaced"},

		// §9  header with table
		{"HDR_TBL_LEFT", "Company Name"},
		{"HDR_TBL_RIGHT", "Doc #12345"},

		// §10 footer with table
		{"FTR_TBL_LEFT", "Legal Notice"},
		{"FTR_TBL_RIGHT", "Page X of Y"},

		// §11  multiple occurrences
		{"MULTI", "REPLACED"},

		// §12 tab inside search string
		{"COL_A\tCOL_B", "MERGED_AB"},

		// §13 newline inside search string
		{"LINE_ONE\nLINE_TWO", "SINGLE_LINE"},

		// §14 deletion (replace with empty)
		{"[DELETE_ME]", ""},

		// §15 short → long expansion
		{"TINY", "THIS_IS_MUCH_LONGER_THAN_BEFORE"},

		// §16 long → short contraction
		{"VERY_LONG_PLACEHOLDER_TEXT_HERE", "Short"},

		// §17 Cyrillic single-run
		{"ЗАМЕНИТЬ", "ГОТОВО"},
		{"Шаблон", "Результат"},

		// §18 cross-run Cyrillic
		{"КРОССРАН", "OK"},

		// §19 replacement at paragraph start
		{"STARTWORD", "REPLACED_START"},

		// §20 replacement at paragraph end
		{"ENDWORD", "REPLACED_END"},

		// §21 no-op: old == ""
		{"", "should_never_appear"},

		// §22 no-op: old == new
		{"NOOP_SAME", "NOOP_SAME"},

		// §23 comment-annotated text
		{"COMMENTED_TEXT", "COMMENT_REPLACED"},

		// §24 table: full row replacement
		{"ROW_NAME", "Alice Johnson"},
		{"ROW_ROLE", "Lead Engineer"},
		{"ROW_DEPT", "Platform"},

		// §25 table: header row replacement
		{"TH_COL1", "Employee"},
		{"TH_COL2", "Department"},
		{"TH_COL3", "Status"},

		// §29 replacement inside comment body
		{"COMMENT_BODY_OLD", "COMMENT_BODY_NEW"},
	}
}

// ============================================================================
// Document builder — creates the "before" document from scratch
// ============================================================================

func buildBeforeDocument() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	// ---- headers / footers ----
	sect, err := doc.Sections().Get(0)
	if err != nil {
		return nil, fmt.Errorf("getting default section: %w", err)
	}

	// Primary header: highlighted text + table with placeholders
	hdr := sect.Header()
	buildHighlightedParagraph(hdr, "HEADER_PLACEHOLDER")
	hdrTbl, err := hdr.AddTable(1, 2, 9000)
	if err != nil {
		return nil, err
	}
	hc0, _ := hdrTbl.CellAt(0, 0)
	hc0.SetText("HDR_TBL_LEFT")
	hc1, _ := hdrTbl.CellAt(0, 1)
	hc1.SetText("HDR_TBL_RIGHT")

	// Primary footer: highlighted text + table with placeholders
	ftr := sect.Footer()
	buildHighlightedParagraph(ftr, "FOOTER_PLACEHOLDER")
	ftrTbl, err := ftr.AddTable(1, 2, 9000)
	if err != nil {
		return nil, err
	}
	fc0, _ := ftrTbl.CellAt(0, 0)
	fc0.SetText("FTR_TBL_LEFT")
	fc1, _ := ftrTbl.CellAt(0, 1)
	fc1.SetText("FTR_TBL_RIGHT")

	// First-page header/footer
	if err := sect.SetDifferentFirstPageHeaderFooter(true); err != nil {
		return nil, err
	}
	buildHighlightedParagraph(sect.FirstPageHeader(), "FIRST_HDR")
	buildHighlightedParagraph(sect.FirstPageFooter(), "FIRST_FTR")

	// ================================================================
	// Legend
	// ================================================================
	heading(doc, "ReplaceText — Visual Regression Test", 0)

	lp1, _ := doc.AddParagraph("")
	addPlain(lp1, "How to read this document:  ")
	r, _ := lp1.AddRun("Yellow highlight")
	_ = r.SetBold(boolPtr(true))
	setHighlightYellow(r)
	addPlain(lp1, " = placeholder (will be replaced).  ")
	r, _ = lp1.AddRun("Green highlight")
	_ = r.SetBold(boolPtr(true))
	setHighlightGreen(r)
	addPlain(lp1, " = expected result.")

	lp2, _ := doc.AddParagraph("")
	addPlain(lp2, "In ")
	addBold(lp2, "02_after_replace.docx")
	addPlain(lp2, ": yellow text should now contain the expected value. "+
		"Compare yellow (actual) vs green (expected) in each spec line — they must match.")

	para(doc, "")

	// ================================================================
	// §1 — Simple placeholder replacement
	// ================================================================
	heading(doc, "1. Simple Placeholder Replacement", 1)
	spec(doc, "{{NAME}}", "Иван Петров")
	spec(doc, "{{DATE}}", "January 15, 2025")
	spec(doc, "{{COMPANY}}", "Acme Corp")

	testPara(doc, "Name: ", "{{NAME}}", "")
	testPara(doc, "Date: ", "{{DATE}}", "")
	testPara(doc, "Company: ", "{{COMPANY}}", "")

	// ================================================================
	// §2 — Cross-run replacement (formatting preserved)
	// ================================================================
	heading(doc, "2. Cross-Run Replacement (Formatting Preserved)", 1)
	spec(doc, "CROSSRUN_REPLACE", "DONE")
	note(doc, "Placeholder split across 3 runs: 'CROSS' (bold red) + 'RUN_RE' (italic blue) + 'PLACE' (normal). "+
		"Result appears in first run. Middle/last runs become empty but keep their rPr.")

	p2, _ := doc.AddParagraph("")
	addPlain(p2, "Before marker: ")
	cr1, _ := p2.AddRun("CROSS")
	_ = cr1.SetBold(boolPtr(true))
	c1 := docx.NewRGBColor(0xFF, 0, 0)
	_ = cr1.Font().Color().SetRGB(&c1)
	setHighlightYellow(cr1)
	cr2, _ := p2.AddRun("RUN_RE")
	_ = cr2.SetItalic(boolPtr(true))
	c2 := docx.NewRGBColor(0, 0, 0xFF)
	_ = cr2.Font().Color().SetRGB(&c2)
	setHighlightYellow(cr2)
	cr3, _ := p2.AddRun("PLACE")
	setHighlightYellow(cr3)
	addPlain(p2, " — after marker")

	// ================================================================
	// §3 — Table cell replacement
	// ================================================================
	heading(doc, "3. Table Cell Replacement", 1)
	spec(doc, "CELL_OLD", "CELL_NEW")
	note(doc, "3×3 table. 'CELL_OLD' appears in 3 cells. Other cells untouched.")

	tbl3, _ := doc.AddTable(3, 3)
	fillTable(tbl3, [][]string{
		{"Header A", "Header B", "Header C"},
		{"CELL_OLD value 1", "Normal text", "CELL_OLD value 2"},
		{"Row 3 Col 1", "CELL_OLD value 3", "Row 3 Col 3"},
	})

	// ================================================================
	// §4 — Nested table
	// ================================================================
	heading(doc, "4. Nested Table Replacement", 1)
	spec(doc, "NESTED_OLD", "NESTED_NEW")
	note(doc, "Table inside a cell. 'NESTED_OLD' in inner table rows. Outer cell untouched.")

	outer, _ := doc.AddTable(1, 2)
	oc0, _ := outer.CellAt(0, 0)
	oc0.SetText("Outer cell — no replacement")
	oc1, _ := outer.CellAt(0, 1)
	inner, _ := oc1.AddTable(2, 1)
	ic0, _ := inner.CellAt(0, 0)
	ic0.SetText("NESTED_OLD — row 1")
	ic1, _ := inner.CellAt(1, 0)
	ic1.SetText("NESTED_OLD — row 2")

	// ================================================================
	// §5 — Merged cells
	// ================================================================
	heading(doc, "5. Merged Cell Replacement", 1)
	spec(doc, "MERGED_CELL_TEXT", "MERGED_REPLACED")
	note(doc, "A1+B1 merged horizontally. Placeholder replaced exactly once, not duplicated.")

	tbl5, _ := doc.AddTable(2, 3)
	a1, _ := tbl5.CellAt(0, 0)
	b1, _ := tbl5.CellAt(0, 1)
	merged, _ := a1.Merge(b1)
	merged.SetText("MERGED_CELL_TEXT — spans two columns")
	cc, _ := tbl5.CellAt(0, 2)
	cc.SetText("Normal C1")
	for c := 0; c < 3; c++ {
		cl, _ := tbl5.CellAt(1, c)
		cl.SetText(fmt.Sprintf("Row 2, Col %d", c+1))
	}

	// ================================================================
	// §6 — Header replacement (text)
	// ================================================================
	heading(doc, "6. Header Replacement (Text)", 1)
	spec(doc, "HEADER_PLACEHOLDER", "Real Header Title")
	note(doc, "Primary page header contains yellow-highlighted placeholder. Visible on page 2+.")

	// ================================================================
	// §7 — Footer replacement (text)
	// ================================================================
	heading(doc, "7. Footer Replacement (Text)", 1)
	spec(doc, "FOOTER_PLACEHOLDER", "Page 1 — Confidential")
	note(doc, "Primary page footer contains yellow-highlighted placeholder. Visible on page 2+.")

	// ================================================================
	// §8 — First-page header/footer
	// ================================================================
	heading(doc, "8. First-Page Header/Footer", 1)
	spec(doc, "FIRST_HDR", "First Page Header — Replaced")
	spec(doc, "FIRST_FTR", "First Page Footer — Replaced")
	note(doc, "First page has separate header/footer. Primary header/footer visible on page 2 onward.")

	// ================================================================
	// §9 — Header table replacement
	// ================================================================
	heading(doc, "9. Header Table Replacement", 1)
	spec(doc, "HDR_TBL_LEFT", "Company Name")
	spec(doc, "HDR_TBL_RIGHT", "Doc #12345")
	note(doc, "Primary header has a 1×2 table. Both cells have placeholders. "+
		"Verifies ReplaceText reaches tables inside headers.")

	// ================================================================
	// §10 — Footer table replacement
	// ================================================================
	heading(doc, "10. Footer Table Replacement", 1)
	spec(doc, "FTR_TBL_LEFT", "Legal Notice")
	spec(doc, "FTR_TBL_RIGHT", "Page X of Y")
	note(doc, "Primary footer has a 1×2 table. Both cells have placeholders. "+
		"Verifies ReplaceText reaches tables inside footers.")

	// Force page 2 so primary header/footer become visible.
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	heading(doc, "— Page 2 (primary header/footer visible here) —", 2)

	// ================================================================
	// §11 — Multiple occurrences
	// ================================================================
	heading(doc, "11. Multiple Occurrences", 1)
	spec(doc, "MULTI", "REPLACED")
	note(doc, "4 occurrences across 2 paragraphs — all must be replaced.")

	p11a, _ := doc.AddParagraph("")
	addHighlighted(p11a, "MULTI")
	addPlain(p11a, " is here, and ")
	addHighlighted(p11a, "MULTI")
	addPlain(p11a, " is there, and ")
	addHighlighted(p11a, "MULTI")
	addPlain(p11a, " is everywhere.")

	testPara(doc, "Another paragraph also has ", "MULTI", " in it.")

	// ================================================================
	// §12 — Tab inside search string
	// ================================================================
	heading(doc, "12. Tab Inside Search String", 1)
	spec(doc, "COL_A⟶COL_B (⟶ = tab)", "MERGED_AB")
	note(doc, "The <w:tab> between A and B is consumed by the match. Trailing tab+COL_C remains.")
	para(doc, "COL_A\tCOL_B\tCOL_C")

	// ================================================================
	// §13 — Newline inside search string
	// ================================================================
	heading(doc, "13. Newline Inside Search String", 1)
	spec(doc, "LINE_ONE⏎LINE_TWO (⏎ = br)", "SINGLE_LINE")
	note(doc, "The <w:br> is consumed. Trailing br+LINE_THREE remains.")
	para(doc, "LINE_ONE\nLINE_TWO\nLINE_THREE")

	// ================================================================
	// §14 — Deletion
	// ================================================================
	heading(doc, "14. Deletion (Replace with Empty String)", 1)
	spec(doc, "[DELETE_ME]", "⟨empty⟩")
	note(doc, "Placeholder disappears completely. Surrounding text stays. "+
		"Yellow highlight vanishes from result.")

	testPara(doc, "Before: ", "[DELETE_ME]", " — bracket text should vanish.")
	testPara(doc, "Also here: ", "[DELETE_ME]", " gone.")

	// ================================================================
	// §15 — Short → long
	// ================================================================
	heading(doc, "15. Short to Long Expansion", 1)
	spec(doc, "TINY", "THIS_IS_MUCH_LONGER_THAN_BEFORE")

	testPara(doc, "Expand this: ", "TINY", " — done.")

	// ================================================================
	// §16 — Long → short
	// ================================================================
	heading(doc, "16. Long to Short Contraction", 1)
	spec(doc, "VERY_LONG_PLACEHOLDER_TEXT_HERE", "Short")

	testPara(doc, "Contract this: ", "VERY_LONG_PLACEHOLDER_TEXT_HERE", " — done.")

	// ================================================================
	// §17 — Cyrillic / UTF-8
	// ================================================================
	heading(doc, "17. Cyrillic / UTF-8 (Single Run)", 1)
	spec(doc, "ЗАМЕНИТЬ", "ГОТОВО")
	spec(doc, "Шаблон", "Результат")
	note(doc, "Multibyte byte offsets must be handled correctly.")

	testPara(doc, "Нужно ", "ЗАМЕНИТЬ", " это слово.")
	testPara(doc, "", "Шаблон", " документа — версия 1.0")

	// ================================================================
	// §18 — Cross-run Cyrillic
	// ================================================================
	heading(doc, "18. Cross-Run Cyrillic", 1)
	spec(doc, "КРОССРАН", "OK")
	note(doc, "Split: 'КРОСС' (bold) + 'РАН' (italic). Both runs keep formatting.")

	p18, _ := doc.AddParagraph("")
	addPlain(p18, "Before: ")
	kr1, _ := p18.AddRun("КРОСС")
	_ = kr1.SetBold(boolPtr(true))
	setHighlightYellow(kr1)
	kr2, _ := p18.AddRun("РАН")
	_ = kr2.SetItalic(boolPtr(true))
	setHighlightYellow(kr2)
	addPlain(p18, " — кросс-рановая кириллица")

	// ================================================================
	// §19 — Replacement at paragraph start
	// ================================================================
	heading(doc, "19. Replacement at Paragraph Start", 1)
	spec(doc, "STARTWORD", "REPLACED_START")

	testPara(doc, "", "STARTWORD", " is at the very beginning.")

	// ================================================================
	// §20 — Replacement at paragraph end
	// ================================================================
	heading(doc, "20. Replacement at Paragraph End", 1)
	spec(doc, "ENDWORD", "REPLACED_END")

	testPara(doc, "This paragraph ends with ", "ENDWORD", "")

	// ================================================================
	// §21 — No-op: old == ""
	// ================================================================
	heading(doc, "21. No-Op: Empty Search String", 1)
	note(doc, "ReplaceText(\"\", ...) returns 0. 'should_never_appear' must NOT appear. "+
		"Paragraph must be identical in both files.")
	para(doc, "Nothing should change here. Absolutely identical in both files.")

	// ================================================================
	// §22 — No-op: old == new
	// ================================================================
	heading(doc, "22. No-Op: old == new", 1)
	spec(doc, "NOOP_SAME", "NOOP_SAME")
	note(doc, "Returns 0. No XML modification. Text identical in both files.")

	testPara(doc, "The word ", "NOOP_SAME", " should remain as-is in both files.")

	// ================================================================
	// §23 — Comment-annotated text
	// ================================================================
	heading(doc, "23. Comment-Annotated Text", 1)
	spec(doc, "COMMENTED_TEXT", "COMMENT_REPLACED")
	note(doc, "A comment is attached to the placeholder. After replacement "+
		"comment markers and the comment body must survive.")

	p23, _ := doc.AddParagraph("")
	addPlain(p23, "Commented word: ")
	p23run, _ := p23.AddRun("COMMENTED_TEXT")
	setHighlightYellow(p23run)
	initials := "VR"
	if _, err := doc.AddComment([]*docx.Run{p23run}, "This is a test comment — must survive replacement", "Visual Regtest", &initials); err != nil {
		return nil, fmt.Errorf("adding comment: %w", err)
	}

	// ================================================================
	// §24 — Table: data row replacement
	// ================================================================
	heading(doc, "24. Table Data Row Replacement", 1)
	spec(doc, "ROW_NAME", "Alice Johnson")
	spec(doc, "ROW_ROLE", "Lead Engineer")
	spec(doc, "ROW_DEPT", "Platform")
	note(doc, "3×3 table with header row. Second row has 3 placeholders. "+
		"Header and third row untouched.")

	tbl24, _ := doc.AddTable(3, 3)
	fillTable(tbl24, [][]string{
		{"Name", "Role", "Department"},
		{"ROW_NAME", "ROW_ROLE", "ROW_DEPT"},
		{"Bob Smith", "Designer", "Creative"},
	})

	// ================================================================
	// §25 — Table: header row replacement
	// ================================================================
	heading(doc, "25. Table Header Row Replacement", 1)
	spec(doc, "TH_COL1", "Employee")
	spec(doc, "TH_COL2", "Department")
	spec(doc, "TH_COL3", "Status")
	note(doc, "Table header row cells contain placeholders. Data rows untouched.")

	tbl25, _ := doc.AddTable(3, 3)
	fillTable(tbl25, [][]string{
		{"TH_COL1", "TH_COL2", "TH_COL3"},
		{"John", "Engineering", "Active"},
		{"Jane", "Marketing", "On Leave"},
	})

	// ================================================================
	// §26 — Multiple different placeholders in one paragraph
	// ================================================================
	heading(doc, "26. Multiple Different Placeholders in One Paragraph", 1)
	spec(doc, "{{NAME}}", "Иван Петров")
	spec(doc, "{{COMPANY}}", "Acme Corp")
	spec(doc, "{{DATE}}", "January 15, 2025")
	note(doc, "Three separate ReplaceText calls each hit this paragraph. "+
		"All replaced, surrounding text intact.")

	p26, _ := doc.AddParagraph("")
	addPlain(p26, "Dear ")
	addHighlighted(p26, "{{NAME}}")
	addPlain(p26, ", your order from ")
	addHighlighted(p26, "{{COMPANY}}")
	addPlain(p26, " on ")
	addHighlighted(p26, "{{DATE}}")
	addPlain(p26, " is confirmed.")

	// ================================================================
	// §27 — Unchanged paragraph (no match)
	// ================================================================
	heading(doc, "27. Unchanged Paragraph (No Match)", 1)
	note(doc, "No placeholders here. Must be byte-identical in both files.")
	para(doc, "The quick brown fox jumps over the lazy dog. 0123456789. No markers whatsoever.")

	// ================================================================
	// §28 — Replacement inside one formatted run, sibling runs untouched
	// ================================================================
	heading(doc, "28. Replacement Inside One Formatted Run", 1)
	spec(doc, "CELL_OLD", "CELL_NEW")
	note(doc, "'CELL_OLD' sits inside a bold run followed by italic text. "+
		"After replacement bold stays bold, italic stays italic.")

	p28, _ := doc.AddParagraph("")
	fr1, _ := p28.AddRun("Bold CELL_OLD text")
	_ = fr1.SetBold(boolPtr(true))
	setHighlightYellow(fr1)
	fr2, _ := p28.AddRun(" then italic text")
	_ = fr2.SetItalic(boolPtr(true))

	// ================================================================
	// §29 — Replacement inside comment body
	// ================================================================
	heading(doc, "29. Replacement Inside Comment Body", 1)
	spec(doc, "COMMENT_BODY_OLD", "COMMENT_BODY_NEW")
	note(doc, "The comment body itself contains a placeholder. ReplaceText must reach "+
		"into word/comments.xml and replace it. The annotated body text stays unchanged.")

	p29, _ := doc.AddParagraph("")
	addPlain(p29, "This text has a comment whose body contains a placeholder: ")
	p29run, _ := p29.AddRun("see comment")
	setHighlightYellow(p29run)
	initials29 := "CB"
	if _, err := doc.AddComment(
		[]*docx.Run{p29run},
		"Review status: COMMENT_BODY_OLD — needs update",
		"Comment Body Test", &initials29,
	); err != nil {
		return nil, fmt.Errorf("adding comment for §29: %w", err)
	}

	return doc, nil
}

// ============================================================================
// Helpers: formatting
// ============================================================================

var (
	colorDarkRed   = docx.NewRGBColor(0xCC, 0x00, 0x00)
	colorDarkGreen = docx.NewRGBColor(0x00, 0x66, 0x00)
	colorGray      = docx.NewRGBColor(0x88, 0x88, 0x88)
	colorLightGray = docx.NewRGBColor(0x99, 0x99, 0x99)
)

func setHighlightYellow(r *docx.Run) {
	hl := enum.WdColorIndexYellow
	_ = r.Font().SetHighlightColor(&hl)
}

func setHighlightGreen(r *docx.Run) {
	hl := enum.WdColorIndexBrightGreen
	_ = r.Font().SetHighlightColor(&hl)
}

// ============================================================================
// Helpers: paragraph builders
// ============================================================================

func heading(doc *docx.Document, text string, level int) {
	if _, err := doc.AddHeading(text, level); err != nil {
		log.Fatalf("AddHeading(%q): %v", text, err)
	}
}

func para(doc *docx.Document, text string) {
	if _, err := doc.AddParagraph(text); err != nil {
		log.Fatalf("AddParagraph: %v", err)
	}
}

// spec adds a specification line:
//
//	▸ ~~old~~ (yellow, strikethrough, dark red)  →  new (green, bold, dark green)
//
// In the "before" doc this shows what will happen.
// In the "after" doc the spec is unchanged (no placeholders in it) and serves
// as the expected-value reference.
func spec(doc *docx.Document, old, new string) {
	p, _ := doc.AddParagraph("")

	// Prefix arrow
	pfx, _ := p.AddRun("    ▸ ")
	_ = pfx.Font().Color().SetRGB(&colorLightGray)

	// Old value: yellow highlight + strikethrough + dark red
	rOld, _ := p.AddRun(old)
	setHighlightYellow(rOld)
	_ = rOld.Font().SetStrike(boolPtr(true))
	_ = rOld.Font().Color().SetRGB(&colorDarkRed)

	// Arrow
	arr, _ := p.AddRun("   →   ")
	_ = arr.Font().Color().SetRGB(&colorLightGray)

	// New value: green highlight + bold + dark green
	rNew, _ := p.AddRun(new)
	setHighlightGreen(rNew)
	_ = rNew.SetBold(boolPtr(true))
	_ = rNew.Font().Color().SetRGB(&colorDarkGreen)
}

// note adds a gray description/instruction line.
func note(doc *docx.Document, text string) {
	p, _ := doc.AddParagraph("")
	r, _ := p.AddRun(text)
	_ = r.Font().Color().SetRGB(&colorGray)
}

// testPara builds: prefix + highlighted(placeholder) + suffix.
// The placeholder run gets yellow highlight — visually shows what will be replaced.
func testPara(doc *docx.Document, prefix, placeholder, suffix string) {
	p, _ := doc.AddParagraph("")
	if prefix != "" {
		addPlain(p, prefix)
	}
	addHighlighted(p, placeholder)
	if suffix != "" {
		addPlain(p, suffix)
	}
}

// ============================================================================
// Helpers: run builders
// ============================================================================

func addPlain(p *docx.Paragraph, text string) {
	_, _ = p.AddRun(text)
}

func addBold(p *docx.Paragraph, text string) {
	r, _ := p.AddRun(text)
	_ = r.SetBold(boolPtr(true))
}

func addHighlighted(p *docx.Paragraph, text string) {
	r, _ := p.AddRun(text)
	setHighlightYellow(r)
}

// ============================================================================
// Helpers: header/footer paragraph with highlighted placeholder
// ============================================================================

type paragraphAdder interface {
	AddParagraph(text string, style ...docx.StyleRef) (*docx.Paragraph, error)
}

func buildHighlightedParagraph(hf paragraphAdder, text string) {
	p, err := hf.AddParagraph("")
	if err != nil {
		log.Fatalf("AddParagraph in header/footer: %v", err)
	}
	r, _ := p.AddRun(text)
	setHighlightYellow(r)
}

// ============================================================================
// Helpers: table fill
// ============================================================================

func fillTable(tbl *docx.Table, data [][]string) {
	for r, row := range data {
		for c, val := range row {
			cell, err := tbl.CellAt(r, c)
			if err != nil {
				log.Fatalf("CellAt(%d,%d): %v", r, c, err)
			}
			cell.SetText(val)
		}
	}
}
