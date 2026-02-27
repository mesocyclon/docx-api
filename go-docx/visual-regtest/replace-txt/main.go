// replace-txt generates a pair of .docx files to visually verify ReplaceText:
//
//	01_before_replace.docx — original document with placeholders and test patterns
//	02_after_replace.docx  — same document reopened, replacements applied, re-saved
//
// The "before" file is created from scratch, saved to disk, then reopened via
// docx.OpenBytes (full serialization roundtrip) before applying replacements.
// This ensures the test exercises the real read→modify→write pipeline.
//
// Open both files side-by-side in Word / LibreOffice to verify each scenario.
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

		// §9  multiple occurrences
		{"MULTI", "REPLACED"},

		// §10 tab inside search string
		{"COL_A\tCOL_B", "MERGED_AB"},

		// §11 newline inside search string
		{"LINE_ONE\nLINE_TWO", "SINGLE_LINE"},

		// §12 deletion (replace with empty)
		{"[DELETE_ME]", ""},

		// §13 short → long expansion
		{"TINY", "THIS_IS_MUCH_LONGER_THAN_BEFORE"},

		// §14 long → short contraction
		{"VERY_LONG_PLACEHOLDER_TEXT_HERE", "Short"},

		// §15 Cyrillic single-run
		{"ЗАМЕНИТЬ", "ГОТОВО"},
		{"Шаблон", "Результат"},

		// §16 cross-run Cyrillic
		{"КРОССРАН", "OK"},

		// §17 replacement at paragraph start
		{"STARTWORD", "REPLACED_START"},

		// §18 replacement at paragraph end
		{"ENDWORD", "REPLACED_END"},

		// §19 no-op: old == ""
		{"", "should_never_appear"},

		// §20 no-op: old == new
		{"NOOP_SAME", "NOOP_SAME"},

		// §21 comment-annotated text
		{"COMMENTED_TEXT", "COMMENT_REPLACED"},
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

	// ---- headers, footers, first-page setup ----
	sect, err := doc.Sections().Get(0)
	if err != nil {
		return nil, fmt.Errorf("getting default section: %w", err)
	}

	if _, err := sect.Header().AddParagraph("HEADER_PLACEHOLDER"); err != nil {
		return nil, err
	}
	if _, err := sect.Footer().AddParagraph("FOOTER_PLACEHOLDER"); err != nil {
		return nil, err
	}
	if err := sect.SetDifferentFirstPageHeaderFooter(true); err != nil {
		return nil, err
	}
	if _, err := sect.FirstPageHeader().AddParagraph("FIRST_HDR"); err != nil {
		return nil, err
	}
	if _, err := sect.FirstPageFooter().AddParagraph("FIRST_FTR"); err != nil {
		return nil, err
	}

	// ---- document title ----
	heading(doc, "ReplaceText — Visual Regression Test", 0)
	check(doc, "Open 01_before_replace.docx and 02_after_replace.docx side-by-side to compare every section below.")

	// ================================================================
	// §1 — Simple placeholder replacement
	// ================================================================
	heading(doc, "1. Simple Placeholder Replacement", 1)
	check(doc, "{{NAME}} → 'Иван Петров', {{DATE}} → 'January 15, 2025', {{COMPANY}} → 'Acme Corp'.")
	para(doc, "Name: {{NAME}}")
	para(doc, "Date: {{DATE}}")
	para(doc, "Company: {{COMPANY}}")

	// ================================================================
	// §2 — Cross-run replacement (formatting preserved)
	// ================================================================
	heading(doc, "2. Cross-Run Replacement (Formatting Preserved)", 1)
	check(doc, "'CROSSRUN_REPLACE' split: 'CROSS' (bold red) + 'RUN_RE' (italic blue) + 'PLACE' (normal). "+
		"After → 'DONE' in first run. Middle/last runs become empty but keep their rPr.")

	p2, _ := doc.AddParagraph("")
	r1, _ := p2.AddRun("CROSS")
	_ = r1.SetBold(boolPtr(true))
	c1 := docx.NewRGBColor(0xFF, 0, 0)
	_ = r1.Font().Color().SetRGB(&c1)
	r2, _ := p2.AddRun("RUN_RE")
	_ = r2.SetItalic(boolPtr(true))
	c2 := docx.NewRGBColor(0, 0, 0xFF)
	_ = r2.Font().Color().SetRGB(&c2)
	_, _ = p2.AddRun("PLACE")

	// ================================================================
	// §3 — Table cell replacement
	// ================================================================
	heading(doc, "3. Table Cell Replacement", 1)
	check(doc, "'CELL_OLD' in three cells → 'CELL_NEW'. Other cells untouched.")

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
	check(doc, "'NESTED_OLD' in table-inside-cell → 'NESTED_NEW'. Outer cell untouched.")

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
	check(doc, "A1+B1 merged horizontally. 'MERGED_CELL_TEXT' replaced exactly once, not duplicated.")

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
	// §6 — Header replacement
	// ================================================================
	heading(doc, "6. Header Replacement", 1)
	check(doc, "Page header: 'HEADER_PLACEHOLDER' → 'Real Header Title'. Visible on page 2+.")

	// ================================================================
	// §7 — Footer replacement
	// ================================================================
	heading(doc, "7. Footer Replacement", 1)
	check(doc, "Page footer: 'FOOTER_PLACEHOLDER' → 'Page 1 — Confidential'. Visible on page 2+.")

	// ================================================================
	// §8 — First-page header/footer
	// ================================================================
	heading(doc, "8. First-Page Header/Footer", 1)
	check(doc, "First page header: 'FIRST_HDR' → 'First Page Header — Replaced'. "+
		"First page footer: 'FIRST_FTR' → 'First Page Footer — Replaced'. "+
		"Primary header/footer visible on page 2 onward.")

	// Force page 2 to show primary header/footer.
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	heading(doc, "— Page 2 (primary header/footer visible here) —", 2)

	// ================================================================
	// §9 — Multiple occurrences
	// ================================================================
	heading(doc, "9. Multiple Occurrences", 1)
	check(doc, "'MULTI' appears 3× in one paragraph, 1× in another → all 4 become 'REPLACED'.")
	para(doc, "MULTI is here, and MULTI is there, and MULTI is everywhere.")
	para(doc, "Another paragraph also has MULTI in it.")

	// ================================================================
	// §10 — Tab inside search string
	// ================================================================
	heading(doc, "10. Tab Inside Search String", 1)
	check(doc, "'COL_A<tab>COL_B' → 'MERGED_AB'. The <w:tab> between A and B is consumed. Trailing '<tab>COL_C' remains.")
	para(doc, "COL_A\tCOL_B\tCOL_C")

	// ================================================================
	// §11 — Newline inside search string
	// ================================================================
	heading(doc, "11. Newline Inside Search String", 1)
	check(doc, "'LINE_ONE<br>LINE_TWO' → 'SINGLE_LINE'. The <w:br> is consumed. Trailing '<br>LINE_THREE' remains.")
	para(doc, "LINE_ONE\nLINE_TWO\nLINE_THREE")

	// ================================================================
	// §12 — Deletion
	// ================================================================
	heading(doc, "12. Deletion (Replace with Empty String)", 1)
	check(doc, "'[DELETE_ME]' disappears. Surrounding text stays.")
	para(doc, "Before: [DELETE_ME] — bracket text should vanish.")
	para(doc, "Also here: [DELETE_ME] gone.")

	// ================================================================
	// §13 — Short → long
	// ================================================================
	heading(doc, "13. Short to Long Expansion", 1)
	check(doc, "'TINY' → 'THIS_IS_MUCH_LONGER_THAN_BEFORE'.")
	para(doc, "Expand this: TINY — done.")

	// ================================================================
	// §14 — Long → short
	// ================================================================
	heading(doc, "14. Long to Short Contraction", 1)
	check(doc, "'VERY_LONG_PLACEHOLDER_TEXT_HERE' → 'Short'.")
	para(doc, "Contract this: VERY_LONG_PLACEHOLDER_TEXT_HERE — done.")

	// ================================================================
	// §15 — Cyrillic / UTF-8
	// ================================================================
	heading(doc, "15. Cyrillic / UTF-8 (Single Run)", 1)
	check(doc, "'ЗАМЕНИТЬ' → 'ГОТОВО'. 'Шаблон' → 'Результат'. Multibyte byte offsets must be correct.")
	para(doc, "Нужно ЗАМЕНИТЬ это слово.")
	para(doc, "Шаблон документа — версия 1.0")

	// ================================================================
	// §16 — Cross-run Cyrillic
	// ================================================================
	heading(doc, "16. Cross-Run Cyrillic", 1)
	check(doc, "'КРОССРАН' split: 'КРОСС' (bold) + 'РАН' (italic) → 'OK'. Both runs keep formatting.")

	p16, _ := doc.AddParagraph("")
	kr1, _ := p16.AddRun("КРОСС")
	_ = kr1.SetBold(boolPtr(true))
	kr2, _ := p16.AddRun("РАН")
	_ = kr2.SetItalic(boolPtr(true))
	_, _ = p16.AddRun(" — кросс-рановая кириллица")

	// ================================================================
	// §17 — Replacement at paragraph start
	// ================================================================
	heading(doc, "17. Replacement at Paragraph Start", 1)
	check(doc, "'STARTWORD' at position 0 → 'REPLACED_START'.")
	para(doc, "STARTWORD is at the very beginning.")

	// ================================================================
	// §18 — Replacement at paragraph end
	// ================================================================
	heading(doc, "18. Replacement at Paragraph End", 1)
	check(doc, "'ENDWORD' at the very end → 'REPLACED_END'.")
	para(doc, "This paragraph ends with ENDWORD")

	// ================================================================
	// §19 — No-op: old == ""
	// ================================================================
	heading(doc, "19. No-Op: Empty Search String", 1)
	check(doc, "ReplaceText(\"\", ...) returns 0. 'should_never_appear' must NOT appear. Paragraph unchanged.")
	para(doc, "Nothing should change here. Absolutely identical in both files.")

	// ================================================================
	// §20 — No-op: old == new
	// ================================================================
	heading(doc, "20. No-Op: old == new", 1)
	check(doc, "ReplaceText(\"NOOP_SAME\", \"NOOP_SAME\") returns 0. No XML modification.")
	para(doc, "The word NOOP_SAME should remain as-is in both files.")

	// ================================================================
	// §21 — Comment-annotated text
	// ================================================================
	heading(doc, "21. Comment-Annotated Text", 1)
	check(doc, "'COMMENTED_TEXT' has a comment attached. After → 'COMMENT_REPLACED'. "+
		"Comment markers (commentRangeStart/End) and the comment body must survive.")

	p21, _ := doc.AddParagraph("")
	p21run, _ := p21.AddRun("COMMENTED_TEXT")
	initials := "VR"
	if _, err := doc.AddComment([]*docx.Run{p21run}, "This is a test comment — must survive replacement", "Visual Regtest", &initials); err != nil {
		return nil, fmt.Errorf("adding comment: %w", err)
	}

	// ================================================================
	// §22 — Multiple different placeholders in one paragraph
	// ================================================================
	heading(doc, "22. Multiple Different Placeholders in One Paragraph", 1)
	check(doc, "Three separate ReplaceText calls each hit this paragraph. "+
		"All three placeholders replaced, surrounding text intact.")
	para(doc, "Dear {{NAME}}, your order from {{COMPANY}} on {{DATE}} is confirmed.")

	// ================================================================
	// §23 — Unchanged paragraph
	// ================================================================
	heading(doc, "23. Unchanged Paragraph (No Match)", 1)
	check(doc, "No markers here. Must be byte-identical in both files.")
	para(doc, "The quick brown fox jumps over the lazy dog. 0123456789. No markers whatsoever.")

	// ================================================================
	// §24 — Replacement inside one formatted run, sibling runs untouched
	// ================================================================
	heading(doc, "24. Replacement Inside One Formatted Run", 1)
	check(doc, "'CELL_OLD' sits inside a bold run followed by italic text. "+
		"After → 'CELL_NEW'. Bold stays bold, italic stays italic.")

	p24, _ := doc.AddParagraph("")
	fr1, _ := p24.AddRun("Bold CELL_OLD text")
	_ = fr1.SetBold(boolPtr(true))
	fr2, _ := p24.AddRun(" then italic text")
	_ = fr2.SetItalic(boolPtr(true))

	return doc, nil
}

// ============================================================================
// Shorthand helpers
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

// check adds a green "✓ Check: " prefix followed by gray description text.
func check(doc *docx.Document, text string) {
	p, err := doc.AddParagraph("")
	if err != nil {
		log.Fatalf("AddParagraph: %v", err)
	}
	r1, _ := p.AddRun("✓ Check: ")
	_ = r1.SetBold(boolPtr(true))
	green := docx.NewRGBColor(0x00, 0x80, 0x00)
	_ = r1.Font().Color().SetRGB(&green)

	r2, _ := p.AddRun(text)
	gray := docx.NewRGBColor(0x66, 0x66, 0x66)
	_ = r2.Font().Color().SetRGB(&gray)
}

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
