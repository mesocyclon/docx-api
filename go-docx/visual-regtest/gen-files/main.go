// gen-files generates .docx documents from scratch using the public API.
//
// Each gen* function produces one standalone document that exercises a
// specific area of the library. The output is written to --output.
//
// Run:  go run ./visual-regtest/gen-files --output ./visual-regtest/gen-files/out
// Or:   make gen-files   (from the visual-regtest directory)
package main

import (
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/vortex/go-docx/pkg/docx"
	"github.com/vortex/go-docx/pkg/docx/enum"
)

// TestCase is one generated document.
type TestCase struct {
	Name string                       // output filename (without .docx)
	Gen  func() (*docx.Document, error) // generator
}

// FileResult captures the outcome of one generation.
type FileResult struct {
	Name    string `json:"name"`
	OK      bool   `json:"ok"`
	Error   string `json:"error,omitempty"`
	Elapsed string `json:"elapsed"`
}

// helpers
func boolPtr(v bool) *bool                { return &v }
func intPtr(v int) *int                   { return &v }
func strPtr(v string) *string             { return &v }
func int64Ptr(v int64) *int64             { return &v }

func main() {
	outputDir := flag.String("output", "", "directory for generated .docx files")
	flag.Parse()

	if *outputDir == "" {
		log.Fatal("--output is required")
	}
	if err := os.MkdirAll(*outputDir, 0o755); err != nil {
		log.Fatalf("creating output dir: %v", err)
	}

	tests := []TestCase{
		{"01_headings", genHeadings},
		{"02_paragraph_styles", genParagraphStyles},
		{"03_font_bold_italic_underline", genFontBasic},
		{"04_font_advanced", genFontAdvanced},
		{"05_font_color", genFontColor},
		{"06_font_size", genFontSize},
		{"07_paragraph_alignment", genParagraphAlignment},
		{"08_paragraph_indent", genParagraphIndent},
		{"09_paragraph_spacing", genParagraphSpacing},
		{"10_line_spacing", genLineSpacing},
		{"11_tab_stops", genTabStops},
		{"12_page_breaks", genPageBreaks},
		{"13_run_breaks", genRunBreaks},
		{"14_table_basic", genTableBasic},
		{"15_table_merged_cells", genTableMergedCells},
		{"16_table_alignment", genTableAlignment},
		{"17_table_nested", genTableNested},
		{"18_table_cell_valign", genTableCellVerticalAlign},
		{"19_table_add_row_col", genTableAddRowCol},
		{"20_sections_multi", genSectionsMulti},
		{"21_section_landscape", genSectionLandscape},
		{"22_section_margins", genSectionMargins},
		{"23_header_footer", genHeaderFooter},
		{"24_header_footer_first_page", genHeaderFooterFirstPage},
		{"25_comments", genComments},
		{"26_core_properties", genCoreProperties},
		{"27_custom_styles", genCustomStyles},
		{"28_mixed_content", genMixedContent},
		{"29_paragraph_format_flow", genParagraphFormatFlow},
		{"30_settings_odd_even", genSettingsOddEven},
		{"31_font_highlight", genFontHighlight},
		{"32_font_name", genFontName},
		{"33_underline_styles", genUnderlineStyles},
		{"34_table_row_height", genTableRowHeight},
		{"35_table_column_width", genTableColumnWidth},
		{"36_section_header_distance", genSectionHeaderDistance},
		{"37_insert_paragraph_before", genInsertParagraphBefore},
		{"38_inline_image", genInlineImage},
		{"39_multiple_runs", genMultipleRuns},
		{"40_paragraph_clear_set_text", genParagraphClearSetText},
		{"41_table_cell_set_text", genTableCellSetText},
		{"42_table_style", genTableStyle},
		{"43_table_bidi", genTableBidi},
		{"44_section_continuous_break", genSectionContinuousBreak},
		{"45_font_subscript_superscript", genFontSubSuperscript},
		{"46_tab_and_newline_in_text", genTabAndNewlineInText},
		{"47_large_document", genLargeDocument},
	}

	var results []FileResult
	for _, tc := range tests {
		start := time.Now()
		fname := tc.Name + ".docx"
		dstPath := filepath.Join(*outputDir, fname)

		doc, err := tc.Gen()
		if err != nil {
			results = append(results, FileResult{Name: fname, OK: false, Error: fmt.Sprintf("gen: %v", err), Elapsed: time.Since(start).String()})
			log.Printf("FAIL %s: %v", fname, err)
			continue
		}
		if err := doc.SaveFile(dstPath); err != nil {
			results = append(results, FileResult{Name: fname, OK: false, Error: fmt.Sprintf("save: %v", err), Elapsed: time.Since(start).String()})
			log.Printf("FAIL %s: save: %v", fname, err)
			continue
		}
		results = append(results, FileResult{Name: fname, OK: true, Elapsed: time.Since(start).String()})
		log.Printf("OK   %s (%s)", fname, time.Since(start))
	}

	// Write manifest.
	manifestPath := filepath.Join(*outputDir, "manifest.json")
	data, _ := json.MarshalIndent(results, "", "  ")
	if err := os.WriteFile(manifestPath, data, 0o644); err != nil {
		log.Fatalf("writing manifest: %v", err)
	}

	okCount := 0
	for _, r := range results {
		if r.OK {
			okCount++
		}
	}
	log.Printf("done: %d/%d succeeded", okCount, len(results))
	if okCount != len(results) {
		os.Exit(1)
	}
}

// ============================================================================
// Test generators
// ============================================================================

// 01 — All heading levels 0–9
func genHeadings() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	for level := 0; level <= 9; level++ {
		text := fmt.Sprintf("Heading Level %d", level)
		if level == 0 {
			text = "Document Title (Level 0)"
		}
		if _, err := doc.AddHeading(text, level); err != nil {
			return nil, fmt.Errorf("heading level %d: %w", level, err)
		}
		if _, err := doc.AddParagraph(fmt.Sprintf("Body text after heading level %d. Lorem ipsum dolor sit amet.", level)); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 02 — Paragraphs with built-in styles
func genParagraphStyles() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	styles := []string{
		"Normal", "Title", "Subtitle",
		"Heading 1", "Heading 2", "Heading 3",
		"List Bullet", "List Number",
		"Quote", "Intense Quote",
		"No Spacing",
	}
	for _, s := range styles {
		if _, err := doc.AddParagraph(
			fmt.Sprintf("This paragraph uses the %q style.", s),
			docx.StyleName(s),
		); err != nil {
			return nil, fmt.Errorf("style %q: %w", s, err)
		}
	}
	return doc, nil
}

// 03 — Bold, Italic, Underline on runs
func genFontBasic() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Formatting: Bold / Italic / Underline", 1); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}

	// Bold
	r1, err := para.AddRun("Bold text. ")
	if err != nil {
		return nil, err
	}
	if err := r1.SetBold(boolPtr(true)); err != nil {
		return nil, err
	}

	// Italic
	r2, err := para.AddRun("Italic text. ")
	if err != nil {
		return nil, err
	}
	if err := r2.SetItalic(boolPtr(true)); err != nil {
		return nil, err
	}

	// Underline
	r3, err := para.AddRun("Underlined text. ")
	if err != nil {
		return nil, err
	}
	u := docx.UnderlineSingle()
	if err := r3.SetUnderline(&u); err != nil {
		return nil, err
	}

	// Bold + Italic
	r4, err := para.AddRun("Bold & Italic. ")
	if err != nil {
		return nil, err
	}
	if err := r4.SetBold(boolPtr(true)); err != nil {
		return nil, err
	}
	if err := r4.SetItalic(boolPtr(true)); err != nil {
		return nil, err
	}

	// Bold + Italic + Underline
	r5, err := para.AddRun("Bold, Italic & Underlined.")
	if err != nil {
		return nil, err
	}
	if err := r5.SetBold(boolPtr(true)); err != nil {
		return nil, err
	}
	if err := r5.SetItalic(boolPtr(true)); err != nil {
		return nil, err
	}
	if err := r5.SetUnderline(&u); err != nil {
		return nil, err
	}

	return doc, nil
}

// 04 — Advanced font properties: strikethrough, all-caps, small-caps, shadow, emboss, outline, hidden
func genFontAdvanced() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Advanced Font Properties", 1); err != nil {
		return nil, err
	}

	type fontTest struct {
		label  string
		apply  func(f *docx.Font) error
	}
	tests := []fontTest{
		{"Strikethrough", func(f *docx.Font) error { return f.SetStrike(boolPtr(true)) }},
		{"Double Strikethrough", func(f *docx.Font) error { return f.SetDoubleStrike(boolPtr(true)) }},
		{"ALL CAPS", func(f *docx.Font) error { return f.SetAllCaps(boolPtr(true)) }},
		{"Small Caps", func(f *docx.Font) error { return f.SetSmallCaps(boolPtr(true)) }},
		{"Shadow", func(f *docx.Font) error { return f.SetShadow(boolPtr(true)) }},
		{"Emboss", func(f *docx.Font) error { return f.SetEmboss(boolPtr(true)) }},
		{"Outline", func(f *docx.Font) error { return f.SetOutline(boolPtr(true)) }},
		{"Imprint (engrave)", func(f *docx.Font) error { return f.SetImprint(boolPtr(true)) }},
		{"Hidden text", func(f *docx.Font) error { return f.SetHidden(boolPtr(true)) }},
	}

	for _, tt := range tests {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		r, err := para.AddRun(tt.label)
		if err != nil {
			return nil, err
		}
		if err := tt.apply(r.Font()); err != nil {
			return nil, fmt.Errorf("%s: %w", tt.label, err)
		}
	}
	return doc, nil
}

// 05 — Font color (RGB)
func genFontColor() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Colors", 1); err != nil {
		return nil, err
	}

	colors := []struct {
		label string
		r, g, b byte
	}{
		{"Red text", 0xFF, 0x00, 0x00},
		{"Green text", 0x00, 0x80, 0x00},
		{"Blue text", 0x00, 0x00, 0xFF},
		{"Orange text", 0xFF, 0xA5, 0x00},
		{"Purple text", 0x80, 0x00, 0x80},
		{"Dark cyan text", 0x00, 0x80, 0x80},
	}

	for _, c := range colors {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(c.label)
		if err != nil {
			return nil, err
		}
		rgb := docx.NewRGBColor(c.r, c.g, c.b)
		if err := run.Font().Color().SetRGB(&rgb); err != nil {
			return nil, err
		}
	}

	// Theme color
	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	run, err := para.AddRun("Theme color: Accent1")
	if err != nil {
		return nil, err
	}
	tc := enum.MsoThemeColorIndexAccent1
	if err := run.Font().Color().SetThemeColor(&tc); err != nil {
		return nil, err
	}

	return doc, nil
}

// 06 — Font sizes
func genFontSize() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Sizes", 1); err != nil {
		return nil, err
	}

	sizes := []float64{8, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72}
	for _, sz := range sizes {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(fmt.Sprintf("%.0fpt text", sz))
		if err != nil {
			return nil, err
		}
		length := docx.Pt(sz)
		if err := run.Font().SetSize(&length); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 07 — Paragraph alignment
func genParagraphAlignment() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Alignment", 1); err != nil {
		return nil, err
	}

	aligns := []struct {
		name string
		val  enum.WdParagraphAlignment
	}{
		{"Left aligned", enum.WdParagraphAlignmentLeft},
		{"Center aligned", enum.WdParagraphAlignmentCenter},
		{"Right aligned", enum.WdParagraphAlignmentRight},
		{"Justified – Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.", enum.WdParagraphAlignmentJustify},
		{"Distributed – Short text fills full width", enum.WdParagraphAlignmentDistribute},
	}

	for _, a := range aligns {
		para, err := doc.AddParagraph(a.name)
		if err != nil {
			return nil, err
		}
		if err := para.SetAlignment(&a.val); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 08 — Paragraph indentation
func genParagraphIndent() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Indentation", 1); err != nil {
		return nil, err
	}

	// Left indent
	p1, err := doc.AddParagraph("Left indent 720 twips (0.5 inch)")
	if err != nil {
		return nil, err
	}
	if err := p1.ParagraphFormat().SetLeftIndent(intPtr(720)); err != nil {
		return nil, err
	}

	// Right indent
	p2, err := doc.AddParagraph("Right indent 720 twips (0.5 inch)")
	if err != nil {
		return nil, err
	}
	if err := p2.ParagraphFormat().SetRightIndent(intPtr(720)); err != nil {
		return nil, err
	}

	// Both left + right indent
	p3, err := doc.AddParagraph("Left + Right indent 1440 twips (1 inch each)")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetLeftIndent(intPtr(1440)); err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetRightIndent(intPtr(1440)); err != nil {
		return nil, err
	}

	// First-line indent
	p4, err := doc.AddParagraph("First-line indent 360 twips (0.25 inch). Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
	if err != nil {
		return nil, err
	}
	if err := p4.ParagraphFormat().SetFirstLineIndent(intPtr(360)); err != nil {
		return nil, err
	}

	// Hanging indent (negative first-line)
	p5, err := doc.AddParagraph("Hanging indent: left=720, firstLine=-360. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt.")
	if err != nil {
		return nil, err
	}
	if err := p5.ParagraphFormat().SetLeftIndent(intPtr(720)); err != nil {
		return nil, err
	}
	if err := p5.ParagraphFormat().SetFirstLineIndent(intPtr(-360)); err != nil {
		return nil, err
	}

	return doc, nil
}

// 09 — Space before/after
func genParagraphSpacing() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Spacing", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("Space before = 480 twips (24pt)")
	if err != nil {
		return nil, err
	}
	if err := p1.ParagraphFormat().SetSpaceBefore(intPtr(480)); err != nil {
		return nil, err
	}

	p2, err := doc.AddParagraph("Space after = 480 twips (24pt)")
	if err != nil {
		return nil, err
	}
	if err := p2.ParagraphFormat().SetSpaceAfter(intPtr(480)); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("Space before=240 and after=240 (12pt each)")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetSpaceBefore(intPtr(240)); err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetSpaceAfter(intPtr(240)); err != nil {
		return nil, err
	}

	if _, err := doc.AddParagraph("Normal spacing after."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 10 — Line spacing
func genLineSpacing() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Line Spacing", 1); err != nil {
		return nil, err
	}

	loremShort := "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation."

	// Single
	p1, err := doc.AddParagraph("Single spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls1 := docx.LineSpacingMultiple(1.0)
	if err := p1.ParagraphFormat().SetLineSpacing(&ls1); err != nil {
		return nil, err
	}

	// 1.5
	p2, err := doc.AddParagraph("1.5 line spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls2 := docx.LineSpacingMultiple(1.5)
	if err := p2.ParagraphFormat().SetLineSpacing(&ls2); err != nil {
		return nil, err
	}

	// Double
	p3, err := doc.AddParagraph("Double line spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls3 := docx.LineSpacingMultiple(2.0)
	if err := p3.ParagraphFormat().SetLineSpacing(&ls3); err != nil {
		return nil, err
	}

	// Exact 18pt = 360 twips
	p4, err := doc.AddParagraph("Exact 18pt spacing: " + loremShort)
	if err != nil {
		return nil, err
	}
	ls4 := docx.LineSpacingTwips(360)
	if err := p4.ParagraphFormat().SetLineSpacing(&ls4); err != nil {
		return nil, err
	}

	// Using SetLineSpacingRule
	p5, err := doc.AddParagraph("SetLineSpacingRule(Double): " + loremShort)
	if err != nil {
		return nil, err
	}
	if err := p5.ParagraphFormat().SetLineSpacingRule(enum.WdLineSpacingDouble); err != nil {
		return nil, err
	}

	return doc, nil
}

// 11 — Tab stops
func genTabStops() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Tab Stops", 1); err != nil {
		return nil, err
	}

	// Left tab at 2 inches
	p1, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts1 := p1.ParagraphFormat().TabStops()
	if _, err := ts1.AddTabStop(2880, enum.WdTabAlignmentLeft, enum.WdTabLeaderSpaces); err != nil {
		return nil, err
	}
	r1, err := p1.AddRun("Before tab")
	if err != nil {
		return nil, err
	}
	r1.AddTab()
	r1.AddText("After left tab at 2\"")

	// Center tab at 3.25 inches with dot leader
	p2, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts2 := p2.ParagraphFormat().TabStops()
	if _, err := ts2.AddTabStop(4680, enum.WdTabAlignmentCenter, enum.WdTabLeaderDots); err != nil {
		return nil, err
	}
	r2, err := p2.AddRun("Item")
	if err != nil {
		return nil, err
	}
	r2.AddTab()
	r2.AddText("Centered with dots")

	// Right tab at 6 inches with dash leader
	p3, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts3 := p3.ParagraphFormat().TabStops()
	if _, err := ts3.AddTabStop(8640, enum.WdTabAlignmentRight, enum.WdTabLeaderDashes); err != nil {
		return nil, err
	}
	r3, err := p3.AddRun("Left text")
	if err != nil {
		return nil, err
	}
	r3.AddTab()
	r3.AddText("$99.99")

	// Decimal tab
	p4, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	ts4 := p4.ParagraphFormat().TabStops()
	if _, err := ts4.AddTabStop(4320, enum.WdTabAlignmentDecimal, enum.WdTabLeaderSpaces); err != nil {
		return nil, err
	}
	r4, err := p4.AddRun("")
	if err != nil {
		return nil, err
	}
	r4.AddTab()
	r4.AddText("123.456 (decimal tab)")

	return doc, nil
}

// 12 — Page breaks
func genPageBreaks() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Page 1", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content on page 1."); err != nil {
		return nil, err
	}

	// AddPageBreak
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Page 2 (after AddPageBreak)", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content on page 2."); err != nil {
		return nil, err
	}

	// PageBreakBefore paragraph property
	p3, err := doc.AddParagraph("Page 3 — via PageBreakBefore property. This paragraph should start on a new page.")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetPageBreakBefore(boolPtr(true)); err != nil {
		return nil, err
	}

	return doc, nil
}

// 13 — Run-level breaks
func genRunBreaks() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Run-Level Breaks", 1); err != nil {
		return nil, err
	}

	// Line break
	p1, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	r1, err := p1.AddRun("Before line break")
	if err != nil {
		return nil, err
	}
	if err := r1.AddBreak(enum.WdBreakTypeLine); err != nil {
		return nil, err
	}
	r1.AddText("After line break (same paragraph)")

	// Column break
	p2, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	r2, err := p2.AddRun("Before column break")
	if err != nil {
		return nil, err
	}
	if err := r2.AddBreak(enum.WdBreakTypeColumn); err != nil {
		return nil, err
	}
	r2.AddText("After column break")

	// Page break in a run
	p3, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	r3, err := p3.AddRun("Before page break (in run)")
	if err != nil {
		return nil, err
	}
	if err := r3.AddBreak(enum.WdBreakTypePage); err != nil {
		return nil, err
	}
	r3.AddText("After page break (same run, new page)")

	return doc, nil
}

// 14 — Basic table
func genTableBasic() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Basic Table (3x3)", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(3, 3)
	if err != nil {
		return nil, err
	}
	for r := 0; r < 3; r++ {
		for c := 0; c < 3; c++ {
			cell, err := tbl.CellAt(r, c)
			if err != nil {
				return nil, err
			}
			cell.SetText(fmt.Sprintf("Row %d, Col %d", r+1, c+1))
		}
	}

	if _, err := doc.AddParagraph("Paragraph after the table."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 15 — Table with merged cells
func genTableMergedCells() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table with Merged Cells", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(4, 4)
	if err != nil {
		return nil, err
	}
	// Label all cells first
	for r := 0; r < 4; r++ {
		for c := 0; c < 4; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("(%d,%d)", r, c))
		}
	}

	// Horizontal merge: row 0, cols 0–1
	c00, _ := tbl.CellAt(0, 0)
	c01, _ := tbl.CellAt(0, 1)
	merged, err := c00.Merge(c01)
	if err != nil {
		return nil, fmt.Errorf("h-merge: %w", err)
	}
	merged.SetText("Merged (0,0)-(0,1)")

	// Vertical merge: col 3, rows 1–3
	c13, _ := tbl.CellAt(1, 3)
	c33, _ := tbl.CellAt(3, 3)
	merged2, err := c13.Merge(c33)
	if err != nil {
		return nil, fmt.Errorf("v-merge: %w", err)
	}
	merged2.SetText("Merged (1,3)-(3,3)")

	// Block merge: rows 2–3, cols 0–1
	c20, _ := tbl.CellAt(2, 0)
	c31, _ := tbl.CellAt(3, 1)
	merged3, err := c20.Merge(c31)
	if err != nil {
		return nil, fmt.Errorf("block-merge: %w", err)
	}
	merged3.SetText("Block merge (2,0)-(3,1)")

	return doc, nil
}

// 16 — Table alignment (left, center, right)
func genTableAlignment() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Alignment", 1); err != nil {
		return nil, err
	}

	aligns := []struct {
		label string
		val   enum.WdTableAlignment
	}{
		{"Left", enum.WdTableAlignmentLeft},
		{"Center", enum.WdTableAlignmentCenter},
		{"Right", enum.WdTableAlignmentRight},
	}

	for _, a := range aligns {
		if _, err := doc.AddParagraph(a.label + " aligned table:"); err != nil {
			return nil, err
		}
		tbl, err := doc.AddTable(2, 2)
		if err != nil {
			return nil, err
		}
		if err := tbl.SetAlignment(&a.val); err != nil {
			return nil, err
		}
		if err := tbl.SetAutofit(false); err != nil {
			return nil, err
		}
		for r := 0; r < 2; r++ {
			for c := 0; c < 2; c++ {
				cell, _ := tbl.CellAt(r, c)
				cell.SetText(fmt.Sprintf("%s R%dC%d", a.label, r, c))
			}
		}
	}
	return doc, nil
}

// 17 — Nested table (table inside a cell)
func genTableNested() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Nested Table", 1); err != nil {
		return nil, err
	}

	outer, err := doc.AddTable(2, 2)
	if err != nil {
		return nil, err
	}
	c00, _ := outer.CellAt(0, 0)
	c00.SetText("Outer (0,0) — has nested table below")

	inner, err := c00.AddTable(2, 2)
	if err != nil {
		return nil, fmt.Errorf("nested table: %w", err)
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := inner.CellAt(r, c)
			cell.SetText(fmt.Sprintf("Inner %d,%d", r, c))
		}
	}

	c01, _ := outer.CellAt(0, 1)
	c01.SetText("Outer (0,1)")
	c10, _ := outer.CellAt(1, 0)
	c10.SetText("Outer (1,0)")
	c11, _ := outer.CellAt(1, 1)
	c11.SetText("Outer (1,1)")

	return doc, nil
}

// 18 — Table cell vertical alignment
func genTableCellVerticalAlign() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Cell Vertical Alignment", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(1, 3)
	if err != nil {
		return nil, err
	}

	// Set row height
	rows := tbl.Rows()
	row, _ := rows.Get(0)
	if err := row.SetHeight(intPtr(1440)); err != nil { // 1 inch
		return nil, err
	}
	rule := enum.WdRowHeightRuleExactly
	if err := row.SetHeightRule(&rule); err != nil {
		return nil, err
	}

	valigns := []struct {
		label string
		val   enum.WdCellVerticalAlignment
	}{
		{"Top", enum.WdCellVerticalAlignmentTop},
		{"Center", enum.WdCellVerticalAlignmentCenter},
		{"Bottom", enum.WdCellVerticalAlignmentBottom},
	}

	for i, va := range valigns {
		cell, _ := tbl.CellAt(0, i)
		cell.SetText(va.label)
		if err := cell.SetVerticalAlignment(&va.val); err != nil {
			return nil, err
		}
	}

	return doc, nil
}

// 19 — Table: AddRow / AddColumn
func genTableAddRowCol() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table — AddRow / AddColumn", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 2)
	if err != nil {
		return nil, err
	}
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("Original R%dC%d", r, c))
		}
	}

	// Add row
	newRow, err := tbl.AddRow()
	if err != nil {
		return nil, err
	}
	cells := newRow.Cells()
	for i, c := range cells {
		c.SetText(fmt.Sprintf("New row C%d", i))
	}

	// Add column
	if _, err := tbl.AddColumn(1440); err != nil {
		return nil, err
	}
	// Fill new column
	for r := 0; r < 3; r++ {
		cell, err := tbl.CellAt(r, 2)
		if err != nil {
			continue
		}
		cell.SetText(fmt.Sprintf("New col R%d", r))
	}

	return doc, nil
}

// 20 — Multiple sections
func genSectionsMulti() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Section 1", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content in section 1."); err != nil {
		return nil, err
	}

	// New page section
	if _, err := doc.AddSection(enum.WdSectionStartNewPage); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Section 2 (New Page)", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content in section 2."); err != nil {
		return nil, err
	}

	// Odd page section
	if _, err := doc.AddSection(enum.WdSectionStartOddPage); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Section 3 (Odd Page)", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content in section 3."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 21 — Landscape section
func genSectionLandscape() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	if sections.Len() > 0 {
		sect, _ := sections.Get(sections.Len() - 1)
		if err := sect.SetOrientation(enum.WdOrientationLandscape); err != nil {
			return nil, err
		}
		// Swap width and height for landscape
		if err := sect.SetPageWidth(intPtr(docx.Inches(11).Twips())); err != nil {
			return nil, err
		}
		if err := sect.SetPageHeight(intPtr(docx.Inches(8.5).Twips())); err != nil {
			return nil, err
		}
	}

	if _, err := doc.AddHeading("Landscape Page", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This entire document is in landscape orientation."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 22 — Custom margins
func genSectionMargins() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)
	if err := sect.SetTopMargin(intPtr(docx.Inches(2).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetBottomMargin(intPtr(docx.Inches(2).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetLeftMargin(intPtr(docx.Inches(1.5).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetRightMargin(intPtr(docx.Inches(1.5).Twips())); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Custom Margins", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Top=2\", Bottom=2\", Left=1.5\", Right=1.5\""); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 23 — Header and Footer
func genHeaderFooter() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)

	// Header
	header := sect.Header()
	if _, err := header.AddParagraph("Header: Document Title — Confidential"); err != nil {
		return nil, err
	}

	// Footer
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Footer: Page X of Y — © 2025 Company"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Document with Header and Footer", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Check the header and footer areas of this document."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 24 — First-page header/footer
func genHeaderFooterFirstPage() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)
	if err := sect.SetDifferentFirstPageHeaderFooter(true); err != nil {
		return nil, err
	}

	// Primary header/footer
	header := sect.Header()
	if _, err := header.AddParagraph("Primary Header (pages 2+)"); err != nil {
		return nil, err
	}
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Primary Footer (pages 2+)"); err != nil {
		return nil, err
	}

	// First page header/footer
	firstHeader := sect.FirstPageHeader()
	if _, err := firstHeader.AddParagraph("FIRST PAGE HEADER"); err != nil {
		return nil, err
	}
	firstFooter := sect.FirstPageFooter()
	if _, err := firstFooter.AddParagraph("FIRST PAGE FOOTER"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("First Page — Different Header/Footer", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This is the first page."); err != nil {
		return nil, err
	}
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Second Page", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This is the second page with the primary header/footer."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 25 — Comments
func genComments() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Document with Comments", 1); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	run1, err := para.AddRun("This text has a comment. ")
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddComment([]*docx.Run{run1}, "This is an important comment.", "Reviewer", strPtr("R")); err != nil {
		return nil, err
	}

	run2, err := para.AddRun("This text also has a comment.")
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddComment([]*docx.Run{run2}, "Another comment by a different author.", "Editor", strPtr("E")); err != nil {
		return nil, err
	}

	return doc, nil
}

// 26 — Core properties
func genCoreProperties() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	cp, err := doc.CoreProperties()
	if err != nil {
		return nil, err
	}
	if err := cp.SetAuthor("Jane Doe"); err != nil {
		return nil, err
	}
	if err := cp.SetTitle("Test Document — Core Properties"); err != nil {
		return nil, err
	}
	if err := cp.SetSubject("QA Testing"); err != nil {
		return nil, err
	}
	if err := cp.SetKeywords("test, docx, go-docx, core-properties"); err != nil {
		return nil, err
	}
	if err := cp.SetCategory("Testing"); err != nil {
		return nil, err
	}
	if err := cp.SetComments("Generated by gen-files visual regression test."); err != nil {
		return nil, err
	}
	if err := cp.SetLastModifiedBy("Build System"); err != nil {
		return nil, err
	}
	if err := cp.SetLanguage("en-US"); err != nil {
		return nil, err
	}
	if err := cp.SetVersion("1.0.0"); err != nil {
		return nil, err
	}
	if err := cp.SetRevision(42); err != nil {
		return nil, err
	}
	cp.SetCreated(time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC))
	cp.SetModified(time.Date(2025, 6, 15, 12, 0, 0, 0, time.UTC))

	if _, err := doc.AddHeading("Core Properties Test", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Open File → Properties to verify metadata."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 27 — Custom styles
func genCustomStyles() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	styles, err := doc.Styles()
	if err != nil {
		return nil, err
	}

	// Custom paragraph style
	ps, err := styles.AddStyle("CustomParagraph", enum.WdStyleTypeParagraph, false)
	if err != nil {
		return nil, err
	}
	sz := docx.Pt(14)
	if err := ps.Font().SetSize(&sz); err != nil {
		return nil, err
	}
	if err := ps.Font().SetBold(boolPtr(true)); err != nil {
		return nil, err
	}
	rgb := docx.NewRGBColor(0x00, 0x66, 0xCC)
	if err := ps.Font().Color().SetRGB(&rgb); err != nil {
		return nil, err
	}
	if err := ps.ParagraphFormat().SetSpaceBefore(intPtr(240)); err != nil {
		return nil, err
	}
	if err := ps.ParagraphFormat().SetSpaceAfter(intPtr(120)); err != nil {
		return nil, err
	}

	// Custom character style
	cs, err := styles.AddStyle("CustomChar", enum.WdStyleTypeCharacter, false)
	if err != nil {
		return nil, err
	}
	if err := cs.Font().SetItalic(boolPtr(true)); err != nil {
		return nil, err
	}
	rgb2 := docx.NewRGBColor(0xCC, 0x00, 0x00)
	if err := cs.Font().Color().SetRGB(&rgb2); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Custom Styles", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Using custom paragraph style:", docx.StyleName("CustomParagraph")); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Another paragraph with custom style.", docx.StyleName("CustomParagraph")); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := para.AddRun("Normal text with "); err != nil {
		return nil, err
	}
	if _, err := para.AddRun("custom character style", docx.StyleName("CustomChar")); err != nil {
		return nil, err
	}
	if _, err := para.AddRun(" applied."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 28 — Mixed content: headings + paragraphs + tables + page breaks
func genMixedContent() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	// Title
	if _, err := doc.AddHeading("Mixed Content Document", 0); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This document exercises mixed content types."); err != nil {
		return nil, err
	}

	// Section: text
	if _, err := doc.AddHeading("Text Section", 1); err != nil {
		return nil, err
	}
	for i := 1; i <= 3; i++ {
		if _, err := doc.AddParagraph(fmt.Sprintf("Paragraph %d: Lorem ipsum dolor sit amet.", i)); err != nil {
			return nil, err
		}
	}

	// Section: table
	if _, err := doc.AddHeading("Table Section", 1); err != nil {
		return nil, err
	}
	tbl, err := doc.AddTable(3, 4)
	if err != nil {
		return nil, err
	}
	// Header row
	headers := []string{"Name", "Age", "City", "Score"}
	for i, h := range headers {
		cell, _ := tbl.CellAt(0, i)
		cell.SetText(h)
	}
	data := [][]string{
		{"Alice", "30", "New York", "95"},
		{"Bob", "25", "London", "87"},
	}
	for r, row := range data {
		for c, val := range row {
			cell, _ := tbl.CellAt(r+1, c)
			cell.SetText(val)
		}
	}

	if _, err := doc.AddParagraph(""); err != nil {
		return nil, err
	}

	// Page break then more content
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("After Page Break", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Final content after page break."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 29 — Paragraph format: widow control, keep together, keep with next
func genParagraphFormatFlow() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Flow Control", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("WidowControl = true: This paragraph has widow/orphan control enabled. Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
	if err != nil {
		return nil, err
	}
	if err := p1.ParagraphFormat().SetWidowControl(boolPtr(true)); err != nil {
		return nil, err
	}

	p2, err := doc.AddParagraph("KeepTogether = true: This paragraph's lines are kept together on the same page. Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
	if err != nil {
		return nil, err
	}
	if err := p2.ParagraphFormat().SetKeepTogether(boolPtr(true)); err != nil {
		return nil, err
	}

	p3, err := doc.AddParagraph("KeepWithNext = true: This paragraph stays with the next.")
	if err != nil {
		return nil, err
	}
	if err := p3.ParagraphFormat().SetKeepWithNext(boolPtr(true)); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("This paragraph follows the keep-with-next paragraph."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 30 — Settings: odd/even page header/footer
func genSettingsOddEven() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	settings, err := doc.Settings()
	if err != nil {
		return nil, err
	}
	if err := settings.SetOddAndEvenPagesHeaderFooter(true); err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)

	header := sect.Header()
	if _, err := header.AddParagraph("Odd Page Header"); err != nil {
		return nil, err
	}
	evenHeader := sect.EvenPageHeader()
	if _, err := evenHeader.AddParagraph("Even Page Header"); err != nil {
		return nil, err
	}
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Odd Page Footer"); err != nil {
		return nil, err
	}
	evenFooter := sect.EvenPageFooter()
	if _, err := evenFooter.AddParagraph("Even Page Footer"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Odd/Even Headers/Footers", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Page 1 (odd)"); err != nil {
		return nil, err
	}
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Page 2 (even)"); err != nil {
		return nil, err
	}
	if _, err := doc.AddPageBreak(); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Page 3 (odd)"); err != nil {
		return nil, err
	}

	return doc, nil
}

// 31 — Font highlight colors
func genFontHighlight() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Highlight Colors", 1); err != nil {
		return nil, err
	}

	highlights := []struct {
		label string
		color enum.WdColorIndex
	}{
		{"Yellow highlight", enum.WdColorIndexYellow},
		{"Green highlight", enum.WdColorIndexBrightGreen},
		{"Turquoise highlight", enum.WdColorIndexTurquoise},
		{"Pink highlight", enum.WdColorIndexPink},
		{"Blue highlight", enum.WdColorIndexBlue},
		{"Red highlight", enum.WdColorIndexRed},
		{"Dark Blue highlight", enum.WdColorIndexDarkBlue},
		{"Teal highlight", enum.WdColorIndexTeal},
		{"Violet highlight", enum.WdColorIndexViolet},
		{"Gray 50% highlight", enum.WdColorIndexGray50},
		{"Gray 25% highlight", enum.WdColorIndexGray25},
	}

	for _, h := range highlights {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(h.label)
		if err != nil {
			return nil, err
		}
		if err := run.Font().SetHighlightColor(&h.color); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 32 — Font name
func genFontName() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Font Names", 1); err != nil {
		return nil, err
	}

	fonts := []string{
		"Arial", "Times New Roman", "Courier New",
		"Verdana", "Georgia", "Trebuchet MS",
		"Comic Sans MS", "Impact", "Calibri",
	}

	for _, f := range fonts {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(fmt.Sprintf("The quick brown fox (%s)", f))
		if err != nil {
			return nil, err
		}
		if err := run.Font().SetName(&f); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 33 — Various underline styles
func genUnderlineStyles() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Underline Styles", 1); err != nil {
		return nil, err
	}

	styles := []struct {
		label string
		val   docx.UnderlineVal
	}{
		{"Single", docx.UnderlineSingle()},
		{"Double", docx.UnderlineStyle(enum.WdUnderlineDouble)},
		{"Thick", docx.UnderlineStyle(enum.WdUnderlineThick)},
		{"Dotted", docx.UnderlineStyle(enum.WdUnderlineDotted)},
		{"Dash", docx.UnderlineStyle(enum.WdUnderlineDash)},
		{"Dot-Dash", docx.UnderlineStyle(enum.WdUnderlineDotDash)},
		{"Dot-Dot-Dash", docx.UnderlineStyle(enum.WdUnderlineDotDotDash)},
		{"Wavy", docx.UnderlineStyle(enum.WdUnderlineWavy)},
		{"Words only", docx.UnderlineStyle(enum.WdUnderlineWords)},
		{"Dash Long", docx.UnderlineStyle(enum.WdUnderlineDashLong)},
		{"Wavy Double", docx.UnderlineStyle(enum.WdUnderlineWavyDouble)},
	}

	for _, s := range styles {
		para, err := doc.AddParagraph("")
		if err != nil {
			return nil, err
		}
		run, err := para.AddRun(s.label + " underline")
		if err != nil {
			return nil, err
		}
		val := s.val
		if err := run.SetUnderline(&val); err != nil {
			return nil, err
		}
	}
	return doc, nil
}

// 34 — Table row height
func genTableRowHeight() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Row Heights", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(3, 2)
	if err != nil {
		return nil, err
	}

	heights := []struct {
		twips int
		rule  enum.WdRowHeightRule
		label string
	}{
		{360, enum.WdRowHeightRuleExactly, "Exact 0.25\""},
		{720, enum.WdRowHeightRuleAtLeast, "AtLeast 0.5\""},
		{1440, enum.WdRowHeightRuleExactly, "Exact 1.0\""},
	}

	rows := tbl.Rows()
	for i, h := range heights {
		row, _ := rows.Get(i)
		if err := row.SetHeight(intPtr(h.twips)); err != nil {
			return nil, err
		}
		if err := row.SetHeightRule(&h.rule); err != nil {
			return nil, err
		}
		cell, _ := tbl.CellAt(i, 0)
		cell.SetText(h.label)
		cell2, _ := tbl.CellAt(i, 1)
		cell2.SetText(fmt.Sprintf("%d twips", h.twips))
	}

	return doc, nil
}

// 35 — Table column width
func genTableColumnWidth() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Column Widths", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 3)
	if err != nil {
		return nil, err
	}
	if err := tbl.SetAutofit(false); err != nil {
		return nil, err
	}

	cols, _ := tbl.Columns()
	widths := []int{1440, 2880, 4320} // 1", 2", 3"
	for i, w := range widths {
		col, _ := cols.Get(i)
		if err := col.SetWidth(intPtr(w)); err != nil {
			return nil, err
		}
	}

	cell00, _ := tbl.CellAt(0, 0)
	cell00.SetText("1 inch")
	cell01, _ := tbl.CellAt(0, 1)
	cell01.SetText("2 inches")
	cell02, _ := tbl.CellAt(0, 2)
	cell02.SetText("3 inches")
	cell10, _ := tbl.CellAt(1, 0)
	cell10.SetText("narrow")
	cell11, _ := tbl.CellAt(1, 1)
	cell11.SetText("medium")
	cell12, _ := tbl.CellAt(1, 2)
	cell12.SetText("wide")

	return doc, nil
}

// 36 — Section header/footer distance
func genSectionHeaderDistance() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}

	sections := doc.Sections()
	sect, _ := sections.Get(sections.Len() - 1)
	if err := sect.SetHeaderDistance(intPtr(docx.Inches(0.3).Twips())); err != nil {
		return nil, err
	}
	if err := sect.SetFooterDistance(intPtr(docx.Inches(0.3).Twips())); err != nil {
		return nil, err
	}

	header := sect.Header()
	if _, err := header.AddParagraph("Header close to edge (0.3\")"); err != nil {
		return nil, err
	}
	footer := sect.Footer()
	if _, err := footer.AddParagraph("Footer close to edge (0.3\")"); err != nil {
		return nil, err
	}

	if _, err := doc.AddHeading("Header/Footer Distance", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Header and footer are 0.3 inches from the edge."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 37 — InsertParagraphBefore
func genInsertParagraphBefore() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("InsertParagraphBefore", 1); err != nil {
		return nil, err
	}

	p1, err := doc.AddParagraph("Paragraph A (added first)")
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Paragraph C (added second)"); err != nil {
		return nil, err
	}

	// Insert before paragraph A
	if _, err := p1.InsertParagraphBefore("Paragraph B (inserted before A)"); err != nil {
		return nil, err
	}

	return doc, nil
}

// 38 — Inline image (generated PNG)
func genInlineImage() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Inline Image", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Below is a programmatically generated image:"); err != nil {
		return nil, err
	}

	// Generate a simple 200x100 PNG with a gradient
	img := image.NewRGBA(image.Rect(0, 0, 200, 100))
	for y := 0; y < 100; y++ {
		for x := 0; x < 200; x++ {
			img.Set(x, y, color.RGBA{
				R: uint8(x * 255 / 200),
				G: uint8(y * 255 / 100),
				B: 128,
				A: 255,
			})
		}
	}
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		return nil, fmt.Errorf("encoding PNG: %w", err)
	}

	reader := bytes.NewReader(buf.Bytes())
	w := int64(docx.Inches(2).Emu())
	h := int64(docx.Inches(1).Emu())
	if _, err := doc.AddPicture(reader, &w, &h); err != nil {
		return nil, fmt.Errorf("add picture: %w", err)
	}

	if _, err := doc.AddParagraph("Image above should be 2\" × 1\" with a gradient."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 39 — Paragraph with multiple styled runs
func genMultipleRuns() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Multiple Runs in One Paragraph", 1); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}

	// Normal
	if _, err := para.AddRun("Normal "); err != nil {
		return nil, err
	}
	// Red bold
	r1, err := para.AddRun("Red Bold ")
	if err != nil {
		return nil, err
	}
	if err := r1.SetBold(boolPtr(true)); err != nil {
		return nil, err
	}
	rgb1 := docx.NewRGBColor(0xFF, 0, 0)
	if err := r1.Font().Color().SetRGB(&rgb1); err != nil {
		return nil, err
	}
	// Blue italic
	r2, err := para.AddRun("Blue Italic ")
	if err != nil {
		return nil, err
	}
	if err := r2.SetItalic(boolPtr(true)); err != nil {
		return nil, err
	}
	rgb2 := docx.NewRGBColor(0, 0, 0xFF)
	if err := r2.Font().Color().SetRGB(&rgb2); err != nil {
		return nil, err
	}
	// Large green
	r3, err := para.AddRun("Large Green ")
	if err != nil {
		return nil, err
	}
	sz := docx.Pt(24)
	if err := r3.Font().SetSize(&sz); err != nil {
		return nil, err
	}
	rgb3 := docx.NewRGBColor(0, 0x80, 0)
	if err := r3.Font().Color().SetRGB(&rgb3); err != nil {
		return nil, err
	}
	// Small underline
	r4, err := para.AddRun("Small Underline")
	if err != nil {
		return nil, err
	}
	sz2 := docx.Pt(8)
	if err := r4.Font().SetSize(&sz2); err != nil {
		return nil, err
	}
	u := docx.UnderlineSingle()
	if err := r4.SetUnderline(&u); err != nil {
		return nil, err
	}

	return doc, nil
}

// 40 — Paragraph.Clear / Paragraph.SetText
func genParagraphClearSetText() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Paragraph Clear & SetText", 1); err != nil {
		return nil, err
	}

	para, err := doc.AddParagraph("Original text that will be replaced.")
	if err != nil {
		return nil, err
	}
	// Clear and set new text
	para.Clear()
	if err := para.SetText("Replaced text via SetText()."); err != nil {
		return nil, err
	}

	// Verify the text matches
	if _, err := doc.AddParagraph(fmt.Sprintf("Paragraph text reads: %q", para.Text())); err != nil {
		return nil, err
	}

	return doc, nil
}

// 41 — Table cell SetText
func genTableCellSetText() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table Cell SetText", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 2)
	if err != nil {
		return nil, err
	}

	// First fill
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("First (%d,%d)", r, c))
		}
	}

	// Overwrite with SetText
	for r := 0; r < 2; r++ {
		for c := 0; c < 2; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("Replaced (%d,%d)", r, c))
		}
	}

	return doc, nil
}

// 42 — Table with style
func genTableStyle() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table with Style", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(4, 3, docx.StyleName("Table Grid"))
	if err != nil {
		return nil, err
	}
	headers := []string{"Product", "Quantity", "Price"}
	for i, h := range headers {
		cell, _ := tbl.CellAt(0, i)
		cell.SetText(h)
	}
	data := [][]string{
		{"Widget A", "100", "$5.99"},
		{"Widget B", "250", "$3.49"},
		{"Widget C", "50", "$12.00"},
	}
	for r, row := range data {
		for c, val := range row {
			cell, _ := tbl.CellAt(r+1, c)
			cell.SetText(val)
		}
	}

	return doc, nil
}

// 43 — Table bidi (RTL direction)
func genTableBidi() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Table BiDi (RTL)", 1); err != nil {
		return nil, err
	}

	tbl, err := doc.AddTable(2, 3)
	if err != nil {
		return nil, err
	}
	if err := tbl.SetTableDirection(boolPtr(true)); err != nil {
		return nil, err
	}

	for r := 0; r < 2; r++ {
		for c := 0; c < 3; c++ {
			cell, _ := tbl.CellAt(r, c)
			cell.SetText(fmt.Sprintf("R%dC%d", r, c))
		}
	}

	if _, err := doc.AddParagraph("Table above has RTL (bidiVisual) direction."); err != nil {
		return nil, err
	}
	return doc, nil
}

// 44 — Section with continuous break
func genSectionContinuousBreak() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Continuous Section Break", 1); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content before continuous break."); err != nil {
		return nil, err
	}

	if _, err := doc.AddSection(enum.WdSectionStartContinuous); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content after continuous section break (same page)."); err != nil {
		return nil, err
	}

	// Even page
	if _, err := doc.AddSection(enum.WdSectionStartEvenPage); err != nil {
		return nil, err
	}
	if _, err := doc.AddParagraph("Content after even-page section break."); err != nil {
		return nil, err
	}

	return doc, nil
}

// 45 — Subscript and Superscript
func genFontSubSuperscript() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Subscript & Superscript", 1); err != nil {
		return nil, err
	}

	// Superscript: E=mc²
	p1, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := p1.AddRun("E = mc"); err != nil {
		return nil, err
	}
	r1, err := p1.AddRun("2")
	if err != nil {
		return nil, err
	}
	if err := r1.Font().SetSuperscript(boolPtr(true)); err != nil {
		return nil, err
	}

	// Subscript: H₂O
	p2, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := p2.AddRun("H"); err != nil {
		return nil, err
	}
	r2, err := p2.AddRun("2")
	if err != nil {
		return nil, err
	}
	if err := r2.Font().SetSubscript(boolPtr(true)); err != nil {
		return nil, err
	}
	if _, err := p2.AddRun("O"); err != nil {
		return nil, err
	}

	// Mixed: x² + y₁
	p3, err := doc.AddParagraph("")
	if err != nil {
		return nil, err
	}
	if _, err := p3.AddRun("x"); err != nil {
		return nil, err
	}
	rSup, err := p3.AddRun("2")
	if err != nil {
		return nil, err
	}
	if err := rSup.Font().SetSuperscript(boolPtr(true)); err != nil {
		return nil, err
	}
	if _, err := p3.AddRun(" + y"); err != nil {
		return nil, err
	}
	rSub, err := p3.AddRun("1")
	if err != nil {
		return nil, err
	}
	if err := rSub.Font().SetSubscript(boolPtr(true)); err != nil {
		return nil, err
	}

	return doc, nil
}

// 46 — Tab and newline characters in text
func genTabAndNewlineInText() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Tab and Newline in Text", 1); err != nil {
		return nil, err
	}

	// Tab in paragraph text
	if _, err := doc.AddParagraph("Column1\tColumn2\tColumn3"); err != nil {
		return nil, err
	}

	// Newline in paragraph text
	if _, err := doc.AddParagraph("Line 1\nLine 2\nLine 3"); err != nil {
		return nil, err
	}

	// Mixed tabs and newlines
	if _, err := doc.AddParagraph("Name:\tJohn\nAge:\t30\nCity:\tNew York"); err != nil {
		return nil, err
	}

	return doc, nil
}

// 47 — Large document (stress test)
func genLargeDocument() (*docx.Document, error) {
	doc, err := docx.New()
	if err != nil {
		return nil, err
	}
	if _, err := doc.AddHeading("Large Document Stress Test", 0); err != nil {
		return nil, err
	}

	for chapter := 1; chapter <= 5; chapter++ {
		if _, err := doc.AddHeading(fmt.Sprintf("Chapter %d", chapter), 1); err != nil {
			return nil, err
		}
		for section := 1; section <= 3; section++ {
			if _, err := doc.AddHeading(fmt.Sprintf("Section %d.%d", chapter, section), 2); err != nil {
				return nil, err
			}
			for p := 1; p <= 5; p++ {
				if _, err := doc.AddParagraph(fmt.Sprintf(
					"Chapter %d, Section %d, Paragraph %d: Lorem ipsum dolor sit amet, consectetur adipiscing elit. "+
						"Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
					chapter, section, p,
				)); err != nil {
					return nil, err
				}
			}
		}

		// Add a table in each chapter
		tbl, err := doc.AddTable(3, 3)
		if err != nil {
			return nil, err
		}
		for r := 0; r < 3; r++ {
			for c := 0; c < 3; c++ {
				cell, _ := tbl.CellAt(r, c)
				cell.SetText(fmt.Sprintf("Ch%d R%dC%d", chapter, r, c))
			}
		}
		if _, err := doc.AddParagraph(""); err != nil {
			return nil, err
		}
	}

	return doc, nil
}
